# -*- coding: utf-8 -*-
import re
from pathlib import Path
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ALLOWED_COLUMNS = [
    "送货日期","订单号","件号","物料名称","物料规格","单位","数量","未税单价","未税总价","印刷",
    "Pos","Material","Quantity","Unit","Price","Amount",
    "Sales order ref","Sales order no","Sales order item",
    "轿厢净开门宽度_LL_mm","轿门净高_HH_mm","门尺寸_LL*HH",
    "DIM_CAR_BOX_INNER_LENGTH_mm","DIM_CAR_BOX_INNER_WIDTH_mm","DIM_CAR_BOX_INNER_HEIGHT_mm",
    "长宽高","包装箱类型","POS_CAR_GLASS_WALL_C"
]

ALIASES = {
    "长*宽*高":"长宽高","L*W*H":"长宽高","尺寸":"长宽高","规格":"物料规格",
    "物料号":"Material","料号":"件号","订单编号":"订单号","采购单号":"订单号",
    "交货日":"送货日期","交货日期":"送货日期",
}

def canonical_column_name(name: str) -> str:
    return ALIASES.get(str(name).strip(), str(name).strip())

def clean_text(value):
    return str(value or "").replace("\xa0", " ").strip()

def normalize_lines(text):
    lines=[]
    for raw in (text or "").replace("\xa0"," ").replace("\r","\n").split("\n"):
        line=re.sub(r"[ \t]+"," ",raw).strip()
        if line:lines.append(line)
    return lines

def flatten_text(text):
    return " ".join(normalize_lines(text))

def compact_text(text):
    return re.sub(r"\s+","",str(text or ""))

def clean_number(value):
    if value is None:return ""
    text=str(value).strip().replace(",","")
    if not text:return ""
    try:
        n=float(text)
        return int(n) if n.is_integer() else n
    except:
        return text

def as_plain_number_text(v):
    v=clean_number(v)
    if v=="":return ""
    return str(int(v)) if isinstance(v,float) and v.is_integer() else str(v)

def split_quantity_unit(v):
    s=str(v or "").strip().replace(" ","")
    if not s:return "","",""
    m=re.match(r"^([\d,]+(?:\.\d+)?)([A-Za-z\u4e00-\u9fa5]+)?$",s)
    if not m:return s,"",s
    return clean_number(m.group(1)),m.group(2) or "",s

def split_sales_order_ref(v):
    s=str(v or "").strip()
    if "/" in s:
        a,b=s.split("/",1)
        return a.strip(),b.strip(),s
    return s,"",s

def find_one(pattern,text,default="",flags=re.I):
    m=re.search(pattern,text or "",flags)
    return m.group(1).strip() if m else default

def find_money_after_label(label,text):
    return clean_number(find_one(rf"{re.escape(label)}\s*[:：]?\s*([0-9,]+(?:\.\d+)?)",text))

def format_lwh(l,w,h):
    a,b,c=as_plain_number_text(l),as_plain_number_text(w),as_plain_number_text(h)
    return f"{a}*{b}*{c}" if a and b and c else ""

def format_ll_hh(ll,hh):
    a,b=as_plain_number_text(ll),as_plain_number_text(hh)
    return f"{a}*{b}" if a and b else ""

# ===================== 提取逻辑（修正版，去掉Pos前缀干扰） =====================
def extract_package_type(block):
    # 只匹配“包装箱类型”后面的值，去掉前面多余的数字
    v = find_one(r"包装箱类型\s*[:：]?\s*([^\n]+)", block)
    if v:
        # 去掉开头的数字+空格（比如 "1 1普通" → "1普通"，"5 5, Heavy-duty ISPM15" → "5, Heavy-duty ISPM15"）
        v = re.sub(r"^\d+\s+", "", v)
        return v.strip()
    return ""

def extract_pos_car_glass_wall_c(block):
    v = find_one(r"POS_CAR_GLASS_WALL_C\s*[:：]?\s*([^\n]+)", block)
    if v:
        return v.strip()
    return ""

def extract_dim_value(block,kind):
    kind=kind.upper()
    for pat in [
        rf"DIM[_ \s]*CAR[_ \s]*BOX[_ \s]*INNER[_ \s]*{kind}\s*[:：]?\s*([0-9,]+(?:\.\d+)?)",
        rf"{kind}\s*[:：]?\s*([0-9,]+(?:\.\d+)?)\s*mm"
    ]:
        v=find_one(pat,block)
        if v:return clean_number(v)
    return clean_number(find_one(rf"DIMCARBOXINNER{kind}([0-9,]+(?:\.\d+)?)",compact_text(block)))

def extract_text_from_pdf(p):
    parts=[]
    with pdfplumber.open(p) as f:
        for i,page in enumerate(f.pages,1):
            t=page.extract_text(x_tolerance=1,y_tolerance=3) or ""
            parts.append(f"\n---PAGE {i}---\n{t}")
    return "\n".join(parts)

def parse_pdf(p):
    t=extract_text_from_pdf(p)
    rows = parse_kone_pdf(p,t)
    return rows

# ===================== KONE 解析（按行解析，解决字段乱串） =====================
def parse_kone_pdf(p,full):
    lines = normalize_lines(full)
    po=find_one(r"Purchase\s+order\s+No\.?\s*([0-9]+)", full)
    rows=[]

    # 1. 先收集所有物料行的索引
    item_lines = []
    for idx, line in enumerate(lines):
        m=re.match(r"^(\d+)\s+([A-Z0-9]{6,}(?:V\d+)?)\s+(\d{2}\.\d{2}\.\d{4})\s+([0-9,.]+\s*[A-Za-z]+)\s+([0-9,.]+)\s+([0-9,.]+)", line.strip())
        if m:
            item_lines.append((idx, m))

    # 2. 为每个物料行，只提取它前后10行内的字段
    for i, (start_idx, m) in enumerate(item_lines):
        # 确定当前物料块的范围：当前行到下一个物料行之前，或末尾
        end_idx = item_lines[i+1][0] if i+1 < len(item_lines) else len(lines)
        block_lines = lines[start_idx:end_idx]
        block_text = "\n".join(block_lines)

        pos,mat,date,qty_raw,price,amt = m.groups()
        qty,unit,_=split_quantity_unit(qty_raw)
        sales=find_one(r"Sales\s+order\s+ref\.?\s*([0-9/]+)", block_text)
        sales_no,sales_item,_=split_sales_order_ref(sales)

        # 只在当前物料块内提取字段
        ll=extract_dim_value(block_text,"LL")
        hh=extract_dim_value(block_text,"HH")
        L=extract_dim_value(block_text,"LENGTH")
        W=extract_dim_value(block_text,"WIDTH")
        H=extract_dim_value(block_text,"HEIGHT")
        pkg=extract_package_type(block_text)
        glass=extract_pos_car_glass_wall_c(block_text)

        row={
            "送货日期":date,"订单号":po,"件号":mat,"物料名称":mat,"物料规格":format_ll_hh(ll,hh) or format_lwh(L,W,H),
            "单位":unit,"数量":qty,"未税单价":clean_number(price),"未税总价":clean_number(amt),"印刷":"KONE",
            "Pos":pos,"Material":mat,"Quantity":qty,"Unit":unit,"Price":clean_number(price),"Amount":clean_number(amt),
            "Sales order ref":sales,"Sales order no":sales_no,"Sales order item":sales_item,
            "轿厢净开门宽度_LL_mm":ll,"轿门净高_HH_mm":hh,"门尺寸_LL*HH":format_ll_hh(ll,hh),
            "DIM_CAR_BOX_INNER_LENGTH_mm":L,"DIM_CAR_BOX_INNER_WIDTH_mm":W,"DIM_CAR_BOX_INNER_HEIGHT_mm":H,
            "长宽高":format_lwh(L,W,H),"包装箱类型":pkg,"POS_CAR_GLASS_WALL_C":glass
        }
        rows.append(row)
    return rows

# ===================== 给 app.py 调用的必须函数 =====================
def get_row_value(row, column):
    column = canonical_column_name(column)
    if column not in ALLOWED_COLUMNS:
        return ""
    return row.get(column, "")

# ===================== 导出 Excel =====================
def write_excel(rows,cols,out,errs=None):
    errs=errs or []
    cols=[canonical_column_name(c) for c in cols if canonical_column_name(c) in ALLOWED_COLUMNS]
    wb=Workbook()
    ws=wb.active
    ws.title="数据"
    hf=PatternFill("solid",fgColor="D9EAF7")
    hfnt=Font(bold=True)
    bd=Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))
    for i,c in enumerate(cols,1):
        cell=ws.cell(1,i,c)
        cell.fill=hf;cell.font=hfnt;cell.border=bd;cell.alignment=Alignment(horizontal="center")
    for r,row in enumerate(rows,2):
        for i,c in enumerate(cols,1):
            cell=ws.cell(r,i,row.get(c,""))
            cell.border=bd;cell.alignment=Alignment(wrap_text=True)
    for i,c in enumerate(cols,1):
        ws.column_dimensions[get_column_letter(i)].width=max(12,min(50,len(str(c))+6))
    log=wb.create_sheet("日志")
    log.cell(1,1,"文件");log.cell(1,2,"结果")
    log.cell(1,1).fill=hf;log.cell(1,2).fill=hf
    if errs:
        for i,e in enumerate(errs,2):
            log.cell(i,1,e.get("file",""));log.cell(i,2,e.get("error",""))
    else:
        log.cell(2,1,"全部");log.cell(2,2,"成功")
    Path(out).parent.mkdir(parents=True,exist_ok=True)
    wb.save(out)
