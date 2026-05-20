# PDF采购订单提取到Excel - 字段白名单版

这是一个可部署到 Render 的完整网页项目。用户无需注册，打开网址即可上传 PDF、勾选字段、预览结果并下载 Excel。

## 当前可选字段

前端勾选区和后端导出都只允许以下字段：

- 送货日期
- 订单号
- 件号
- 物料名称
- 物料规格
- 单位
- 数量
- 未税单价
- 未税总价
- 印刷
- Pos
- Material
- Quantity
- Unit
- Price
- Amount
- Sales order ref
- Sales order no
- Sales order item
- 轿厢净开门宽度_LL_mm
- 轿门净高_HH_mm
- 门尺寸_LL*HH
- DIM_CAR_BOX_INNER_LENGTH_mm
- DIM_CAR_BOX_INNER_WIDTH_mm
- DIM_CAR_BOX_INNER_HEIGHT_mm
- 长宽高

尺寸字段顺序说明：门尺寸按 `LL宽、HH高、LL*HH`；木箱尺寸按 `L长、W宽、H高、长宽高`。

## 本地运行

```bash
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000
```

浏览器打开：

```text
http://127.0.0.1:8000
```

## Render 免费部署

1. 新建 GitHub 仓库。
2. 上传本项目所有文件。
3. 登录 Render。
4. New + -> Web Service。
5. 连接 GitHub 仓库。
6. 配置：

```text
Runtime: Python
Build Command: pip install -r requirements.txt
Start Command: uvicorn app:app --host 0.0.0.0 --port $PORT
```

也可以使用仓库里的 `render.yaml` 自动识别配置。

## 用户使用流程

1. 打开部署后的网址。
2. 上传一个或多个 PDF。
3. 勾选导出列。
4. 点击“开始提取并预览”。
5. 确认预览无误后点击“下载Excel”。

## 注意

- 适合文字型 PDF。
- 扫描件或图片版 PDF 需要先 OCR。
- 当前后端会过滤所有不在字段白名单里的列，即使用户篡改前端请求也不会导出额外字段。
