# ProductPlanning

泛 FOF 型多资产理财产品的规划文档、渠道路演报告及文档生成脚本。

## 目录内容

- `*.md`：产品规划、报告框架及相关文字材料。
- `*.docx`：正式规划书、渠道路演报告及历史版本。
- `build_channel_roadshow_docx.py`：跨平台 Word 报告的基础样式和图表生成模块。
- `build_channel_roadshow_v2_docx.py`：V2.0 渠道路演报告生成脚本。
- `.roadshow_assets/`：报告使用的示例图表。
- `.source_doc_media/`：从原始材料提取并由报告引用的图表素材。

PDF 预览、页面渲染结果、Office 临时文件和本机缓存不纳入版本管理。

## 生成报告

安装依赖后，在项目目录运行：

```bash
python build_channel_roadshow_v2_docx.py
```

脚本将生成 `泛FOF型多资产理财产品_渠道路演报告_V2.0.docx`。
