# 微博评论爬取与词频统计

这是一个基于 Python 的微博评论爬取与词频统计工具。  
主要功能包括：
- 自动爬取微博关键词下的评论、地址及点赞数，并保存至 Excel 或 txt 文件。
- 支持对评论内容进行词频统计，并过滤常见停用词。

---

## 功能特点
- 🔍 **微博评论爬取**：根据关键词，批量爬取微博评论及其相关信息，包括评论内容、位置、点赞数等。
- 📊 **词频统计**：对爬取到的评论内容进行分词，并统计词频，支持自定义停用词过滤。
- 📄 **多格式存储**：
  - 将评论内容、地址和点赞数保存至 Excel 文件 (`weibo_comment.xlsx`)。
  - 将评论内容保存至 TXT 文件 (`weibo_comment.txt`) 并进行词频统计。

---

## 环境依赖
在开始之前，请确保你已安装以下依赖库：
- `openpyxl`: 用于 Excel 文件操作
- `requests`: 用于 HTTP 请求
- `beautifulsoup4`: 用于网页解析
- `jieba`: 用于中文分词
- `pandas`: 用于数据处理

你可以通过以下命令安装所需依赖：

```bash
pip install openpyxl requests beautifulsoup4 jieba pandas
```

---

## 支援
如果可以希望能给个喝水费：
![收款码](https://github.com/user-attachments/assets/ea80b45c-dde0-466e-ad85-e1eadd972b6b){: style="width:600px; height:400px;" }
