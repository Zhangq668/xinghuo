# 薪火相传 —— 红色文化 AIGC 教育创作平台（MVP）

面向「AI+教育 · 红色教育创新与文化传承」方向，用 AIGC 技术生成红色故事脚本、红色动漫及文创作品的 Web 应用。

## 技术栈

- 界面 + 逻辑：Streamlit（Python，一个文件搞定）
- 文本生成：DeepSeek 大模型（脚本、讲解）
- 图像生成：智谱 CogView（红色动漫、文创作品）

## 快速开始

双击 `run.bat`，或手动执行：

```bash
streamlit run app.py
```

浏览器自动打开 `http://localhost:8501`。

## 两种模式

| 模式 | 是否需要 API Key | 说明 |
|------|-----------------|------|
| 演示模式（默认） | 不需要 | 脚本用内置示例，动漫/文创用上传图片 |
| 真实模式 | 需要 | 填 DeepSeek Key 生成脚本，填智谱 CogView Key 生成动漫/文创图 |

## 获取 API Key

### DeepSeek（文本）
1. 打开 [platform.deepseek.com](https://platform.deepseek.com/)，注册登录；
2. 在「API Keys」页创建 Key（形如 `sk-xxxx`）。

### 智谱 CogView（图像）
1. 打开 [open.bigmodel.cn](https://open.bigmodel.cn/)，注册登录；
2. 创建 API Key（形如 `xxx.yyy`）；
3. 接口地址 `https://open.bigmodel.cn/api/paas/v4/images/generations`，模型 `cogview-3`。

## 三个功能模块

1. **智能创作**：输入红色主题 → 生成红色故事脚本（DeepSeek）、红色动漫/文创图（智谱 CogView）。
2. **红色资源库**：内置红色史料，支持关键词检索（对应附件1「红色资源整合与智能检索」）。
3. **AI 红色讲解员**：对话式讲解红色历史（DeepSeek）。

## 后续可扩展

- 用「知稷」平台的 RAG 能力，把内置史料库升级为可扩展的红色知识库 + 史实校验。
