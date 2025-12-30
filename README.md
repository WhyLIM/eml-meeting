# EML 邮件学术报告信息提取工具

使用 LLM 从 EML 邮件文件中自动提取学术报告/培训信息，可以指定文件夹来批量导出到 Excel。

## 提取字段

| 字段          | 说明                   |
| ------------- | ---------------------- |
| 文件名        | EML 文件名             |
| 培训/会议名称 | 讲座主题               |
| 开始时间      | 格式: yyyy-MM-dd hh:mm |
| 结束时间      | 格式: yyyy-MM-dd hh:mm |
| 学时(小时)    | 持续时长（数字）       |
| 地点          | 讲座地点               |
| 讲座目的      | 讲座目的（50字以内）   |
| 讲座内容      | 讲座内容（50字以内）   |
| 提取状态      | 成功/失败              |

## 安装

### 依赖安装

```bash
pip install -r requirements.txt
```

或手动安装：

```bash
pip install openpyxl python-dotenv requests
```

### 环境变量配置

创建 `.env` 文件并配置 API 密钥：

```env
ZAI_API_KEY=your_zai_api_key
OPENAI_API_KEY=your_openai_api_key
DEEPSEEK_API_KEY=your_deepseek_api_key
GEMINI_API_KEY=your_gemini_api_key
```

## 使用方法

### 基本用法

```bash
python main.py
```

### 命令行参数

| 参数             | 说明                                            | 默认值             |
| ---------------- | ----------------------------------------------- | ------------------ |
| `-c, --config` | 配置文件路径                                    | config.yaml        |
| `-i, --input`  | 输入EML文件或目录                               | （从配置文件读取） |
| `-o, --output` | 输出文件路径                                    | （从配置文件读取） |
| `--api`        | API提供商 (zai-plan/zai/openai/deepseek/gemini) | （从配置文件读取） |
| `--model`      | 指定模型名称                                    | （从配置文件读取） |

### 配置文件格式 (config.yaml)

```yaml
input_dir: messages_package
output_file: output/result.xlsx
api_provider: zai-plan
model: glm-4.5
```

## 辅助脚本

### 合并 Excel 文件

将多个 Excel 文件合并为一个：

```bash
python merge_excel.py
```

指定输入输出：

```bash
python merge_excel.py -i output -o output/merged.xlsx
```

### 按重复拆分 Excel

将 Excel 按培训/会议名称重复情况拆分为三个文件：

```bash
python split_by_duplicate.py
```

指定路径：

```bash
python split_by_duplicate.py -i output/merged_clean.xlsx -u output/unique.xlsx -f output/duplicates_first.xlsx -s output/duplicates_second.xlsx
```

输出文件：

- `unique.xlsx` - 不重复记录
- `duplicates_first.xlsx` - 重复记录的第一行
- `duplicates_second.xlsx` - 重复记录的第二行及更多

## API 配置

由于我买了智谱的 Coding Plan，只对 zai-plan 进行了测试，使用 glm-4.5 模型，它的并发限制是 10 个请求/s，但实际使用时发现设置为 5 才能稳定不报错。

其它模型提供商的接口没有进行测试，使用其它模型可能需要自行修改一些代码。

## 项目结构

```
eml-parser/
├── main.py              # 主程序入口
├── eml_parser.py         # EML 文件解析
├── extractor.py          # LLM 信息提取
├── llm_client.py         # LLM API 客户端
├── config.py            # 配置管理
├── config_loader.py      # YAML 配置加载
├── merge_excel.py        # Excel 合并
├── split_by_duplicate.py # Excel 拆分
├── config.yaml          # 配置文件
├── .env                # 环境变量（自行创建）
└── requirements.txt      # 依赖列表
```

## 许可证

MIT License
