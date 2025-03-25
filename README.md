# PPTGenius 智能幻灯

PPTGenius 是一个基于大语言模型的智能幻灯片生成系统。项目名称中的 "Genius" 代表智慧与创造力，突出了项目的技术深度和自动化能力。通过简单的输入，即可快速生成专业的幻灯片内容。

## ✨ 主要特点

- 🤖 智能生成：基于大语言模型，自动生成专业的幻灯片内容
- 📝 结构化输出：自动生成大纲、内容，并转换为 PPT 格式
- 🎨 灵活定制：支持自定义模板和样式
- 📊 多媒体支持：支持图片和表格的插入
- 🔄 实时预览：支持实时查看生成的内容
- 🚀 快速部署：支持 Docker 容器化部署

## 🛠️ 技术栈

- 后端：Python Flask
- 前端：HTML, CSS, JavaScript
- 核心功能：Markdown 转 PPT
- 部署：Docker
- 依赖管理：pip

## 📦 安装

### 方式一：直接安装

1. 克隆项目
```bash
git clone https://github.com/yourusername/PPTGenius.git
cd PPTGenius
```

2. 创建虚拟环境（这里用的python3.9.21）
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或
.\venv\Scripts\activate  # Windows
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

4. 配置环境
- 复制 `config.py.example` 为 `config.py`
- 修改 `config.py` 中的配置信息

5. 运行应用
```bash
python app.py
```

### 方式二：Docker 部署

1. 构建镜像
```bash
docker build -t pptgenius .
```

2. 运行容器
```bash
docker run -d -p 5000:5000 -v $(pwd)/uploads:/app/uploads pptgenius
```

3. 访问应用
打开浏览器访问 `http://localhost:5000`

## 🎯 使用方法

1. 输入基本信息
   - 角色：选择你的角色（如：讲师、学生等）
   - 主题：输入演讲主题
   - 子主题数量：选择需要生成的子主题数量

2. 生成内容
   - 点击"生成主题"获取相关子主题
   - 选择或输入子主题后点击"生成大纲"
   - 编辑大纲后点击"生成内容"
   - 编辑内容后点击"生成PPT"

3. 自定义内容
   - 支持上传图片和 Excel 表格
   - 支持编辑生成的内容
   - 支持使用自定义模板

## 📁 项目结构

```
PPTGenius/
├── app.py              # 主应用文件
├── md2ppt.py           # Markdown 转 PPT 工具（测试环境使用）
├── config.py           # 配置文件
├── requirements.txt    # 项目依赖
├── templates/          # HTML 模板
├── static/            # 静态文件
├── uploads/           # 上传文件目录
└── Dockerfile         # Docker 配置文件
```

## 🔧 配置说明

在 `config.py` 中可以配置以下内容：

- LLM API 配置
- 文件上传配置
- 应用运行配置

## 📄 开源协议

本项目采用 MIT 协议开源。详见 [LICENSE](LICENSE) 文件。

## 👤 作者

- **lisiyu**
  - 邮箱：lhchlsy2000@163.com
  - 微信：![WeChat](wechat.jpg)

## 🙏 致谢

感谢以下开源项目的贡献：

- [Auto-PPT](https://github.com/limaoyi1/Auto-PPT)：提供了宝贵的参考和灵感

## 📝 更新日志

### v1.0.0
- 初始版本发布
- 支持基本的 PPT 生成功能
- 支持图片和表格上传
- 支持 Docker 部署

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request 来帮助改进项目。

## 📞 联系方式

如有问题或建议，请提交 Issue 或联系作者。 