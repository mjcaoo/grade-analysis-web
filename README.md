# 成绩分析Web服务

基于原有的 `grade_analysis_final.py` 成绩分析程序实现的简易Web服务，提供便捷的文件上传、在线分析和结果下载功能。

## 功能特性

- 📁 **文件上传**: 支持上传多个Excel成绩文件和主要课程列表文件
- 🔍 **智能分析**: 自动处理成绩数据，计算学分加权平均分和综合排名
- 📊 **结果展示**: 实时显示分析结果，包括学生排名、统计信息等
- 💾 **结果下载**: 支持下载完整的Excel分析报告
- 🌐 **Web界面**: 现代化的响应式Web界面，操作简单直观

## 项目结构

```
zongce/
├── app.py                      # Flask Web应用主文件
├── grade_analyzer.py           # 成绩分析核心模块
├── grade_analysis_final.py     # 原始分析脚本
├── requirements.txt            # Python依赖包
├── templates/
│   └── index.html             # Web界面模板
├── uploads/                    # 上传文件存储目录
├── results/                    # 分析结果存储目录
└── data/                      # 示例数据文件
```

## 安装与运行

### 1. 环境准备

确保已安装Python 3.8+，然后安装依赖：

```bash
pip install -r requirements.txt
```

### 2. 启动服务

#### 方法：手动启动
```bash
python app.py
```

### 3. 访问服务

启动后可通过以下地址访问：
- 本地访问：http://127.0.0.1:5000
- 局域网访问：http://你的IP地址:5000

## 使用说明

### 1. 示例文件下载
- 首次使用建议先下载示例文件了解正确格式
- 在网站首页可下载：
  - **成绩文件示例**：展示正确的成绩文件格式
  - **主要课程列表示例**：展示主要课程文件格式

### 2. 文件上传
- **成绩文件**: 上传包含学生成绩的Excel文件（支持多个文件）
- **主要课程文件**: 必需上传主要课程列表文件，用于区分主要课程和其他课程

### 3. 文件格式要求
- 文件格式：`.xlsx` 或 `.xls`
- 成绩文件必须包含：
  - 学号列（列名包含"学号"）
  - 姓名列（列名包含"姓名"）
  - 课程成绩列（格式：课程名称【学分数】）

### 4. 分析过程
1. 上传所需文件
2. 点击"开始分析"按钮
3. 等待分析完成
4. 查看结果或下载完整报告

## API接口

### 主要端点

- `GET /` - 主页界面
- `POST /upload` - 文件上传
- `POST /analyze` - 执行分析
- `GET /results` - 获取分析结果
- `GET /download/<filename>` - 下载结果文件
- `GET /download_result` - 直接下载分析结果
- `GET /sample/<filename>` - 下载示例文件
- `GET /api/status` - 服务状态检查

### 示例API调用

```javascript
// 文件上传
const formData = new FormData();
formData.append('grade_files', file1);
formData.append('grade_files', file2);
formData.append('main_course_file', mainCourseFile);

fetch('/upload', {
    method: 'POST',
    body: formData
})
.then(response => response.json())
.then(data => console.log(data));

// 执行分析
fetch('/analyze', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'}
})
.then(response => response.json())
.then(data => console.log(data));

// 下载示例文件
window.location.href = '/sample/课程成绩-示例.xlsx';
window.location.href = '/sample/主要课程列表-示例.xlsx';
```

## 技术架构

- **后端**: Flask + Python
- **前端**: HTML5 + CSS3 + JavaScript
- **数据处理**: Pandas + NumPy
- **Excel操作**: OpenPyXL

## 核心功能说明

### 成绩转换规则
- 优秀 → 95分
- 良好 → 85分  
- 中等 → 75分
- 及格 → 65分
- 不及格 → 55分
- 数字成绩直接使用（0-100分范围）

### 综合分析逻辑
1. 主要课程：全部计入综合成绩
2. 其他课程：仅取成绩最高的4门课程
3. 计算方式：学分加权平均分 = Σ(课程成绩×学分) / Σ学分
4. 排序方式：按学分加权平均分降序排列

## 安全说明

- 上传的文件仅在当前会话中有效
- 服务器会为每个会话创建独立的文件存储空间
- 建议在生产环境中配置适当的安全策略

## 故障排除

### 常见问题

1. **服务启动失败**
   - 检查Python环境和依赖包安装
   - 确认5000端口未被占用

2. **文件上传失败**
   - 检查文件格式是否为Excel格式
   - 确认文件大小不超过限制

3. **分析失败**
   - 检查Excel文件格式是否正确
   - 确认包含必要的列（学号、姓名、课程成绩）

## 开发说明

本项目基于原有的命令行成绩分析工具重构而成，主要改进：

1. **模块化设计**: 将核心分析逻辑封装为 `GradeAnalyzer` 类
2. **Web化界面**: 提供直观的Web操作界面
3. **会话管理**: 支持多用户同时使用
4. **错误处理**: 完善的错误处理和用户提示
5. **结果管理**: 自动生成和管理分析结果文件

## 版本历史

- v1.0.0 - 初始版本，实现基本的Web化功能

## 贡献

欢迎提交Issue和Pull Request来改进这个项目。

## 许可

本项目基于MIT许可证开源。