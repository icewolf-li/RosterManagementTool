# 花名册管理工具 - 项目上下文文档

## 项目概述

**项目名称**：花名册管理工具  
**项目类型**：Python GUI 桌面应用程序  
**当前版本**：v1.2.1  
**主要用途**：管理计应(中职升本)2501班的学生文件和花名册，提供文件夹创建、文件重命名、汇总表生成等功能。

## 技术栈

- **编程语言**：Python 3.12
- **GUI框架**：Tkinter（Python标准库）
- **数据处理**：pandas
- **Excel处理**：openpyxl, xlsxwriter
- **打包工具**：PyInstaller
- **版本控制**：Git
- **CI/CD**：GitHub Actions

## 项目结构

```
计应花名册/
├── init.py                    # 主程序源代码（GUI应用，283行）
├── tkinter_test.py            # Tkinter测试脚本
├── requirements.txt           # Python依赖配置
├── init.spec                  # PyInstaller打包配置
├── build.cmd                  # Windows构建脚本
├── icon.ico                   # 应用程序图标
├── 花名册.xlsx               # 学生花名册数据文件
├── Readme.md                  # 项目说明文档
├── IFLOW.md                   # iFlow上下文文件（本文件）
├── .gitignore                 # Git忽略规则
├── .venv/                     # Python虚拟环境
├── 计应(中职升本)2501班/      # 学生数据目录（40个学生文件夹）
│   ├── 彭恒基/
│   ├── 覃艳玲/
│   └── ...（38个其他学生文件夹）
├── build/                     # 构建输出目录
├── dist/                      # 可执行文件输出目录
└── .github/
    └── workflows/
        └── build.yml          # GitHub Actions CI/CD配置
```

## 核心功能

### 1. 创建文件夹结构
- 自动创建主文件夹：`计应(中职升本)2501班/`
- 为40名学生创建个人子文件夹
- 支持实时进度显示

### 2. 文件重命名
支持4种命名方式：
- **学号**：`2531020130101（1）.jpg`
- **姓名**：`彭恒基（1）.jpg`
- **学号-姓名**：`2531020130101-彭恒基（1）.jpg`
- **姓名-学号**：`彭恒基-2531020130101（1）.jpg`

### 3. 创建汇总表
- 收集非空文件夹的学生信息
- 生成Excel格式的汇总表（`汇总表.xlsx`）
- 按学号排序，包含学号和姓名两列

### 4. 清理空文件夹
- 遍历所有学生文件夹
- 自动删除没有文件的空文件夹
- 支持实时进度显示

## 构建和运行

### 本地开发运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python init.py
```

### 本地构建可执行文件

```bash
# 使用构建脚本（推荐）
build.cmd

# 或直接使用PyInstaller
pyinstaller init.spec
```

构建产物：
- 可执行文件：`dist/花名册管理工具.exe`
- 单文件打包，无需额外依赖

### GitHub Actions 自动构建

**触发条件**：
- Push到main/master分支
- Pull Request到main/master分支
- 推送版本标签（v*.*.*）

**构建环境**：
- 操作系统：Windows-latest
- Python版本：3.12
- 依赖缓存：启用pip缓存加速构建

**构建流程**：
1. **build作业**（每次运行）：
   - 检出代码
   - 设置Python 3.12环境
   - 安装依赖（pandas, openpyxl, pyinstaller）
   - 执行PyInstaller构建
   - 上传构建产物（保留30天）

2. **release作业**（仅标签推送）：
   - 下载构建产物
   - 创建GitHub Release
   - 上传可执行文件到Release

## 开发约定

### 代码结构
- **主程序文件**：`init.py`（283行代码）
- **GUI类**：`RosterManagerApp`
- **入口点**：`if __name__ == "__main__"` 块位于文件末尾

### 数据格式
- **学生数量**：40名
- **学号范围**：2531020130101 到 2531020130140
- **主目录**：`计应(中职升本)2501班/`
- **输出汇总表**：`汇总表.xlsx`

### 文件命名规则
- 学号格式：10位数字（2531020130101）
- 文件序号：括号内的数字（1, 2, 3...）
- 支持文件扩展名：.jpg, .png, .jpeg等

### GUI界面规范
- 窗口大小：300x600像素
- 窗口标题："花名册管理工具 v1.2.1"
- 中文界面
- 实时日志输出区域

## 关键文件说明

### init.py
主程序源代码文件，包含：
- `RosterManagerApp` 类：GUI应用程序主类
- `check_excel_file()` 函数：创建文件夹结构
- `rename_file()` 函数：文件重命名（支持4种方式）
- `create_summary_file()` 函数：创建Excel汇总表
- `delete_empty_folders()` 函数：删除空文件夹
- 学生数据：`name_list`（40个姓名）和 `name_to_id`（姓名到学号的映射）

### requirements.txt
项目依赖配置：
```
pandas          # 数据处理
openpyxl        # Excel文件读写
pyinstaller     # 打包工具
xlsxwriter      # Excel写入支持
```

### init.spec
PyInstaller打包配置：
- 单文件模式：`--onefile`
- 无控制台窗口：`--windowed`
- 应用图标：`icon.ico`
- 隐式导入：pandas, tkinter
- UPX压缩：启用

### build.cmd
Windows批处理构建脚本：
```cmd
pyinstaller init.spec
```

### .github/workflows/build.yml
GitHub Actions CI/CD配置：
- 自动化构建流程
- 版本标签自动发布
- 构建产物保留30天
- 支持最新Windows环境

## 版本信息

- **当前版本**：v1.2.1
- **Git远程仓库**：https://github.com/icewolf-li/-----.git
- **最近提交**：fa420f6d "完善"

## 学生数据

项目包含40名学生的数据，学号从2531020130101到2531020130140。学生姓名包括：
彭恒基、覃艳玲、曾祥娟、曾莹、陈慧华、陈金泉、方兴林、韩善枨、黄亮唐、黄攀峰、劳高校、劳明权、黎嘉嘉、李国成、李可祺、林炎丽、刘艾玉蕾、刘东鑫、刘芳妤、刘佳雯、吕佳恒、马倩汝、马枝文、宋文婷、孙培相、孙齐旺、覃绪尧、唐雅怡、韦彬明、韦晓、吴振国、幸国梁、徐柳静、杨凤洁、杨树奎、姚江业、雍雨轩、袁玉兰、张博涵、周博岩

## 注意事项

1. **Python版本**：建议使用Python 3.12
2. **虚拟环境**：项目使用.venv虚拟环境，激活后再运行
3. **文件编码**：所有Python文件使用UTF-8编码
4. **GUI依赖**：Tkinter是Python标准库，无需额外安装
5. **Excel文件**：确保花名册.xlsx文件存在且格式正确
6. **构建清理**：重新构建前建议删除build和dist目录

## 常见问题

### Q: 如何修改学生名单？
A: 修改 `init.py` 文件中的 `name_list` 和 `name_to_id` 变量。

### Q: 如何添加新的重命名方式？
A: 在 `rename_file()` 函数中添加新的命名逻辑，并在GUI下拉菜单中添加选项。

### Q: 构建失败怎么办？
A: 检查Python版本（需要3.12），确保所有依赖已安装，尝试删除build和dist目录后重新构建。

### Q: 如何修改应用图标？
A: 替换 `icon.ico` 文件，确保文件名不变。

## 开发建议

1. **测试**：在修改代码后，先运行 `python init.py` 测试功能
2. **构建验证**：本地构建成功后再推送到GitHub触发CI/CD
3. **版本管理**：使用语义化版本号（vX.Y.Z）
4. **文档更新**：重大功能更新时同步更新Readme.md
5. **代码规范**：遵循Python PEP 8代码风格指南