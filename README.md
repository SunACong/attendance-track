# 考勤追踪系统

## 项目简介
基于Python的考勤管理系统，旨在简化企业考勤数据处理流程，支持多种考勤方式集成与自动化报表生成。系统提供直观的Web界面，方便管理员进行考勤数据管理、异常处理和统计分析。

## 主要功能
- 多源考勤数据整合（PC端打卡、OA系统打卡、请假登记、离岗登记、倒班表、出差登记表）
- 各个模块充分解耦逻辑分离
- 考勤异常自动识别与提醒
- 多维度考勤报表生成与导出
- 打包为exe支持未包含python环境的电脑（支持win7）


## 项目结构
```
.
├── README.md         - 项目说明文档
├── app.py            - 应用入口，提供Web界面与交互逻辑
├── all.py            - 公共工具函数库，包含数据处理、日期计算等通用方法
├── requirements.txt  - 项目依赖包列表
├── processLGDJ.py    - 离岗登记数据处理模块
├── processPCKQ.py    - PC端打卡数据处理模块
├── processQJDJ.py    - 请假登记数据处理与审批流程
├── processShift.py   - 倒班数据处理模块
├── processYDKQ.py    - OA系统打卡数据同步与处理
├── processCCKQ.py    - 出差登记模块处理
└── run.spec          - PyInstaller打包配置文件
```
## 快速开始
### 环境要求
- Python 3.8+ 
- 网络连接（首次运行需下载依赖）

### 安装步骤
1. 克隆或下载项目到本地
2. 打开命令行终端，进入项目目录
3. 安装依赖包
```bash
pip install -r requirements.txt
```

### 运行项目
1. 启动应用
```bash
python app.py
```
2. Tkinter界面会自动弹出


## 打包项目
如需将应用打包为独立可执行文件（适用于无Python环境的电脑）：

```bash
pyinstaller --noconsole --onefile app.py
```

打包完成后，可执行文件位于`dist`目录下

## 故障排除
### 常见问题
- **依赖安装失败**：尝试使用国内镜像源 `pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`
- **数据处理错误**：检查输入数据格式是否符合系统要求（可参考示例数据格式）
- **打包失败**：确保已安装所有依赖，特别是pyinstaller `pip install pyinstaller`


## 联系方式
如有任何问题或建议，请联系：
- 邮箱：2683885184@qq.com
- 项目地址：https://github.com/sunacong/attendance-track