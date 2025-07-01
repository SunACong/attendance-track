# 考勤追踪系统

## 项目简介
基于Python和Streamlit的自动化考勤管理系统，旨在简化企业考勤数据处理流程，支持多种考勤方式集成与自动化报表生成。系统提供直观的Web界面，方便管理员进行考勤数据管理、异常处理和统计分析。

## 主要功能
- 多源考勤数据整合（PC端打卡、OA系统打卡、请假登记、离岗登记）
- 自动化考勤数据清洗与校验
- 直观的考勤状态可视化仪表盘
- 自定义考勤规则配置
- 考勤异常自动识别与提醒
- 多维度考勤报表生成与导出


## 项目结构
```
.
├── README.md         - 项目说明文档
├── app.py            - Streamlit应用入口，提供Web界面与交互逻辑
├── all.py            - 公共工具函数库，包含数据处理、日期计算等通用方法
├── requirements.txt  - 项目依赖包列表
├── processLGDJ.py    - 离岗登记数据处理模块
├── processPCKQ.py    - PC端打卡数据处理模块
├── processQJDJ.py    - 请假登记数据处理与审批流程
├── processYDKQ.py    - OA系统打卡数据同步与处理
├── run.py            - 应用启动脚本
└── run.spec          - PyInstaller打包配置文件
```
## 快速开始
### 环境要求
- Python 3.8+ 
- 网络连接（首次运行需下载依赖）
- 管理员权限（部分数据处理操作）

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
streamlit run app.py
```
2. 系统会自动打开浏览器，访问以下地址（如未自动打开可手动访问）
```
http://localhost:8501/
```

### 使用示例
1. **数据导入**：在首页点击"导入考勤数据"按钮，选择对应的数据文件
2. **数据处理**：系统自动处理选中的考勤数据类型（PC打卡/OA打卡/请假等）
3. **查看报表**：在左侧导航栏选择"考勤报表"，设置时间范围和部门筛选条件
4. **导出数据**：点击报表页面的"导出Excel"按钮保存考勤结果


## 打包项目
如需将应用打包为独立可执行文件（适用于无Python环境的电脑）：

1. 清理历史打包文件
```bash
rmdir /s /q build dist __pycache__
```
2. 执行打包命令
```bash
pyinstaller run.spec
```
3. 打包完成后，可执行文件位于`dist`目录下

## 故障排除
### 常见问题
- **依赖安装失败**：尝试使用国内镜像源 `pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`
- **端口占用**：修改启动命令为 `streamlit run app.py --server.port 8502` 更换端口
- **数据处理错误**：检查输入数据格式是否符合系统要求（可参考示例数据格式）
- **打包失败**：确保已安装所有依赖，特别是pyinstaller `pip install pyinstaller`

## 许可证
本项目采用MIT许可证 - 详情参见LICENSE文件

## 贡献指南
1. Fork本仓库
2. 创建特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 打开Pull Request

## 联系方式
如有任何问题或建议，请联系：
- 邮箱：example@company.com
- 项目地址：https://github.com/username/attendance-tracking