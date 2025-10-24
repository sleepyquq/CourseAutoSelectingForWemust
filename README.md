# 澳门科技大学选课自动化脚本

一个基于Selenium的自动化选课脚本，专门用于澳门科技大学(EAS)选课系统的课程抢选。

## 功能特点

- 🚀 **自动化选课**：自动完成选课流程，无需手动操作
- 🔄 **循环重试**：自动刷新页面并重试，提高抢课成功率
- 🎯 **精确定位**：使用XPath精确定位页面元素
- 🌐 **Edge浏览器支持**：基于Microsoft Edge浏览器，兼容性好
- 📊 **详细日志**：提供详细的操作日志，便于调试和监控

## 项目结构

```
selectAuto/
├── auto_select_edge.py    # 主脚本文件
├── msedgedriver.exe       # Edge浏览器驱动
└── README.md              # 项目说明文档
```

## 环境要求

- Python 3.6+
- Microsoft Edge浏览器
- Windows操作系统
- 网络连接

## 依赖安装

```bash
pip install selenium
```

## 使用方法

### 1. 准备工作

1. 确保已安装Microsoft Edge浏览器
2. 下载对应版本的EdgeDriver（已包含在项目中）
3. 安装Python依赖包

### 2. 配置说明

脚本中的关键配置项：

- **选课URL**：`https://eas-course-enrollment-wmweb.must.edu.mo/add-select`
- **目标课程**：人文藝術类课程（第3个加选按钮）
- **用户数据目录**：`C:\\Users\\ovo\\AppData\\Local\\Microsoft\\Edge\\User Data`

### 3. 运行脚本

```bash
python auto_select_edge.py
```

### 4. 操作流程

脚本会自动执行以下操作：

1. 启动Edge浏览器
2. 跳转到选课页面
3. 点击"人文藝術"选课分支
4. 查找并点击第3个"加選"按钮
5. 刷新页面，循环重试

## 自定义配置

### 修改目标课程

要抢选其他课程，需要修改以下部分：

```python
# 1. 修改选课分支（第54行）
branch_btn = driver.find_element(By.XPATH, "//li[contains(@class, 'ivu-menu-item')]//span[contains(text(), '人文藝術')]")

# 2. 修改加选按钮索引（第68行）
select_btn = all_add_btns[2]  # 修改索引号选择不同课程
```

### 修改用户数据目录

```python
# 第19行，修改为你的Edge用户数据目录
options.add_argument("--user-data-dir=C:\\Users\\你的用户名\\AppData\\Local\\Microsoft\\Edge\\User Data")
```

### 调整等待时间

```python
# 调整各种等待时间以适应网络环境
time.sleep(30)  # 初始加载等待时间
time.sleep(5)   # 分支页面加载等待时间
time.sleep(10)  # 页面刷新等待时间
```

## 注意事项

⚠️ **重要提醒**：

1. **合法使用**：请确保在选课系统开放时间内合法使用
2. **账号安全**：使用前请确保已登录选课系统
3. **网络稳定**：建议在网络稳定的环境下运行
4. **浏览器版本**：确保EdgeDriver版本与Edge浏览器版本匹配
5. **系统权限**：如遇到权限问题，请以管理员身份运行

## 故障排除

### 常见问题

1. **浏览器启动失败**
   - 关闭所有Edge浏览器窗口
   - 结束所有msedge.exe进程
   - 以管理员身份重新运行脚本

2. **找不到页面元素**
   - 检查网络连接
   - 确认选课系统是否正常开放
   - 检查页面结构是否发生变化

3. **EdgeDriver版本不匹配**
   - 下载与Edge浏览器版本匹配的EdgeDriver
   - 替换项目中的msedgedriver.exe文件

### 调试模式

脚本已启用调试模式，可以通过以下方式调试：

```python
# 启用远程调试端口
options.add_argument("--remote-debugging-port=9222")
```

在浏览器中访问 `http://localhost:9222` 进行调试。

## 技术实现

- **自动化框架**：Selenium WebDriver
- **浏览器**：Microsoft Edge
- **定位方式**：XPath选择器
- **等待策略**：显式等待 + 固定时间等待
- **异常处理**：try-catch异常捕获和重试机制

## 免责声明

本项目仅供学习和研究使用，使用者需要：

- 遵守学校选课系统的使用规则
- 承担使用本脚本的所有风险和责任
- 不得用于任何违法或不当用途

## 许可证

本项目采用MIT许可证，详见LICENSE文件。

## 贡献

欢迎提交Issue和Pull Request来改进这个项目！


---

**祝您选课顺利！** 🎓

---
ps:也许之后我有空了会继续完善本项目为完整的可视化图形界面操作软件


