# DOU Trend Update Workstation

A web-based tool for appending new Hawi build power breakdown data to the DOU Trend Exec Progress spreadsheet.

---

## 快速启动 / Quick Start

### 方法一：双击启动（推荐）

1. 确保已安装 **Python 3.10+**（[下载](https://www.python.org/downloads/)）
2. 双击运行 `start.bat`
3. 脚本会自动安装依赖并启动服务器
4. 浏览器打开显示的地址即可使用

### 方法二：命令行启动

```cmd
cd c:\Project\DoU
pip install -r Execution/web_station/requirements.txt
python Execution/web_station/app.py
```

---

## 团队共享部署 / Team Deployment

让团队成员都能通过浏览器访问，只需在**一台有网络权限的 Windows 机器**上运行服务器：

### 步骤

1. **选择主机**：选一台能访问 `\\grilled\ptt_china_la_golden\...` 的 Windows 机器（可以是你自己的电脑，也可以是团队服务器）

2. **安装依赖**：
   ```cmd
   pip install flask openpyxl waitress
   ```

3. **启动服务器**：
   ```cmd
   python Execution/web_station/app.py
   ```
   或双击 `start.bat`

4. **查看网络地址**：启动后终端会显示：
   ```
   Local   : http://localhost:5000
   Network : http://10.x.x.x:5000   ← 把这个地址发给团队成员
   ```

5. **团队成员**：在浏览器中打开 `http://<主机IP>:5000` 即可使用，无需安装任何软件

### 防火墙设置（如需要）

如果团队成员无法访问，需要在主机上开放 5000 端口：

```cmd
netsh advfirewall firewall add rule name="DOU Workstation" dir=in action=allow protocol=TCP localport=5000
```

---

## 开机自启 / Auto-start on Boot (Optional)

如果希望服务器开机自动启动，可以将 `start.bat` 添加到 Windows 任务计划程序：

1. 打开「任务计划程序」→「创建基本任务」
2. 触发器选「计算机启动时」
3. 操作选「启动程序」→ 选择 `start.bat`
4. 勾选「无论用户是否登录都要运行」

---

## 文件说明 / File Structure

```
Execution/
├── web_station/
│   ├── app.py              # Flask 后端（核心逻辑）
│   ├── templates/
│   │   └── index.html      # 前端页面
│   ├── requirements.txt    # Python 依赖
│   ├── start.bat           # Windows 一键启动脚本
│   ├── test_api.py         # API 测试脚本
│   └── README.md           # 本文档
├── append_build.py         # 命令行版本（同等功能）
├── verify_result.py        # 结果验证脚本
├── Hawi_GDoUv6_Exec_Progress_Joy.xlsx   # 目标 Excel 文件
└── SM8975_GDoUv6_Power_Breakdown_template_v2 (Output).xlsx  # 源数据（每次从服务器复制）
```

---

## 使用说明 / Usage

1. **PowerFtrace 路径**：粘贴构建的 PowerFtrace 文件夹路径
   - ✅ 正确：`\\grilled\...\Hawi.LA.1.0-00541-PERF.INT-1\PowerFtrace`
   - ❌ 错误（含文件名）：`\\grilled\...\PowerFtrace\SM8975_GDoUv6_Power_Breakdown_template_v2 (Output).xlsx`
   - 💡 如果不小心粘贴了含文件名的路径，系统会自动修正

2. **Build Label**：输入构建标签，格式 `M<编号>_<日期>`，例如 `M541_0401`

3. **Browse Builds**：点击可浏览服务器上的所有构建，点击自动填入路径和标签

4. **Run Update**：执行 6 步流程（复制→读取→检测→写入→保存→对比）

5. **结果页面**：
   - 6 个指标卡（DoU Current、Previous、Projection、Δ vs Projection 等）
   - 表格 1：当前构建 vs 上一构建（Battery Adjusted）
   - 表格 2：当前构建 vs Projection 目标
   - 下载按钮：下载更新后的 Excel 文件

---

## 依赖 / Dependencies

| 包 | 用途 |
|---|---|
| `flask` | Web 框架 |
| `openpyxl` | Excel 读写 |
| `waitress` | 生产级 WSGI 服务器（Windows 推荐） |
| `requests` | API 测试用 |