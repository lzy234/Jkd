# Windows Excel插件实现方案（COM加载项）

## 1. 项目概述

本项目开发一个Windows平台的Excel COM加载项，实现与远程订单系统的交互功能，包括登录、获取订单数据和修改订单信息（派送人、备注）。

## 2. 技术选型

### 2.1 Excel插件开发技术

三种主要技术路线对比：

| 特性 | VBA | VSTO | COM加载项 |
|------|-----|------|----------|
| 开发环境 | Excel内置VBA编辑器 | Visual Studio | Visual Studio |
| 编程语言 | VBA | C#/.NET | C#/.NET |
| 分发方式 | .xlsm文件 | ClickOnce/MSI | 注册表+DLL文件 |
| 稳定性 | 中等 | 旁加载常出问题 | 高（传统可靠） |
| 兼容性 | 所有版本 | Office 2010+ | 所有版本 |
| 维护性 | 较低 | 高 | 高 |
| 部署复杂度 | 低 | 中 | 中 |

**推荐选择：COM加载项**
- 更稳定可靠，避免旁加载问题
- 兼容性好，适用于所有Excel版本
- 传统的注册方式，部署后不易出问题
- 完整的.NET支持，可使用现代C#特性
- 通过注册表注册，启动更快

### 2.2 开发环境需求

- Visual Studio 2019/2022
- .NET Framework 4.7.2或更高版本
- Excel 2010或更高版本（开发测试用）
- Windows SDK（用于COM注册）

### 2.3 COM加载项关键特性

- 实现IDTExtensibility2接口
- 通过COM互操作访问Excel对象模型
- 使用注册表注册加载项
- 程序集需要标记为ComVisible
- 需要强名称签名

## 3. 插件架构设计

### 3.1 插件模块划分

```
Excel COM加载项
├── Connect类（入口点，实现IDTExtensibility2）
├── UI层
│   ├── CommandBar（菜单和工具栏）
│   ├── 自定义窗体（登录和设置）
│   └── 工作表（数据展示）
├── 业务逻辑层
│   ├── 订单管理服务
│   ├── 用户认证服务
│   └── 数据同步服务
└── 数据访问层
    ├── API客户端
    ├── 数据转换器
    └── 缓存管理
```

### 3.2 关键技术组件

1. **COM互操作**：使用Excel COM对象模型
2. **CommandBar**：创建自定义菜单和工具栏
3. **Windows窗体**：用于登录界面和高级设置
4. **工作表交互**：使用Excel Interop API操作单元格和数据
5. **HTTP客户端**：使用HttpClient或RestSharp发起API请求
6. **JSON处理**：使用Newtonsoft.Json解析API响应
7. **注册表操作**：自动注册/注销加载项

## 4. COM加载项核心实现

### 4.1 Connect类（加载项入口）

```csharp
using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddin
{
    [ComVisible(true)]
    [Guid("YOUR-GUID-HERE")] // 使用guidgen.exe生成唯一GUID
    [ProgId("ExcelOrderAddin.Connect")]
    public class Connect : IDTExtensibility2
    {
        private Excel.Application _excelApp;
        private CommandBarButton _loginButton;
        private CommandBarButton _getOrdersButton;
        private CommandBarButton _refreshButton;
        
        private AuthenticationService _authService;
        private OrderService _orderService;
        private OrderUpdateService _orderUpdateService;
        
        // 加载项连接到Excel时调用
        public void OnConnection(object application, ext_ConnectMode connectMode, 
                                object addInInst, ref Array custom)
        {
            try
            {
                _excelApp = (Excel.Application)application;
                
                // 初始化服务
                _authService = new AuthenticationService();
                _orderService = new OrderService(_authService);
                _orderUpdateService = new OrderUpdateService(_authService);
                
                // 创建自定义菜单
                CreateCustomMenu();
                
                // 注册事件处理器
                _excelApp.SheetChange += OnSheetChange;
                
                Logger.LogInfo("订单管理插件已加载");
            }
            catch (Exception ex)
            {
                Logger.LogError("插件加载失败", ex);
                System.Windows.Forms.MessageBox.Show($"插件加载失败: {ex.Message}");
            }
        }
        
        // 加载项断开连接时调用
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                // 清理事件处理器
                if (_excelApp != null)
                {
                    _excelApp.SheetChange -= OnSheetChange;
                }
                
                // 清理菜单
                RemoveCustomMenu();
                
                // 释放COM对象
                if (_excelApp != null)
                {
                    Marshal.ReleaseComObject(_excelApp);
                    _excelApp = null;
                }
                
                Logger.LogInfo("订单管理插件已卸载");
            }
            catch (Exception ex)
            {
                Logger.LogError("插件卸载失败", ex);
            }
        }
        
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
        
        // 创建自定义菜单
        private void CreateCustomMenu()
        {
            try
            {
                // 获取或创建菜单栏
                CommandBar menuBar = _excelApp.CommandBars["Worksheet Menu Bar"];
                
                // 创建"订单管理"菜单
                CommandBarPopup orderMenu = null;
                try
                {
                    orderMenu = (CommandBarPopup)menuBar.Controls["订单管理"];
                }
                catch
                {
                    orderMenu = (CommandBarPopup)menuBar.Controls.Add(
                        MsoControlType.msoControlPopup, 
                        Type.Missing, 
                        Type.Missing, 
                        Type.Missing, 
                        true);
                    orderMenu.Caption = "订单管理(&O)";
                }
                
                // 添加登录按钮
                _loginButton = (CommandBarButton)orderMenu.Controls.Add(
                    MsoControlType.msoControlButton, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    true);
                _loginButton.Caption = "登录";
                _loginButton.Click += OnLoginClick;
                
                // 添加获取订单按钮
                _getOrdersButton = (CommandBarButton)orderMenu.Controls.Add(
                    MsoControlType.msoControlButton, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    true);
                _getOrdersButton.Caption = "获取订单";
                _getOrdersButton.Click += OnGetOrdersClick;
                
                // 添加刷新按钮
                _refreshButton = (CommandBarButton)orderMenu.Controls.Add(
                    MsoControlType.msoControlButton, 
                    Type.Missing, 
                    Type.Missing, 
                    Type.Missing, 
                    true);
                _refreshButton.Caption = "刷新数据";
                _refreshButton.Click += OnRefreshClick;
            }
            catch (Exception ex)
            {
                Logger.LogError("创建菜单失败", ex);
            }
        }
        
        // 移除自定义菜单
        private void RemoveCustomMenu()
        {
            try
            {
                CommandBar menuBar = _excelApp.CommandBars["Worksheet Menu Bar"];
                CommandBarControl orderMenu = menuBar.Controls["订单管理(&O)"];
                if (orderMenu != null)
                {
                    orderMenu.Delete(true);
                }
            }
            catch { }
        }
        
        // 菜单事件处理
        private void OnLoginClick(CommandBarButton ctrl, ref bool cancelDefault)
        {
            var loginForm = new LoginForm(_authService);
            loginForm.ShowDialog();
        }
        
        private async void OnGetOrdersClick(CommandBarButton ctrl, ref bool cancelDefault)
        {
            try
            {
                if (!_authService.IsAuthenticated)
                {
                    System.Windows.Forms.MessageBox.Show("请先登录");
                    return;
                }
                
                var orders = await _orderService.GetOrdersAsync("send", 1, 100);
                
                // 获取活动工作表
                Excel.Worksheet activeSheet = (Excel.Worksheet)_excelApp.ActiveSheet;
                _orderService.DisplayOrdersToWorksheet(activeSheet, orders);
                
                System.Windows.Forms.MessageBox.Show($"成功获取{orders.Count}条订单");
            }
            catch (Exception ex)
            {
                Logger.LogError("获取订单失败", ex);
                System.Windows.Forms.MessageBox.Show($"获取订单失败: {ex.Message}");
            }
        }
        
        private async void OnRefreshClick(CommandBarButton ctrl, ref bool cancelDefault)
        {
            // 刷新逻辑
            OnGetOrdersClick(ctrl, ref cancelDefault);
        }
        
        // 工作表变更事件处理
        private void OnSheetChange(object sh, Excel.Range target)
        {
            var sheet = sh as Excel.Worksheet;
            if (sheet == null || sheet.Name != "订单数据") return;
            
            try
            {
                // 判断修改的列
                if (target.Column == 7) // 备注列
                {
                    string orderUuid = ((Excel.Range)sheet.Cells[target.Row, 1]).Value?.ToString();
                    string newMemo = target.Value?.ToString() ?? "";
                    
                    if (!string.IsNullOrEmpty(orderUuid))
                    {
                        System.Threading.Tasks.Task.Run(async () =>
                        {
                            bool success = await _orderUpdateService.UpdateOrderMemoAsync(orderUuid, newMemo);
                            
                            _excelApp.Invoke((Action)(() =>
                            {
                                ((Excel.Range)sheet.Cells[target.Row, 8]).Value = success ? "更新成功" : "更新失败";
                            }));
                        });
                    }
                }
                else if (target.Column == 6) // 派送人列
                {
                    // 类似备注列的处理逻辑
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("处理单元格变更失败", ex);
            }
        }
    }
}
```

### 4.2 程序集配置（AssemblyInfo.cs）

```csharp
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("ExcelOrderAddin")]
[assembly: AssemblyDescription("Excel订单管理COM加载项")]
[assembly: AssemblyCompany("Your Company")]
[assembly: AssemblyProduct("ExcelOrderAddin")]
[assembly: AssemblyCopyright("Copyright © 2025")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

// 设置为COM可见
[assembly: ComVisible(true)]

// 类型库的GUID
[assembly: Guid("YOUR-TYPE-LIB-GUID-HERE")]
```

### 4.3 登录窗体

```csharp
using System;
using System.Windows.Forms;

namespace ExcelOrderAddin
{
    public partial class LoginForm : Form
    {
        private readonly AuthenticationService _authService;
        
        private TextBox txtUsername;
        private TextBox txtPassword;
        private Button btnLogin;
        private Button btnCancel;
        private Label lblStatus;
        
        public LoginForm(AuthenticationService authService)
        {
            _authService = authService;
            InitializeComponents();
        }
        
        private void InitializeComponents()
        {
            this.Text = "订单系统登录";
            this.Width = 400;
            this.Height = 250;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // 用户名
            var lblUsername = new Label
            {
                Text = "用户名:",
                Left = 30,
                Top = 30,
                Width = 80
            };
            this.Controls.Add(lblUsername);
            
            txtUsername = new TextBox
            {
                Left = 120,
                Top = 30,
                Width = 230
            };
            this.Controls.Add(txtUsername);
            
            // 密码
            var lblPassword = new Label
            {
                Text = "密码:",
                Left = 30,
                Top = 70,
                Width = 80
            };
            this.Controls.Add(lblPassword);
            
            txtPassword = new TextBox
            {
                Left = 120,
                Top = 70,
                Width = 230,
                PasswordChar = '*'
            };
            this.Controls.Add(txtPassword);
            
            // 状态标签
            lblStatus = new Label
            {
                Left = 30,
                Top = 110,
                Width = 320,
                ForeColor = System.Drawing.Color.Red
            };
            this.Controls.Add(lblStatus);
            
            // 登录按钮
            btnLogin = new Button
            {
                Text = "登录",
                Left = 120,
                Top = 150,
                Width = 100
            };
            btnLogin.Click += BtnLogin_Click;
            this.Controls.Add(btnLogin);
            
            // 取消按钮
            btnCancel = new Button
            {
                Text = "取消",
                Left = 250,
                Top = 150,
                Width = 100
            };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);
        }
        
        private async void BtnLogin_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text) || 
                string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                lblStatus.Text = "请输入用户名和密码";
                return;
            }
            
            btnLogin.Enabled = false;
            lblStatus.Text = "正在登录...";
            lblStatus.ForeColor = System.Drawing.Color.Blue;
            
            try
            {
                bool success = await _authService.LoginAsync(txtUsername.Text, txtPassword.Text);
                
                if (success)
                {
                    lblStatus.Text = "登录成功!";
                    lblStatus.ForeColor = System.Drawing.Color.Green;
                    
                    await System.Threading.Tasks.Task.Delay(1000);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    lblStatus.Text = "登录失败，请检查用户名和密码";
                    lblStatus.ForeColor = System.Drawing.Color.Red;
                    btnLogin.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"登录错误: {ex.Message}";
                lblStatus.ForeColor = System.Drawing.Color.Red;
                btnLogin.Enabled = true;
                Logger.LogError("登录异常", ex);
            }
        }
    }
}
```

### 4.4 服务实现（与之前相同）

AuthenticationService、OrderService、OrderUpdateService的实现与之前VSTO版本相同，这里不再重复。

## 5. 注册和部署

### 5.1 COM注册配置

在项目属性中配置：
1. 勾选"为COM互操作注册"
2. 签名选项卡中创建强名称密钥文件

### 5.2 注册表注册脚本

创建注册脚本 `Register.reg`:

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\ExcelOrderAddin.Connect]
"Description"="Excel订单管理加载项"
"FriendlyName"="订单管理"
"LoadBehavior"=dword:00000003
"CommandLineSafe"=dword:00000000
```

LoadBehavior说明：
- 0 = 断开连接，未加载
- 1 = 已加载
- 2 = 启动时加载
- 3 = 启动时加载并已加载（推荐）
- 8 = 按需加载
- 16 = 首次加载

### 5.3 安装脚本

创建 `Install.bat`:

```batch
@echo off
echo 正在安装Excel订单管理加载项...

REM 复制DLL到目标目录
set INSTALL_DIR=%APPDATA%\ExcelOrderAddin
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /Y "ExcelOrderAddin.dll" "%INSTALL_DIR%\"
copy /Y "Newtonsoft.Json.dll" "%INSTALL_DIR%\"

REM 注册COM组件
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase "%INSTALL_DIR%\ExcelOrderAddin.dll"

REM 导入注册表
regedit /s Register.reg

echo 安装完成！
pause
```

### 5.4 卸载脚本

创建 `Uninstall.bat`:

```batch
@echo off
echo 正在卸载Excel订单管理加载项...

set INSTALL_DIR=%APPDATA%\ExcelOrderAddin

REM 注销COM组件
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /u "%INSTALL_DIR%\ExcelOrderAddin.dll"

REM 删除注册表项
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\ExcelOrderAddin.Connect" /f

REM 删除文件（可选）
REM rmdir /S /Q "%INSTALL_DIR%"

echo 卸载完成！
pause
```

### 5.5 生成安装包

使用Inno Setup或WiX Toolset创建专业安装包：

**Inno Setup脚本示例** (`Setup.iss`):

```iss
[Setup]
AppName=Excel订单管理加载项
AppVersion=1.0
DefaultDirName={userappdata}\ExcelOrderAddin
DefaultGroupName=Excel订单管理
OutputDir=.\Setup
OutputBaseFilename=ExcelOrderAddin_Setup
Compression=lzma2
SolidCompression=yes

[Files]
Source: "bin\Release\ExcelOrderAddin.dll"; DestDir: "{app}"
Source: "bin\Release\Newtonsoft.Json.dll"; DestDir: "{app}"

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\ExcelOrderAddin.Connect"; ValueType: string; ValueName: "Description"; ValueData: "Excel订单管理加载项"
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\ExcelOrderAddin.Connect"; ValueType: string; ValueName: "FriendlyName"; ValueData: "订单管理"
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\ExcelOrderAddin.Connect"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3

[Run]
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/codebase ""{app}\ExcelOrderAddin.dll"""; Flags: runhidden

[UninstallRun]
Filename: "{dotnet40}\RegAsm.exe"; Parameters: "/u ""{app}\ExcelOrderAddin.dll"""; Flags: runhidden
```

## 6. 数据模型定义

（与之前VSTO版本相同，此处省略）

## 7. 异常处理与日志记录

（与之前VSTO版本相同，此处省略）

## 8. 调试和测试

### 8.1 调试配置

在项目属性 -> 调试：
- 启动外部程序：`C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`
- 命令行参数：（留空）
- 工作目录：（留空）

### 8.2 调试步骤

1. 确保项目已勾选"为COM互操作注册"
2. 以管理员权限运行Visual Studio
3. 按F5启动调试，会自动启动Excel
4. 在Excel中查看加载项是否已加载
5. 设置断点进行调试

### 8.3 常见问题排查

- **加载项未出现**：检查注册表是否正确注册
- **COM注册失败**：以管理员权限运行RegAsm
- **版本冲突**：确保.NET Framework版本匹配
- **事件不触发**：检查事件处理器是否正确注册

## 9. 与VSTO的主要区别

| 方面 | VSTO | COM加载项 |
|------|------|----------|
| 入口点 | ThisAddIn类 | Connect类（IDTExtensibility2） |
| UI创建 | Ribbon XML + Designer | CommandBar API |
| 注册方式 | ClickOnce/MSI | 注册表+RegAsm |
| 部署复杂度 | 中（旁加载问题多） | 低（传统可靠） |
| 生命周期管理 | 自动 | 手动（需实现接口方法） |
| 稳定性 | 较差（旁加载） | 好 |

## 10. 总结

COM加载项虽然相对传统，但在稳定性和兼容性方面优于VSTO，特别是在避免旁加载问题方面具有明显优势。通过本方案，可以创建一个稳定可靠的Excel订单管理插件。