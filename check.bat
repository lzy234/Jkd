@echo off
chcp 65001 >nul
echo ============================================
echo 环境检查
echo ============================================
echo.

echo [1] 检查 Node.js...
node --version
if %errorlevel% neq 0 (
    echo ❌ Node.js 未安装！请先安装 Node.js
    goto :end
) else (
    echo ✓ Node.js 已安装
)
echo.

echo [2] 检查 npm...
npm --version
if %errorlevel% neq 0 (
    echo ❌ npm 未安装！
    goto :end
) else (
    echo ✓ npm 已安装
)
echo.

echo [3] 检查 node_modules...
if exist "node_modules\" (
    echo ✓ 依赖已安装
) else (
    echo ❌ 依赖未安装，请运行 install.bat
)
echo.

echo [4] 检查关键文件...
if exist "src\taskpane.js" (
    echo ✓ src\taskpane.js 存在
) else (
    echo ❌ src\taskpane.js 不存在
)

if exist "manifest.xml" (
    echo ✓ manifest.xml 存在
) else (
    echo ❌ manifest.xml 不存在
)

if exist "webpack.config.js" (
    echo ✓ webpack.config.js 存在
) else (
    echo ❌ webpack.config.js 不存在
)
echo.

echo [5] 检查端口 3000...
netstat -ano | findstr :3000 >nul
if %errorlevel% equ 0 (
    echo ✓ 端口 3000 正在监听（开发服务器可能正在运行）
) else (
    echo ℹ 端口 3000 未被占用（开发服务器未运行）
)
echo.

:end
echo ============================================
echo 检查完成
echo ============================================
echo.
echo 如果所有项都显示 ✓，可以运行 dev.bat 启动开发服务器
echo.
pause

