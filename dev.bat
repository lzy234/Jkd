@echo off
chcp 65001 >nul
echo ============================================
echo Excel订单管理插件 - 开发服务器
echo ============================================
echo.
echo 正在启动开发服务器...
echo 服务器地址: https://localhost:3000
echo.
echo 提示: 首次访问需要信任自签名证书
echo.
call npm start

