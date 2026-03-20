@echo off
echo 正在启动扫描仪服务 (方案一)...
echo 服务将在 http://localhost:9527 上运行

:: 检查端口是否被占用
netstat -ano | findstr :9527 > nul
if %errorlevel% equ 0 (
    echo 端口 9527 已被占用，正在查找占用进程...
    for /f "tokens=5" %%a in ('netstat -ano ^| findstr :9527') do set PID=%%a
    echo 占用端口的进程ID: %PID%
    echo 正在终止该进程...
    taskkill /F /PID %PID%
    if %errorlevel% equ 0 (
        echo 进程已成功终止
    ) else (
        echo 终止进程失败，请手动关闭占用端口的程序
        pause
        exit /b 1
    )
)

:: 启动服务
echo 正在启动服务...
start LocalScanServiceV2.exe

:: 等待服务启动
echo 等待服务启动中...
ping localhost -n 3 > nul

:: 打开浏览器
echo 正在打开浏览器...
start http://localhost:9527

echo 服务已启动，请在浏览器中查看
echo 按任意键退出...
pause > nul