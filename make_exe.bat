@echo off
chcp 65001
setlocal
cd /d "%~dp0"

:: 尝试结束已运行的进程
taskkill /F /IM AutoNetworker.exe /T >nul 2>nul

where python >nul 2>nul
if errorlevel 1 (
  echo Python 未安装，无法打包
  exit /b 1
)
python -m pip --version >nul 2>nul
if errorlevel 1 (
  python -m ensurepip --upgrade
)
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --upgrade pip
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --upgrade PyQt6 pyinstaller requests selenium websocket-client openai keyboard numpy

if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"

python -m PyInstaller -D -w --name AutoNetworker --collect-all PyQt6 --collect-all selenium --collect-all websocket --collect-all openai --collect-all keyboard main.py

if exist "dist\AutoNetworker\AutoNetworker.exe" (
  echo 构建成功: dist\AutoNetworker\AutoNetworker.exe
  echo 等待文件释放锁...
  timeout /t 3 /nobreak >nul
  
  echo 正在将编译结果推送到 GitHub 仓库...
  set "RELEASE_DIR=%TEMP%\AutoNetworker_Release"
  if exist "%RELEASE_DIR%" rd /s /q "%RELEASE_DIR%"
  mkdir "%RELEASE_DIR%"
  
  :: 将编译好的文件复制到系统临时目录
  xcopy /E /Y /I "dist\AutoNetworker\*" "%RELEASE_DIR%\" >nul
  
  :: 切换到临时目录进行 git 操作，避免在当前项目下产生嵌套的 .git
  pushd "%RELEASE_DIR%"
  git init
  git branch -M main
  git add .
  git commit -m "Auto Release Build"
  git remote add origin https://github.com/jeffrey0328/AutoNetworker.git
  git push -f origin main
  
  if errorlevel 1 (
      echo 推送失败，请检查网络或 Git 权限。
  ) else (
      echo 成功将发布文件强推至 https://github.com/jeffrey0328/AutoNetworker
  )
  popd
  
  :: 清理临时目录
  rd /s /q "%RELEASE_DIR%"
  
  echo.
  echo 是否要为此次提交创建 GitHub Release 标签? (Y/N)
  set /p upload_choice=
  if /I "%upload_choice%"=="Y" (
    echo 正在调用 Release 脚本...
    python upload_release.py
  ) else (
    echo 跳过创建 Release。
  )
  
  exit /b 0
)
echo 构建失败
exit /b 1
