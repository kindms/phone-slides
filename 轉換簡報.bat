@echo off
chcp 65001 >nul
rem 把 PPT 檔拖到這個檔案上：轉成手機簡報並上傳
cd /d "%~dp0"
if "%~1"=="" (
  echo 請把 PPT 檔案拖曳到「轉換簡報.bat」上面
  pause
  exit /b
)
py -3.14 tools\convert.py "%~1" || (pause & exit /b)
echo 上傳到網路中…
git -C "%~dp0." add d
git -C "%~dp0." commit -q -m "新增簡報：%~n1"
git -C "%~dp0." push -q
echo.
echo 完成！約 1～2 分鐘後連結就能打開。連結請看 簡報連結.txt
pause
