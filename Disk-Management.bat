@echo off
rem ============================================================
rem  ディスク容量管理運用表 作成バッチ
rem
rem  ・管理者権限が無い場合は UAC で自動昇格して自分自身を再実行する
rem  ・PowerShell を ExecutionPolicy Bypass で起動し
rem    同じフォルダの Disk-Management.ps1 を実行する
rem  ・結果は このフォルダ直下に Disk-Management-YYYYMMDDHHMM.HTML として出力される
rem  ・引数はそのまま PowerShell スクリプトへ渡す
rem      例) Disk-Management.bat -Exclude MAGNET.mirai.local
rem ============================================================
setlocal
cd /d "%~dp0"

rem --- 管理者権限チェック (fltmc は管理者以外では失敗する) ---
fltmc >nul 2>&1
if errorlevel 1 (
    echo 管理者権限が必要です。UAC で昇格して再実行します...
    set "ARGS=%*"
    if defined ARGS (
        powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList '%*' -WorkingDirectory '%~dp0' -Verb RunAs"
    ) else (
        powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -WorkingDirectory '%~dp0' -Verb RunAs"
    )
    exit /b
)

if not exist "%~dp0Disk-Management.ps1" (
    echo [ERROR] Disk-Management.ps1 が見つかりません: %~dp0
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Disk-Management.ps1" -OutputDir "%~dp0." %*
set "RC=%errorlevel%"

echo.
if "%RC%"=="0" (
    echo 完了しました。全機器の導通・容量取得が正常です。
) else if "%RC%"=="2" (
    echo 完了しました。導通NG・認証エラー・取得失敗・ドライブ未検出があります。HTML の内容を確認してください。
) else (
    echo [ERROR] スクリプトが異常終了しました。^(終了コード %RC%^)
)
pause
endlocal & exit /b %RC%
