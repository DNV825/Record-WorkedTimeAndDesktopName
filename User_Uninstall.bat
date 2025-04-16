@echo off
:: カレントディレクトリをバッチファイルと同じ場所にするイディオム。
:: 管理者権限ありでバッチファイルを起動すると、カレントディレクトリが別の場所に移動してしまうので、この処理が必要となる。
cd /d %~dp0
powershell -NoLogo -ExecutionPolicy RemoteSigned -File .\Setup.ps1 -ActionType Uninstall
pause