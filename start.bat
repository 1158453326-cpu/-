@echo off
start /B "ContractLens" C:\Users\11584\.workbuddy\binaries\python\envs\contractlens\Scripts\streamlit.exe run "E:\WorkBuddy Program\2026-07-11-22-03-21\contract_lens.py" --server.address 0.0.0.0 --server.headless true
timeout /t 3 /nobreak >nul
start /B "Ngrok" ngrok http 8502
echo 网站已后台启动，窗口可以关了
echo ngrok 链接在 http://127.0.0.1:4040 查看
timeout /t 5 /nobreak >nul
