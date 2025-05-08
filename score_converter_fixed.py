name: Build EXE with PyInstaller

on:
  push:
    paths:
      - "**.py"          # 只要 .py 有變動就重打包
      - requirements.txt
  workflow_dispatch:      # 也可手動按按鈕觸發

jobs:
  build-win:
    runs-on: windows-latest          # 在雲端 Win Server 打包
    steps:
    - name: 取得程式碼
      uses: actions/checkout@v4

    - name: 設定 Python 3.12
      uses: actions/setup-python@v5
      with:
        python-version: "3.12"

    - name: 安裝依賴
      run: |
        python -m pip install --upgrade pip
        pip install -r requirements.txt

    - name: 使用 PyInstaller 打包
      run: |
        pyinstaller --onefile score_converter_fixed.py

    - name: 上傳 EXE 成品
      uses: actions/upload-artifact@v4
      with:
        name: score_converter_win
        path: dist/score_converter_fixed.exe
