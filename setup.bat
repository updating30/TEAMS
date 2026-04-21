@echo off
chcp 65001 > nul
echo ============================================
echo  仮想環境セットアップ
echo ============================================
echo.

echo [1/2] 仮想環境を作成中...
python -m venv venv
if errorlevel 1 (
    echo エラー: Python が見つかりません。Python 3.10 以上をインストールしてください。
    pause
    exit /b 1
)

echo [2/2] 依存パッケージをインストール中...
venv\Scripts\pip install --upgrade pip -q
venv\Scripts\pip install -r requirements.txt

echo.
echo ============================================
echo  セットアップ完了！
echo  次回からは run.bat を実行してください。
echo ============================================
pause
