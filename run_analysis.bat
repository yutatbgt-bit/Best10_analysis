@echo off
setlocal

echo ==================================================
echo  ベスト10突合レポート生成スクリプト実行ツール
echo ==================================================
echo.

set /p STORE_NAME="参照する店舗名を入力してください（例: 全店, 王子店, 全シート）[デフォルト: 全店]: "
if "%STORE_NAME%"=="" set "STORE_NAME=全店"

rem 「全シート」入力時はCLI引数 all に変換
if "%STORE_NAME%"=="全シート" set "STORE_NAME=all"

echo.
echo 処理を開始します... (店舗: %STORE_NAME%)
echo.

uv run src/main.py --store %STORE_NAME%

echo.
echo ==================================================
if "%STORE_NAME%"=="all" (
    echo  全店舗の一括出力が完了しました。
    echo  結果は「output_data/」フォルダをご確認ください。
) else (
    echo  処理が完了しました。
    echo  結果は「output_data/best10_with_current_sales_%STORE_NAME%.xlsx」をご確認ください。
)
echo ==================================================
echo.
pause
endlocal
