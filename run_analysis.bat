@echo off
echo ==================================================
echo  ベスト10実績＆今月売上突合スクリプトを実行します
echo ==================================================
echo.

set /p STORE_NAME="参照する店舗名を入力してください（例: 全店, 王子店, 芦屋店）[デフォルト: 全店]: "
if "%STORE_NAME%"=="" set STORE_NAME=全店

uv run src/main.py --store %STORE_NAME%

echo.
echo ==================================================
echo  処理が完了しました。
echo  生成された「output_data/best10_with_current_sales_%%STORE_NAME%%.xlsx」をご確認ください。
echo ==================================================
echo.
pause
