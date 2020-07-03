@ECHO OFF
ECHO "Las librerias se estan instalando"
pip install pip
python -m pip install --upgrade pip
pip uninstall pandas
pip install pandas==1.0.1
pip install selenium
pip_install webdriver_manager
pip install openpyxl
echo.
echo.
echo.
echo.
echo.
echo.
echo.

PAUSE