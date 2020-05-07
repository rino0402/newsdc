:top
python -m pip install --force-reinstall --upgrade pip
python -m pip install --upgrade pip
if not errorlevel == 0 python -m pip install --force-reinstall --upgrade pip
pip install pip-review
pip install slackweb
pip install pyodbc
pip install beautifulsoup4
pip install lxml
rem Microsoft Build Tools 2015
rem https://www.microsoft.com/ja-JP/download/confirmation.aspx?id=48159
rem https://download.microsoft.com/download/5/F/7/5F7ACAEB-8363-451F-9425-68A90F98B238/visualcppbuildtools_full.exe
pip install openpyxl
pip install xlrd
pip install pandas
if errorlevel == 0 exit/b
goto :top
pip install fbprophet

exit/
Package    Version	Package    Version
---------- -------  ---------- -------
et-xmlfile 1.0.1    et-xmlfile 1.0.1  
jdcal      1.3      jdcal      1.4    
openpyxl   2.5.0    openpyxl   2.5.9  
pip        10.0.1   pip        18.1   
pyodbc     4.0.22   pyodbc     4.0.24 
setuptools 28.8.0   setuptools 39.0.1 
slackweb   1.0.5    slackweb   1.0.5  
xlrd       1.1.0    xlrd       1.1.0  











