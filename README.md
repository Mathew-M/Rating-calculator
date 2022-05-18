# Rating_calculator
"Rating_calculator" is a rating calculator for the record of board games such as Shogi or Chess. Calculating Method is base on [Shogi Club 24](https://www.shogidojo.net/)

# Functions
- Calculating rate
- Adding results/users
- Ranking
- Sorting
- Filtering
- Dropbox linkage (upload/download)
- Bom converting (It prevents garbled text which arise between utf-9 and shift-JIS)
- Excel sync

# Requirement
```
altgraph==0.17.2
certifi==2021.10.8        
charset-normalizer==2.0.11
dropbox==11.27.0
future==0.18.2
idna==3.3
importlib-metadata==4.10.0
llvmlite==0.38.0
openpyxl==3.0.9
pandas==1.3.5
pefile==2021.9.3
ping3==3.0.2
ply==3.11
pyinstaller==4.7
pyinstaller-hooks-contrib==2021.4
python-dateutil==2.8.2
pytz==2021.3
pywin32-ctypes==0.2.0
requests==2.27.1
six==1.16.0
stone==3.3.1
typing-extensions==4.0.1
urllib3==1.26.8
zipp==3.6.0
```

# Quick start
1. Add users
Push Add User button and fill in default rates and usernames in popup.
2. Add results
Push Add Result button and fill in winners and losers in the popup.
OR Edit "result_add.xlsx" and Push "Add by file" button. (tip: the scope is under "##", so overwriting is usable.)
3. Set dropbox_access_token on line18 (option)
```dropbox_access_token="hogehogehogehogehoge"```
4. Set up_down_password on line22 (option)
```updown_pass="hogehoge"```
5. Upload/ Download (option)
Push Upload/ Download button and fill in password.

# Others
It is assumed that users will convert this program to an exe file by using pyinstaller after doing the third and forth operation in the Quick start.

# License
This software is release under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).



