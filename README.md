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

# Requirement
- python>=3.8.1
- dropbox>=11.27.0
- pandas>=1.3.5
- ping3>=3.0.2

# Quick start
1. Add users
Push Add User button and fill in default rates and usernames in popup.
2. Add results
Push Add Result button and fill in winners and losers in the popup.
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



