first clone this repo using git clone https://github.com/mehulrana5/GoogleFinance-Automation.git
then go to cd ./GoogleFinance-Automation
in this you have a data.xlsx file which you have to open outside vscode
update the excel file and save it 
then do python -m venv venv
then do .\venv\Scripts\activate
then do pip install -r requirements.txt
then go to https://googlechromelabs.github.io/chrome-for-testing/#stable to download chromedriver that is compatible to your google chrome on destop extract the chromedriver.exe file from the zip file and move it to ./GoogleFinance-Automation dir
then create .env file and add
    GF_EMAIL=""
    GF_PW=""
then do python.exe .\main.py
