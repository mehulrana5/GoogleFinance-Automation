# Google Finance Automation

## Setup Instructions

### 1. Clone the Repository
First, clone this repository using the following command:

```bash
git clone https://github.com/mehulrana5/GoogleFinance-Automation.git
```

### 2. Navigate to the Project Directory
```bash
cd ./GoogleFinance-Automation
```

### 3. Update data.xlsx
Open the data.xlsx file (located in the project folder) outside of VSCode (e.g., in Excel). Update it with your data and save the changes.

### 4. Set Up a Virtual Environment
```bash
python -m venv venv
```

### 5. Activate the Virtual Environment
  windows
  ```bash
  .\venv\Scripts\activate
  ```
  Linux/macOS
  ```bash
  source venv/bin/activate
  ```

### 6. Install Dependencies
```bash
pip install -r requirements.txt
```

### 7. Download and Configure ChromeDriver
Go to [ChromeDriver download page](https://googlechromelabs.github.io/chrome-for-testing/#stable) to download the `chromedriver` that is compatible with your version of Google Chrome. Extract the `chromedriver.exe` file from the zip file and move it to the `./GoogleFinance-Automation` directory.

### 8. Configure Environment Variables
Create a `.env` file in the project root directory, and add your Google Finance login credentials (If you face an issue while login then create a new acc):
```bash
GF_EMAIL="your_google_finance_email"
GF_PW="your_google_finance_password"
```

### 9. Run the Automation Script
```bash
python .\main.py
```