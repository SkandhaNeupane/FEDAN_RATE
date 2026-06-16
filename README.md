FEDAN USD Rate Scraper & WhatsApp Notifier

Automatically monitors the USD buying rate published by FEDAN twice daily, logs every rate to an Excel file, and sends an alert to a WhatsApp group — all without any manual intervention once set up.


How It Works

The script runs continuously in the background and wakes up during two daily windows:

WindowTime (NST)WhatsApp Message FormatMorning10:00 – 10:59 AMFEDAN(10AM) - June 16, 2026: 152.50Evening2:00 – 2:59 PMFEDAN(2PM) - June 16, 2026: 152.86


Checks the FEDAN website every 1 minute within each window
Skips Saturdays automatically
If FEDAN corrects a rate mid-session, a follow-up alert is sent marked (Updated)
Each session supports up to 3 rate changes per day
Rate and delivery timestamp are saved to Excel only after confirmed delivery
Failed sends are retried automatically on the next check



Requirements


Windows PC or laptop
Python 3.8+
Microsoft Edge (already installed on Windows 10/11)
Internet connection
WhatsApp account (for scanning QR on first run)



Installation

1. Clone the repository

bashgit clone https://github.com/your-username/fedan-scraper.git
cd fedan-scraper

2. Create and activate a virtual environment

bashpython -m venv venv
venv\Scripts\activate

3. Install dependencies

bashpip install requests beautifulsoup4 openpyxl selenium

4. Create the working directory

The script stores the Excel file, log, and browser profile here:

bashmkdir C:\fedan_rate

Copy fedan_scraper.py and run_fedan.bat into C:\fedan_rate\.


Configuration

Open fedan_scraper.py and update these values near the top if needed:

pythonBASE_DIR       = r"C:\fedan_rate"        # Folder for Excel, log, and browser profile
GROUP_NAME     = "FEDAN"                 # Exact name of your WhatsApp group
MORNING_HOUR   = 10                      # Hour for morning window (24h format)
EVENING_HOUR   = 14                      # Hour for evening window (24h format)
CHECK_INTERVAL = 60                      # Seconds between scrape checks


Running the Script

Option A — Run directly (for testing)

bashcd C:\fedan_rate
venv\Scripts\activate
python fedan_scraper.py

Option B — Run via batch file (recommended for 24/7 use)

Double-click run_fedan.bat. This keeps a terminal window open and automatically restarts the script if it ever crashes:

bat@echo off
:loop
python C:\fedan_rate\fedan_scraper.py
timeout /t 10
goto loop


First Run — WhatsApp QR Scan

The first time the script sends a WhatsApp alert, Microsoft Edge will open and load WhatsApp Web. The terminal will show:

>>> Please scan the QR code in the Edge window, then press Enter here.


Open WhatsApp on your phone
Go to Settings → Linked Devices → Link a Device
Scan the QR code shown in the Edge window
Press Enter in the terminal


The session is saved in C:\fedan_rate\edge_profile\ — you will not need to scan again unless WhatsApp logs out.


Excel Output

The script creates C:\fedan_rate\fedan_rates.xlsx automatically on first run.

ColumnContentsS.N.Serial numberDateDate of the entryMorning Rate 1–3Up to 3 morning rates per dayMessage Delivered (×3)Timestamp of confirmed WhatsApp deliveryEvening Rate 1–3Up to 3 evening rates per dayMessage Delivered (×3)Timestamp of confirmed WhatsApp delivery

A blank delivery cell means the alert was not successfully sent for that rate.


Running 24/7 on a Spare Laptop

For unattended operation:


Disable Windows sleep: Settings → System → Power & Sleep → set both to Never
Run on startup via Task Scheduler (so the script starts automatically after a reboot):

Open Task Scheduler → Create Task
General: Name it FEDAN Rate Scraper → check Run whether user is logged on or not and Run with highest privileges
Triggers: At startup, delay 2 minutes
Actions: Start run_fedan.bat
Conditions: Uncheck On AC power only → check Start only if network available
Settings: If already running, do not start a new instance



Do not close the Edge window that opens at send time — it will be reused for the evening session



Troubleshooting

SymptomLikely CauseFixUSD row not found in table at exactly 10:00 or 14:00FEDAN hasn't published the rate yetNormal — script retries every minuteEdge init failedEdge or its driver is locked by an old sessionRestart the script; it kills orphaned Edge processes automaticallyWhatsApp QR asked againSession expired or edge_profile\ was deletedRescan the QR codeExcel file not createdC:\fedan_rate\ folder doesn't existRun mkdir C:\fedan_rateScript stops after a crashRunning directly without run_fedan.batUse run_fedan.bat which auto-restarts on crash

Logs are written to C:\fedan_rate\fedan_scraper.log. If a WhatsApp send fails, a screenshot is saved as C:\fedan_rate\whatsapp_error.png for debugging.


Project Structure

C:\fedan_rate\
├── fedan_scraper.py      # Main script
├── run_fedan.bat         # Auto-restart launcher
├── fedan_rates.xlsx      # Excel output (auto-created)
├── fedan_scraper.log     # Log file (auto-created)
├── edge_profile\         # Saved WhatsApp Web session (auto-created)
└── whatsapp_error.png    # Screenshot on send failure (if any)


Notes


The WhatsApp group name in GROUP_NAME must match the group name exactly as it appears in WhatsApp
Do not manually edit the rate cells in the Excel file while the script is running — the Excel file is the script's memory and editing rates may trigger false "Updated" alerts
The edge_profile\ folder contains your WhatsApp session — do not share or commit it to version control (it is already in .gitignore)
