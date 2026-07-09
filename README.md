# 🗓️ Apple Calender - Exchange → iCloud Calendar Sync (macOS AppleScript)

An AppleScript that automatically synchronizes your Exchange calendar - configured in Apple Calendar - with an iCould calendar.

It only copies new or changed appointments (delta sync), covers the last n days and the next 30 days, and optionally maintains a log in the Documents folder.

I created this to meet company requirements, to NOT log in to exchange from non-company devices. 

---

## ✨ Features

- 🧠 UID-Sync + Update + Delete
- 🕓 **Time window**: 10 days backward, 30 days forward
- 📁 **Local storage** of the sync status in `~/Documents/ExchangeSync/lastSync.txt`
- 🪵 **Log file** with timestamp and results in `~/Documents/ExchangeSync/sync.log`
- 🚀 **Automatic background run** via `launchd` (optional)


---

## 📦 Installation

1. **Copy script**
- Open the **Script Editor** on macOS
- Paste the contents of `exchange_sync.scpt`
- Adjust the two calendar names in the upper area to:
```applescript
set sourceCalendarName to "Exchange"
set targetCalendarName to "Private"
```
- Save the script to:
```
~/Library/Scripts/exchange_sync.scpt
```

2. **Folder for Sync Data**
- The script automatically creates the folder:
```
~/Documents/ExchangeSync/
```
- This contains:
- `lastSync.txt` → saves the date of the last successful run
- `sync.log` → contains all synchronization entries
  
### ⚙️ Automatic run (recommended)

To run the script automatically on a regular basis (e.g., every 2 hours), move the .plist script to e.g.
```
~/Library/LaunchAgents/com.exchange.sync.plist
```

Execute in Terminal:
```
launchctl load ~/Library/LaunchAgents/com.exchange.sync.plist
```

## 🧠 How it works
On the first run, the script creates the file lastSync.txt and synchronizes all appointments from the last 90 days.
On each subsequent run:
Reads the date from lastSync.txt
Checks which appointments have been added or changed in the Exchange calendar since then
Copies these to the target calendar
Saves the current date again in lastSync.txt
Writes a line to sync.log with the timestamp and the number of copied events

