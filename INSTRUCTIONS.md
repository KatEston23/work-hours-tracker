# 🕒 Work Hours Tracker - Executable

## 📋 Description
This executable allows you to easily track your work hours and generate a professional Excel calendar with colors and formatting.

## 🚀 How to use the executable

### Option 1: Run from file explorer
1. Go to the `dist/` folder
2. Double-click on `Work_Hours_Tracker.exe`
3. A console window will open

### Option 2: Run from command line
1. Open PowerShell or CMD
2. Navigate to the folder where the executable is located
3. Run: `.\dist\Work_Hours_Tracker.exe`

## 📝 Usage Instructions

1. **Time Entry**: The program will ask you to enter your clock-in and clock-out times
   - Format: `HH.MM` (example: `08.30` for 8:30 AM)
   - Automatically alternates between clock-in and clock-out
   - Type `done` when finished

2. **Usage Example**:
   ```
   📅 Today's date: 2025-01-18
   Enter your time entries in HH.MM format (example: 08.42).
   Type 'done' when finished.

   Time entry 1: 08.00    (Clock In)
   Time entry 2: 12.00    (Clock Out - Lunch)
   Time entry 3: 13.00    (Clock In - Back from lunch)
   Time entry 4: 17.30    (Clock Out)
   Time entry 5: done
   ```

## 📊 Generated Excel File

The program creates/updates the `work_hours_history.xlsx` file with:

### 🎨 Calendar Format:
- **Monthly tabs**: One tab per month with data
- **Color coding**:
  - 🔵 Light blue: Weekdays
  - 🔴 Light red: Weekends
  - 🟢 Green: Monthly totals (hours worked)
  - 🟡 Gold: Extra hours
  - 🔷 Dark blue: Headers

### 📈 Information displayed:
- **W:** Total hours worked each day
- **E:** Extra hours (if any)
- **Monthly totals** at the bottom of each calendar

## 📁 Important Files

- `Work_Hours_Tracker.exe`: The main executable
- `work_hours_history.xlsx`: Excel file with your data (created automatically)

## 💡 Tips

1. **Portability**: You can copy just the `Work_Hours_Tracker.exe` file to any Windows computer
2. **Backup**: The Excel file is saved in the same folder where you run the program
3. **Previous data**: The program preserves and maintains all previous data
4. **No installation**: No need to install Python or any other dependencies

## ⚠️ Important

- Run the program from the folder where you want the Excel file to be saved
- Make sure you have write permissions in the folder
- The Excel file can be opened with Microsoft Excel, LibreOffice Calc, or Google Sheets

## 🔧 Troubleshooting

- **If it doesn't open**: Make sure Windows isn't blocking the file
- **If there are errors**: Run from PowerShell/CMD to see error messages
- **Antivirus**: Some antivirus software may flag the executable as suspicious (this is normal for .exe files compiled with PyInstaller)

Enjoy your new work hours tracking system! 🎉
