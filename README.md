# Work Hours Tracker

A Python application that helps you track and calculate your daily work hours, with automatic Excel report generation and overtime calculations.

## 📋 Features

- **Daily Work Hours Tracking**: Record multiple time entries per day
- **Automatic Calculations**: Calculate total worked hours and overtime
- **Excel Report Generation**: Creates formatted Excel files with work history
- **Calendar Integration**: Monthly calendar view with work hours
- **Portable Executable**: Standalone .exe file for easy distribution
- **User-Friendly Interface**: Simple command-line interface with clear instructions

## 🚀 Quick Start

### For End Users (Windows)
1. Download the latest release from the [Releases](https://github.com/yourusername/work-hours-tracker/releases) page
2. Extract the ZIP file
3. Double-click `Run_Work_Hours_Tracker.bat` or `Work_Hours_Tracker_COMPILED.exe`
4. Follow the on-screen instructions

### For Developers

#### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

#### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/work-hours-tracker.git
   cd work-hours-tracker
   ```

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # On Windows
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the application:
   ```bash
   python main.py
   ```

## 📦 Building the Executable

To create a standalone executable:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Build the executable:
   ```bash
   pyinstaller --onefile --name Work_Hours_Tracker_COMPILED main.py
   ```

3. The executable will be created in the `dist/` folder

## 📁 Project Structure

```
Work_Hours_Tracker/
├── main.py                 # Main application code
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── INSTRUCTIONS.md        # Detailed usage instructions
├── .gitignore            # Git ignore rules
├── dist/                 # Distribution files (executable, etc.)
│   ├── Work_Hours_Tracker_COMPILED.exe
│   ├── README_For_Users.txt
│   ├── INSTALLATION_GUIDE.txt
│   └── Run_Work_Hours_Tracker.bat
└── build/                # Build files (generated)
```

## 🎯 Usage

1. **Date Selection**: Choose today, yesterday, or enter a custom date
2. **Time Entries**: Enter your work times in HH.MM format (e.g., 08.42 for 8:42 AM)
3. **Calculation**: The program automatically calculates:
   - Total worked hours
   - Standard workday hours (8 hours)
   - Overtime hours
4. **Excel Report**: A formatted Excel file is created with your work history

## 📊 Output

The program generates an Excel file (`work_hours_history.xlsx`) containing:
- Daily work hours summary
- Monthly calendar view
- Overtime calculations
- Formatted tables with color coding

## 🔧 Configuration

- **Standard Workday**: 8 hours (configurable in the code)
- **Time Format**: HH.MM (24-hour format)
- **Excel Filename**: `work_hours_history.xlsx`

## 📝 Dependencies

- `pandas`: Data manipulation and analysis
- `openpyxl`: Excel file creation and formatting
- `datetime`: Date and time handling

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

If you encounter any issues:
1. Check the [Issues](https://github.com/yourusername/work-hours-tracker/issues) page
2. Create a new issue with detailed information
3. Include your operating system and Python version

## 🔄 Version History

- **v1.0.0**: Initial release with basic work hours tracking
- **v1.1.0**: Added Excel report generation
- **v1.2.0**: Added calendar view and improved formatting

---

**Note**: Replace `yourusername` in the GitHub URLs with your actual GitHub username when you create the repository.
