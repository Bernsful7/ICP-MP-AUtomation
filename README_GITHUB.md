# ICP/MP Expert Automation Suite

A comprehensive automation toolkit for Agilent ICP/MP Expert instruments, providing GUI, CLI, and API interfaces for instrument control, worksheet management, and LIMS integration.

## 🎯 Features

### **GUI Application** (`gui_automation_app.py`)
- **Complete Instrument Control**: Plasma, pump, purge, and measurement controls
- **Comprehensive Worksheet Management**: Create, open, save, and manage worksheets
- **LIMS Export**: Direct export to LIMS format with configurable paths
- **Sample Management**: Queue management with automatic sample detection
- **Real-time Status**: Live instrument status monitoring
- **Auto-connection**: Automatic connection at startup with graceful error handling

### **CLI Tools**
- **`automation_cli.py`**: Command-line interface for scripted automation
- **`automation_demo_app.py`**: Demonstration application with examples
- **`advanced_automation_app.py`**: Advanced automation features

### **Core Components**
- **`Automation.py`**: Python wrapper for Agilent Automation SDK
- **Complete SDK Integration**: Full coverage of available automation methods
- **Robust Error Handling**: Graceful socket disconnection and recovery
- **Cross-platform Support**: Windows with IronPython 2.7

## 🚀 Quick Start

### Prerequisites
- **Windows OS** (required for Agilent SDK)
- **IronPython 2.7** (for .NET Framework integration)
- **MP Expert Software** (running and accessible)

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/YOUR_USERNAME/icp-mp-expert-automation.git
   cd icp-mp-expert-automation
   ```

2. **Install IronPython 2.7:**
   - Download from: https://github.com/IronLanguages/ironpython2/releases
   - Install to default location or update paths in batch files

3. **Run the GUI application:**
   ```bash
   ipy.exe gui_automation_app.py
   ```

## 📱 GUI Application

### **Window Layout (1100x800)**
```
┌─────────────────── 1100px ──────────────────┐
│ Connection(300x80)  Status(320x180)  Sample(280x320) │ 
│                                                      │
│ Control(300x250)   Worksheet(320x130) Advanced(280x110) │
│                                                      │
│ Activity Log (920x250)                              │
└──────────────────── 800px ─────────────────────────┘
```

### **Key Features:**
- ✅ **Auto-connection** at startup (127.0.0.1:8000)
- ✅ **Sample Detection** from opened worksheets
- ✅ **File Type Support** (.mpts templates, .mpws worksheets)
- ✅ **LIMS Integration** with configurable export paths
- ✅ **Real-time Logging** with activity tracking

## 🛠️ Development

### **Project Structure**
```
├── gui_automation_app.py          # Main GUI application
├── Automation.py                  # SDK wrapper
├── automation_cli.py              # CLI interface
├── automation_demo_app.py         # Demo application
├── advanced_automation_app.py     # Advanced features
├── docs/                          # Documentation
│   ├── INSTALLATION_GUIDE.md      # Setup instructions
│   ├── IRONPYTHON_GUIDE.md        # IronPython setup
│   └── CODE_REVIEW_REPORT.md      # Code review notes
├── tests/                         # Test files
│   ├── test_gui_syntax.py         # Syntax validation
│   ├── test_ironpython.py         # IronPython compatibility
│   └── comprehensive_test.py      # Full test suite
└── lib/                           # Required DLLs
    ├── Agilent.AtomicSpectroscopy.SDK.Protocol.dll
    ├── TransportProtocol.dll
    ├── XdrSocketClient.dll
    └── xdr.dll
```

### **Setting Up Development Environment**

1. **Install dependencies:**
   ```bash
   # IronPython 2.7 required for .NET integration
   # Standard Python packages are included in Lib/ folder
   ```

2. **Configure IronPython path:**
   ```bash
   # Run the setup script
   setup_ironpython_path.bat
   ```

3. **Test the installation:**
   ```bash
   ipy.exe test_ironpython.py
   ```

## 📋 Usage Examples

### **GUI Application**
```bash
# Launch the main GUI
ipy.exe gui_automation_app.py
```

### **CLI Automation**
```bash
# Command-line automation
ipy.exe automation_cli.py --connect --plasma-on --start
```

### **Demo Application**
```bash
# Run demonstration
ipy.exe automation_demo_app.py
```

## 🔧 Configuration

### **Connection Settings**
- **Default Host**: 127.0.0.1
- **Default Port**: 8000
- **Auto-connect**: Enabled at startup

### **File Types**
- **Templates**: `.mpts` files
- **Worksheets**: `.mpws` files
- **Export**: Configurable LIMS format

### **SDK Methods Supported**
- **Instrument Control**: PlasmaOn/Off, PumpSlow/Fast/Off, PurgeOn/Off
- **Measurements**: Start, Stop, SelectSolution
- **Worksheets**: WorksheetNew, WorksheetOpen, WorksheetSaveAs, WorksheetSaveClose
- **Data Export**: Export, DeleteResults
- **UI Control**: ShowUI, HideUI

## 🧪 Testing

```bash
# Syntax validation
ipy.exe test_gui_syntax.py

# IronPython compatibility
ipy.exe test_ironpython.py

# Comprehensive testing
ipy.exe comprehensive_test.py
```

## 📚 Documentation

- **[Installation Guide](INSTALLATION_GUIDE.md)**: Complete setup instructions
- **[IronPython Guide](IRONPYTHON_GUIDE.md)**: IronPython configuration
- **[Code Review](CODE_REVIEW_REPORT.md)**: Development notes and reviews
- **[Automation Getting Started](Automation_Getting_Started.pdf)**: Official Agilent documentation

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Troubleshooting

### **Common Issues:**

1. **"Module not found" errors**:
   - Ensure IronPython 2.7 is installed
   - Check that required DLLs are in the same directory

2. **Connection failures**:
   - Verify MP Expert is running
   - Check host/port settings (default: 127.0.0.1:8000)
   - Ensure firewall allows connections

3. **UI not showing**:
   - Try the ShowUI button in the GUI
   - Check if MP Expert is minimized

### **Getting Help:**
- Check the [Issues](https://github.com/YOUR_USERNAME/icp-mp-expert-automation/issues) page
- Review documentation in the `docs/` folder
- Run diagnostic tests in the `tests/` folder

## 🔗 Related Resources

- [IronPython Official Site](https://ironpython.net/)
- [Agilent MP Expert Documentation](https://www.agilent.com/)
- [.NET Framework Documentation](https://docs.microsoft.com/en-us/dotnet/)

---

**Built with ❤️ for the analytical chemistry community**
