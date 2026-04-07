# 📊 Use Marco in Excel — UiPath Automation

A UiPath RPA project that automates Excel file formatting and editing using macros, eliminating manual repetitive work on spreadsheets.

---

## 📋 Overview

This automation runs a UiPath workflow that opens an Excel file, executes a predefined macro (VBA), and applies formatting or edits automatically — no human interaction required.

**Use case:** Automatically format invoice files (`invoice_edit.xlsm`) according to a defined template, saving time and reducing human error.

---

## 🗂️ Project Structure

```
Use-Marco-in-Excel/
├── Main.xaml            # Main UiPath workflow entry point
├── entry-points.json    # Workflow entry point configuration
├── project.json         # UiPath project metadata
├── project.uiproj       # UiPath Studio project file
├── invoice_edit.xlsm    # Excel file with embedded macro
├── format.txt           # Formatting rules/reference
├── .gitignore
└── README.md
```

---

## ⚙️ Prerequisites

Before running this project, make sure you have:

| Requirement | Version |
|---|---|
| UiPath Studio | 2022.4 or later |
| Microsoft Excel | 2016 or later |
| .NET Framework | 4.6.1+ |

> ⚠️ **Important:** Excel macros must be **enabled**. Go to Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros.

---

## 🚀 Getting Started

### For Developers

1. **Clone the repository**
   ```bash
   git clone https://github.com/nanashi193/Use-Marco-in-Excel.git
   ```

2. **Open in UiPath Studio**
   - Launch UiPath Studio
   - Click **Open** → navigate to the cloned folder
   - Select `project.json`

3. **Run the workflow**
   - Open `Main.xaml`
   - Click **Run** (F5) or **Debug** (F7)

### For Business Users

1. Make sure UiPath Robot is installed and running on your machine.
2. Ask your developer/IT team to publish this process to UiPath Orchestrator.
3. Trigger the job from Orchestrator or UiPath Assistant.
4. The formatted Excel file will be saved automatically.

---

## 🔄 How It Works

```
Start
  └─► Open Excel file (invoice_edit.xlsm)
        └─► Run Excel Macro (VBA)
              └─► Apply formatting rules
                    └─► Save file
                          └─► Close Excel → Done ✅
```

1. UiPath opens the target `.xlsm` file using the **Excel Application Scope** activity.
2. The **Invoke VBA** or **Execute Macro** activity triggers the embedded macro.
3. The macro applies formatting defined in `format.txt` (e.g., column widths, fonts, borders, number formats).
4. The file is saved and closed automatically.

---

## 📝 Configuration

Edit `project.json` or the workflow variables in `Main.xaml` to customize:

| Variable | Description | Default |
|---|---|---|
| `filePath` | Path to the Excel file | `invoice_edit.xlsm` |
| `macroName` | Name of the VBA macro to run | *(set in Main.xaml)* |
| `sheetName` | Target worksheet name | *(set in Main.xaml)* |

---

## 🛠️ Troubleshooting

**Macro not running?**
- Ensure macros are enabled in Excel Trust Center settings.
- The file must be saved as `.xlsm` (macro-enabled workbook), not `.xlsx`.

**UiPath cannot find the file?**
- Check that the `filePath` variable points to the correct absolute or relative path.

**Excel stays open after error?**
- Enable the **Kill Process** activity in the exception handler to force-close Excel.

---

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit your changes: `git commit -m "feat: add your feature"`
4. Push and open a Pull Request

---

## 👤 Author

**nanashi193** — [GitHub Profile](https://github.com/nanashi193)

---

## 📄 License

This project is for internal/demo use. No license specified.