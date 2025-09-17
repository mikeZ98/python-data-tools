# Python Data Tools 🛠️🐍

A collection of small but powerful Python utilities for **automation, data processing, plotting, and PLC communication**.
Each tool lives in its own folder with a short README and can be run independently.

## 📂 Contents
- **backup/** – GUI file/folder backup tool (Tkinter + watchdog)
- **plots/** – merge CSV/Excel/DBF files and generate interactive plots
- **plc/** – PLC utilities (Siemens S7 over snap7) + Modbus demo notebook
- **planner/** – Excel planner & aggregator (merging, pivots, charts)

## 🚀 Quickstart
```bash
git clone https://github.com/<your_user>/python-data-tools.git
cd python-data-tools
python -m venv .venv && source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
```

Run any tool, e.g. **Backup**:
```bash
python backup/backup.py
```

## 🛠️ Requirements (key libs)
- pandas, numpy
- matplotlib, plotly
- watchdog
- openpyxl, xlsxwriter
- python-snap7 (for PLC)
- Jupyter (for Modbus notebook)

See [requirements.txt](requirements.txt) for the full list.

## 📄 License
Released under the MIT License.
