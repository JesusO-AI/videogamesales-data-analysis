# Videogames Sales Data Analytics Pipeline
Python data pipeline that processes global video games sales data, generates multi-dimensional summaries using pandas, creates custom visualizations with matplotlib, and compiles everything into an executive-ready Excel dashboard using XLsxWriter.

## How to run the project
To run pipeline locally and generate the dashboard, please follow these steps:

## 1 Prerequisites
Make sure you have python installed, then install the required data science dependencies via your terminal.
```bash
pip install pandas matplotlib xlsxwriter
```

## 2 Project layout
Ensure your local project structure matches the following setup:
- Main.py(Core data script)
- data_to_analyze/vgsales.csv(Raw input dataset)
- result/(Target directory where spreadsheet and dashboard are generated)

## Execution
Run pipeline from terminal
```bash
python main.py
```

When execution is completed, a fully formatted spreadsheet named analysed.xlsx along with those trend charts will be automatically generated inside your directory.

## Pipeline features and breakdowns
The Main.py script automatically identifies the raw data to transform it into several targeted analytical blocks.
- Market analysis: Shows total software copy sales sliced by Platform, Genre, Publisher, Year, and Game Title.
- Released volatility: Track which publishing giants push the highest volume of special games to the market.
- Historical trends: Isolates Top-Categories over time to see where revenue market shifted over the years.

## Tech Stack and libraries used
- Python(Core Language)
- Pandas(Data Manipulation, Grouping and aggregation)
- Matplotlib(Data Visualization and Trend Plotting)
- XLsxWriter(Excel spreadsheet Design, Cell Borders and conditional formatting)
