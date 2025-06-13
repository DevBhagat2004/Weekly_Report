# Automated Weekly Reports

A Python automation tool that processes raw data and generates formatted Excel reports with charts and summary statistics.

## Features

- **Data Cleaning**: Automatically cleans and processes raw CSV data
- **Excel Generation**: Creates professionally formatted Excel reports
- **Multiple Worksheets**: Summary, Raw Data, and Charts sheets
- **Data Visualization**: Automatic chart generation for key metrics
- **Configurable**: Easy to customize via JSON configuration
- **Logging**: Comprehensive logging for debugging and monitoring
- **Error Handling**: Robust error handling and data validation

## Installation

1. Clone or download the project files
2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Quick Start

1. **Prepare your data**: Place your raw data CSV file in the project directory (default: `raw_data.csv`)
2. **Run the script**:
```bash
python weekly_report_generator.py
```
3. **Find your report**: The generated Excel file will be saved with timestamp (e.g., `weekly_report_20241213.xlsx`)

## Project Structure

```
automated-weekly-reports/
├── weekly_report_generator.py  # Main script
├── requirements.txt           # Python dependencies
├── config.json               # Configuration settings
├── README.md                 # This file
├── raw_data.csv             # Your input data (created automatically if missing)
├── weekly_report_YYYYMMDD.xlsx  # Generated reports
└── weekly_report.log        # Log files
```

## Data Format

Your CSV file should contain columns like:
- **Date**: Date column (YYYY-MM-DD format preferred)
- **Product**: Product names or categories
- **Sales**: Sales figures
- **Units**: Unit quantities
- **Revenue**: Revenue amounts
- **Region**: Geographic regions

### Sample Data Format:
```csv
Date,Product,Sales,Units,Revenue,Region
2024-12-06,Product A,15,50,2500.00,North
2024-12-07,Product B,12,30,1800.00,South
2024-12-08,Product C,20,75,3750.00,East
```

## Configuration

Edit `config.json` to customize the tool:

```json
{
    "input_file": "your_data.csv",
    "output_file": "custom_report_{date}.xlsx",
    "data_columns": ["Date", "Product", "Sales", "Units", "Revenue", "Region"],
    "date_column": "Date",
    "numeric_columns": ["Sales", "Units", "Revenue"]
}
```

## Generated Report Structure

The Excel report contains three worksheets:

### 1. Summary Sheet
- Report generation date and time
- Key performance metrics
- Total, average, min, and max values
- Top performing products and regions

### 2. Raw Data Sheet
- Cleaned and filtered data
- Professional formatting
- Proper number formatting for currency and quantities

### 3. Charts Sheet
- Bar charts for top products by revenue
- Visual representations of key metrics
- Configurable chart types and colors

## Usage Examples

### Basic Usage
```python
from weekly_report_generator import WeeklyReportGenerator

# Initialize generator
generator = WeeklyReportGenerator()

# Run complete process
generator.run_full_process()
```

### Custom Configuration
```python
# Use custom config file
generator = WeeklyReportGenerator('custom_config.json')

# Load specific data file
generator.load_raw_data('monthly_data.csv')

# Clean data
generator.clean_data()

# Generate custom report
generator.create_excel_report('custom_report.xlsx')
```

## Advanced Features

### Logging
The script creates detailed logs in `weekly_report.log`:
- Data loading progress
- Cleaning operations
- Error messages and warnings
- Report generation status

### Data Cleaning Process
1. **Date Conversion**: Converts date strings to datetime objects
2. **Numeric Cleaning**: Removes currency symbols and converts to numbers
3. **Missing Data**: Removes rows with critical missing values
4. **Time Filtering**: Filters data for the current week
5. **Validation**: Validates data integrity

### Error Handling
- Multiple encoding detection for CSV files
- Graceful handling of missing columns
- Comprehensive exception catching
- Detailed error logging

## Customization Options

### Adding New Metrics
Modify the `generate_summary_stats()` method to include custom calculations:

```python
# Add custom metric
summary['Custom_Metric'] = self.cleaned_data['Revenue'].std()
```

### Custom Charts
Extend the `_create_charts_sheet()` method for additional visualizations:

```python
# Add pie chart
from openpyxl.chart import PieChart
pie_chart = PieChart()
# Configure chart...
```

### Different Data Sources
The tool can be extended to read from:
- Excel files (.xlsx, .xls)
- Database connections
- APIs
- Multiple CSV files

## Scheduling Automation

### Windows Task Scheduler
1. Open Task Scheduler
2. Create Basic Task
3. Set trigger for weekly execution
4. Set action to run Python script

### Linux/Mac Cron Job
```bash
# Edit crontab
crontab -e

# Add weekly execution (every Monday at 9 AM)
0 9 * * 1 /usr/bin/python3 /path/to/weekly_report_generator.py
```

### GitHub Actions (CI/CD)
```yaml
name: Weekly Report
on:
  schedule:
    - cron: '0 9 * * 1'  # Every Monday at 9 AM
jobs:
  generate-report:
    runs-on: ubuntu-latest
    steps:
    - uses: actions/checkout@v2
    - name: Set up Python
      uses: actions/setup-python@v2
      with:
        python-version: '3.9'
    - name: Install dependencies
      run: pip install -r requirements.txt
    - name: Generate report
      run: python weekly_report_generator.py
```

## Troubleshooting

### Common Issues

**1. "File not found" error**
- Ensure `raw_data.csv` exists in the project directory
- Check file path in `config.json`

**2. "Encoding error" when reading CSV**
- The script tries multiple encodings automatically
- If issues persist, save CSV as UTF-8

**3. "No data after cleaning" warning**
- Check date format in your data
- Verify numeric columns don't contain text

**4. Excel file won't open**
- Ensure no other program has the file open
- Check disk space and permissions

### Debug Mode
Enable detailed logging by modifying the logging level:
```python
logging.basicConfig(level=logging.DEBUG)
```

## Requirements

- Python 3.7+
- pandas 1.5.0+
- openpyxl 3.1.0+
- numpy 1.24.0+

## License

This project is open source and available under the MIT License.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review the log files
3. Create an issue with detailed error information

## Future Enhancements

- Web interface for non-technical users
- Email integration for automatic report distribution
- Dashboard with real-time metrics
- Integration with business intelligence tools
- Multi-format export (PDF, PowerPoint)
- Advanced statistical analysis
- Machine learning insights
