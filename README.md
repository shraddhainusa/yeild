# Yield Performance Dashboard

A comprehensive Streamlit dashboard for analyzing and visualizing yield performance metrics from Excel data.

## 📊 Features

### 1. Executive Summary
- Average MRP Plan %
- Average Actual %
- Average Variance %
- Total Number of Orders
- Total Start Quantity
- Total Delivered Quantity
- Average Weighted Average %

### 2. Yield History Overall
- Interactive table with all overall yield data
- Filters for Material and Variance range
- Bar chart comparing MRP Plan vs Actual by Material
- Visual highlighting of materials where Actual is below MRP Plan

### 3. Monthly Yield Trend
- Data automatically reshaped from wide to long format
- Monthly columns reformatted into: Material, Period, MRP Plan, Actual, Variance
- Line chart for MRP Plan vs Actual over time
- Bar/Line chart for Variance trends
- Material filter for focused analysis

### 4. 90 Days Rolling History
- 9 Key Performance Indicator cards:
  - Total GDPW
  - Average In House Production
  - Total Orders
  - Total Order Quantity
  - Total Goods Received
  - Total Yield Confirmed
  - Total Scrap Confirmed
  - Average Actual Yield %
  - Average Routing Yield %
  - Average MRP Yield %
- Interactive filters for Material and Fiscal Year Week
- Data table with all metrics
- Multiple visualization charts

### 5. 90 Day Pivot
- Pivot summary table display
- Comparison chart: Average MRP Yield % vs Average Actual Yield %
- Interactive visualization for trend analysis

## 🎨 Design Features

- ✅ Clean and professional layout
- ✅ Sidebar filters for easy navigation
- ✅ Clear chart titles and labels
- ✅ Proper percentage formatting
- ✅ Responsive design that works on all screen sizes
- ✅ Comprehensive code comments

## 📋 Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. Clone or download this repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Place the Excel file (`dummy_yield_history_data_updated.xlsx`) in the same directory as `app.py`

## 🚀 Running the Dashboard

```bash
streamlit run app.py
```

The dashboard will open in your default browser at `http://localhost:8501`

## 📁 File Structure

```
.
├── app.py                                    # Main Streamlit application
├── requirements.txt                          # Python dependencies
├── README.md                                 # This file
└── dummy_yield_history_data_updated.xlsx    # Excel data source
```

## 📊 Excel Data Format

The Excel workbook should contain the following sheets:

### 1. Yield History Overall
Contains columns:
- Material
- MRP Plan %
- Actual %
- Variance %
- Number of Orders
- Start Quantity
- Delivered Quantity
- Weighted Average %

### 2. Yield History Monthly
**Wide Format** (automatically converted to long format by the app):
- Material | Period1_MRP Plan | Period1_Actual | Period1_Variance | Period2_MRP Plan | ...

### 3. 90 Days Rolling History
Contains columns:
- Material
- GDPW
- In House Production
- Orders
- Order Quantity
- Goods Received
- Yield Confirmed
- Scrap Confirmed
- Actual Yield %
- Routing Yield %
- MRP Yield %
- Fiscal Year Week

### 4. 90 Day Pivot
Contains summary pivot data with:
- Category/Material column
- Average MRP Yield %
- Average Actual Yield %

## 🔍 Key Dashboard Sections

### Sidebar Navigation
- **Executive Summary**: Overview metrics
- **Yield History Overall**: Detailed yield analysis with filters
- **Monthly Yield Trend**: Time-series trend analysis
- **90 Days Rolling History**: Recent performance indicators
- **90 Day Pivot**: Summary statistics

## 🎯 Usage Tips

1. **Filters**: Use sidebar filters to focus on specific materials or time periods
2. **Charts**: Hover over charts to see detailed values
3. **Tables**: Click column headers to sort or use the search function
4. **Responsive**: The dashboard auto-adjusts to different screen sizes

## 📈 Performance Metrics

- **MRP Plan %**: Planned yield percentage
- **Actual %**: Achieved yield percentage
- **Variance %**: Difference between actual and plan
- **GDPW**: Goods Produced Per Week
- **Yield Confirmed**: Confirmed yielded quantity

## ⚙️ Dependencies

- **streamlit** (1.35.0): Web app framework
- **pandas** (2.0.3): Data manipulation and analysis
- **numpy** (1.24.3): Numerical computing
- **plotly** (5.17.0): Interactive charting
- **openpyxl** (3.10.10): Excel file handling

## 🐛 Troubleshooting

### File Not Found
- Ensure `dummy_yield_history_data_updated.xlsx` is in the same directory as `app.py`

### Sheet Not Found
- Verify that all required sheets exist in the Excel file with correct names:
  - Yield History Overall
  - Yield History Monthly
  - 90 Days Rolling History
  - 90 Day Pivot

### Empty Charts
- Check that column names in Excel match expected patterns (e.g., contain "MRP", "Actual", "Yield")
- Ensure data is properly formatted (numbers, not text)

## 📝 Notes

- The dashboard automatically caches data for improved performance
- All percentages are displayed with 2 decimal places
- Large numbers are formatted with thousands separators
- The monthly data reshaping handles flexible date formats

## 👨‍💼 Support

For issues or questions, please check:
1. Excel file format matches the required structure
2. All sheet names are exact matches (case-sensitive)
3. Column headers contain the expected keywords

## 📄 License

This project is provided as-is for data analysis purposes.

## ✨ Future Enhancements

- Add export to PDF functionality
- Implement data refresh scheduling
- Add more advanced statistical analysis
- Create custom date range selection
- Add anomaly detection
