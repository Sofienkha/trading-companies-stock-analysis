import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# Read the CSV file
def csv_to_excel():
    try:
        # Read the CSV data
        df = pd.read_csv('trading_companies_stock_data.csv')
        
        # Create a new Excel writer
        writer = pd.ExcelWriter('trading_companies_stock_data.xlsx', engine='xlsxwriter')
        
        # Write the dataframe to an Excel sheet
        df.to_excel(writer, sheet_name='Stock Data', index=False)
        
        # Get the workbook and the worksheet
        workbook = writer.book
        worksheet = writer.sheets['Stock Data']
        
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Write the column headers with the defined format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        # Set column widths
        worksheet.set_column('A:A', 6)  # Ticker
        worksheet.set_column('B:B', 25)  # Company Name
        worksheet.set_column('C:L', 12)  # Other columns
        
        # Add conditional formatting to highlight stocks with high P/E ratios
        worksheet.conditional_format('E2:E32', {'type': 'cell',
                                              'criteria': '>',
                                              'value': 30,
                                              'format': workbook.add_format({'bg_color': '#FFC7CE'})})
        
        # Add conditional formatting to highlight stocks with high dividend yields
        worksheet.conditional_format('H2:H32', {'type': 'cell',
                                              'criteria': '>',
                                              'value': 3,
                                              'format': workbook.add_format({'bg_color': '#C6EFCE'})})
        
        # Create a new sheet for charts
        price_chart = workbook.add_worksheet('Price Charts')
        
        # Create a column chart for current prices
        chart1 = workbook.add_chart({'type': 'column'})
        
        # Configure the series for the top 10 companies by price
        top_prices = df.sort_values('Current Price ($)', ascending=False).head(10)
        
        # Add the series to the chart
        chart1.add_series({
            'name': 'Current Stock Price',
            'categories': ['Stock Data', 1, df.columns.get_loc('Ticker'), 
                          len(top_prices), df.columns.get_loc('Ticker')],
            'values': ['Stock Data', 1, df.columns.get_loc('Current Price ($)'), 
                      len(top_prices), df.columns.get_loc('Current Price ($)')],
            'data_labels': {'value': True}
        })
        
        # Configure the chart
        chart1.set_title({'name': 'Top 10 Companies by Stock Price'})
        chart1.set_x_axis({'name': 'Company Ticker'})
        chart1.set_y_axis({'name': 'Stock Price ($)'})
        chart1.set_style(11)
        
        # Insert the chart into the worksheet
        price_chart.insert_chart('A2', chart1, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Create a pie chart for market cap distribution
        chart2 = workbook.add_chart({'type': 'pie'})
        
        # Get the top 5 companies by market cap
        top_cap = df.sort_values('Market Cap ($B)', ascending=False).head(5)
        
        # Add the series to the chart
        chart2.add_series({
            'name': 'Market Cap Distribution',
            'categories': ['Stock Data', 1, df.columns.get_loc('Ticker'), 
                          len(top_cap), df.columns.get_loc('Ticker')],
            'values': ['Stock Data', 1, df.columns.get_loc('Market Cap ($B)'), 
                      len(top_cap), df.columns.get_loc('Market Cap ($B)')],
            'data_labels': {'percentage': True}
        })
        
        # Configure the chart
        chart2.set_title({'name': 'Top 5 Companies by Market Cap'})
        chart2.set_style(10)
        
        # Insert the chart into the worksheet
        price_chart.insert_chart('A20', chart2, {'x_scale': 1.2, 'y_scale': 1.2})
        
        # Create a new sheet for financial analysis
        analysis = workbook.add_worksheet('Financial Analysis')
        
        # Add summary statistics
        analysis.write('A1', 'Financial Metrics Summary', workbook.add_format({'bold': True, 'font_size': 14}))
        analysis.write('A3', 'Metric', header_format)
        analysis.write('B3', 'Average', header_format)
        analysis.write('C3', 'Median', header_format)
        analysis.write('D3', 'Min', header_format)
        analysis.write('E3', 'Max', header_format)
        
        # Calculate statistics for key metrics
        metrics = ['Current Price ($)', 'Market Cap ($B)', 'P/E Ratio', 'Dividend Yield (%)', 'EPS ($)', 'Beta']
        
        for i, metric in enumerate(metrics):
            analysis.write(4+i, 0, metric)
            analysis.write(4+i, 1, df[metric].mean())
            analysis.write(4+i, 2, df[metric].median())
            analysis.write(4+i, 3, df[metric].min())
            analysis.write(4+i, 4, df[metric].max())
        
        # Add a scatter plot for P/E ratio vs Market Cap
        chart3 = workbook.add_chart({'type': 'scatter'})
        
        # Configure series
        chart3.add_series({
            'name': 'P/E Ratio vs Market Cap',
            'categories': ['Stock Data', 1, df.columns.get_loc('Market Cap ($B)'), 
                          len(df), df.columns.get_loc('Market Cap ($B)')],
            'values': ['Stock Data', 1, df.columns.get_loc('P/E Ratio'), 
                      len(df), df.columns.get_loc('P/E Ratio')],
            'marker': {'type': 'circle', 'size': 8}
        })
        
        # Configure the chart
        chart3.set_title({'name': 'P/E Ratio vs Market Cap'})
        chart3.set_x_axis({'name': 'Market Cap ($B)'})
        chart3.set_y_axis({'name': 'P/E Ratio'})
        chart3.set_style(11)
        
        # Insert the chart into the worksheet
        analysis.insert_chart('G3', chart3, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # Save the Excel file
        writer.close()
        
        print(f"Excel file created successfully: trading_companies_stock_data.xlsx")
        return True
    except Exception as e:
        print(f"Error creating Excel file: {e}")
        return False

if __name__ == "__main__":
    csv_to_excel()
