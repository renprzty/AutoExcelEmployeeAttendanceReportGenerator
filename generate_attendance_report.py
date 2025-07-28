import pandas as pd
from sqlalchemy import create_engine
import os

def generate_attendance_report():
    # Connect to database (replace with your DB credentials)
    engine = create_engine("sqlite:///hr.db")

    # Query attendance data
    query = """
    SELECT employee_id, name, date, status
    FROM attendance_logs
    WHERE strftime('%Y-%m', date) = '2025-07'
    """
    df = pd.read_sql(query, engine)
    df['date'] = pd.to_datetime(df['date'])

    # Save raw data
    raw_data = df.copy()

    # Create summary table
    summary = df.groupby(['employee_id', 'name', 'status']).size().unstack(fill_value=0)
    summary['Total Days'] = summary.sum(axis=1)
    summary['Attendance %'] = (summary.get('Present', 0) / summary['Total Days'])
    summary['Attendance %'] = summary['Attendance %'].round(2)

    summary = summary.reset_index()

    # Sort for top attendance
    top5 = summary.sort_values(by='Attendance %', ascending=False).head(5)

    # Export to Excel
    output_file = 'employee_attendance_report.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Sheet 1: Raw Data
        raw_data.to_excel(writer, sheet_name='Raw Attendance', index=False)

        # Sheet 2: Summary
        summary.to_excel(writer, sheet_name='Monthly Summary', index=False)

        # Excel objects
        workbook = writer.book
        summary_sheet = writer.sheets['Monthly Summary']
        raw_sheet = writer.sheets['Raw Attendance']

        # Format % in summary
        percentage_format = workbook.add_format({'num_format': '0%'})
        summary_sheet.set_column('F:F', 12, percentage_format)

        # Format date in raw data
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        raw_sheet.set_column('C:C', 20, date_format)

        # Autofilter in summary
        summary_sheet.autofilter(0, 0, summary.shape[0], summary.shape[1] - 1)

        # Conditional formatting (highlight if < 80%)
        summary_sheet.conditional_format(f'F2:F{len(summary)+1}', {
            'type': 'cell',
            'criteria': '<',
            'value': 0.8,
            'format': workbook.add_format({'font_color': 'red', 'bold': True})
        })

        # Chart
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Top Attendance %',
            'categories': ['Monthly Summary', 1, 1, 5, 1],
            'values':     ['Monthly Summary', 1, 5, 5, 5],
        })
        chart.set_title({'name': 'Top 5 Employee Attendance (%)'})
        chart.set_y_axis({'name': 'Attendance %'})
        chart.set_x_axis({'name': 'Employee Name'})
        summary_sheet.insert_chart('H2', chart)

    print(f"Attendance report generated: {os.path.abspath(output_file)}")

generate_attendance_report()
