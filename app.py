import os
import csv
import pandas as pd
from flask import Flask, request, render_template, send_file
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_table():
    try:
        # Get the raw table input from the form
        table_data = request.form['table_data'].strip().split('\n')
        table = [row.split(',') for row in table_data]

        # Validate the header
        header = table[0]
        if header[-1].strip().lower() != "total":
            return "Error: The last column must be named 'Total'. Please go back and correct the input."

        # Calculate totals for each row
        for i in range(1, len(table)):
            try:
                scores = [int(x.strip()) for x in table[i][1:-1]]
                table[i][-1] = str(sum(scores))
            except ValueError:
                return f"Invalid scores in row {i + 1}. Please check the input."

        # Save files to in-memory buffers
        excel_buffer = BytesIO()
        csv_buffer = BytesIO()
        md_buffer = BytesIO()

        # DataFrame for saving to XLSX
        df = pd.DataFrame(table[1:], columns=header)
        df.to_excel(excel_buffer, index=False, engine='xlsxwriter')
        excel_buffer.seek(0)

        # Save as CSV
        csv_writer = csv.writer(csv_buffer)
        csv_writer.writerows(table)
        csv_buffer.seek(0)

        # Save as Markdown
        md_content = f"| {' | '.join(header)} |\n"
        md_content += f"|{'|'.join(['-' * len(h) for h in header])}|\n"
        for row in table[1:]:
            md_content += f"| {' | '.join(row)} |\n"
        md_buffer.write(md_content.encode('utf-8'))
        md_buffer.seek(0)

        # Respond with the files for download
        return (
            send_file(excel_buffer, as_attachment=True, download_name="completed_table.xlsx"),
            send_file(csv_buffer, as_attachment=True, download_name="completed_table.csv"),
            send_file(md_buffer, as_attachment=True, download_name="completed_table.md")
        )

    except Exception as e:
        return f"An error occurred: {e}"

if __name__ == '__main__':
    port = int(os.getenv("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
