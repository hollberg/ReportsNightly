import pandas as pd
import plotly.express as px
import re
from datetime import datetime, timedelta
from SQL import queries  # Stores Mitch's SQL Scripts

QUERIES = queries.QUERIES

cursor = queries.cursor
engine = queries.engine

# 1. Load data
df_funds = pd.read_sql_query(QUERIES['open_daf_funds'], engine)
df_files = pd.read_sql_query(QUERIES['fund_files'], engine)

# Check the column names
print("Columns in df_funds:", df_funds.columns)
print("Columns in df_files:", df_files.columns)

# 2. Update df_funds with counts of files / file types
file_counts = df_files.groupby('funit_id').size().reset_index(name='total_files')


# 3. Count files with "fund agreement" in the name
def has_fund_agreement(name):
    return bool(re.search(r'fund\s+agreement', name, re.IGNORECASE))


agreement_counts = df_files.groupby('funit_id').apply(
    lambda x: x['file_name'].apply(has_fund_agreement).sum()
).reset_index(name='fund_agreement_in_name')

# 4. Count signed fund agreements
signed_agreements = df_files[df_files['File Category'] == 'Fund - Fund Agreement - SIGNED'].groupby(
    'funit_id').size().reset_index(name='fund_agreement_signed')

# Merge all the new information with df_funds
df_funds = df_funds.merge(file_counts, on='funit_id', how='left')
df_funds = df_funds.merge(agreement_counts, on='funit_id', how='left')
df_funds = df_funds.merge(signed_agreements, on='funit_id', how='left')

# Fill NaN values with 0 for funds without any files
df_funds[['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']] = df_funds[
    ['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']].fillna(0)

# Convert columns to integer type
df_funds[['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']] = df_funds[
    ['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']].astype(int)

print(df_funds[['funit_id', 'total_files', 'fund_agreement_in_name', 'fund_agreement_signed']])

# Get today's date -2 (CSuite data is on a 2 day lag)
date_minus_two = datetime.today() - timedelta(days=2)

# Create the summary DataFrame using the correct column name
summary_df = df_funds.groupby('Steward Name').agg({
    'funit_id': 'count',
    'fund_agreement_in_name': lambda x: (x > 0).sum(),
    'fund_agreement_signed': lambda x: (x > 0).sum()
}).reset_index()

summary_df.columns = ['Steward Name', 'Total Funds', '# Funds with "Fund Agreement" in filename',
                      '# Funds with Fund Agreement signed']

# Create the bar plot using Plotly Express
fig = px.bar(summary_df, x='Steward Name',
             y=['Total Funds', '# Funds with "Fund Agreement" in filename', '# Funds with Fund Agreement signed'],
             title="Fund Summary by Steward", barmode='group')

# Save the plot and the data as an ASPX file
with open("fund_steward_summary.aspx", "w") as f:
    f.write(fig.to_html(full_html=False, include_plotlyjs='cdn'))

    f.write("""
    <div id="button-container">
        <button onclick="clearSelection()">Clear Selection</button>
        <button onclick="exportTableToCSV()">Export to CSV</button>
        <button onclick="exportTableToExcel()">Export to Excel</button>
    </div>
    <div id="table-container" style="max-height: 400px; overflow-y: auto;"></div>
    <script src="https://unpkg.com/tablesort@5.2.1/dist/tablesort.min.js"></script>
    <script src="https://unpkg.com/tablefilter@0.2.20/dist/tablefilter/tablefilter.min.js"></script>
    <script>
    let selectedStewards = [];
    const fundsData = """ + df_funds.rename(columns=lambda x: x.replace(' ', '_')).to_json(orient='records') + """;

    function showTable() {
        const tableContainer = document.getElementById('table-container');
        tableContainer.innerHTML = '';  // Clear previous table

        const filteredData = fundsData.filter(fund => selectedStewards.includes(fund['Steward_Name']));

        if (filteredData.length === 0) {
            tableContainer.innerHTML = '<p>No funds available for this steward.</p>';
            return;
        }

        const table = document.createElement('table');
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
        table.setAttribute('id', 'fundTable');

        const headerRow = document.createElement('tr');
        const headers = ['Fund_Name', 'funit_id', 'total_files', 'fund_agreement_in_name', 'fund_agreement_signed'];
        headers.forEach(headerText => {
            const header = document.createElement('th');
            header.style.border = '1px solid black';
            header.style.padding = '8px';
            header.style.backgroundColor = '#f2f2f2';
            header.style.position = 'sticky';
            header.style.top = '0';
            header.style.zIndex = '1';
            header.appendChild(document.createTextNode(headerText.replace('_', ' ')));
            headerRow.appendChild(header);
        });
        table.appendChild(headerRow);

        filteredData.forEach(rowData => {
            const row = document.createElement('tr');
            headers.forEach(header => {
                const cell = document.createElement('td');
                cell.style.border = '1px solid black';
                cell.style.padding = '8px';
                if (header === 'funit_id') {
                    const link = document.createElement('a');
                    link.href = `https://cfga.fcsuite.com/erp/funit/display?id=${rowData[header]}`;
                    link.textContent = rowData[header];
                    link.target = '_blank';
                    cell.appendChild(link);
                } else {
                    cell.appendChild(document.createTextNode(rowData[header] || ''));
                }
                row.appendChild(cell);
            });
            table.appendChild(row);
        });

        tableContainer.appendChild(table);

        // Apply sorting and filtering
        new Tablesort(table);
        const tf = new TableFilter(table, {
            base_path: 'https://unpkg.com/tablefilter@0.2.20/dist/tablefilter/',
            paging: {
                length: 10
            },
            rows_counter: true,
            btn_reset: true
        });
        tf.init();

        // Resize the plot when the table is shown
        const plotDiv = document.getElementsByClassName('plotly-graph-div')[0];
        plotDiv.style.height = '50%';
        Plotly.Plots.resize(plotDiv);
    }

    function clearSelection() {
        selectedStewards = [];
        const plotDiv = document.getElementsByClassName('plotly-graph-div')[0];
        Plotly.restyle(plotDiv, 'marker.color', 'rgba(0, 116, 217, 0.7)');
        document.getElementById('table-container').innerHTML = '';
        plotDiv.style.height = '100%';
        Plotly.Plots.resize(plotDiv);
    }

    function exportTableToCSV() {
        const tableContainer = document.getElementById('table-container');
        const rows = tableContainer.getElementsByTagName('tr');
        let csvContent = "";
        for (let i = 0; i < rows.length; i++) {
            const cols = rows[i].querySelectorAll('td, th');
            let row = [];
            for (let j = 0; j < cols.length; j++) {
                row.push(cols[j].innerText);
            }
            csvContent += row.join(",") + "\\n";
        }

        const blob = new Blob([csvContent], { type: 'text/csv' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'table_data.csv';
        link.click();
    }

    function exportTableToExcel() {
        const tableContainer = document.getElementById('table-container');
        const rows = tableContainer.getElementsByTagName('tr');
        let csvContent = "";
        for (let i = 0; i < rows.length; i++) {
            const cols = rows[i].querySelectorAll('td, th');
            let row = [];
            for (let j = 0; j < cols.length; j++) {
                row.push(cols[j].innerText);
            }
            csvContent += row.join("\\t") + "\\n";
        }

        const blob = new Blob([csvContent], { type: 'application/vnd.ms-excel' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'table_data.xls';
        link.click();
    }

    const plot = document.getElementsByClassName('plotly-graph-div')[0];
    plot.on('plotly_click', function(data){
        const steward = data.points[0].x;
        if (selectedStewards.includes(steward)) {
            selectedStewards = selectedStewards.filter(item => item !== steward);
        } else {
            selectedStewards.push(steward);
        }
        Plotly.restyle(plot, 'marker.color', selectedStewards.map(s => ({
            x: steward,
            marker: { color: 'rgba(255, 0, 0, 0.7)', line: { width: 2, color: 'black' } }
        })));
        showTable();
    });

    </script>
    """)

print("ASPX file with interactive table created successfully.")
