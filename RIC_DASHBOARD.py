import dash
from dash import dcc, html, Input, Output
import pandas as pd
import plotly.graph_objects as go

# Load the Excel file
file_path = 'C:/Users/Win 10/Desktop/filtered_data.xlsx'
data = pd.ExcelFile(file_path)

# Extract the list of RICs (first sheet)
ric_list = data.parse(data.sheet_names[0]).iloc[:, 0].str.replace('.jk', '', regex=False).tolist()

# Function to load data for a specific RIC
def load_ric_data(ric):
    # Loại bỏ đuôi '.jk' khỏi RIC để tìm đúng tên sheet
    sheet_name = ric.split(".")[0]
    debug_message = ""

    if sheet_name in data.sheet_names:
        df = data.parse(sheet_name=sheet_name)
        df.columns = df.columns.str.strip()  # Loại bỏ khoảng trắng trong tên cột

        # Kiểm tra và xử lý cột Date
        if 'Date' in df.columns:
            try:
                df['Date'] = pd.to_datetime(df['Date'], format="%m/%d/%Y %I:%M:%S %p", errors='coerce')
            except Exception as e:
                debug_message += f"Error parsing Date column in {ric}: {e}\n"
            df = df.dropna(subset=['Date'])  # Loại bỏ các dòng có giá trị Date không hợp lệ

        # Đảm bảo các cột số liệu được chuyển đổi chính xác
        numeric_columns = ['Price Open', 'Price High', 'Price Low', 'Price Close', 'Volume']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')  # Chuyển đổi thành số
                df = df.dropna(subset=[col])  # Loại bỏ các dòng có dữ liệu không hợp lệ trong cột số

        debug_message += f"Loaded sheet for {ric}. Columns: {list(df.columns)}\n"
        return df, debug_message
    else:
        debug_message += f"RIC '{ric}' not found. Expected sheet name: {sheet_name}. Available sheets: {data.sheet_names}\n"
        return pd.DataFrame(), debug_message

# Initialize the Dash app
app = dash.Dash(__name__)

# Layout of the app
app.layout = html.Div([
    html.Div([
        html.Label('Select RIC (Company):'),
        dcc.Dropdown(
            id='ric-dropdown',
            options=[{'label': ric, 'value': ric} for ric in ric_list],
            value=ric_list[0],  # Default to the first RIC
            multi=False
        )
    ]),
    html.Div([
        html.Label('Select Price Type:'),
        dcc.Checklist(
            id='price-type-checklist',
            options=[
                {'label': 'Price Open', 'value': 'Price Open'},
                {'label': 'Price High', 'value': 'Price High'},
                {'label': 'Price Low', 'value': 'Price Low'},
                {'label': 'Price Close', 'value': 'Price Close'},
                {'label': 'Volume', 'value': 'Volume'}
            ],
            value=['Price Close'],  # Default to Price Close
            inline=True
        )
    ]),
    dcc.Graph(id='stock-price-graph'),
    html.Div(id='debug-output', style={'whiteSpace': 'pre-line', 'marginTop': '20px'})
])

# Callback to update the graph based on user selection
@app.callback(
    [
        Output('stock-price-graph', 'figure'),
        Output('debug-output', 'children')
    ],
    [
        Input('ric-dropdown', 'value'),
        Input('price-type-checklist', 'value')
    ]
)
def update_graph(selected_ric, selected_price_types):
    # Load data for the selected RIC
    ric_data, debug_message = load_ric_data(selected_ric)

    if ric_data.empty:
        # Trả về thông báo lỗi nếu không có dữ liệu
        return go.Figure(layout={
            'title': f'No data available for {selected_ric}',
            'xaxis_title': 'Date',
            'yaxis_title': 'Price'
        }), debug_message + "No data available."

    # Tạo biểu đồ
    fig = go.Figure()
    for price_type in selected_price_types:
        if price_type in ric_data.columns:
            fig.add_trace(go.Scatter(
                x=ric_data['Date'],
                y=ric_data[price_type],
                mode='lines',
                name=price_type
            ))
        else:
            debug_message += f"Column '{price_type}' not found in {selected_ric}.\n"

    # Cập nhật layout
    fig.update_layout(
        title=f'Stock Prices for {selected_ric}',
        xaxis_title='Date',
        yaxis_title='Price',
        template='plotly_white'
    )
    return fig, debug_message

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
