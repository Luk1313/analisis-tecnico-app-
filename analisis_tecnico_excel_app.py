import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas_ta as ta

# === Cargar archivo Excel ===
df = pd.read_excel('Datos históricos.xlsx', sheet_name='datos')
df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d.%m.%Y', errors='coerce')
df = df.sort_values('Fecha')

# === Calcular indicadores ===
df['%_Cambio'] = df['Último'].pct_change() * 100
df['MA20'] = df['Último'].rolling(window=20).mean()
df['MA50'] = df['Último'].rolling(window=50).mean()
df['RSI'] = ta.rsi(df['Último'], length=14)
macd = ta.macd(df['Último'])
df = pd.concat([df, macd], axis=1)

# === Crear gráfico con subtramas ===
fig = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.03,
                    row_heights=[0.5, 0.2, 0.3],
                    subplot_titles=("Velas + MA", "% de Cambio", "RSI + MACD"))

fig.add_trace(go.Candlestick(
    x=df['Fecha'], open=df['Apertura'], high=df['Máximo'],
    low=df['Mínimo'], close=df['Último'], name='Velas'), row=1, col=1)

fig.add_trace(go.Scatter(x=df['Fecha'], y=df['MA20'], name='MA20', line=dict(color='blue')), row=1, col=1)
fig.add_trace(go.Scatter(x=df['Fecha'], y=df['MA50'], name='MA50', line=dict(color='orange')), row=1, col=1)

fig.add_trace(go.Bar(
    x=df['Fecha'], y=df['%_Cambio'], name='% de Cambio',
    marker_color=df['%_Cambio'].apply(lambda x: 'green' if x >= 0 else 'red')), row=2, col=1)

fig.add_trace(go.Scatter(x=df['Fecha'], y=df['RSI'], name='RSI', line=dict(color='purple')), row=3, col=1)

fig.add_trace(go.Scatter(x=df['Fecha'], y=df['MACD_12_26_9'], name='MACD', line=dict(color='green')), row=3, col=1)
fig.add_trace(go.Scatter(x=df['Fecha'], y=df['MACDs_12_26_9'], name='Señal', line=dict(color='red')), row=3, col=1)
fig.add_trace(go.Bar(x=df['Fecha'], y=df['MACDh_12_26_9'], name='Histograma MACD',
                     marker_color=df['MACDh_12_26_9'].apply(lambda x: 'green' if x >= 0 else 'red')), row=3, col=1)

fig.update_layout(
    template="plotly_white",
    height=900,
    xaxis_rangeslider_visible=False,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)

fig.show()
