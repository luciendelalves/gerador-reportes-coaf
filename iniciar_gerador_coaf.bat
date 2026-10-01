@echo off
cd /d "%~dp0"

if not exist ".streamlit" mkdir ".streamlit"
if not exist ".streamlit\config.toml" (
    > ".streamlit\config.toml" echo [server]
    >>".streamlit\config.toml" echo address = "127.0.0.1"
    >>".streamlit\config.toml" echo headless = false
    >>".streamlit\config.toml" echo maxUploadSize = 50
    >>".streamlit\config.toml" echo(
    >>".streamlit\config.toml" echo [browser]
    >>".streamlit\config.toml" echo gatherUsageStats = false
    >>".streamlit\config.toml" echo(
    >>".streamlit\config.toml" echo [theme]
    >>".streamlit\config.toml" echo base = "dark"
    >>".streamlit\config.toml" echo primaryColor = "#4C8BF5"
    >>".streamlit\config.toml" echo backgroundColor = "#0E1117"
    >>".streamlit\config.toml" echo secondaryBackgroundColor = "#1C2333"
    >>".streamlit\config.toml" echo textColor = "#F5F7FA"
    >>".streamlit\config.toml" echo font = "sans serif"
)

py -m streamlit run gerador_coaf.py
