# ANBIMA Debenture Price Scraper

Python script that reads a list of debentures from an Excel file, looks each one up on [ANBIMA Data](https://data.anbima.com.br) and writes back:

- unit price on the curve (PU curva) and the latest event
- ANBIMA indicative market price (PU de mercado)
- indicative market rate

Built with Selenium and openpyxl.

## How to run

1. Install Selenium and openpyxl, plus a ChromeDriver matching your Chrome version.
2. List the debenture codes in column A of an Excel file.
3. Update the ChromeDriver and Excel paths at the top of `ativos.py` and run it.
