#!/bin/bash

$GREEN = "Green"
$YELLOW = "Yellow"
$NC = "White"

function passKey {
    Write-Host "🏗️ Creating passKey" -ForegroundColor $YELLOW
    # password
    Set-Content -Path "models/passKey.py" -Value @"
import msoffcrypto
import openpyxl
import sys
import os

def remove_excel_password(input_file, output_file=None):
    if output_file is None:
        output_file = input_file  # Overwrite the input file if no output file is specified

    while True:  # Keep looping until the password is correct or the user gives up
        # Prompt for the password
        password = input(f"Enter password for '{input_file}' (or type 'quit' to exit): ")

        if password.lower() == 'quit':
            print("Exiting without removing password.")
            return False

        try:
            # Decrypt the file using msoffcrypto-tool
            decrypted_file = output_file
            with open(input_file, "rb") as file:
                office_file = msoffcrypto.OfficeFile(file)
                office_file.load_key(password=password)  # Load the password
                with open(decrypted_file, "wb") as decrypted:
                    office_file.decrypt(decrypted)  # Decrypt and save the file

            print(f"Password removed successfully. File saved to '{output_file}'.")
            return True  # Password was correct, exit the function
        except msoffcrypto.exceptions.InvalidKeyError:
            print("Incorrect password. Please try again.")
            # The loop continues to the next iteration, prompting for the password again
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            return False  # Some other error occurred, exit the function

def add_fk_id_estado(input_file, output_file):
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(input_file)
        ws = wb.active

        # Find or create the "fkIdEstado" column
        header_row = 1
        fk_id_col = None

        # Search for existing column
        for col in range(1, ws.max_column + 1):
            if ws.cell(row=header_row, column=col).value == "fkIdEstado":
                fk_id_col = col
                break

        # If column doesn't exist, add it at the end
        if fk_id_col is None:
            fk_id_col = ws.max_column + 1
            ws.cell(row=header_row, column=fk_id_col, value="fkIdEstado")

        # Fill the column with value 1 (starting from row 2)
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=fk_id_col, value=1)

        # Save the modified workbook
        wb.save(output_file)
        print(f"File with 'fkIdEstado' column added saved to '{output_file}'.")
        return True
    except Exception as e:
        print(f"Error while adding 'fkIdEstado' column: {e}")
        return False

if __name__ == "__main__":
    # Prompt user to place the file in src/ directory
    input("Pega el archivo de Excel file en la carpeta 'src/' y asegura de nombrarlo 'dataHistoricaPBI.xlsx'. y presiona Enter cuando este listo...")

    # Check if file exists
    input_excel_file = "src/dataHistoricaPBI.xlsx"
    if not os.path.exists(input_excel_file):
        print(f"Error: File '{input_excel_file}' not found. Please make sure:")
        print("1. The 'src/' directory exists")
        print("2. The file is placed in the 'src/' directory")
        print("3. The file is named 'dataHistoricaPBI.xlsx'")
        sys.exit(1)

    # First output file (decrypted)
    output_excel_file = "src/data.xlsx"

    # Second output file with fkIdEstado column
    output_with_fk_file = "src/fk1Data.xlsx"

    if remove_excel_password(input_excel_file, output_excel_file):
        # If decryption succeeded, create the modified file
        if add_fk_id_estado(output_excel_file, output_with_fk_file):
            print("Both files created successfully:")
            print(f"- Decrypted file: {output_excel_file}")
            print(f"- File with fkIdEstado column: {output_with_fk_file}")
        else:
            print("Process partially completed - decrypted file created but failed to add fkIdEstado column.")
    else:
        print("Process failed - could not decrypt the input file.")
"@

Write-Host "🏗️ Creating FKID01 Filter" -ForegroundColor $YELLOW
    # toJson
    Set-Content -Path "models/toJson.py" -Value @"
import pandas as pd

def excel_to_json(excel_file, json_file):

    try:
        # Read the Excel file into a Pandas DataFrame
        df = pd.read_excel(excel_file)

        # Convert the DataFrame to a JSON string
        json_data = df.to_json(orient='records', indent=4)  

        # Write the JSON string to a file
        with open(json_file, 'w') as f:
            f.write(json_data)

        print(f"Successfully converted '{excel_file}' to '{json_file}'")

    except FileNotFoundError:
        print(f"Error: Excel file '{excel_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage:
if __name__ == "__main__":
    excel_file_path = "src/fk1Data.xlsx"  
    json_file_path = "tables/trends/data2.json"  

    excel_to_json(excel_file_path, json_file_path)
"@

}

function createCats {
    Write-Host "🏗️ Creating Categories" -ForegroundColor $YELLOW
    
    # Banks
    Set-Content -Path "models/cats.py" -Value @"
import pandas as pd
from datetime import datetime

# Shared constants and functions
TRM_DICT = {
    2020: 3432.50,
    2021: 3981.16,
    2022: 4810.20,
    2023: 4780.38,
    2024: 4409.00
}

CURRENCY_RATES = {
    2020: {
        'EUR': 1.141, 'GBP': 1.280, 'AUD': 0.690, 'CAD': 0.746,
        'HNL': 0.0406, 'AWG': 0.558, 'DOP': 0.0172, 'PAB': 1.000,
        'CLP': 0.00126, 'CRC': 0.00163, 'ARS': 0.0119, 'ANG': 0.558,
        'COP': 0.00026,  'BBD': 0.50, 'MXN': 0.0477, 'BOB': 0.144, 'BSD': 1.00,
        'GYD': 0.0048, 'UYU': 0.025, 'DKK': 0.146, 'KYD': 1.20, 'BMD': 1.00, 
        'VEB': 0.0000000248, 'VES': 0.000000248, 'BRL': 0.187, 'NIO': 0.0278
    },
    2021: {
        'EUR': 1.183, 'GBP': 1.376, 'AUD': 0.727, 'CAD': 0.797,
        'HNL': 0.0415, 'AWG': 0.558, 'DOP': 0.0176, 'PAB': 1.000,
        'CLP': 0.00118, 'CRC': 0.00156, 'ARS': 0.00973, 'ANG': 0.558,
        'COP': 0.00027, 'BBD': 0.50, 'MXN': 0.0492, 'BOB': 0.141, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.024, 'DKK': 0.155, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0.00000000002, 'VES': 0.00000002, 'BRL': 0.192, 'NIO': 0.0285
    },
    2022: {
        'EUR': 1.051, 'GBP': 1.209, 'AUD': 0.688, 'CAD': 0.764,
        'HNL': 0.0408, 'AWG': 0.558, 'DOP': 0.0181, 'PAB': 1.000,
        'CLP': 0.00117, 'CRC': 0.00155, 'ARS': 0.00597, 'ANG': 0.558,
        'COP': 0.00021, 'BBD': 0.50, 'MXN': 0.0497, 'BOB': 0.141, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.025, 'DKK': 0.141, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.196, 'NIO': 0.0267
    },
    2023: {
        'EUR': 1.096, 'GBP': 1.264, 'AUD': 0.676, 'CAD': 0.741,
        'HNL': 0.0406, 'AWG': 0.558, 'DOP': 0.0177, 'PAB': 1.000,
        'CLP': 0.00121, 'CRC': 0.00187, 'ARS': 0.00275, 'ANG': 0.558,
        'COP': 0.00022, 'BBD': 0.50, 'MXN': 0.0564, 'BOB': 0.143, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.025, 'DKK': 0.148, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.194, 'NIO': 0.0267
    },
    2024: {
        'EUR': 1.093, 'GBP': 1.267, 'AUD': 0.674, 'CAD': 0.742,
        'HNL': 0.0405, 'AWG': 0.558, 'DOP': 0.0170, 'PAB': 1.000,
        'CLP': 0.00111, 'CRC': 0.00192, 'ARS': 0.00121, 'ANG': 0.558,
        'COP': 0.00022, 'BBD': 0.50, 'MXN': 0.0547, 'BOB': 0.142, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.024, 'DKK': 0.147, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.190, 'NIO': 0.0260 }
}

def get_trm(year):
    """Gets TRM for a given year from the dictionary"""
    return TRM_DICT.get(year)

def get_exchange_rate(currency_code, year):
    """Gets exchange rate for a given currency and year from the dictionary"""
    year_rates = CURRENCY_RATES.get(year)
    if year_rates:
        return year_rates.get(currency_code)
    return None

def get_currency_code(moneda_text):
    """Extracts the currency code from the 'Texto Moneda' field"""
    currency_mapping = {
        'HNL -Lempira hondureño': 'HNL',
        'EUR - Euro': 'EUR',
        'AWG - Florín holandés o de Aruba': 'AWG',
        'DOP - Peso dominicano': 'DOP',
        'PAB -Balboa panameña': 'PAB', 
        'CLP - Peso chileno': 'CLP',
        'CRC - Colón costarricense': 'CRC',
        'ARS - Peso argentino': 'ARS',
        'AUD - Dólar australiano': 'AUD',
        'ANG - Florín holandés': 'ANG',
        'CAD -Dólar canadiense': 'CAD',
        'GBP - Libra esterlina': 'GBP',
        'USD - Dolar estadounidense': 'USD',
        'COP - Peso colombiano': 'COP',
        'BBD - Dólar de Barbados o Baja': 'BBD',
        'MXN - Peso mexicano': 'MXN',
        'BOB - Boliviano': 'BOB',
        'BSD - Dolar bahameño': 'BSD',
        'GYD - Dólar guyanés': 'GYD',
        'UYU - Peso uruguayo': 'UYU',
        'DKK - Corona danesa': 'DKK',
        'KYD - Dólar de las Caimanes': 'KYD',
        'BMD - Dólar de las Bermudas': 'BMD',
        'VEB - Bolívar venezolano': 'VEB',  
        'VES - Bolívar soberano': 'VES',  
        'BRL - Real brasilero': 'BRL',  
        'NIO - Córdoba nicaragüense': 'NIO',
    }
    return currency_mapping.get(moneda_text)

def get_valid_year(row, periodo_df):
    """Extracts a valid year, handling missing values and format variations."""
    try:
        fkIdPeriodo = pd.to_numeric(row['fkIdPeriodo'], errors='coerce')
        if pd.isna(fkIdPeriodo):  # Handle missing fkIdPeriodo
            print(f"Warning: Missing fkIdPeriodo at index {row.name}. Skipping row.")
            return None

        matching_row = periodo_df[periodo_df['Id'] == fkIdPeriodo]
        if matching_row.empty:
            print(f"Warning: No matching Id found in periodoBR.xlsx for fkIdPeriodo {fkIdPeriodo} at index {row.name}. Skipping row.")
            return None

        year_str = matching_row['Año'].iloc[0]

        try:
            year = int(year_str)  # Try direct conversion to integer
            return year
        except (ValueError, TypeError):
            try:
                year = pd.to_datetime(year_str, errors='coerce').year  # Try datetime conversion, handle errors gracefully
                if pd.isna(year):  # check for NaT which occurs when conversion fails.
                    raise ValueError  # If conversion failed re-raise a ValueError.
                return year

            except ValueError:
                print(f"Warning: Invalid year format '{year_str}' for fkIdPeriodo {fkIdPeriodo} at index {row.name}. Skipping row.")
                return None

    except Exception as e:
        print(f"Error in get_valid_year for fkIdPeriodo {fkIdPeriodo} at index {row.name}: {e}")
        return None

def analyze_banks(file_path, output_file_path, periodo_file_path):
    """Analyze bank account data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario',
        'Nombre', 'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Banco - Entidad', 'Banco - Tipo Cuenta', 'Texto Moneda',
        'Banco - fkIdPaís', 'Banco - Nombre País',
        'Banco - Saldo', 'Banco - Comentario'
    ]
    
    banks_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Banco', maintain_columns].copy()
    banks_df = banks_df[banks_df['fkIdEstado'] != 1]
    
    banks_df['Banco - Saldo COP'] = 0.0
    banks_df['TRM Aplicada'] = None
    banks_df['Tasa USD'] = None
    banks_df['Año Declaración'] = None 
    
    for index, row in banks_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index} and fkIdPeriodo {row['fkIdPeriodo']}. Skipping row.")
                banks_df.loc[index, 'Año Declaración'] = "Año no encontrado"
                continue 
                
            banks_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                banks_df.loc[index, 'Banco - Saldo COP'] = float(row['Banco - Saldo'])
                banks_df.loc[index, 'TRM Aplicada'] = 1.0
                banks_df.loc[index, 'Tasa USD'] = None
                continue
                
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Banco - Saldo']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    banks_df.loc[index, 'Banco - Saldo COP'] = cop_amount
                    banks_df.loc[index, 'TRM Aplicada'] = trm
                    banks_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            banks_df.loc[index, 'Año Declaración'] = "Error de procesamiento"
            continue
    
    banks_df.to_excel(output_file_path, index=False)

def analyze_debts(file_path, output_file_path, periodo_file_path):
    """Analyze debts data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Pasivos - Entidad Personas',
        'Pasivos - Tipo Obligación', 'fkIdMoneda', 'Texto Moneda',
        'Pasivos - Valor', 'Pasivos - Comentario', 'Pasivos - Valor COP'
    ]
    
    debts_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Pasivo', maintain_columns].copy()
    debts_df = debts_df[debts_df['fkIdEstado'] != 1]
    
    debts_df['Pasivos - Valor COP'] = 0.0
    debts_df['TRM Aplicada'] = None
    debts_df['Tasa USD'] = None
    debts_df['Año Declaración'] = None 
    
    for index, row in debts_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            debts_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                debts_df.loc[index, 'Pasivos - Valor COP'] = float(row['Pasivos - Valor'])
                debts_df.loc[index, 'TRM Aplicada'] = 1.0
                debts_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Pasivos - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    debts_df.loc[index, 'Pasivos - Valor COP'] = cop_amount
                    debts_df.loc[index, 'TRM Aplicada'] = trm
                    debts_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            debts_df.loc[index, 'Año Declaración'] = "Error de procesamiento"
            continue

    debts_df.to_excel(output_file_path, index=False)

def analyze_goods(file_path, output_file_path, periodo_file_path):
    """Analyze goods/patrimony data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Patrimonio - Activo', 'Patrimonio - % Propiedad',
        'Patrimonio - Propietario', 'Patrimonio - Valor Comercial',
        'Patrimonio - Comentario',
        'Patrimonio - Valor Comercial COP', 'Texto Moneda'
    ]
    
    goods_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Patrimonio', maintain_columns].copy()
    goods_df = goods_df[goods_df['fkIdEstado'] != 1]
    
    goods_df['Patrimonio - Valor COP'] = 0.0
    goods_df['TRM Aplicada'] = None
    goods_df['Tasa USD'] = None
    goods_df['Año Declaración'] = None 
    
    for index, row in goods_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
                
            goods_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                goods_df.loc[index, 'Patrimonio - Valor COP'] = float(row['Patrimonio - Valor Comercial'])
                goods_df.loc[index, 'TRM Aplicada'] = 1.0
                goods_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Patrimonio - Valor Comercial']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    goods_df.loc[index, 'Patrimonio - Valor COP'] = cop_amount
                    goods_df.loc[index, 'TRM Aplicada'] = trm
                    goods_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
        
    goods_df['Patrimonio - Valor Corregido'] = goods_df['Patrimonio - Valor COP'] * (goods_df['Patrimonio - % Propiedad'] / 100)
    
    # Rename columns for consistency
    rename_dict = {
        'Patrimonio - Valor Corregido': 'Bienes - Valor Corregido',
        'Patrimonio - Valor Comercial COP': 'Bienes - Valor Comercial COP',
        'Patrimonio - Comentario': 'Bienes - Comentario',
        'Patrimonio - Valor Comercial': 'Bienes - Valor Comercial',
        'Patrimonio - Propietario': 'Bienes - Propietario',
        'Patrimonio - % Propiedad': 'Bienes - % Propiedad',
        'Patrimonio - Activo': 'Bienes - Activo',
        'Patrimonio - Valor COP': 'Bienes - Valor COP'
    }
    goods_df = goods_df.rename(columns=rename_dict)
    
    goods_df.to_excel(output_file_path, index=False)

def analyze_incomes(file_path, output_file_path, periodo_file_path):
    """Analyze income data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Ingresos - fkIdConcepto', 'Ingresos - Texto Concepto',
        'Ingresos - Valor', 'Ingresos - Comentario', 'Ingresos - Otros',
        'Ingresos - Valor_COP', 'Texto Moneda'
    ]

    incomes_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Ingreso', maintain_columns].copy()
    incomes_df = incomes_df[incomes_df['fkIdEstado'] != 1]
    
    incomes_df['Ingresos - Valor COP'] = 0.0
    incomes_df['TRM Aplicada'] = None
    incomes_df['Tasa USD'] = None
    incomes_df['Año Declaración'] = None 
    
    for index, row in incomes_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            incomes_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                incomes_df.loc[index, 'Ingresos - Valor COP'] = float(row['Ingresos - Valor'])
                incomes_df.loc[index, 'TRM Aplicada'] = 1.0
                incomes_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Ingresos - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    incomes_df.loc[index, 'Ingresos - Valor COP'] = cop_amount
                    incomes_df.loc[index, 'TRM Aplicada'] = trm
                    incomes_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
    
    incomes_df.to_excel(output_file_path, index=False)

def analyze_investments(file_path, output_file_path, periodo_file_path):
    """Analyze investment data"""
    df = pd.read_excel(file_path)
    periodo_df = pd.read_excel(periodo_file_path)

    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Inversiones - Tipo Inversión', 'Inversiones - Entidad',
        'Inversiones - Valor', 'Inversiones - Comentario',
        'Inversiones - Valor COP', 'Texto Moneda'
    ]
    
    invest_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Inversión', maintain_columns].copy()
    invest_df = invest_df[invest_df['fkIdEstado'] != 1]
    
    invest_df['Inversiones - Valor COP'] = 0.0
    invest_df['TRM Aplicada'] = None
    invest_df['Tasa USD'] = None
    invest_df['Año Declaración'] = None 
    
    for index, row in invest_df.iterrows():
        try:
            year = get_valid_year(row, periodo_df)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            invest_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                invest_df.loc[index, 'Inversiones - Valor COP'] = float(row['Inversiones - Valor'])
                invest_df.loc[index, 'TRM Aplicada'] = 1.0
                invest_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Inversiones - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    invest_df.loc[index, 'Inversiones - Valor COP'] = cop_amount
                    invest_df.loc[index, 'TRM Aplicada'] = trm
                    invest_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
    
    invest_df.to_excel(output_file_path, index=False)

def run_all_analyses():
    """Run all analysis functions with their respective file paths"""
    file_path = 'src/data.xlsx'
    periodo_file_path = 'src/periodoBR.xlsx'
    
    analyze_banks(file_path, 'tables/cats/banks.xlsx', periodo_file_path)
    analyze_debts(file_path, 'tables/cats/debts.xlsx', periodo_file_path)
    analyze_goods(file_path, 'tables/cats/goods.xlsx', periodo_file_path)
    analyze_incomes(file_path, 'tables/cats/incomes.xlsx', periodo_file_path)
    analyze_investments(file_path, 'tables/cats/investments.xlsx', periodo_file_path)

if __name__ == "__main__":
    run_all_analyses()
"@
}

function createNets {
    Write-Host "🏗️ Creating Nets" -ForegroundColor $YELLOW
    # Banks
    Set-Content -Path "models/nets.py" -Value @"
import pandas as pd

# Common columns used across all analyses
COMMON_COLUMNS = [
    'Usuario', 'Nombre', 'Compañía', 'Cargo',
    'fkIdPeriodo', 'fkIdEstado',
    'Año Creación', 'Año Envío',
    'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
    'Año Declaración'
]

# Base groupby columns for summaries
BASE_GROUPBY = ['Usuario', 'Nombre', 'Compañía', 'Cargo', 'fkIdPeriodo', 'Año Declaración', 'Año Creación']

def analyze_banks(file_path, output_file_path):
    """Analyze bank accounts data"""
    df = pd.read_excel(file_path)

    # Specific columns for banks
    bank_columns = [
        'Banco - Entidad', 'Banco - Tipo Cuenta',
        'Banco - fkIdPaís', 'Banco - Nombre País',
        'Banco - Saldo', 'Banco - Comentario',
        'Banco - Saldo COP'
    ]
    
    df = df[COMMON_COLUMNS + bank_columns]
    
    # Create a temporary combination column for counting
    df_temp = df.copy()
    df_temp['Bank_Account_Combo'] = df['Banco - Entidad'] + "|" + df['Banco - Tipo Cuenta']
    
    # Perform all aggregations
    summary = df_temp.groupby(BASE_GROUPBY).agg(
        **{
            'Cant_Bancos': pd.NamedAgg(column='Banco - Entidad', aggfunc='nunique'),
            'Cant_Cuentas': pd.NamedAgg(column='Bank_Account_Combo', aggfunc='nunique'),
            'Banco - Saldo COP': pd.NamedAgg(column='Banco - Saldo COP', aggfunc='sum')
        }
    ).reset_index()

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_debts(file_path, output_file_path):
    """Analyze debts data"""
    df = pd.read_excel(file_path)

    # Specific columns for debts
    debt_columns = [
        'Pasivos - Entidad Personas', 'Pasivos - Tipo Obligación', 
        'Pasivos - Valor', 'Pasivos - Comentario',
        'Pasivos - Valor COP', 'Texto Moneda'
    ]
    
    df = df[COMMON_COLUMNS + debt_columns]
    
    # Calculate total Pasivos and count occurrences
    summary = df.groupby(BASE_GROUPBY).agg({      
        'Pasivos - Valor COP': 'sum',
        'Pasivos - Entidad Personas': 'count'
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Pasivos - Entidad Personas': 'Cant_Deudas',
        'Pasivos - Valor COP': 'Total Pasivos'
    })

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_goods(file_path, output_file_path):
    """Analyze goods/assets data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for goods
    goods_columns = [
        'Bienes - Activo', 'Bienes - % Propiedad',
        'Bienes - Propietario', 'Bienes - Valor Comercial',
        'Bienes - Comentario', 'Bienes - Valor Comercial COP',
        'Bienes - Valor Corregido'
    ]
    
    df = df[COMMON_COLUMNS + goods_columns]

    summary = df.groupby(BASE_GROUPBY).agg({
        'Bienes - Valor Corregido': 'sum',
        'Bienes - Activo': 'count' 
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Bienes - Activo': 'Cant_Bienes',
        'Bienes - Valor Corregido': 'Total Bienes'
    })

    summary.to_excel(output_file_path, index=False) 
    return summary

def analyze_incomes(file_path, output_file_path):
    """Analyze income data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for incomes
    income_columns = [
        'Ingresos - fkIdConcepto', 'Ingresos - Texto Concepto',
        'Ingresos - Valor', 'Ingresos - Comentario',
        'Ingresos - Otros', 'Ingresos - Valor COP',
        'Texto Moneda'
    ]

    df = df[COMMON_COLUMNS + income_columns]
    
    # Calculate Ingresos and count occurrences
    summary = df.groupby(BASE_GROUPBY).agg({
        'Ingresos - Valor COP': 'sum',
        'Ingresos - Texto Concepto': 'count'
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Ingresos - Texto Concepto': 'Cant_Ingresos',
        'Ingresos - Valor COP': 'Total Ingresos'
    })

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_investments(file_path, output_file_path):
    """Analyze investments data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for investments
    invest_columns = [
        'Inversiones - Tipo Inversión', 'Inversiones - Entidad',
        'Inversiones - Valor', 'Inversiones - Comentario',
        'Inversiones - Valor COP', 'Texto Moneda'
    ]
    
    df = df[COMMON_COLUMNS + invest_columns]
    
    # Calculate total Inversiones and count occurrences
    summary = df.groupby(BASE_GROUPBY + ['Inversiones - Tipo Inversión']).agg( 
        {'Inversiones - Valor COP': 'sum',
         'Inversiones - Tipo Inversión': 'count'}
    ).rename(columns={
        'Inversiones - Tipo Inversión': 'Cant_Inversiones',
        'Inversiones - Valor COP': 'Total Inversiones'
    }).reset_index()
    
    summary.to_excel(output_file_path, index=False)
    return summary 

def calculate_assets(banks_file, goods_file, invests_file, output_file):
    """Calculate total assets by combining banks, goods and investments"""
    banks = pd.read_excel(banks_file)
    goods = pd.read_excel(goods_file)
    invests = pd.read_excel(invests_file)

    # Group investments by base columns (summing across types)
    invests_grouped = invests.groupby(BASE_GROUPBY).agg({
        'Total Inversiones': 'sum',
        'Cant_Inversiones': 'sum'
    }).reset_index()

    # Merge all three dataframes
    merged = pd.merge(goods, banks, on=BASE_GROUPBY, how='outer')
    merged = pd.merge(merged, invests_grouped, on=BASE_GROUPBY, how='outer')
    merged.fillna(0, inplace=True)

    # Calculate total assets
    merged['Total Activos'] = (
        merged['Total Bienes'] + 
        merged['Banco - Saldo COP'] + 
        merged['Total Inversiones']
    )

    # Reorder and rename columns
    final_columns = BASE_GROUPBY + [
        'Total Bienes', 'Cant_Bienes',
        'Banco - Saldo COP', 'Cant_Bancos', 'Cant_Cuentas',
        'Total Inversiones', 'Cant_Inversiones',
        'Total Activos'
    ]
    merged = merged[final_columns]

    merged.to_excel(output_file, index=False)
    return merged

def calculate_net_worth(debts_file, assets_file, output_file):
    """Calculate net worth by combining assets and debts"""
    debts = pd.read_excel(debts_file)
    assets = pd.read_excel(assets_file)

    # Merge the summaries
    merged = pd.merge(
        assets, 
        debts, 
        on=BASE_GROUPBY, 
        how='outer'
    )
    merged.fillna(0, inplace=True)
    
    # Calculate net worth
    merged['Total Patrimonio'] = merged['Total Activos'] - merged['Total Pasivos']
    
    # Final column order
    final_columns = BASE_GROUPBY + [
        'Total Activos',
        'Cant_Bienes',
        'Cant_Bancos',
        'Cant_Cuentas',
        'Cant_Inversiones',
        'Total Pasivos',
        'Cant_Deudas',
        'Total Patrimonio'
    ]
    merged = merged[final_columns]
    
    merged.to_excel(output_file, index=False)
    return merged

def run_all_analyses():
    """Run all analyses in sequence with default file paths"""
    # Individual analyses
    bank_summary = analyze_banks(
        'tables/cats/banks.xlsx',
        'tables/nets/bankNets.xlsx'
    )
    
    debt_summary = analyze_debts(
        'tables/cats/debts.xlsx',
        'tables/nets/debtNets.xlsx'
    )
    
    goods_summary = analyze_goods(
        'tables/cats/goods.xlsx',
        'tables/nets/goodNets.xlsx'
    )
    
    income_summary = analyze_incomes(
        'tables/cats/incomes.xlsx',
        'tables/nets/incomeNets.xlsx'
    )
    
    invest_summary = analyze_investments(
        'tables/cats/investments.xlsx',
        'tables/nets/investNets.xlsx'
    )
    
    # Combined analyses
    assets_summary = calculate_assets(
        'tables/nets/bankNets.xlsx',
        'tables/nets/goodNets.xlsx',
        'tables/nets/investNets.xlsx',
        'tables/nets/assetNets.xlsx'
    )
    
    net_worth_summary = calculate_net_worth(
        'tables/nets/debtNets.xlsx',
        'tables/nets/assetNets.xlsx',
        'tables/nets/worthNets.xlsx'
    )
    
    return {
        'bank_summary': bank_summary,
        'debt_summary': debt_summary,
        'goods_summary': goods_summary,
        'income_summary': income_summary,
        'invest_summary': invest_summary,
        'assets_summary': assets_summary,
        'net_worth_summary': net_worth_summary
    }

if __name__ == '__main__':
    # Run all analyses when script is executed
    results = run_all_analyses()
    print("All nets analyses completed successfully!")
"@
}

function createTrends {
    Write-Host "🏗️ Creating Trends" -ForegroundColor $YELLOW

# trends
   
Set-Content -Path "models/trends/trends.py" -Value @"
import pandas as pd

def get_trend_symbol(value):
    
    try:
        value_float = float(value.strip('%')) / 100
        if pd.isna(value_float):
            return "➡️"  # neutral for first entry
        elif value_float > 0.1:  # more than 10% increase
            return "📈"
        elif value_float < -0.1:  # more than 10% decrease
            return "📉"
        else:
            return "➡️"  # relatively stable
    except Exception as e:
        return "➡️"  # handle any conversion errors

def calculate_variation(df, column):

    # Sort by Usuario and Año Declaración
    df = df.sort_values(by=['Usuario', 'Año Declaración'])
    
    # Calculate absolute and relative variations
    absolute_col = f'{column} Var. Abs.'
    relative_col = f'{column} Var. Rel.'
    
    df[absolute_col] = df.groupby('Usuario')[column].diff()
    
    # Fill missing values with forward fill and calculate percentage change
    df[relative_col] = (
        df.groupby('Usuario')[column]
        .ffill()  # Forward fill missing values
        .pct_change(fill_method=None) * 100  # Explicitly set fill_method=None
    )
    
    # Format relative variation as a percentage string
    df[relative_col] = df[relative_col].apply(lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%")
    
    return df

def embed_trend_symbols(df, columns):

    for col in columns:
        absolute_col = f'{col} Var. Abs.'
        relative_col = f'{col} Var. Rel.'
        
        # Add trend symbols to the variation values
        df[absolute_col] = df.apply(
            lambda row: f"{row[absolute_col]} {get_trend_symbol(row[relative_col])}", axis=1
        )
        df[relative_col] = df.apply(
            lambda row: f"{row[relative_col]} {get_trend_symbol(row[relative_col])}", axis=1
        )
    
    return df

def calculate_leverage(df):

    df['Apalancamiento'] = (df['Patrimonio'] / df['Activos']) * 100
    return df

def calculate_debt_level(df):

    df['Endeudamiento'] = (df['Pasivos'] / df['Activos']) * 100
    return df

def process_asset_data(df_assets):

    # Group by Usuario and Año Declaración to calculate total bank balance, number of banks, total assets, total investments, and number of investments
    df_assets_grouped = df_assets.groupby(['Usuario', 'Año Declaración']).agg(
        BancoSaldo=('Banco - Saldo COP', 'sum'),
        Bienes=('Total Bienes', 'sum'),
        Inversiones=('Total Inversiones', 'sum')
    ).reset_index()

    # Calculate variations for BancoSaldo
    df_assets_grouped = calculate_variation(df_assets_grouped, 'BancoSaldo')

    # Calculate variations for Bienes
    df_assets_grouped = calculate_variation(df_assets_grouped, 'Bienes')

    # Calculate variations for Inversiones
    df_assets_grouped = calculate_variation(df_assets_grouped, 'Inversiones')

    # Embed trend symbols into variation values
    df_assets_grouped = embed_trend_symbols(df_assets_grouped, ['BancoSaldo', 'Bienes', 'Inversiones'])

    return df_assets_grouped

def process_income_data(df_income):

    # Group by Usuario and Año Declaración to calculate total income and number of incomes
    df_income_grouped = df_income.groupby(['Usuario', 'Año Declaración']).agg(
        Ingresos=('Total Ingresos', 'sum'),
        Cant_Ingresos=('Cant_Ingresos', 'sum')
    ).reset_index()

    # Calculate variations for Ingresos
    df_income_grouped = calculate_variation(df_income_grouped, 'Ingresos')

    # Embed trend symbols into variation values
    df_income_grouped = embed_trend_symbols(df_income_grouped, ['Ingresos'])

    return df_income_grouped

def save_results(df, excel_filename="tables/trends/trends.xlsx"):

    try:
        # Save as Excel
        df.to_excel(excel_filename, index=False)
        print(f"Data saved to {excel_filename}")
    except Exception as e:
        print(f"Error saving file: {e}")

def main():
    try:
        # Load the worth Excel file
        df = pd.read_excel("tables/nets/worthNets.xlsx")
        print("Worth file loaded successfully.")
        
        # Rename columns
        df = df.rename(columns={
            'Total Activos': 'Activos',
            'Total Pasivos': 'Pasivos',
            'Total Patrimonio': 'Patrimonio'
        })
        
        # Calculate Leverage
        df = calculate_leverage(df)
        
        # Calculate Debt Level
        df = calculate_debt_level(df)
        
        # Columns to calculate variations for
        columns_to_analyze = ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento']
        
        # Calculate variations for each column
        for column in columns_to_analyze:
            df = calculate_variation(df, column)
        
        # Embed trend symbols into variation values
        df = embed_trend_symbols(df, columns_to_analyze)
        
        # Load the asset data
        df_assets = pd.read_excel("tables/nets/assetNets.xlsx")
        print("Asset file loaded successfully.")
        
        # Process the asset data
        df_assets_processed = process_asset_data(df_assets)
        
        # Load the income data
        df_income = pd.read_excel("tables/nets/incomeNets.xlsx")
        print("Income file loaded successfully.")
        
        # Process the income data
        df_income_processed = process_income_data(df_income)
        
        # Merge the processed asset and income data with the worth DataFrame
        df = pd.merge(df, df_assets_processed, on=['Usuario', 'Año Declaración'], how='left')
        df = pd.merge(df, df_income_processed, on=['Usuario', 'Año Declaración'], how='left')
        
        # Save results
        save_results(df)
    except FileNotFoundError:
        print("Error: One of the required files was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
"@

# overTrends
   
Set-Content -Path "models/trends/overTrends.py" -Value @"
import pandas as pd

def get_trend_symbol(value):
    """Determine the trend symbol based on the percentage change."""
    try:
        value_float = float(value.strip('%')) / 100
        if pd.isna(value_float):
            return "➡️"
        elif value_float > 0.1:
            return "📈"
        elif value_float < -0.1:
            return "📉"
        else:
            return "➡️"
    except Exception:
        return "➡️"

def calculate_yearly_variations(df):
    """Calculate absolute and relative variations for all specified columns."""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    columns_to_analyze = [
        'Activos', 'Pasivos', 'Patrimonio', 
        'Apalancamiento', 'Endeudamiento',
        'BancoSaldo', 'Bienes', 'Inversiones', 'Ingresos',
        'Cant_Deudas', 'Cant_Bancos', 'Cant_Cuentas', 'Cant_Bienes', 
        'Cant_Inversiones', 'Cant_Ingresos'
    ]
    
    # Precompute all variations first (avoids fragmentation)
    new_columns = {}
    
    for column in columns_to_analyze:
        if column not in df.columns:
            print(f"Warning: Column '{column}' not found. Skipping.")
            continue
        
        grouped = df.groupby('Usuario')[column]
        
        for year in [2021, 2022, 2023, 2024]:
            # Absolute variation
            abs_col = f'{year} {column} Var. Abs.'
            new_columns[abs_col] = grouped.diff()
            
            # Relative variation
            rel_col = f'{year} {column} Var. Rel.'
            pct_change = grouped.pct_change(fill_method=None) * 100
            new_columns[rel_col] = pct_change.apply(
                lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%"
            )
    
    # Add all new columns at once
    df = pd.concat([df, pd.DataFrame(new_columns)], axis=1)
    
    # Add trend symbols
    for column in columns_to_analyze:
        if column not in df.columns:
            continue
            
        for year in [2021, 2022, 2023, 2024]:
            abs_col = f'{year} {column} Var. Abs.'
            rel_col = f'{year} {column} Var. Rel.'
            
            if abs_col in df.columns:
                df[abs_col] = df.apply(
                    lambda row: f"{row[abs_col]:.2f} {get_trend_symbol(row[rel_col])}" 
                    if pd.notna(row[abs_col]) else "N/A ➡️",
                    axis=1
                )
            if rel_col in df.columns:
                df[rel_col] = df.apply(
                    lambda row: f"{row[rel_col]} {get_trend_symbol(row[rel_col])}", 
                    axis=1
                )
    
    return df

def save_results(df, excel_path, json_path):
    """Save DataFrame to both Excel and JSON formats."""
    try:
        # Save to Excel
        df.to_excel(excel_path, index=False)
        print(f"Excel file saved to: {excel_path}")
        
        # Save to JSON (ensure proper encoding for symbols)
        df.to_json(json_path, orient='records', indent=4, force_ascii=False)
        print(f"JSON file saved to: {json_path}")
    except Exception as e:
        print(f"Error saving files: {e}")

def main():
    try:
        # Load data
        df = pd.read_excel("tables/trends/trends.xlsx")
        print("Trends file loaded successfully.")
        
        # Calculate variations
        df = calculate_yearly_variations(df)
        
        # Save results
        save_results(
            df,
            excel_path="tables/trends/overTrends.xlsx",
            json_path="tables/trends/data.json"
        )
    
    except FileNotFoundError:
        print("Error: Input file 'trends.xlsx' not found.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
"@
}

function createIndex {
    Write-Host "🏗️ Creating Server" -ForegroundColor $YELLOW
    # server
    Set-Content -Path "models/server.py" -Value @"
import http.server
import socketserver
import os
import webbrowser
import threading

def start_server(port=8000):
  """Starts a simple HTTP server in a background thread."""
  Handler = http.server.SimpleHTTPRequestHandler

  with socketserver.TCPServer(("", port), Handler) as httpd:
    print(f"Serving at port {port}")
    httpd.serve_forever()

def main():
    """Automates the process of generating tables and opening index.html in a browser."""

    # Assuming the PowerShell script's tasks are already done and index.html is ready

    print("Hold CTRL and click http://localhost:8000/")

    # Get the absolute path to index.html
    index_path = os.path.abspath("index.html")

    # Check if index.html exists
    if os.path.exists(index_path):
        # Start the HTTP server in a separate thread
        server_thread = threading.Thread(target=start_server)
        server_thread.daemon = True  # Allow the main thread to exit even if the server is running
        server_thread.start()

        # Open index.html in the default browser
        webbrowser.open("http://localhost:8000")  # Open via the server URL
    else:
        print("Error: index.html not found in the current directory.")
        return # exit the program if index.html doesn't exist


    # Keep the main thread alive (optional) or perform other tasks
    input("Press Enter to stop the server and exit...\n")  #Wait for user input to terminate


if __name__ == "__main__":
    main()
"@

Write-Host "🏗️ Creating HTML" -ForegroundColor $YELLOW
    # html
    Set-Content -Path "index.html" -Value @"
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>A R P A</title>
    <link rel="stylesheet" href="static/style.css">
    <link rel="shortcut icon" href="favicon.png" type="image/x-icon">
</head>
<body>
    <div class="topnav-container">
        <div class="logoIN"></div>
        <div class="nomPag">A R P A</div>
    </div>
    
    <h1>Bienes y Rentas</h1>
    <div class="filter-form">
        <select id="dataSource">
            <option value="tables/trends/data.json">Datos Completos</option>
            <option value="tables/trends/data2.json">Datos fkID01</option>
        </select>
        <button onclick="changeDataSource()">Cambiar Datos</button>
    </div>
    <div class="filter-form">
        <select id="column">
            <option value="">-- Selecciona columna --</option>
            <optgroup label="Información Personal">
                <option value="Usuario">Nombre</option>
                <option value="Año Declaración">Año Declaración</option>
                <option value="Compañía">Compañía</option>
                <option value="Cargo">Cargo</option>
                <option value="Nombre">Usuario</option>
            </optgroup>
            <optgroup label="Valores Principales">
                <option value="Activos">Activos</option>
                <option value="Pasivos">Pasivos</option>
                <option value="Patrimonio">Patrimonio</option>
                <option value="Apalancamiento">Apalancamiento</option>
                <option value="Endeudamiento">Endeudamiento</option>
                <option value="Cant_Deudas">Cant. Deudas</option>
                <option value="BancoSaldo">Saldo Bancario</option>
                <option value="Cant_Bancos">Cant. Bancos</option>
                <option value="Bienes">Bienes</option>
                <option value="Cant_Bienes">Cant. Bienes</option>
                <option value="Inversiones">Inversiones</option>
                <option value="Cant_Inversiones">Cant. Inversiones</option>
                <option value="Ingresos">Ingresos</option>
                <option value="Cant_Ingresos">Cant. Ingresos</option>
            </optgroup>
            <optgroup label="Variaciones Absolutas">
                <option value="Activos Var. Abs.">Activos Var. Abs.</option>
                <option value="Pasivos Var. Abs.">Pasivos Var. Abs.</option>
                <option value="Patrimonio Var. Abs.">Patrimonio Var. Abs.</option>
                <option value="Apalancamiento Var. Abs.">Apalancamiento Var. Abs.</option>
                <option value="Endeudamiento Var. Abs.">Endeudamiento Var. Abs.</option>
                <option value="BancoSaldo Var. Abs.">BancoSaldo Var. Abs.</option>
                <option value="Bienes Var. Abs.">Bienes Var. Abs.</option>
                <option value="Inversiones Var. Abs.">Inversiones Var. Abs.</option>
                <option value="Ingresos Var. Abs.">Ingresos Var. Abs.</option>
            </optgroup>
            <optgroup label="Variaciones Relativas">
                <option value="Activos Var. Rel.">Activos Var. Rel.</option>
                <option value="Pasivos Var. Rel.">Pasivos Var. Rel.</option>
                <option value="Patrimonio Var. Rel.">Patrimonio Var. Rel.</option>
                <option value="Apalancamiento Var. Rel.">Apalancamiento Var. Rel.</option>
                <option value="Endeudamiento Var. Rel.">Endeudamiento Var. Rel.</option>
                <option value="BancoSaldo Var. Rel.">BancoSaldo Var. Rel.</option>
                <option value="Bienes Var. Rel.">Bienes Var. Rel.</option>
                <option value="Inversiones Var. Rel.">Inversiones Var. Rel.</option>
                <option value="Ingresos Var. Rel.">Ingresos Var. Rel.</option>
            </optgroup>
        </select>
        
        <select id="operator">
            <option value=">">Mayor que</option>
            <option value="<">Menor que</option>
            <option value="=">Igual a</option>
            <option value=">=">Mayor o igual</option>
            <option value="<=">Menor o igual</option>
            <option value="between">Entre</option>
            <option value="contains">Contiene</option>
        </select>
        
        <input type="text" id="value1" placeholder="Valor">
        <input type="text" id="value2" placeholder="y" style="display: none;">
        
        <button onclick="addFilter()">Agregar Filtro</button>
        <button onclick="exportToExcel()">Exportar Excel</button>

        <div class="filter-buttons">
            <button onclick="applyPredeterminedFilter('Patrimonio', '>', '3000000000')">Patrimonio > ,000M</button>
            <button onclick="applyPredeterminedFilter('Patrimonio Var. Rel.', '>', '30')">Patrimonio Var. Rel. > 30%</button>
            <button onclick="applyPredeterminedFilter('Cant_Bienes', '<', '0')">Cant. Bienes < 0</button>
            <button onclick="applyPredeterminedFilter('Endeudamiento Var. Rel.', '>', '50')">Endeudamiento Var. Rel. > 50%</button>
            <button onclick="applyPredeterminedFilter('Ingresos', '>', '50000000')">Ingresos > </button>
            <button onclick="applyPredeterminedFilter('Cant_Deudas', '>=', '5')">Cant. Deudas ≥ 5</button>
            <button onclick="applyPredeterminedFilter('Cant_Bienes', '>=', '6')">Cant. Bienes ≥ 6</button>
        </div>
    </div>
    
    <div id="filters"></div>
    
    <div class="table-scroll-container">
        <table id="results">
            <thead>
                <tr>
                    <th style="position: sticky; left: 0; top: 0; background-color: #f0f0f0; z-index: 2; padding: 8px; border: 1px solid #ddd;"><button onclick="quickFilter('Nombre')">Nombre<span class="sort-icon">↕</span></button></th>
                    <th style="position: sticky; left: 150px; top: 0; background-color: #f0f0f0; z-index: 2; padding: 8px; border: 1px solid #ddd;"><button onclick="quickFilter('Año Declaración')">Año Declaración<span class="sort-icon">↕</span></button></th>
                    <th style="position: sticky; left: 300px; top: 0; background-color: #f0f0f0; z-index: 2; padding: 8px; border: 1px solid #ddd;"><button onclick="quickFilter('Compañía')">Compañía<span class="sort-icon">↕</span></button></th>
                    <th style="position: sticky; left: 420px; top: 0; background-color: #f0f0f0; z-index: 2; padding: 8px; border: 1px solid #ddd;"><button onclick="quickFilter('Cargo')">Cargo<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Usuario')">Usuario<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Activos')">Activos<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Pasivos')">Pasivos<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Patrimonio')">Patrimonio<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Apalancamiento')">Apalancamiento<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Endeudamiento')">Endeudamiento<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Cant_Deudas')">Cant_Deudas<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('BancoSaldo')">BancoSaldo<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Cant_Bancos')">Cant_Bancos<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Bienes')">Bienes<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Cant_Bienes')">Cant_Bienes<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Inversiones')">Inversiones<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Cant_Inversiones')">Cant_Inversiones<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Ingresos')">Ingresos<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Cant_Ingresos')">Cant_Ingresos<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Activos Var. Abs.')">Activos Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Pasivos Var. Abs.')">Pasivos Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Patrimonio Var. Abs.')">Patrimonio Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Apalancamiento Var. Abs.')">Apalancamiento Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Endeudamiento Var. Abs.')">Endeudamiento Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('BancoSaldo Var. Abs.')">BancoSaldo Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Bienes Var. Abs.')">Bienes Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Inversiones Var. Abs.')">Inversiones Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Ingresos Var. Abs.')">Ingresos Var. Abs.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Activos Var. Rel.')">Activos Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Pasivos Var. Rel.')">Pasivos Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Patrimonio Var. Rel.')">Patrimonio Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Apalancamiento Var. Rel.')">Apalancamiento Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Endeudamiento Var. Rel.')">Endeudamiento Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('BancoSaldo Var. Rel.')">BancoSaldo Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Bienes Var. Rel.')">Bienes Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Inversiones Var. Rel.')">Inversiones Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th><button onclick="quickFilter('Ingresos Var. Rel.')">Ingresos Var. Rel.<span class="sort-icon">↕</span></button></th>
                    <th style="position: sticky; right: 0; background-color: #f8f9fa;">Acciones</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td colspan="35" class="loading">Cargando datos...</td>
                </tr>
            </tbody>
        </table>
    </div>

    <!-- SheetJS for Excel export -->
    <script src="https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js"></script>
    <script src="static/script.js"></script>
</body>
</html>
"@

Write-Host "🏗️ Creating CSS" -ForegroundColor $YELLOW
    # css
    Set-Content -Path "static/style.css" -Value @"
@import url('https://fonts.googleapis.com/css2?family=Open+Sans&display=swap');
* {
    font-family: 'Open Sans', sans-serif;
    box-sizing: border-box;
}
body {
    margin: 0;
    padding: 20px;
    background-color: #f8f9fa;
}
.topnav-container {
    display: flex;
    align-items: center;
    margin-bottom: 20px;
}
.logoIN {
    cursor: pointer;
    width: 40px;
    height: 40px;
    background-color: #0b00a2;
    border-radius: 8px;
    display: inline-flex;
    position: relative;
}
.logoIN::before {
    content: "";
    width: 40px;
    height: 40px;
    border-radius: 50%;
    position: absolute;
    top: 30%;
    left: 70%;
    transform: translate(-50%, -50%);
    background-image: linear-gradient(to right, 
        #ffffff 2px, transparent 1.5px,
        transparent 1.5px, #ffffff 1.5px,
        #ffffff 2px, transparent 1.5px);
    background-size: 4px 100%; 
}
.nomPag {
    margin-left: 10px;
    color: #0b00a2;
    font-weight: bold;
    font-size: 1rem;
}
h1 {
    color: #333;
    margin-top: 0;
}
.filter-form {
    background: white;
    padding: 15px;
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    margin-bottom: 20px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
}
.filter-form select, 
.filter-form input, 
.filter-form button {
    padding: 8px 12px;
    border: 1px solid #ddd;
    border-radius: 4px;
    transition: all 0.3s;
}
.filter-form button {
    background-color: #0b00a2;
    color: white;
    border: none;
    cursor: pointer;
}
.filter-form button:hover {
    background-color: #09007a;
}
.filter-form select:focus, 
.filter-form select.highlighted {
    border: 2px solid #0b00a2;
    background-color: #f0f5ff;
}
#filters {
    margin-bottom: 20px;
}
.filter-tag {
    display: inline-block;
    background: #e9ecef;
    padding: 5px 10px;
    border-radius: 20px;
    margin-right: 10px;
    margin-bottom: 10px;
}
.filter-tag button {
    background: none;
    border: none;
    color: #dc3545;
    margin-left: 5px;
    cursor: pointer;
}
.predetermined-filters {
    background: white;
    padding: 15px;
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    margin-bottom: 20px;
}
.predetermined-filters label {
    display: block;
    margin-bottom: 8px;
    cursor: pointer;
}
.table-scroll-container {
    position: relative;
    height: calc(100vh - 300px);
    overflow: auto;
    margin-top: 20px;
    border: 1px solid #ddd;
    border-radius: 8px;
    background: white;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
#results {
    width: 100%;
    border-collapse: collapse;
    min-width: 1200px;
}
#results th, #results td {
    padding: 12px;
    text-align: left;
    border-bottom: 1px solid #ddd;
    white-space: nowrap;
}
#results thead {
    position: sticky;
    top: 0;
    z-index: 10;
    background-color: #f8f9fa;
}
#results th {
    position: sticky;
    top: 0;
    background-color: #f8f9fa;
}
#results tr:hover {
    background-color: #f1f1f1;
}
#results tr:hover td {
    background-color: #f1f1f1;
}
#results td:last-child {
    position: sticky;
    right: 0;
    z-index: 5;
    background-color: white;
}
#results tr:hover td:last-child {
    background-color: #f1f1f1;
}
#results th button {
    padding: 12px;
    background: none;
    border: none;
    width: 100%;
    text-align: left;
    font-weight: bold;
    cursor: pointer;
}

#results th button:hover {
    color: #0b00a2;
}

#results th button .sort-icon {
    margin-left: 5px;
    opacity: 0.5;
}

#results th button:hover .sort-icon {
    opacity: 1;
}

#results th.sorted-asc button,
#results th.sorted-desc button {
    color: #0b00a2;
    font-weight: bold;
}

#results th.sorted-asc button .sort-icon,
#results th.sorted-desc button .sort-icon {
    opacity: 1;
}

.trend-icon {
    font-size: 1.2em;
    margin-left: 3px;
}
.loading {
    text-align: center;
    padding: 20px;
    font-style: italic;
    color: #6c757d;
}
.highlighted-column {
    background-color: #f0f5ff !important;
    font-weight: bold;
}

/* Modal Styles */
.modal-overlay {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0,0,0,0.5);
    display: flex;
    justify-content: center;
    align-items: center;
    z-index: 1000;
}
.modal-content {
    background: white;
    padding: 20px;
    border-radius: 8px;
    width: 90%;
    max-width: 800px;
    max-height: 90vh;
    overflow-y: auto;
    box-shadow: 0 4px 8px rgba(0,0,0,0.2);
}
.modal-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 20px;
    padding-bottom: 10px;
    border-bottom: 1px solid #eee;
}
.close-button {
    background: none;
    border: none;
    font-size: 24px;
    cursor: pointer;
    color: #666;
}
.detail-section {
    margin-bottom: 25px;
    padding-bottom: 15px;
    border-bottom: 1px solid #f0f0f0;
}
.detail-section:last-child {
    border-bottom: none;
}
.detail-grid, .variation-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
    gap: 12px;
}
.detail-item, .variation-item {
    padding: 8px;
    background: #f9f9f9;
    border-radius: 4px;
}
.detail-item strong, .variation-item strong {
    color: #333;
}
.modal-body {
    display: flex;
    flex-direction: column;
    gap: 20px;
}
.detail-section h3 {
    margin-top: 0;
    color: #0b00a2;
    border-bottom: 1px solid #dee2e6;
    padding-bottom: 8px;
}
.variation-item {
    border-left: 3px solid #0b00a2;
}
.filter-buttons {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    margin-left: 10px;
}

.filter-buttons button {
    background-color: #e9ecef;
    color: #495057;
    border: 1px solid #ced4da;
    padding: 5px 10px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 0.9rem;
    transition: all 0.2s;
}

.filter-buttons button:hover {
    background-color: #dee2e6;
}

.filter-buttons button.active {
    background-color: #0b00a2;
    color: white;
    border-color: #0b00a2;
}

@media (max-width: 768px) {
    .filter-form select, 
    .filter-form input, 
    .filter-form button {
        width: 100%;
        margin-right: 0;
    }
    #results thead {
        display: none;
    }
    #results tr {
        display: block;
        margin-bottom: 15px;
        border: 1px solid #ddd;
        border-radius: 8px;
    }
    #results td {
        display: flex;
        justify-content: space-between;
        align-items: center;
        text-align: right;
        padding-left: 50%;
        position: relative;
    }
    #results td::before {
        content: attr(data-label);
        position: absolute;
        left: 12px;
        font-weight: bold;
    }
    .detail-grid, .variation-grid {
        grid-template-columns: 1fr;
    }
}
"@

Write-Host "🏗️ Creating Javascript" -ForegroundColor $YELLOW
    #javascript
Set-Content -Path "static/script.js" -Value @'
// Global variables
let allData = [];
let filteredData = [];
let lastSelectedColumn = '';
let currentFilterColumn = '';
let currentSortColumn = '';
let sortDirection = 'asc';
let processingData = false;
const filters = [];
let currentDataSource = 'tables/trends/data.json';

// DOM elements
const operatorSelect = document.getElementById('operator');
const value2Input = document.getElementById('value2');

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    setupEventListeners();
});

// Load JSON data
async function loadData() {
    try {
        document.querySelector('#results tbody').innerHTML = `
            <tr>
                <td colspan="35" class="loading">Cargando datos...</td>
            </tr>
        `;

        const response = await fetch(currentDataSource);
        if (!response.ok) throw new Error(`Error al cargar ${currentDataSource}`);
        
        allData = await response.json();
        filteredData = [...allData];
        renderTable();
        
    } catch (error) {
        console.error('Error:', error);
        document.querySelector('#results tbody').innerHTML = `
            <tr>
                <td colspan="35">Error al cargar ${currentDataSource}: ${error.message}</td>
            </tr>
        `;
    }
}

function changeDataSource() {
    const dataSourceSelect = document.getElementById('dataSource');
    currentDataSource = dataSourceSelect.value;
    
    // Clear existing filters
    filters.length = 0;
    renderFilters();
    
    // Reset sort
    currentSortColumn = '';
    sortDirection = 'asc';
    
    // Reload data
    loadData();
    
    // Clear any highlights
    document.querySelectorAll('.highlighted-column').forEach(el => {
        el.classList.remove('highlighted-column');
    });
    document.getElementById('column').classList.remove('highlighted');
    currentFilterColumn = '';
    lastSelectedColumn = '';
    
    // Reset column dropdown
    document.getElementById('column').selectedIndex = 0;
}

// Set up event listeners
function setupEventListeners() {
    operatorSelect.addEventListener('change', toggleValue2Input);
    
    document.getElementById('column').addEventListener('change', function() {
        currentFilterColumn = this.value;
        if (this.value) {
            this.classList.add('highlighted');
        } else {
            this.classList.remove('highlighted');
        }
        
        // Auto-focus the value input for quick filtering
        if (this.value && lastSelectedColumn !== this.value) {
            document.getElementById('value1').focus();
        }
        lastSelectedColumn = this.value;
    });
    // listener to excelupload
    document.getElementById('excelUpload').addEventListener('change', handleFileUpload);
}

// excel upload fucntion
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Show loading state
    processingData = true;
    document.querySelector('#results tbody').innerHTML = `
        <tr>
            <td colspan="37" class="loading">Procesando archivo Excel, por favor espere...</td>
        </tr>
    `;

    try {
        const formData = new FormData();
        formData.append('file', file);

        const response = await fetch('/src', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error('Error en el análisis del archivo');
        }

        const result = await response.json();
        
        if (result.success) {
            // Reload the data after processing
            await loadData();
            alert('Archivo procesado correctamente. Los datos han sido actualizados.');
        } else {
            throw new Error(result.message || 'Error desconocido al procesar el archivo');
        }
    } catch (error) {
        console.error('Error:', error);
        document.querySelector('#results tbody').innerHTML = `
            <tr>
                <td colspan="37">Error al procesar el archivo: ${error.message}</td>
            </tr>
        `;
    } finally {
        processingData = false;
        // Reset the file input
        event.target.value = '';
    }
}

// Toggle second value input for 'between' operator
function toggleValue2Input() {
    value2Input.style.display = operatorSelect.value === 'between' ? 'inline-block' : 'none';
}

// Add a new filter
function addFilter() {
    const column = document.getElementById('column').value;
    const operator = operatorSelect.value;
    const value1 = document.getElementById('value1').value.trim();
    let value2 = '';
    
    if (!column || !operator || !value1) {
        alert('Por favor complete todos los campos del filtro');
        return;
    }
    
    if (operator === 'between') {
        value2 = document.getElementById('value2').value.trim();
        if (!value2) {
            alert('Por favor complete el segundo valor para el filtro "Entre"');
            return;
        }
    }
    
    filters.push({ column, operator, value1, value2 });
    renderFilters();
    
    // Highlight the column in the table
    highlightColumn(column);
    
    // Keep the column selected in the dropdown
    lastSelectedColumn = column;
    currentFilterColumn = column;
    
    applyFilters();
}

function quickFilter(columnName) {
    // If clicking the same column, toggle sort direction
    if (currentSortColumn === columnName) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortColumn = columnName;
        sortDirection = 'asc';
    }
    
    // Remove any existing sort indicators
    document.querySelectorAll('#results th').forEach(th => {
        th.classList.remove('sorted-asc', 'sorted-desc');
    });
    
    // Add sort indicator to current column
    const columnMap = {
        'Usuario': 2,
        'Nombre': 0,
        'Compañía': 2,
        'Cargo': 3,
        'Año Declaración': 1,
        'Activos': 5,
        'Pasivos': 6,
        'Patrimonio': 7,
        'Apalancamiento': 8,
        'Endeudamiento': 9,
        'Cant_Deudas': 10,
        'BancoSaldo': 11,
        'Cant_Bancos': 12,
        'Bienes': 13,
        'Cant_Bienes': 14,
        'Inversiones': 15,
        'Cant_Inversiones': 16,
        'Ingresos': 17,
        'Cant_Ingresos': 18,
        'Activos Var. Abs.': 19,
        'Pasivos Var. Abs.': 20,
        'Patrimonio Var. Abs.': 21,
        'Apalancamiento Var. Abs.': 22,
        'Endeudamiento Var. Abs.': 23,
        'BancoSaldo Var. Abs.': 24,
        'Bienes Var. Abs.': 25,
        'Inversiones Var. Abs.': 26,
        'Ingresos Var. Abs.': 27,
        'Activos Var. Rel.': 28,
        'Pasivos Var. Rel.': 29,
        'Patrimonio Var. Rel.': 30,
        'Apalancamiento Var. Rel.': 31,
        'Endeudamiento Var. Rel.': 32,
        'BancoSaldo Var. Rel.': 33,
        'Bienes Var. Rel.': 34,
        'Inversiones Var. Rel.': 35,
        'Ingresos Var. Rel.': 36
    };
    
    const columnIndex = columnMap[columnName];
    if (columnIndex !== undefined) {
        const header = document.querySelector(`#results th:nth-child(${columnIndex + 1})`);
        if (header) {
            header.classList.add(`sorted-${sortDirection}`);
            
            // Update sort icon
            const icon = header.querySelector('.sort-icon');
            if (icon) {
                icon.textContent = sortDirection === 'asc' ? '↑' : '↓';
            }
        }
    }
    
    // Open filter dialog with column pre-selected
    document.getElementById('column').value = columnName;
    document.getElementById('column').classList.add('highlighted');
    document.getElementById('value1').focus();
    
    // Highlight the column
    highlightColumn(columnName);
    currentFilterColumn = columnName;
    lastSelectedColumn = columnName;
    
    // Auto-sort when clicking 
    sortTable(columnName, sortDirection);
}

function sortTable(columnName, direction) {
    filteredData.sort((a, b) => {
        let valA = a[columnName];
        let valB = b[columnName];
        
        // Handle numeric values
        if (!isNaN(parseFloat(valA)) && !isNaN(parseFloat(valB))) {
            valA = parseFloat(valA);
            valB = parseFloat(valB);
            return direction === 'asc' ? valA - valB : valB - valA;
        }
        
        // Handle string values
        if (typeof valA === 'string' && typeof valB === 'string') {
            return direction === 'asc' 
                ? valA.localeCompare(valB) 
                : valB.localeCompare(valA);
        }
        
        return 0;
    });
    
    renderTable();
    
    // Re-apply column highlight if there's a current filter column
    if (currentFilterColumn) {
        highlightColumn(currentFilterColumn);
    }
}

// Highlight table column
function highlightColumn(columnName) {
    // Remove any existing highlights
    const headers = document.querySelectorAll('#results th');
    const cells = document.querySelectorAll('#results td');
    
    headers.forEach(header => header.classList.remove('highlighted-column'));
    cells.forEach(cell => cell.classList.remove('highlighted-column'));
    
    // Find the column index
    const columnMap = {
        'Usuario': 4,
        'Nombre': 0,
        'Compañía': 2,
        'Cargo': 3,
        'Año Declaración': 1,
        'Activos': 5,
        'Pasivos': 6,
        'Patrimonio': 7,
        'Apalancamiento': 8,
        'Endeudamiento': 9,
        'Cant_Deudas': 10,
        'BancoSaldo': 11,
        'Cant_Bancos': 12,
        'Bienes': 13,
        'Cant_Bienes': 14,
        'Inversiones': 15,
        'Cant_Inversiones': 16,
        'Ingresos': 17,
        'Cant_Ingresos': 18,
        'Activos Var. Abs.': 19,
        'Pasivos Var. Abs.': 20,
        'Patrimonio Var. Abs.': 21,
        'Apalancamiento Var. Abs.': 22,
        'Endeudamiento Var. Abs.': 23,
        'BancoSaldo Var. Abs.': 24,
        'Bienes Var. Abs.': 25,
        'Inversiones Var. Abs.': 26,
        'Ingresos Var. Abs.': 27,
        'Activos Var. Rel.': 28,
        'Pasivos Var. Rel.': 29,
        'Patrimonio Var. Rel.': 30,
        'Apalancamiento Var. Rel.': 31,
        'Endeudamiento Var. Rel.': 32,
        'BancoSaldo Var. Rel.': 33,
        'Bienes Var. Rel.': 34,
        'Inversiones Var. Rel.': 35,
        'Ingresos Var. Rel.': 36
    };
    
    const columnIndex = columnMap[columnName];
    if (columnIndex === undefined) return;
    
    // Highlight header
    if (headers[columnIndex]) {
        headers[columnIndex].classList.add('highlighted-column');
    }
    
    // Highlight cells
    document.querySelectorAll(`#results tr td:nth-child(${columnIndex + 1})`).forEach(cell => {
        cell.classList.add('highlighted-column');
    });
}

// Render active filters
function renderFilters() {
    const filtersContainer = document.getElementById('filters');
    filtersContainer.innerHTML = filters.map((filter, index) => `
        <div class="filter-tag">
            ${filter.column} ${getOperatorSymbol(filter.operator)} ${filter.value1}
            ${filter.operator === 'between' ? ` y ${filter.value2}` : ''}
            <button onclick="removeFilter(${index})">×</button>
        </div>
    `).join('');
}

// Get operator symbol for display
function getOperatorSymbol(operator) {
    const symbols = {
        '>': '>',
        '<': '<',
        '=': '=',
        '>=': '≥',
        '<=': '≤',
        'between': 'entre',
        'contains': 'contiene'
    };
    return symbols[operator] || operator;
}

// Remove a filter
function removeFilter(index) {
    filters.splice(index, 1);
    renderFilters();
    applyFilters();
}

// Apply all active filters
function applyFilters() {
    if (filters.length === 0) {
        filteredData = [...allData];
        renderTable();
        return;
    }
    
    filteredData = allData.filter(item => {
        return filters.every(filter => {
            const itemValue = item[filter.column];
            if (itemValue === undefined || itemValue === null) return false;
            
            // Handle percentage values
            let numericValue;
            if (typeof itemValue === 'string' && itemValue.includes('%')) {
                numericValue = parseFloat(itemValue.replace('%', ''));
            } else {
                numericValue = parseFloat(itemValue);
            }
            
            const filterValue1 = parseFloat(filter.value1);
            const filterValue2 = parseFloat(filter.value2);
            
            switch (filter.operator) {
                case '>': return numericValue > filterValue1;
                case '<': return numericValue < filterValue1;
                case '=': return numericValue === filterValue1;
                case '>=': return numericValue >= filterValue1;
                case '<=': return numericValue <= filterValue1;
                case 'between': 
                    return numericValue >= filterValue1 && numericValue <= filterValue2;
                case 'contains':
                    return String(itemValue).toLowerCase().includes(filter.value1.toLowerCase());
                default: return true;
            }
        });
    });
    
    renderTable();
    
    // Re-apply column highlight if there's a current filter column
    if (currentFilterColumn) {
        highlightColumn(currentFilterColumn);
    }
}

// Handle predetermined filters
function applyPredeterminedFilter(column, operator, value1) {
    // Check if this filter already exists
    const existingIndex = filters.findIndex(f => 
        f.column === column && f.operator === operator && f.value1 === value1
    );
    
    if (existingIndex === -1) {
        // Add new filter
        filters.push({ column, operator, value1 });
        
        // Highlight the button
        const buttons = document.querySelectorAll('.filter-buttons button');
        buttons.forEach(button => {
            if (button.textContent.includes(column) && 
                button.textContent.includes(operator) &&
                button.textContent.includes(value1)) {
                button.classList.add('active');
            }
        });
    } else {
        // Remove existing filter
        filters.splice(existingIndex, 1);
        
        // Remove highlight from button
        const buttons = document.querySelectorAll('.filter-buttons button');
        buttons.forEach(button => {
            if (button.textContent.includes(column) && 
                button.textContent.includes(operator) &&
                button.textContent.includes(value1)) {
                button.classList.remove('active');
            }
        });
    }
    
    renderFilters();
    applyFilters();
    highlightColumn(column);
}

// Clear all filters
function clearFilters() {
    filters.length = 0;
    renderFilters();
    filteredData = [...allData];
    renderTable();
    
    // Clear button highlights
    document.querySelectorAll('.filter-buttons button').forEach(button => {
        button.classList.remove('active');
    });
    
    // Clear column highlights
    document.querySelectorAll('.highlighted-column').forEach(el => {
        el.classList.remove('highlighted-column');
    });
    document.getElementById('column').classList.remove('highlighted');
    currentFilterColumn = '';
    lastSelectedColumn = '';
}

// Render the data table
function renderTable() {
    const tbody = document.querySelector('#results tbody');
    
    if (filteredData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="37">No se encontraron resultados con los filtros aplicados</td>
            </tr>
        `;
        return;
    }
    
    tbody.innerHTML = filteredData.map(item => {
        // Function to format cell with color based on value
        const formatCell = (value, isPercentage = false) => {
            if (value === undefined || value === null || value === '') return '';
            
            // Remove trend icons if present
            const cleanValue = String(value).replace(/[📈📉➡️]/g, '').trim();
            
            // Try to parse as number
            const numValue = parseFloat(cleanValue.replace('%', '').replace(/[^\d.-]/g, ''));
            if (isNaN(numValue)) return value;
            
            // Format number
            let formattedValue;
            if (isPercentage) {
                formattedValue = numValue.toFixed(2) + '%';
            } else if (Math.abs(numValue) >= 1000000) {
                formattedValue = '$' + (numValue / 1000000).toFixed(2) + 'M';
            } else {
                formattedValue = new Intl.NumberFormat('es-CO').format(numValue);
            }
            
            // Determine color - only red for negative, black otherwise
            const color = numValue < 0 ? 'color: #dc3545;' : 'color: #000;';
            
            return `<span style="${color}">${formattedValue}</span>`;
        };
        
        return `
            <tr>
                <td style="position: sticky; left: 0; background-color: white; z-index: 2;">${item.Nombre || ''}</td>
                <td style="position: sticky; left: 150px; background-color: white; z-index: 2;">${item['Año Declaración'] || ''}</td>
                <td style="position: sticky; left: 300px; background-color: white; z-index: 2;">${item['Compañía'] || ''}</td>
                <td style="position: sticky; left: 420px; background-color: white; z-index: 2;">${item.Cargo || ''}</td>
                <td>${item.Usuario || ''}</td>
                <td>${formatCell(item.Activos)}</td>
                <td>${formatCell(item.Pasivos)}</td>
                <td>${formatCell(item.Patrimonio)}</td>
                <td>${formatCell(item.Apalancamiento, true)}</td>
                <td>${formatCell(item.Endeudamiento, true)}</td>
                <td>${formatCell(item['Cant_Deudas'])}</td>
                <td>${formatCell(item.BancoSaldo)}</td>
                <td>${formatCell(item.Cant_Bancos)}</td>
                <td>${formatCell(item.Bienes)}</td>
                <td>${formatCell(item.Cant_Bienes)}</td>
                <td>${formatCell(item.Inversiones)}</td>
                <td>${formatCell(item.Cant_Inversiones)}</td>
                <td>${formatCell(item.Ingresos)}</td>
                <td>${formatCell(item.Cant_Ingresos)}</td>
                <td>${formatCell(item['Activos Var. Abs.'])}</td>
                <td>${formatCell(item['Pasivos Var. Abs.'])}</td>
                <td>${formatCell(item['Patrimonio Var. Abs.'])}</td>
                <td>${formatCell(item['Apalancamiento Var. Abs.'], true)}</td>
                <td>${formatCell(item['Endeudamiento Var. Abs.'], true)}</td>
                <td>${formatCell(item['BancoSaldo Var. Abs.'])}</td>
                <td>${formatCell(item['Bienes Var. Abs.'])}</td>
                <td>${formatCell(item['Inversiones Var. Abs.'])}</td>
                <td>${formatCell(item['Ingresos Var. Abs.'])}</td>
                <td>${formatCell(item['Activos Var. Rel.'], true)}</td>
                <td>${formatCell(item['Pasivos Var. Rel.'], true)}</td>
                <td>${formatCell(item['Patrimonio Var. Rel.'], true)}</td>
                <td>${formatCell(item['Apalancamiento Var. Rel.'], true)}</td>
                <td>${formatCell(item['Endeudamiento Var. Rel.'], true)}</td>
                <td>${formatCell(item['BancoSaldo Var. Rel.'], true)}</td>
                <td>${formatCell(item['Bienes Var. Rel.'], true)}</td>
                <td>${formatCell(item['Inversiones Var. Rel.'], true)}</td>
                <td>${formatCell(item['Ingresos Var. Rel.'], true)}</td>
                <td style="position: sticky; right: 0; background-color: white;">
                    <button onclick="viewDetails('${item.Usuario}', ${item['Año Declaración']})" style="background-color: #0b00a2; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer;">Detalles</button>
                </td>
            </tr>
        `;
    }).join('');

    // Scroll to top after rendering
    document.querySelector('.table-scroll-container').scrollTop = 0;
}

// View detailed record
function viewDetails(userId, year) {
    // Convert year to number if it's coming as string
    year = typeof year === 'string' ? parseInt(year) : year;
    
    // Find the record with case-insensitive comparison
    const record = allData.find(item => {
        // Handle potential undefined/null values
        const itemUserId = item.Usuario ? item.Usuario.toString().toLowerCase() : '';
        const itemYear = item['Año Declaración'] ? parseInt(item['Año Declaración']) : null;
        
        return itemUserId === userId.toLowerCase() && itemYear === year;
    });

    if (!record) {
        alert(`Registro no encontrado para:\nUsuario: ${userId}\nAño: ${year}`);
        console.error('Record not found:', { userId, year, allData });
        return;
    }

    // Create modal HTML with all available data
    const modalHTML = `
        <div id="detailModal" class="modal-overlay">
            <div class="modal-content">
                <div class="modal-header">
                    <h2>Detalles Completo - ${record.Nombre} (${year})</h2>
                    <div>
                        <button onclick="exportDetailsToExcel()" style="margin-right: 10px; padding: 5px 10px; background-color: #0b00a2; color: white; border: none; border-radius: 4px; cursor: pointer;">
                            Exportar a Excel
                        </button>
                        <button onclick="closeModal()" class="close-button">×</button>
                    </div>
                </div>
                
                <div class="modal-body">
                    ${renderDetailSection('Información Básica', [
                        { label: 'Nombre', value: record.Nombre },
                        { label: 'Usuario', value: record.Usuario },
                        { label: 'Compañía', value: record['Compañía'] },
                        { label: 'Cargo', value: record.Cargo },
                        { label: 'Año Declaración', value: record['Año Declaración'] },
                        { label: 'Año Creación', value: record['Año Creación'] }
                    ])}
                    
                    ${renderFinancialSection('Resumen Financiero', record)}
                    
                    ${renderVariationSection('Variaciones Recientes', record)}
                    
                    <!-- Yearly Variations Sections -->
                    ${renderYearlyVariationSection('Variaciones con 2021', record, '2021')}
                    ${renderYearlyVariationSection('Variaciones con 2022', record, '2022')}
                    ${renderYearlyVariationSection('Variaciones con 2023', record, '2023')}
                    ${renderYearlyVariationSection('Variaciones con 2024', record, '2024')}
                    
                    <!-- Yearly Count Variations Sections -->
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2021', record, '2021')}
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2022', record, '2022')}
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2023', record, '2023')}
                    ${renderYearlyCountVariationSection('Variaciones de Cantidad con 2024', record, '2024')}
                </div>
            </div>
        </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHTML);
}

// Helper function to render detail sections
function renderDetailSection(title, fields) {
    const filteredFields = fields.filter(field => field.value !== undefined && field.value !== null);
    
    if (filteredFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="detail-grid">
                ${filteredFields.map(field => `
                    <div class="detail-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Helper function to render financial section
function renderFinancialSection(title, record) {
    const financialFields = [
        { label: 'Activos', value: formatNumber(record.Activos) },
        { label: 'Pasivos', value: formatNumber(record.Pasivos) },
        { label: 'Patrimonio', value: formatNumber(record.Patrimonio) },
        { label: 'Apalancamiento', value: record.Apalancamiento ? `${record.Apalancamiento}%` : null },
        { label: 'Endeudamiento', value: record.Endeudamiento ? `${record.Endeudamiento}%` : null },
        { label: 'Saldo Bancario', value: formatNumber(record.BancoSaldo) },
        { label: 'Bienes', value: formatNumber(record.Bienes) },
        { label: 'Inversiones', value: formatNumber(record.Inversiones) },
        { label: 'Ingresos', value: formatNumber(record.Ingresos) }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    return renderDetailSection(title, financialFields);
}

// Helper function to render variation sections
function renderVariationSection(title, record) {
    const variationFields = [
        { label: 'Activos Var. Abs.', value: record['Activos Var. Abs.'] },
        { label: 'Activos Var. Rel.', value: record['Activos Var. Rel.'] },
        { label: 'Pasivos Var. Abs.', value: record['Pasivos Var. Abs.'] },
        { label: 'Pasivos Var. Rel.', value: record['Pasivos Var. Rel.'] },
        { label: 'Patrimonio Var. Abs.', value: record['Patrimonio Var. Abs.'] },
        { label: 'Patrimonio Var. Rel.', value: record['Patrimonio Var. Rel.'] },
        { label: 'Apalancamiento Var. Abs.', value: record['Apalancamiento Var. Abs.'] },
        { label: 'Apalancamiento Var. Rel.', value: record['Apalancamiento Var. Rel.'] },
        { label: 'Endeudamiento Var. Abs.', value: record['Endeudamiento Var. Abs.'] },
        { label: 'Endeudamiento Var. Rel.', value: record['Endeudamiento Var. Rel.'] },
        { label: 'BancoSaldo Var. Abs.', value: record['BancoSaldo Var. Abs.'] },
        { label: 'BancoSaldo Var. Rel.', value: record['BancoSaldo Var. Rel.'] },
        { label: 'Bienes Var. Abs.', value: record['Bienes Var. Abs.'] },
        { label: 'Bienes Var. Rel.', value: record['Bienes Var. Rel.'] },
        { label: 'Inversiones Var. Abs.', value: record['Inversiones Var. Abs.'] },
        { label: 'Inversiones Var. Rel.', value: record['Inversiones Var. Rel.'] },
        { label: 'Ingresos Var. Abs.', value: record['Ingresos Var. Abs.'] },
        { label: 'Ingresos Var. Rel.', value: record['Ingresos Var. Rel.'] }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    if (variationFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="variation-grid">
                ${variationFields.map(field => `
                    <div class="variation-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Helper function to render yearly variation sections
function renderYearlyVariationSection(title, record, year) {
    const yearlyFields = [
        { label: `Activos Var. Abs. ${year}`, value: record[`${year} Activos Var. Abs.`] },
        { label: `Activos Var. Rel. ${year}`, value: record[`${year} Activos Var. Rel.`] },
        { label: `Pasivos Var. Abs. ${year}`, value: record[`${year} Pasivos Var. Abs.`] },
        { label: `Pasivos Var. Rel. ${year}`, value: record[`${year} Pasivos Var. Rel.`] },
        { label: `Patrimonio Var. Abs. ${year}`, value: record[`${year} Patrimonio Var. Abs.`] },
        { label: `Patrimonio Var. Rel. ${year}`, value: record[`${year} Patrimonio Var. Rel.`] },
        { label: `BancoSaldo Var. Abs. ${year}`, value: record[`${year} BancoSaldo Var. Abs.`] },
        { label: `BancoSaldo Var. Rel. ${year}`, value: record[`${year} BancoSaldo Var. Rel.`] },
        { label: `Bienes Var. Abs. ${year}`, value: record[`${year} Bienes Var. Abs.`] },
        { label: `Bienes Var. Rel. ${year}`, value: record[`${year} Bienes Var. Rel.`] },
        { label: `Inversiones Var. Abs. ${year}`, value: record[`${year} Inversiones Var. Abs.`] },
        { label: `Inversiones Var. Rel. ${year}`, value: record[`${year} Inversiones Var. Rel.`] },
        { label: `Ingresos Var. Abs. ${year}`, value: record[`${year} Ingresos Var. Abs.`] },
        { label: `Ingresos Var. Rel. ${year}`, value: record[`${year} Ingresos Var. Rel.`] }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    if (yearlyFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="variation-grid">
                ${yearlyFields.map(field => `
                    <div class="variation-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Helper function to render yearly count variations
function renderYearlyCountVariationSection(title, record, year) {
    const countVariationFields = [
        { label: `Cant. Deudas Var. Abs. ${year}`, value: record[`${year} Cant_Deudas Var. Abs.`] },
        { label: `Cant. Deudas Var. Rel. ${year}`, value: record[`${year} Cant_Deudas Var. Rel.`] },
        { label: `Cant. Bancos Var. Abs. ${year}`, value: record[`${year} Cant_Bancos Var. Abs.`] },
        { label: `Cant. Bancos Var. Rel. ${year}`, value: record[`${year} Cant_Bancos Var. Rel.`] },
        { label: `Cant. Bienes Var. Abs. ${year}`, value: record[`${year} Cant_Bienes Var. Abs.`] },
        { label: `Cant. Bienes Var. Rel. ${year}`, value: record[`${year} Cant_Bienes Var. Rel.`] },
        { label: `Cant. Inversiones Var. Abs. ${year}`, value: record[`${year} Cant_Inversiones Var. Abs.`] },
        { label: `Cant. Inversiones Var. Rel. ${year}`, value: record[`${year} Cant_Inversiones Var. Rel.`] },
        { label: `Cant. Ingresos Var. Abs. ${year}`, value: record[`${year} Cant_Ingresos Var. Abs.`] },
        { label: `Cant. Ingresos Var. Rel. ${year}`, value: record[`${year} Cant_Ingresos Var. Rel.`] }
    ].filter(field => field.value !== undefined && field.value !== null);
    
    if (countVariationFields.length === 0) return '';
    
    return `
        <div class="detail-section">
            <h3>${title}</h3>
            <div class="variation-grid">
                ${countVariationFields.map(field => `
                    <div class="variation-item">
                        <strong>${field.label}:</strong>
                        <span>${field.value}</span>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Format numbers with thousands separator
function formatNumber(num) {
    if (num === undefined || num === null) return 'N/A';
    return new Intl.NumberFormat('es-CO').format(num);
}

// Close modal function
function closeModal() {
    const modal = document.getElementById('detailModal');
    if (modal) modal.remove();
}

// Export details to Excel
function exportDetailsToExcel() {
    const modal = document.getElementById('detailModal');
    if (!modal) return;
    
    // Get all the data from the modal
    const data = [];
    const sections = modal.querySelectorAll('.detail-section');
    
    sections.forEach(section => {
        const sectionTitle = section.querySelector('h3').textContent;
        const items = section.querySelectorAll('.detail-item, .variation-item');
        
        items.forEach(item => {
            const label = item.querySelector('strong')?.textContent.replace(':', '') || '';
            const value = item.querySelector('span')?.textContent || item.textContent.replace(label, '').replace(':', '').trim();
            
            if (label && value) {
                data.push({
                    'Sección': sectionTitle,
                    'Campo': label,
                    'Valor': value
                });
            }
        });
    });
    
    if (data.length === 0) {
        alert('No hay datos para exportar');
        return;
    }
    
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Detalles");
    
    // Get the user name from the modal title
    const modalTitle = modal.querySelector('.modal-header h2').textContent;
    const fileName = modalTitle.replace('Detalles Completo - ', '').replace(/[\/\\?%*:|"<>]/g, '-') + '.xlsx';
    
    // Export to file
    XLSX.writeFile(workbook, fileName);
}

// Export to Excel
function exportToExcel() {
    if (filteredData.length === 0) {
        alert('No hay datos para exportar');
        return;
    }
    
    // Use the original data with trend icons
    const exportData = filteredData.map(item => ({...item}));
    
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Datos Filtrados");
    
    // Export to file
    XLSX.writeFile(workbook, 'datos_filtrados.xlsx');
}
'@
}

function createStructure {
    Write-Host "🏗️ Creating Structure" -ForegroundColor $YELLOW

    # Create Python virtual environment
    python -m venv .venv
    .\.venv\Scripts\Activate.ps1

    # Upgrade pip and install required packages
    python -m pip install --upgrade pip
    python -m pip install pandas python-dotenv openpyxl plotly msoffcrypto-tool

    # Always create subdirectories
    Write-Host "🏗️ Creating directory structure" -ForegroundColor $YELLOW
    $directories = @(
        "static",
        "models/trends",
        "tables/cats",
        "tables/nets",
        "tables/trends"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

}

function generateTables {
    Write-Host "🏗️ Generating Tables" -ForegroundColor $GREEN

    # Python scripts to generate tables
    $scripts = @(
        "models/passKey.py",
        "models/cats.py",
        "models/nets.py",
        "models/toJson.py"
    )

    foreach ($script in $scripts) {
        #Execute the script
        python $script
    }
}

function generaTrends {
    Write-Host "🏗️ Generating Trends" -ForegroundColor $GREEN

    # Python scripts to generate tables
    $scripts = @(
        "models/trends/trends.py",
        "models/trends/overTrends.py"
        
    )

    foreach ($script in $scripts) {
        #Execute the script
        python $script
    }
}

function deleteData {
    param (
        [string]$rootDir = "."  # Define $rootDir as a parameter with a default value
    )

    # Define the path to the file
    $filePath = Join-Path -Path $rootDir -ChildPath "src/data.xlsx"

    # Check if the file exists
    if (Test-Path -Path $filePath -PathType Leaf) {
        try {
            Remove-Item -Path $filePath -Force
            Write-Host "Deleted file: $filePath"
        }
        catch {
            Write-Host "Error deleting file $($filePath): $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "File does not exist: $filePath"
    }
}
function main {
    Write-Host "🏗️ BYR Analyzing" -ForegroundColor $NC

    # Call functions to create structure and generate tables
    createStructure
    passKey
    createCats
    createNets
    createTrends

    # Generate tables
    generateTables
    generaTrends
    deleteData

    Write-Host "🏗️ The tables are generated" -ForegroundColor $YELLOW
    Write-Host "🏗️ Opening index.html in browser..." -ForegroundColor $GREEN
    
    createIndex
    python models/server.py

}

main