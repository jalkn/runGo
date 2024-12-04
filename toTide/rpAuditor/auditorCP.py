import pandas as pd
import numpy as np
from datetime import datetime
import os
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns
from typing import List, Dict

""" Configuracion
1. Crear un entorno virtual
python -m venv venv

2. Activar el entorno virtual
source venv/bin/activate

3. Instalar las librerias necesarias
pip install pandas numpy matplotlib seaborn openpyxl
"""

class ExcelAuditor:
    def __init__(self, input_directory: str, output_directory: str):
        self.input_directory = Path(input_directory)
        self.output_directory = Path(output_directory)
        self.create_output_directory()
        
    def create_output_directory(self):
        #Crear directorio de salida si no existe
        if not self.output_directory.exists():
            self.output_directory.mkdir(parents=True)
    
    def get_excel_files(self) -> List[Path]:
        #Obtener lista de archivos Excel en el directorio de entrada
        return list(self.input_directory.glob("*.xlsx"))
    
    def analyze_file(self, file_path: Path) -> Dict:
        #Analizar un archivo Excel individual
        analysis = {'filename': file_path.name, 'sheets': {}}
        
        for sheet_name in pd.ExcelFile(file_path).sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            analysis['sheets'][sheet_name] = {
                'total_rows': len(df),
                'total_columns': len(df.columns),
                'missing_values': df.isnull().sum().to_dict(),
                'duplicates': df.duplicated().sum(),
                'numeric_columns_stats': {},
                'categorical_columns_stats': {}
            }
            # Análisis por tipo de columna
            for column in df.columns:
                if np.issubdtype(df[column].dtype, np.number):
                    analysis['sheets'][sheet_name]['numeric_columns_stats'][column] = {
                        'mean': df[column].mean(),
                        'std': df[column].std(),
                        'min': df[column].min(),
                        'max': df[column].max()
                    }
                else:
                    analysis['sheets'][sheet_name]['categorical_columns_stats'][column] = {
                        'unique_values': df[column].nunique(),
                        'most_common': df[column].value_counts().head(3).to_dict()
                    }
            
            # Análisis de clientes
            if 'Cliente' in df.columns and 'Producto' in df.columns and 'Cantidad' in df.columns:
                analysis['sheets'][sheet_name]['client_analysis'] = self.analyze_clients(df)
        
        return analysis
    
    def analyze_clients(self, df: pd.DataFrame) -> Dict:
        client_analysis = {}
        
        # Cliente que compra más de cierto producto
        product_max_client = df.groupby(['Producto', 'Cliente'])['Cantidad'].sum().reset_index()
        product_max_client = product_max_client.loc[product_max_client.groupby('Producto')['Cantidad'].idxmax()]
        client_analysis['max_client_per_product'] = dict(zip(product_max_client['Producto'], product_max_client['Cliente']))
        
        # Cliente que más productos compra
        total_purchases = df.groupby('Cliente')['Cantidad'].sum().sort_values(ascending=False)
        client_analysis['top_purchasing_client'] = total_purchases.index[0]
        client_analysis['top_purchasing_client_amount'] = total_purchases.iloc[0]
        
        return client_analysis
    
    def generate_report(self, analysis: Dict):
        #Generar informe de análisis en formato Markdown
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = self.output_directory / f"audit_report_{timestamp}.md"
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(f"# Informe de Auditoría de Datos\n")
            f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            f.write(f"## Resumen del archivo: {analysis['filename']}\n\n")
            
            for sheet_name, sheet_data in analysis['sheets'].items():
                f.write(f"### Hoja: {sheet_name}\n")
                f.write(f"- Total de filas: {sheet_data['total_rows']}\n")
                f.write(f"- Total de columnas: {sheet_data['total_columns']}\n")
                f.write(f"- Registros duplicados: {sheet_data['duplicates']}\n\n")

                f.write("#### Análisis de valores faltantes\n")
                for col, missing in sheet_data['missing_values'].items():
                    f.write(f"- {col}: {missing} valores faltantes\n")
                
                f.write("\n#### Análisis de columnas numéricas\n")
                for col, stats in sheet_data['numeric_columns_stats'].items():
                    f.write(f"\n##### {col}\n")
                    f.write(f"- Media: {stats['mean']:.2f}\n")
                    f.write(f"- Desviación estándar: {stats['std']:.2f}\n")
                    f.write(f"- Mínimo: {stats['min']:.2f}\n")
                    f.write(f"- Máximo: {stats['max']:.2f}\n")
                
                f.write("\n#### Análisis de columnas categóricas\n")
                for col, stats in sheet_data['categorical_columns_stats'].items():
                    f.write(f"\n##### {col}\n")
                    f.write(f"- Valores únicos: {stats['unique_values']}\n")
                    f.write("- Valores más comunes:\n")
                    for val, count in stats['most_common'].items():
                        f.write(f"  * {val}: {count}\n")
                
                if 'client_analysis' in sheet_data:
                    f.write("\n#### Análisis de Clientes\n")
                    f.write("\n##### Cliente que más compra por producto\n")
                    for product, client in sheet_data['client_analysis']['max_client_per_product'].items():
                        f.write(f"- {product}: {client}\n")
                    f.write(f"\n##### Cliente que más productos compra en total\n")
                    f.write(f"- Cliente: {sheet_data['client_analysis']['top_purchasing_client']}\n")
                    f.write(f"- Cantidad total: {sheet_data['client_analysis']['top_purchasing_client_amount']}\n")
                
                f.write("\n")
                
                #compara empleados y proveedores vendor(id de cada hoja)
    
    def run_audit(self):
        #Ejecutar el proceso completo de auditoría
        excel_files = self.get_excel_files()
        if not excel_files:
            print("No se encontraron archivos Excel en el directorio especificado.")
            return
        
        for file_path in excel_files:
            print(f"Analizando {file_path.name}...")
            analysis = self.analyze_file(file_path)
            self.generate_report(analysis)
            print(f"Informe generado para {file_path.name}")

if __name__ == "__main__":
    # Ejemplo de uso
    auditor = ExcelAuditor(
        input_directory="./datos_entrada",
        output_directory="./informes_auditoria"
    )
    auditor.run_audit()