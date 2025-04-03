# A R P A

Este repositorio en su primera version contiene los scripts y filtros necesarios para analizar datos históricos en el archivo dataHistoricaPBI.xlsx.

## 1. Preparación

Ejecute el script principal `run.ps1`. Este script instala las dependencias, crea los scripts de análisis, se abrira el navegador la previsualizacion del entorno de análisis.

```powershell
.\run.ps1
```

## 2. Ejecución del análisis y Visualización de datos

1. El script en la terminal

```
python app.py
```
2. Cargue el archivo dataHistoricaPBI.xlsx dando click en "Subir archivo Excel", ingrese la contraseña y genera la tabla dando click "Analizar Archivo". El navegador mostrará los datos.  

3. Para filtrar los datos:

- Utilice los botones para agregar, visualizar, reiniciar y aplicar filtros.

- Guarde los resultados filtrados en la carpeta de descragas/ con el botón "Guardar Excel".

- Dando click en detalles puede visualizar todos los datos por fila y guardar en excel.


## 3. Resultados

Después "Analizar Archivo",  Tambien la carpeta `tables/` contendrá los resultados del análisis en archivos Excel, organizados en las subcarpetas `cats/`, `nets/` y `trends/`.  La estructura resultante será similar a la siguiente:

```
arpa/
├── models/
│   ├── passKey.py
│   ├── fk1Data.py
│   ├── server.py
│   ├── cats.py
│   ├── nets.py
│   ├── period.py
│   └── trends.py
├── src/
│   ├── dataHistoricaPBI.xlsx
│   ├── data.json
│   ├── fk1Data.json
│   └── periodoPBI.xlsx
├── tables/
│   ├── cats/
│   │   ├── banks.xlsx
│   │   ├── debts.xlsx
│   │   ├── goods.xlsx
│   │   ├── incomes.xlsx
│   │   └── investments.xlsx
│   ├── nets/
│   │   ├── assetNets.xlsx
│   │   ├── bankNets.xlsx
│   │   ├── debtNets.xlsx
│   │   ├── goodNets.xlsx
│   │   ├── incomeNets.xlsx
│   │   ├── investNets.xlsx
│   │   └── worthNets.xlsx
│   └── trends/
│       ├── overTrends.xlsx
│       └── trends.xlsx
├── static/
│   ├── style.css
│   └── script.js
├── favicon.png
├── index.html
├── .gitignore
├── README.md
├── app.py
└── run.ps1
```