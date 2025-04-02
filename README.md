# A R P A

Este repositorio primer version contiene los scripts y filtros necesarios para analizar datos históricos en el archivo `src/dataHistoricaPBI.xlsx`. El periodo a evaluar se define en el archivo `src/periodoPBI.xlsx`.

## 1. Preparación

**Creación de directorios:**

Ejecute el script `set.ps1` para crear la estructura de directorios necesaria:

```powershell
.\set.ps1
```

**Ubicación de los datos:**

Asegurese de pegar los archivos `dataHistoricaPBI.xlsx` y `periodoPBI.xlsx` en la carpeta `src/`.

La estructura del repositorio debe ser la siguiente:

```
byrAnalize/
├── models/
├── tables/
├── src/
│   ├── dataHistoricaPBI.xlsx
│   └── periodoPBI.xlsx
├── favicon.png
├── .gitignore
├── README.md
├── run.ps1
└── set.ps1
```

## 2. Ejecución del análisis

Ejecute el script principal `run.ps1`. Este script instala las dependencias, crea los scripts de análisis, ejecuta el análisis y genera los resultados en la carpeta `tables/`.

```powershell
.\run.ps1
```

## 3. Resultados

Después de ejecutar `run.ps1`, se abrira el navegador con la visualización de los datos. Tambien la carpeta `tables/` contendrá los resultados del análisis en archivos Excel, organizados en las subcarpetas `cats/`, `nets/` y `trends/`.  La estructura resultante será similar a la siguiente:

```
byrAnalize/
├── models/
│   ├── passKey.py
│   ├── server.py
│   ├── cats/
│   │   ├── banks.py
│   │   ├── debts.py
│   │   ├── goods.py
│   │   ├── incomes.py
│   │   └── investments.py
│   ├── nets/
│   │   ├── assetNets.py
│   │   ├── bankNets.py
│   │   ├── debtNets.py
│   │   ├── goodNets.py
│   │   ├── incomeNets.py
│   │   ├── investNets.py
│   │   └── worthNets.py
│   └── trends/
│       ├── overTrends.py
│       ├── trends.py
│       └── filters.py
├── src/
│   ├── dataHistoricaPBI.xlsx
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
│       ├── trends.xlsx
│       └── data.json
├── static/
│   ├── style.css
│   └── script.js
├── favicon.png
├── index.html
├── .gitignore
├── README.md
├── run.ps1
└── set.ps1
```

## 4. Visualización de datos

### 4.1 Visualización en el navegador

1. Ejecute el script en la terminal:

```
python models/server.py
```

El navegador mostrará los datos.  Para filtrar los datos:

2. Utilice los botones para agregar, visualizar, reiniciar y aplicar filtros.

3. Guarde los resultados filtrados en la carpeta de descragas/ con el botón "Guardar Excel".

### 4.2 Visualización en la terminal

1. Ejecute el script en la terminal:

```
python models/trends/filters.py
```

2. El script permite filtrar los datos cargados desde `data.json` y guardar los resultados en un archivo Excel. Siga las instrucciones en la terminal para agregar, visualizar, reiniciar y aplicar filtros.

3. Guarde los resultados filtrados en un archivo Excel.


## 5. Flujo de trabajo con Git

### Actualizar el repositorio

```
git clone https://github.com/jalkn/Byranalize.git
git pull
```