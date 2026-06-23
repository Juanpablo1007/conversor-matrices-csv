# Conversor de matrices de errores a CSV

Aplicacion en Python para convertir matrices de errores en Excel (`.xlsx` o `.xlsm`) a archivos CSV con un formato fijo.

El proyecto incluye:

- Una interfaz grafica para seleccionar matrices y generar CSV sin escribir comandos.
- Un modo por terminal para automatizar conversiones.
- Una version normal del CSV.
- Una version extendida del CSV con la columna `stage_type`.

## Caracteristicas

- Lee archivos Excel con `pandas`.
- Genera CSV en codificacion `utf-8-sig`, compatible con Excel en Windows.
- Escribe todos los campos entre comillas dobles.
- Mantiene los JSON de `user_message` y `message_description` con comillas escapadas usando `\"`.
- Convierte saltos de linea internos en espacios.
- Convierte valores vacios o `NaN` a texto vacio.
- Ignora columnas extra del Excel.
- Tolera nombres de columnas con mayusculas, minusculas, espacios o guiones bajos.
- Detecta automaticamente la hoja real de datos si el Excel trae hojas de plantilla o referencia.
- Completa `itc_homologated_code` a 3 digitos cuando es numerico.
- Permite generar una version extendida con `stage_type=PILOTO`.
- En la version extendida, cambia `response_type=ALERT` por `response_type=WARNING`.

## Requisitos

- Python 3.10 o superior.
- Windows, macOS o Linux.
- Dependencias listadas en `requirements.txt`.

## Instalacion

Descarga o clona este proyecto y entra a la carpeta:

```bash
cd nombre-del-proyecto
```

Instala las dependencias:

```bash
pip install -r requirements.txt
```

En Windows, si `pip` no funciona, prueba:

```powershell
py -m pip install -r requirements.txt
```

## Uso con interfaz grafica

Ejecuta:

```bash
python excel_a_csv_gui.py
```

En Windows, si `python` no funciona:

```powershell
py excel_a_csv_gui.py
```

Pasos:

1. Haz clic en **Agregar matrices**.
2. Selecciona uno o varios archivos Excel.
3. Opcionalmente, elige una carpeta de salida.
4. Haz clic en **CSV normal** o **CSV extendido**.

Si no eliges carpeta de salida, el CSV se guarda junto al Excel original.

Ejemplo:

```text
Matriz de errores.xlsx
Matriz de errores.csv
Matriz de errores_extendido.csv
```

## Uso por terminal

Generar un CSV normal junto al Excel:

```bash
python excel_a_csv.py "ruta/al/archivo.xlsx"
```

Elegir una ruta de salida:

```bash
python excel_a_csv.py "ruta/al/archivo.xlsx" "ruta/salida.csv"
```

Generar la version extendida:

```bash
python excel_a_csv.py "ruta/al/archivo.xlsx" --extendido
```

Forzar una hoja especifica:

```bash
python excel_a_csv.py "ruta/al/archivo.xlsx" "ruta/salida.csv" --hoja "Nombre de la hoja"
```

## Columnas esperadas en el Excel

El Excel puede traer las columnas en cualquier orden. El programa acepta variantes con espacios, guiones bajos y mayusculas/minusculas.

Por ejemplo, estas formas son equivalentes:

```text
technicalCode
Technical Code
technical_code
TECHNICALCODE
```

Columnas reconocidas:

| Columna esperada | Descripcion |
| --- | --- |
| `technicalCode` | Codigo tecnico del error. |
| `channel` | Canal. |
| `origin` | Origen o codigo del servicio. |
| `idLanguage` | Idioma. |
| `technicalException` | Excepcion tecnica. |
| `userMessage` | Mensaje que se muestra al usuario. Puede contener JSON en texto plano. |
| `messageDescription` | Descripcion del mensaje. Puede contener texto largo. |
| `type` | Tipo. |
| `messageType` | Tipo de mensaje. |
| `transactionCode` | Codigo de transaccion. |
| `itcHomologatedCode` | Codigo homologado ITC. |
| `responseType` | Tipo de respuesta. |

Si falta una columna, el CSV se genera igual y esa columna queda vacia.

Si `itcHomologatedCode` viene mal nombrada entre `transactionCode` y `responseType`, el programa intenta tomarla automaticamente como respaldo.

## Formato del CSV normal

Columnas generadas, en este orden:

```text
technical_code,channel,origin,id_language,technical_exception,user_message,message_description,type,message_type,transaction_code,itc_homologated_code,response_type,status,reason
```

Reglas:

- `status` se crea vacia.
- `reason` se crea vacia.
- Todos los valores van entre comillas dobles.
- `itc_homologated_code` se rellena a 3 digitos si es numerico:
  - `1` -> `001`
  - `13` -> `013`
  - `21` -> `021`

## Formato del CSV extendido

Columnas generadas, en este orden:

```text
technical_code,channel,origin,id_language,technical_exception,user_message,message_description,type,message_type,transaction_code,itc_homologated_code,response_type,stage_type,status,reason
```

Reglas adicionales:

- `stage_type` siempre queda con valor `PILOTO`.
- `stage_type` queda antes de `status` y `reason`.
- Si `response_type` es `ALERT`, en el extendido se cambia a `WARNING`.
- El CSV normal no cambia `ALERT`.

## Archivos del proyecto

| Archivo | Uso |
| --- | --- |
| `excel_a_csv.py` | Conversor principal y uso por terminal. |
| `excel_a_csv_gui.py` | Interfaz grafica para convertir matrices. |
| `requirements.txt` | Dependencias del proyecto. |
| `README.md` | Guia de uso. |

## Problemas comunes

### `python` no se reconoce como comando

En Windows, prueba con:

```powershell
py excel_a_csv_gui.py
```

### `pip` no se reconoce como comando

En Windows, prueba con:

```powershell
py -m pip install -r requirements.txt
```

### Falta `pandas` u `openpyxl`

Instala las dependencias:

```bash
pip install -r requirements.txt
```

O en Windows:

```powershell
py -m pip install -r requirements.txt
```

### El CSV se genero en una ruta distinta

Si usas la interfaz y no eliges carpeta de salida, el CSV queda junto al Excel original.

Si usas terminal y no pasas una ruta de salida, el CSV queda junto al Excel original.

## Notas

- El programa no modifica el Excel original.
- Las columnas extra del Excel se ignoran.
- Los archivos CSV se pueden abrir en Excel, editores de texto o herramientas de carga masiva.
