# Generador de TXT Banco Unión

Convierte extractos bancarios del Banco Unión (formato `.xls`) al formato TXT de 110 caracteres por línea requerido por el sistema OVT (Ministerio de Trabajo).

Procesa los tipos de operación **DEP** (depósito), **TEC** (transferencia electrónica) y **CRV** (cobro por ventanilla) con crédito mayor a cero.

---

## Requisitos

| | Linux | Windows |
|---|---|---|
| Python | 3.8 o superior | 3.8 o superior |
| pip | incluido con Python | incluido con Python |

---

## Instalación en Linux

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/generacion-depositos.git
cd generacion-depositos
```

### 2. Verificar versión de Python

```bash
python3 --version
```

Si el sistema tiene Python 2 como predeterminado (verificar con `python --version`), usar `pyenv`:

```bash
# Instalar pyenv si no está disponible
curl https://pyenv.run | bash

# Instalar Python 3.11
pyenv install 3.11.9
pyenv local 3.11.9
python --version  # debe mostrar Python 3.11.9
```

### 3. Crear entorno virtual

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 4. Instalar dependencias

```bash
pip install pandas xlrd lxml html5lib beautifulsoup4
```

### 5. Verificar instalación

```bash
python3 generar_txt_bu.py
```

Debe mostrar el mensaje de uso sin errores.

---

## Instalación en Windows

### 1. Instalar Python 3

Descargar desde [python.org](https://www.python.org/downloads/) y marcar la opción **"Add Python to PATH"** durante la instalación.

Verificar:

```cmd
python --version
```

### 2. Clonar el repositorio

```cmd
git clone https://github.com/tu-usuario/generacion-depositos.git
cd generacion-depositos
```

### 3. Crear entorno virtual

```cmd
python -m venv .venv
.venv\Scripts\activate
```

### 4. Instalar dependencias

```cmd
pip install pandas xlrd lxml html5lib beautifulsoup4
```

### 5. Verificar instalación

```cmd
python generar_txt_bu.py
```

---

## Uso

### Básico — nombre automático

```bash
# Linux
python3 generar_txt_bu.py "ExtractosBancarios (39).xls"

# Windows
python generar_txt_bu.py "ExtractosBancarios (39).xls"
```

El archivo TXT se genera automáticamente con el formato:

```
MTE_011719_DDMMYYYY_HHMM.txt
```

Ejemplo: `MTE_011719_16032026_1430.txt`

Si se procesan varios extractos del mismo día, la hora en el nombre evita sobreescribir archivos previos.

### Avanzado — nombre personalizado

```bash
python3 generar_txt_bu.py "ExtractosBancarios (39).xls" MTE_011719_16032026.txt
```

---

## Formatos de Excel soportados

El script detecta automáticamente el formato del archivo:

| Formato | Descripción |
|---|---|
| XLS binario | Archivo `.xls` nativo exportado desde banca en línea |
| XLS-HTML 13 cols | Archivo `.xls` con cabecera de una sola moneda |
| XLS-HTML 16 cols | Archivo `.xls` con cabecera multinivel (Moneda Origen + Bolivianos) |

---

## Salida esperada

```
📂 Leyendo: ExtractosBancarios (39).xls
   Formato            : HTML (13 cols)
   Total filas leídas : 1547
   Registros (DEP+TEC+CRV>0): 1536
✅ 1536 registros generados | 0 ignorados
📄 Archivo: /ruta/MTE_011719_16032026_1430.txt
✅ Todas las líneas miden exactamente 110 chars
```

---

## Formato TXT generado

Cada línea tiene exactamente **110 caracteres**:

| Posición | Longitud | Contenido |
|---|---|---|
| 0–17 | 18 | Monto en centavos, alineado a la izquierda |
| 18–38 | 21 | Cuenta fija `BUNCC10000006036425  ` |
| 39–46 | 8 | Fecha `YYYYMMDD` |
| 47 | 1 | Tipo `C` (crédito) |
| 48–75 | 28 | Glosa truncada, sin caracteres especiales |
| 76–86 | 11 | Número de documento, alineado a la izquierda |
| 87–109 | 23 | `21 2    ` + 15 espacios |

---

## Dependencias

```
pandas
xlrd
lxml
html5lib
beautifulsoup4
```

---

## Notas

- Los caracteres con tilde, ñ y símbolos especiales (°, ´, etc.) se convierten automáticamente a ASCII para garantizar exactamente 110 caracteres por línea.
- El sistema destino depura duplicados, por lo que es seguro subir un extracto más completo del mismo día.
- Solo se procesan operaciones con tipo `DEP`, `TEC` o `CRV` y crédito mayor a cero. Se ignoran `TFD` y otros tipos.
