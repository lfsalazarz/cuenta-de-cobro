# Instalación:
## Paso 1:
~~~
python -m venv env
~~~

## Paso 2:
### Linux/Mac
~~~
source env/bin/activate
~~~

### Windows
~~~
env\Scripts\activate.bat
~~~

## Paso 3:
~~~
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
~~~

## Paso 4:
Editar config.yml

## Paso 5 (opcional):
Agregar imagen de firma a los .docx

# Uso:
**Activar el env**

Ayuda
~~~
python document.py single --help
python document.py batch --help
~~~
## Una sola cuenta
Cuenta de cobro con la fecha actual
~~~
python document.py single
~~~
Cuenta de cobro con la fecha del archivo de configuración
~~~
python document.py single --custom-date
~~~
Cambiar el nombre del archivo generado
~~~
python document.py single --output cuenta1.docx
~~~
## Cuenta por lotes
Cuentas de cobro por lotes del año 2020
~~~
python document.py batch --year 2020
~~~
Cuentas de cobro por lotes del año 2021 en un directorio especifico
~~~
python document.py batch --year 2021 --output-dir path/to/mydir
~~~
