## Instalación

# app_catalogojoyero

Una aplicación web desarrollada en Flask para gestionar catálogos de joyería mediante hojas de cálculo (Excel) y gestión de imágenes. Cada usuario puede crear, importar, editar y descargar su catálogo, el cual se almacena en archivos Excel y se complementa con imágenes. Además, incluye autenticación de usuarios, recuperación de contraseña mediante token enviado por correo y despliegue en Heroku.

## Características

- **Autenticación de Usuarios:**  
  Registro, inicio de sesión y cierre de sesión.  
  Inicio de sesión admite usuario (nombre) o email de forma indistinta.

- **Recuperación de Contraseña:**  
  Permite solicitar un enlace de recuperación mediante nombre o email; el enlace es enviado al correo registrado.

- **Gestión de Catálogos:**  
  Cada usuario puede tener varias hojas de cálculo (tablas) en las que se almacenan los datos de cada objeto (número, descripción, peso, valor e imágenes).

- **Importación y Creación de Hojas de Cálculo:**  
  Los usuarios pueden crear una nueva tabla vacía o importar un archivo Excel existente.

- **Gestión de Registros:**  
  Se pueden agregar, editar y eliminar registros dentro de la tabla seleccionada.

- **Descarga del Catálogo:**  
  El catálogo se exporta a un archivo Excel y, junto con las imágenes referenciadas, se comprime en un archivo ZIP para su descarga.

- **Despliegue en Heroku:**  
  La aplicación está preparada para desplegarse en Heroku utilizando Gunicorn y un Procfile.

## Estructura del Proyecto

```plaintext
app_catalogojoyero/
├── app.py                # Aplicación principal en Flask
├── Procfile              # Archivo para desplegar en Heroku (ej.: "web: gunicorn app:app")
├── requirements.txt      # Dependencias del proyecto
├── README.md             # Este archivo
├── .gitignore            # Archivos y carpetas a ignorar en Git (por ejemplo, .venv/)
├── spreadsheets/         # Carpeta para almacenar las hojas de cálculo (archivos .xlsx)
│   └── (archivos Excel)
├── imagenes_subidas/      # Carpeta para almacenar las imágenes subidas
│   └── (imágenes)
├── static/               # Archivos estáticos
│   └── styles.css        # Hojas de estilos CSS
└── templates/            # Plantillas HTML de la aplicación
    ├── register.html
    ├── login.html
    ├── forgot_password.html
    ├── reset_password_form.html
    ├── tables.html
    ├── index.html      # Catálogo (accesible en /catalog)
    └── editar.html

Instalación y Configuración
	1.	Clonar el repositorio:

git clone https://github.com/tu_usuario/app_catalogojoyero.git
cd app_catalogojoyero


	2.	Crear y activar un entorno virtual (opcional, pero recomendado):

python -m venv .venv
source .venv/bin/activate  # En Windows: .venv\Scripts\activate


	3.	Instalar las dependencias:

pip install -r requirements.txt


	4.	Configurar variables de entorno:
	•	Define SECRET_KEY en tu entorno de producción.
	•	Configura el URI de MongoDB Atlas en la variable MONGO_URI (ya está incluido en el código, pero puedes ajustarlo mediante variables de entorno si lo prefieres).
	5.	Asegúrate de que el archivo Procfile esté presente:

web: gunicorn app:app



Uso
	•	Desarrollo local:
Ejecuta la aplicación con:

python app.py

Accede en tu navegador a http://127.0.0.1:5000.

	•	Funcionalidades principales:
	•	Regístrate o inicia sesión.
	•	Accede a la sección de “Tablas” para crear o importar una hoja de cálculo.
	•	Selecciona una tabla para trabajar en el catálogo.
	•	Agrega, edita o elimina registros en el catálogo.
	•	Descarga el catálogo completo (Excel + imágenes comprimidas en ZIP).

Despliegue en Heroku
	1.	Inicia sesión en Heroku CLI:

heroku login


	2.	Crea una aplicación en Heroku (si aún no lo has hecho):

heroku create nombre-de-tu-app


	3.	Realiza el push a Heroku:

git push heroku master


	4.	Escala el dyno web:

heroku ps:scale web=1 --app nombre-de-tu-app


	5.	Verifica los logs:

heroku logs --tail --app nombre-de-tu-app



Consideraciones de Seguridad
	•	Las opciones tlsAllowInvalidCertificates y tlsAllowInvalidHostnames están activadas para pruebas en Heroku. En producción, es recomendable solucionarlas de forma que se utilicen certificados válidos.

Licencia

MIT License

---

Este README.md proporciona una descripción general del proyecto, instrucciones de instalación, uso, despliegue y detalles sobre la estructura y consideraciones de seguridad. Puedes personalizarlo según tus necesidades y agregar secciones adicionales si es necesario.# Catálogo Joyas (Múltiples Imágenes)

Este proyecto es una aplicación en __Flask__ que permite gestionar un catálogo de joyas, incluyendo __hasta tres imágenes__ por cada fila/registro. Se guarda la información en un archivo de Excel (`datos.xlsx`) y las imágenes en la carpeta `imagenes_subidas/`.

## Características

1. **Crear nuevos registros**: Se rellenan los campos de Número, Descripción, Peso, Valor y se pueden subir hasta 3 imágenes simultáneas.
2. **Editar registros existentes**: Permite modificar los datos (descripción, peso, valor) y reemplazar o añadir imágenes en cada una de las 3 columnas de imagen (si no subes nada para una columna, se conserva la anterior).
3. **Visualización de imágenes** en la tabla (vista web) con un pequeño **preview** y un enlace para abrir cada imagen en una nueva pestaña.
4. **Descarga del Excel**: Desde la web se puede descargar el archivo `datos.xlsx` que contiene los datos y en las celdas de imágenes se crean hipervínculos locales (texto “Ver Imagen X”).
5. **Persistencia**: Si el archivo `datos.xlsx` existe, se leen sus filas y se muestran en la tabla al cargar la app. Si no existe, se crea cuando agregamos el primer registro.

## Requisitos

- **Python 3.7+** (recomendado 3.10 o superior)
- **Bibliotecas**:
   - Flask
   - openpyxl
   - werkzeug (se incluye normalmente con Flask, pero se menciona por si acaso)

Puedes instalarlas con:

```bash
pip install flask openpyxl werkzeug
```

## Estructura de archivos

La estructura mínima sugerida para el proyecto es:

```ini
app_catalogojoyero/
├── app.py
├── imagenes_subidas/
│   └─ (aquí se guardarán las imágenes subidas)
├── templates/
│   ├── index.html
│   └── editar.html
└── datos.xlsx
```

- `app.py`: código principal Flask.
- `imagenes_subidas/`: carpeta donde se guardan físicamente los archivos de imagen que subas en la aplicación.
- `templates/index.html` y `templates/editar.html`: plantillas HTML para la vista principal y la edición de registros.
- `datos.xlsx`: archivo Excel donde se guardan (y de donde se leen) los datos de cada joya y sus posibles hasta 3 imágenes.

## Cómo ejecutar

1. **Ubícate** en la carpeta raíz del proyecto, por ejemplo:

```bash
cd app_catalogojoyero/
```

2. **Instala** las dependencias si no lo has hecho:

```bash
pip install flask openpyxl werkzeug
```

3. **Ejecuta** la aplicación:

```bash
python app.py
```

4. **Abre** tu navegador en la ruta:

```sh
http://127.0.0.1:5000/
```

5. **Uso básico**:

   - Para **añadir un registro**: En la pantalla principal, rellena “Número”, “Descripción”, “Peso”, “Valor” y, si quieres, sube hasta 3 imágenes. Pulsa **Agregar**.
   - Para **editar un registro**: En la tabla de registros, haz clic en **Editar** junto al número que corresponda. Podrás cambiar la descripción, peso, valor y/o subir nuevas imágenes (si no subes una, se conserva la que hubiera).
   - Para **descargar** el Excel, haz clic en **Descargar Excel (datos.xlsx)**.
   - Las imágenes subidas se almacenan en __`imagenes_subidas/`__ y se referencian en el Excel como hipervínculos.

## Funcionamiento interno

- **Al cargar** la página principal (`GET /`):

   - Si existe `datos.xlsx`, se lee con `openpyxl` (modo `read_only=False` para poder extraer hipervínculos) y se vuelcan las filas en una lista de diccionarios.
   - Cada fila puede tener hasta 3 rutas de imágenes (hipervínculos en columnas E, F, G).
   - Se renderiza la plantilla `index.html` mostrando la tabla y el formulario para agregar nuevo registro.

- **Al crear** un nuevo registro (`POST /`):

   - Se recogen los datos (número, descripción, peso, valor).
   - Se sube un __array__ de archivos (`request.files.getlist("imagenes")`) y se guardan, cada uno en la carpeta `imagenes_subidas/`.
   - Se abre (o crea) `datos.xlsx` y se escribe la nueva fila en la parte final. Las columnas E, F, G guardan el texto “Ver Imagen X” con un hipervínculo a la ruta local de la imagen.

- **Al editar** (`GET /editar/<numero>`):

   - Se busca la fila en la primera columna (Número) que coincida con `<numero>`. Se cargan sus datos y se pre-rellena el formulario de edición.

- **Al enviar** la edición (`POST /editar/<numero>`):

   - Se vuelven a escribir los nuevos valores en la misma fila.
   - Si subes imágenes, se sustituyen los hipervínculos en la columna E, F y/o G de esa fila por los nuevos.
   - Si no subes nada para una de las columnas, se conserva la imagen anterior.

## Personalizar el número de imágenes

En el ejemplo hemos habilitado **3** imágenes (columnas E, F, G). Si prefieres **más o menos**:

- Cambia las columnas correspondientes (por ejemplo, 2 imágenes => E y F)
- Ajusta el `for` en el código que recorre `[5, 6, 7]` o la parte de `row[4], row[5], row[6]`.
- Ajusta el formulario HTML para permitir `multiple` archivos, pero solo el número de imágenes que desees.

## Licencia y usos

Puedes usar este proyecto como base para tu catálogo, adaptarlo, agregar validaciones adicionales, sistemas de usuarios o cualquier otra funcionalidad que requieras. No existe una licencia específica en este repositorio, simplemente es un ejemplo didáctico que puedes personalizar.

---

¡Y con esto tendrás tu catálogo de joyas en Flask, guardando y mostrando **múltiples imágenes** por fila en tu Excel!
