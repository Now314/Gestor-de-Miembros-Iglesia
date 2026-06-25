# Miembros IDEA

Aplicación de escritorio desarrollada en **Python** para la gestión y automatización de miembros de la Iglesia IDEA.

## Características

* Gestión de la lista de miembros.
* Registro y actualización de información de los miembros.
* Generación automática de formularios.
* Exportación de documentos en formato PDF.
* Almacenamiento de fotografías y firmas de los miembros.

---

## Tecnologías

* Python 3.14+
* NumPy 2.5
* Pandas 3
* Openpyxl
* FPDF
* Pillow

---

## Instalación

### Clonar el repositorio

```bash
git clone https://github.com/Now314/Miembros_IDEA.git
cd Miembros_IDEA
```

### Crear un entorno virtual

```bash
python -m venv venv
```

### Activar el entorno virtual

**Windows**

```bash
venv\Scripts\activate
```

**Linux / macOS**

```bash
source venv/bin/activate
```

### Instalar las dependencias

```bash
pip install -r requirements.txt
```

---

## Ejecución

```bash
python miembros_idea.py
```

---

## Estructura del proyecto

```text
Miembros_IDEA/
│
├── docs/
├── pdf/
├── miembros_idea.py
├── requirements.txt
├── README.md
└── .gitignore
```

---

## Carpetas excluidas del repositorio

### `data/`

Esta carpeta **no se incluye** en el repositorio porque contiene información privada de la iglesia.

Debe contener un archivo de Excel con las siguientes columnas:

| Columna          |
| ---------------- |
| Id               |
| Fecha            |
| Nombre           |
| Apellido         |
| Dirección        |
| Teléfono         |
| Email            |
| Fecha_Nacimiento |
| Lugar_Nacimiento |
| Estado_Civil     |
| Fecha_Bautismo   |
| Bautizado_por    |
| Miembro_por      |
| Foto             |
| Firma_Miembro    |
| Firma_Pastor     |
| Elaborado_por    |
| Notas            |

### `assets/`

Esta carpeta tampoco se incluye en el repositorio.

Debe contener la siguiente estructura:

```text
assets/
├── firmas/
└── fotos/
```

---

## Generar el ejecutable

```bash
pyinstaller miembros_idea.py
```

---

## Autor

**Santiago Sánchez**

📧 [manuelssm0704@proton.me](mailto:manuelssm0704@proton.me)

GitHub: https://github.com/Now314

---

## Licencia

Este proyecto se distribuye únicamente con fines educativos y de uso interno.
