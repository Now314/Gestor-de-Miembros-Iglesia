"""Gestor de formularios de miembros IDEA"""

import os
import sys
import ctypes
import shutil
from datetime import date
from pathlib import Path
from fpdf import FPDF
from PIL import Image, ImageDraw
from PIL.Image import Resampling
import pandas as pd

# Rutas Principales
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent
RUTA_DATA = os.path.join(BASE_DIR, "data", "BaseDatos.xlsx")
RUTA_BACKUP = os.path.join(BASE_DIR, "data", "Backup.xlsx")
RUTA_PLANTILLA = os.path.join(BASE_DIR, "pdf", "Formulario_IDEA.png")
RUTA_FOTOS = os.path.join(BASE_DIR, "assets", "fotos")
RUTA_FIRMAS = os.path.join(BASE_DIR, "assets", "firmas")

def main():
    """Inicialización del programa."""
    print("")
    print("¡Bienvenido!")
    print("")
    print("Fieles a Su Nombre, firmes en Su Palabra.")
    print("")
    print("-"*30)

    while True:

        # Lectura del Excel.
        excel_leido = leer_excel(RUTA_DATA)
        # Menú de opciones para el usuario.
        print("\nMenú de opciones:")
        print("1. Mostrar Lista de Miembros IDEA.")
        print("2. Buscar Miembro IDEA.")
        print("3. Generar Formulario de Membresia IDEA.")
        print("4. Agregar Nuevo Miembro IDEA.")
        print("5. Eliminar Miembro IDEA.")
        print("6. Editar Registro Miembro IDEA.")
        print("7. Salir.")

        opcion_elegida = input("Seleccione una opción del menú (1-7): ").strip()
        if opcion_elegida not in ["1", "2", "3", "4", "5", "6","7"]:
            print("Opción no válida. Por favor, seleccione una opción del 1 al 7.")
            continue

        if opcion_elegida == "1":

            print("\nLista de Miembros IDEA:\n")
            for _, fila in excel_leido.iterrows():
                miembro=diccionario_datos(fila)
                mostrar_datos(diccionario=miembro)
            print("Fin de la lista de miembros.")
            print("-"*30)
            print("")

        elif opcion_elegida == "2":

            miembro_selecionado = buscar_miembro(miembros_idea=excel_leido)
            if miembro_selecionado is None:
                continue
            datos_miembro = diccionario_datos(miembro=miembro_selecionado)
            mostrar_datos(datos_miembro)
            print("-"*30)
            print("")

        elif opcion_elegida == "3":

            miembro_selecionado = buscar_miembro(miembros_idea=excel_leido)
            if miembro_selecionado is None:
                continue
            datos_miembro = diccionario_datos(miembro=miembro_selecionado)
            mostrar_datos(diccionario=datos_miembro)
            print("\nGenerando formulario...")
            pdf = generar_formulario_vacio(plantilla=RUTA_PLANTILLA)
            pdf_con_imagenes = rellenar_foto_y_firmas(pdf=pdf, diccionario=datos_miembro)
            pdf_formulario_miembro = rellenar_texto(dic=datos_miembro, pdf=pdf_con_imagenes)
            guardar_formulario_miembro(pdf_completo=pdf_formulario_miembro, dic=datos_miembro)
            print("-"*30)
            print("")

        elif opcion_elegida == "4":
            agregar_miembro(excel_leido)
            print("-"*30)
            print("")
        elif opcion_elegida == "5":
            eliminar_miembro(excel_leido)
            print("-"*30)
            print("")
        elif opcion_elegida == "6":
            editar_miembro(excel_leido)
            print("-"*30)
            print("")
        elif opcion_elegida == "7":
            print("")
            print("¡Hasta luego!")
            print("")
            break

def leer_excel(path_excel):
    """
    Abrir y leer datos ubicados en el excel de miembros IDEA.
    """
    try:
        data = pd.read_excel(path_excel, dtype=str)
        return data
    except FileNotFoundError as exc:
        raise FileNotFoundError("No se encontró la base de datos.") from exc
    except (pd.errors.ParserError, ValueError) as e:
        print("\nERROR leyendo Excel:")
        print(e)
        return pd.DataFrame()

def buscar_miembro(miembros_idea):
    """
    Buscar un miembro por su nombre, mostrando coincidencias y permitiendo seleccionar una.
    """
    # Buscar coincidencias
    if "Nombre" not in miembros_idea.columns:
        raise ValueError("La columna 'Nombre' no existe.")

    coincidencias = pd.DataFrame()

    while True:

        buscar_nombre = input("\nIngrese el nombre o el apellido del miembro: ").strip()

        if buscar_nombre == "":
            print("Debe escribir un nombre.")
            continue
        elif buscar_nombre.lower() == "salir":
            print("Volviendo al menú principal.")
            break

        coincidencias = miembros_idea[
            miembros_idea["Nombre"].str.contains(
                buscar_nombre,
                case=False,
                na=False
            )
            |
            miembros_idea["Apellido"].str.contains(
                buscar_nombre,
                case=False,
                na=False
            )
        ]
        if not coincidencias.empty:

            # Mostrar coincidencias
            print("\nCoincidencias encontradas:\n")
            #Resetear índices para mostrar correctamente.
            lista_de_coincidencias = coincidencias.reset_index(drop=True)

            for i, fila in lista_de_coincidencias.iterrows():

                nombre = fila.get("Nombre", "")
                apellido = fila.get("Apellido", "")
                print(f"[{i}] {nombre} {apellido}")

            # Seleccionar un indice y guardar la selección.
            try:
                indice = int(input("\nSeleccione el índice: "))
            except ValueError:
                print("Debe ingresar un número.")
                continue

            if indice < 0 or indice >= len(lista_de_coincidencias):
                print("Índice inválido.")
                continue

            miembro = lista_de_coincidencias.iloc[indice]
            return miembro

        print("No se encontró ningún miembro.")

def extraer_dato(columna, miembro):
    """
    Obtiene el valor de cualquier columna especificada de los datos del miembro 
    seleccionado para poder guardarlo en una variable.
    """
    dato = miembro.get(columna, "")
    if pd.isna(dato):
        return ""

    return str(dato).strip()

def diccionario_datos(miembro):
    """
    Crea un diccionario con los datos del miembro seleccionado, utilizando la función extraer_dato.
    """
    dic = {
        "Fecha Actual": date.today().strftime("%d/%m/%Y"),
        "Nombre": extraer_dato("Nombre", miembro),
        "Apellido": extraer_dato("Apellido", miembro),
        "Dirección": extraer_dato("Dirección", miembro),
        "Teléfono": extraer_dato("Teléfono", miembro),
        "Email": extraer_dato("Email", miembro),
        "Fecha de Nacimiento": extraer_dato("Fecha_Nacimiento", miembro),
        "Lugar de Nacimiento": extraer_dato("Lugar_Nacimiento", miembro),
        "Estado Civil": extraer_dato("Estado_Civil", miembro),
        "Fecha de Bautismo": extraer_dato("Fecha_Bautismo", miembro),
        "Bautizado por": extraer_dato("Bautizado_por", miembro),
        "Miembro por": extraer_dato("Miembro_por", miembro),
        "Foto": extraer_dato("Foto", miembro),
        "Firma del Miembro": extraer_dato("Firma_Miembro", miembro),
        "Firma del Pastor": extraer_dato("Firma_Pastor", miembro),
        "Elaborado por": extraer_dato("Elaborado_por", miembro),
    }

    return dic

def mostrar_datos(diccionario):
    """
    Muestra los datos extraídos del miembro seleccionado.
    """
    print("\nDatos encontrados:\n")
    for clave, valor_dato in diccionario.items():
        print(f"{clave}: {valor_dato}")

def generar_formulario_vacio(plantilla):
    """
    Genera el formulario de membresía en PDF utilizando los datos del diccionario.
    """
    # Crear PDF.
    formulario_vacio = FPDF(unit="mm", format="letter")
    formulario_vacio.add_page()

    if not os.path.exists(plantilla):
        raise ValueError("No se encontró la plantilla del formulario.")

    # Fondo carta
    formulario_vacio.image(plantilla, x=0, y=0, w=216, h=279)
    return formulario_vacio

def recortar_imagen(ruta_original, ruta_destino, ancho_mm, alto_mm, dpi=300):
    """
    Recortar y redimensionar la imagen manteniendo la relación de aspecto, 
    y aplicar bordes redondeados.
    """
    # mm -> px
    ancho_px = int((ancho_mm / 25.4) * dpi)
    alto_px = int((alto_mm / 25.4) * dpi)

    try:
        img = Image.open(ruta_original)
    except (OSError, ValueError) as e:
        raise ValueError(
            f"Error abriendo imagen: {e}"
        ) from e

    ancho_original, alto_original = img.size
    ratio_original = ancho_original / alto_original
    ratio_destino = ancho_px / alto_px

    # RECORTE HORIZONTAL
    if ratio_original > ratio_destino:

        nuevo_ancho = int(alto_original * ratio_destino)
        izquierda = (ancho_original - nuevo_ancho) // 2
        derecha = izquierda + nuevo_ancho
        img = img.crop((izquierda, 0, derecha, alto_original))

    # RECORTE VERTICAL
    else:

        nuevo_alto = int(ancho_original / ratio_destino)
        arriba = (alto_original - nuevo_alto) // 2
        abajo = arriba + nuevo_alto
        img = img.crop((0, arriba, ancho_original, abajo))

    # REDIMENSIONAR
    img = img.resize((ancho_px, alto_px), Resampling.LANCZOS)

    # BORDES REDONDEADOS
    radio = min(ancho_px, alto_px) // 12
    mascara = Image.new("L", (ancho_px, alto_px), 0)
    draw = ImageDraw.Draw(mascara)
    draw.rounded_rectangle((0, 0, ancho_px, alto_px), radius=radio, fill=255)

    # Convertir a RGBA
    img = img.convert("RGBA")
    # Aplicar transparencia
    img.putalpha(mascara)
    # Guardar PNG transparente
    img.save(ruta_destino, format="PNG")

def rellenar_foto_y_firmas(pdf, diccionario):
    """
    Colocar en el formulario generado la foto del miembro y las firmas del miembro y pastor,
    si existen, utilizando las rutas especificadas en el diccionario.
    """
    # FUENTE
    pdf.set_font("Helvetica", size=11)

    # FOTO
    if diccionario["Foto"] != "":

        ruta_foto = os.path.join(RUTA_FOTOS, diccionario["Foto"])
        if os.path.exists(ruta_foto):

            ruta_temp = os.path.join(BASE_DIR, "temp_foto.png")

            try:
                recortar_imagen(ruta_foto, ruta_temp, 43, 50)
                pdf.image(ruta_temp, x=140, y=148, w=43, h=55)
            except (OSError, ValueError) as e:
                print(f"Error cargando foto: {e}")
            finally:
                # Eliminar temporal
                if os.path.exists(ruta_temp):
                    os.remove(ruta_temp)

        else:
            pdf.rect(140, 148, 43, 55)
            pdf.text(154, 176, "Sin Foto")

    # FIRMA MIEMBRO
    if diccionario["Firma del Miembro"] != "":

        ruta_firma = os.path.join(RUTA_FIRMAS, diccionario["Firma del Miembro"])
        if os.path.exists(ruta_firma):
            pdf.image(ruta_firma, x=48, y=206, w=43, h=11)

    # FIRMA PASTOR
    if diccionario["Firma del Pastor"] != "":

        ruta_firma_pastor = os.path.join(RUTA_FIRMAS, diccionario["Firma del Pastor"])
        if os.path.exists(ruta_firma_pastor):
            pdf.image(ruta_firma_pastor, x=130, y=206, w=45, h=11)

    return pdf

def rellenar_texto(dic, pdf):

    """Dibuja el texto en las coordenadas(x, y) especificadas."""
    def pegar_texto(x, y, data):
        pdf.text(x, y, str(data))

    # Fecha
    pegar_texto(46, 87, dic["Fecha Actual"])
    # Nombre y apellido
    pegar_texto(50, 102, dic["Nombre"])
    pegar_texto(140, 102, dic["Apellido"])
    # Dirección
    pegar_texto(50, 113, dic["Dirección"])
    # Teléfono y email
    pegar_texto(50, 123.5, dic["Teléfono"])
    pegar_texto(126, 123.5, dic["Email"])
    # Fecha nacimiento y lugar
    pegar_texto(70, 135, dic["Fecha de Nacimiento"])
    pegar_texto(135, 135, dic["Lugar de Nacimiento"])
    # Estado civil
    pegar_texto(55, 152.5, dic["Estado Civil"])
    # Fecha bautismo
    pegar_texto(68, 163.5, dic["Fecha de Bautismo"])
    # Bautizado por
    pegar_texto(65, 174.5, dic["Bautizado por"])
    # Recibido como miembro por
    pegar_texto(84, 185, dic["Miembro por"])
    # Elaborado por
    pegar_texto(125, 234, dic["Elaborado por"])

    return pdf

def guardar_formulario_miembro(pdf_completo, dic):
    """
    Guardar el formulario generado en PDF, utilizando el nombre del miembro para el nombre 
    del archivo. El PDF se guardará en el escritorio del usuario.
    """
    def limpiar_nombre(nombre):
        """Limpia el nombre para usarlo en el nombre del archivo."""
        caracteres = r'\/:*?"<>|'
        for c in caracteres:
            nombre = nombre.replace(c, "_")
        return nombre

    csidl_desktopdirectory = 0x10
    buf = ctypes.create_unicode_buffer(260)

    ctypes.windll.shell32.SHGetFolderPathW(
        None,
        csidl_desktopdirectory,
        None,
        0,
        buf
    )

    desktop = Path(buf.value)

    # Crear nombre del PDF y exportar a escritorio
    titulo_pdf = limpiar_nombre(f"Formulario_{dic['Nombre']}_{dic['Apellido']}.pdf")
    ruta_salida = desktop / titulo_pdf

    try:
        pdf_completo.output(ruta_salida)
        print("\n¡Formulario generado correctamente!")
        print(ruta_salida)
    except (OSError, ValueError) as e:
        print("\nError guardando PDF:")
        print(e)

def agregar_miembro(excel):
    """
    Añadir un miembro a la lista de miembros IDEA.
    """
    nuevo_miembro = {
        "Id":"",
        "Fecha": date.today().strftime("%d/%m/%Y"),
        "Nombre": "",
        "Apellido": "",
        "Dirección": "",
        "Teléfono": "",
        "Email": "",
        "Fecha_Nacimiento": "",
        "Lugar_Nacimiento": "",
        "Estado_Civil": "",
        "Fecha_Bautismo": "",
        "Bautizado_por":"",
        "Miembro_por": "",
        "Foto": "",
        "Firma_Miembro": "",
        "Firma_Pastor": "",
        "Elaborado_por": "",
        "Notas": ""
    }
    if excel.empty:
        nuevo_miembro["Id"] = "1"
    else:
        ultimo_id = pd.to_numeric(
            excel["Id"],
            errors="coerce"
        ).max()
        nuevo_miembro["Id"] = str(int(ultimo_id) + 1)

    nuevo_miembro_copy = nuevo_miembro.copy()

    for key in nuevo_miembro_copy:

        if key in ("Id", "Fecha"):
            continue

        if key in ("Firma_Miembro", "Firma_Pastor"):
            print(f"{key}: suba la imagen directamente y actualice el Excel de forma manual.")
            continue

        if key == "Foto":
            nuevo_miembro["Foto"] = f"{nuevo_miembro['Nombre']}_{nuevo_miembro['Apellido']}.jpg"
            print(f"Foto asignada: {nuevo_miembro['Foto']}")
            print("Suba la imagen a la carpeta Fotos.")
            continue

        nuevo_miembro[key] = input(f"Por favor agregue {key}: ").strip()

    try:
        excel.loc[len(excel)] = nuevo_miembro
        guardar_cambios(excel)
        print("\nMiembro agregado correctamente.")
    except (KeyError, ValueError) as e:
        print("\nERROR añadiendo información al Excel:")
        print(e)

def eliminar_miembro(excel):
    """
    Eliminar el registro de un miembro del Excel.
    """

    confirmacion = input(
        "¿Confirma que quiere eliminar un registro del programa? (s/n): "
    ).strip().lower()

    if confirmacion != "s":
        print("Operación cancelada.")
        return

    miembro_para_eliminar = buscar_miembro(excel)

    if miembro_para_eliminar is None:
        return

    id_miembro = miembro_para_eliminar["Id"]

    # Eliminar miembro
    excel = excel[excel["Id"] != id_miembro]

    # Reiniciar índices de pandas
    excel = excel.reset_index(drop=True)

    # Renumerar IDs
    excel["Id"] = range(1, len(excel) + 1)

    # Convertir a texto para mantener consistencia
    excel["Id"] = excel["Id"].astype(str)

    guardar_cambios(excel)

    print("Miembro eliminado correctamente.")

def editar_miembro(excel):
    """
    Editar un registro ya existente en el Excel.
    """

    miembro_para_editar = buscar_miembro(excel)
    if miembro_para_editar is None:
        return

    id_miembro = miembro_para_editar["Id"]
    indice = excel[excel["Id"] == id_miembro].index[0]

    for columna in excel.columns:

        if columna == "Id":
            continue

        valor_actual = excel.loc[indice, columna]

        print(f"\n{columna}: {valor_actual}")

        confirmacion = input(
            "¿Desea editar esta información? (s/n): "
        ).strip().lower()

        if confirmacion != "s":
            continue

        nuevo_dato = input(
            "Agregue la nueva información aquí --> "
        ).strip()

        excel.loc[indice, columna] = nuevo_dato

        print("Dato guardado correctamente.")

    guardar_cambios(excel)

    print("Miembro editado correctamente.")

def guardar_cambios(cambios_realizados):
    """
    Guardar cualquier cambio que se realice en el Excel desde el programa.
    """

    try:

        if os.path.exists(RUTA_DATA):
            shutil.copy2(RUTA_DATA, RUTA_BACKUP)

        cambios_realizados.to_excel(RUTA_DATA, index=False)

    except OSError as e:
        print("\nERROR guardando cambios:")
        print(e)

if __name__ == "__main__":
    main()
