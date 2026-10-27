
import io
import re
import zipfile
import unicodedata
import datetime as dt
import pandas as pd
import streamlit as st
import holidays
from docx import Document
from pypdf import PdfReader

st.title("Revisor de cuentas de cobro | Demostración")
st.warning(
    "Solo para pruebas con datos ficticios. "
    "No cargue documentos institucionales ni datos personales reales."
)

def normalizar(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return "".join(
        c for c in texto if not unicodedata.combining(c)
    ).upper().strip()

@st.cache_data
def base_demo():
    return pd.DataFrame([
        {
            "NOMBRE": "ANA MARIA PRUEBA",
            "CEDULA": "999000001",
            "CONTRATO": "DEF-DEMO-001",
            "REGIONAL": "ANTIOQUIA",
            "PROYECTO": "EXIGIBILIDAD",
            "HONORARIOS": 5900000
        },
        {
            "NOMBRE": "LAURA PATRICIA EJEMPLO",
            "CEDULA": "999000002",
            "CONTRATO": "DEF-DEMO-002",
            "REGIONAL": "QUINDIO",
            "PROYECTO": "VICTIMAS",
            "HONORARIOS": 5900000
        }
    ])

def documentos(archivos):
    salida = []
    for archivo in archivos:
        contenido = archivo.getvalue()
        if archivo.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(contenido)) as z:
                    for info in z.infolist():
                        nombre = info.filename
                        if (
                            info.is_dir()
                            or ".." in nombre.split("/")
                            or nombre.startswith("/")
                            or info.file_size > 20000000
                        ):
                            continue
                        if nombre.lower().endswith(
                            (".pdf", ".docx", ".xlsx", ".xlsm")
                        ):
                            salida.append((nombre, z.read(info)))
            except (zipfile.BadZipFile, ValueError):
                st.error("No se pudo abrir uno de los ZIP.")
        else:
            salida.append((archivo.name, contenido))
    return salida

def extraer(nombre, contenido):
    try:
        extension = nombre.lower().rsplit(".", 1)[-1]
        if extension == "pdf":
            return "\n".join(
                p.extract_text() or ""
                for p in PdfReader(io.BytesIO(contenido)).pages
            )
        if extension == "docx":
            d = Document(io.BytesIO(contenido))
            partes = [p.text for p in d.paragraphs]
            partes += [
                c.text
                for t in d.tables
                for fila in t.rows
                for c in fila.cells
            ]
            return "\n".join(partes)
        if extension in ("xlsx", "xlsm"):
            hojas = pd.read_excel(
                io.BytesIO(contenido),
                sheet_name=None,
                header=None,
                dtype=str
            )
            return "\n".join(
                str(v)
                for hoja in hojas.values()
                for v in hoja.fillna("").to_numpy().flatten()
                if str(v).strip()
            )
    except Exception:
        return ""
    return ""

def identificar(textos, base):
    contenido = normalizar("\n".join(textos.values()))
    resultados = []

    for _, r in base.iterrows():
        pruebas = []
        if normalizar(r["CONTRATO"]) in contenido:
            pruebas.append("contrato")
        if re.search(
            r"(?<!\d)" + re.escape(str(r["CEDULA"])) + r"(?!\d)",
            contenido
        ):
            pruebas.append("cédula")
        if normalizar(r["NOMBRE"]) in contenido:
            pruebas.append("nombre")
        if pruebas:
            resultados.append((len(pruebas), r, pruebas))

    resultados.sort(key=lambda x: x[0], reverse=True)

    if not resultados:
        return None, "No se encontró coincidencia con la base ficticia."

    mejor = resultados[0]
    if mejor[0] < 2 or (
        len(resultados) > 1 and resultados[1][0] == mejor[0]
    ):
        return None, (
            "Identificación no concluyente. "
            "Se requieren dos datos coincidentes y un resultado único."
        )

    return mejor[1], (
        "Coincidencia preliminar mediante "
        + ", ".join(mejor[2]) + "."
    )

st.success(
    "Base ficticia incorporada automáticamente. "
    "No necesitas subir un Excel."
)

with st.expander("Ver contratistas ficticias para la prueba"):
    st.dataframe(
        base_demo(),
        hide_index=True,
        use_container_width=True
    )

periodo = st.date_input(
    "Mes de la cuenta",
    value=dt.date(2026, 10, 1)
)

archivos = st.file_uploader(
    "Cargar F62, F68, planilla, matrices o ZIP de demostración",
    type=["pdf", "docx", "xlsx", "xlsm", "zip"],
    accept_multiple_files=True
)

if st.button(
    "Revisar cuenta automáticamente",
    type="primary",
    use_container_width=True
):
    if not archivos:
        st.warning("Primero carga documentos de demostración.")
        st.stop()

    contenidos = documentos(archivos)
    if not contenidos:
        st.error("No se encontraron archivos compatibles.")
        st.stop()

    textos = {
        n: extraer(n, b)
        for n, b in contenidos
    }

    contratista, mensaje = identificar(textos, base_demo())

    if contratista is None:
        st.warning(mensaje)
    else:
        st.success(
            "Contratista identificada: "
            + str(contratista["NOMBRE"])
        )
        st.caption(
            "Contrato: " + str(contratista["CONTRATO"])
            + " | Regional: " + str(contratista["REGIONAL"])
            + " | Proyecto: " + str(contratista["PROYECTO"])
        )
        st.caption(mensaje)

    hallazgos = []

    def agregar(documento, tipo, observacion):
        hallazgos.append({
            "Documento": documento,
            "Estado": tipo,
            "Observación": observacion
        })

    nombres = [normalizar(n) for n, _ in contenidos]

    requeridos = {
        "F62": any("F62" in n for n in nombres),
        "F68": any("F68" in n for n in nombres),
        "Planilla PDF": any(
            "PLANILLA" in n and n.endswith(".PDF")
            for n in nombres
        ),
        "Matriz de casos": any(
            "CASO" in n and n.endswith((".XLSX", ".XLSM"))
            for n in nombres
        ),
        "Matriz PYD Excel": any(
            "PYD" in n and n.endswith((".XLSX", ".XLSM"))
            for n in nombres
        )
    }

    for etiqueta, existe in requeridos.items():
        if not existe:
            agregar(
                "Expediente",
                "Verificar",
                "No se identificó " + etiqueta
                + ". Revisar si el archivo tiene otro nombre."
            )

    for nombre, contenido in contenidos:
        texto = textos[nombre]
        clave = normalizar(nombre)
        plano = normalizar(texto)

        if not texto.strip():
            agregar(
                nombre,
                "Revisión manual",
                "No se pudo extraer texto; puede ser un escaneo "
                "o un archivo protegido."
            )

        if "F68" in clave and not re.search(
            r"\bV\s*3\b|VERSION\s*0?3",
            plano + " " + clave
        ):
            agregar(
                nombre,
                "Revisión manual",
                "No se pudo confirmar la versión V3 del F68."
            )

        if "F62" in clave or "F68" in clave:
            if "LISA CRISTINA" not in plano:
                agregar(
                    nombre,
                    "Revisión manual",
                    "Verificar que la supervisora registrada sea "
                    "la Dra. Lisa Cristina Gómez."
                )

            if "YINA NATALIA" in plano:
                agregar(
                    nombre,
                    "Corregir",
                    "Revisar el nombre de la supervisora registrado."
                )

            agregar(
                nombre,
                "Revisión manual",
                "Comprobar visualmente la firma de la contratista "
                "y la correspondencia de obligaciones y soportes."
            )

            if (
                contratista is not None
                and normalizar(contratista["CONTRATO"]) not in plano
            ):
                agregar(
                    nombre,
                    "Corregir",
                    "No se encontró el contrato esperado "
                    "en el texto extraído."
                )

        if "CASO" in clave and clave.endswith((".XLSX", ".XLSM")):
            for campo in [
                "NOMBRE", "APELLIDO", "CEDULA",
                "TELEFONO", "CORREO", "DIRECCION"
            ]:
                if campo in plano:
                    agregar(
                        nombre,
                        "Privacidad",
                        "Revisar posibles columnas con "
                        + campo.lower()
                        + " de víctimas; eliminar datos identificables."
                    )

        if "PLANILLA" in clave and clave.endswith(".PDF"):
            if "2360000" not in re.sub(r"\D", "", texto):
                agregar(
                    nombre,
                    "Revisión manual",
                    "No se confirmó IBC de $2.360.000; "
                    "comprobar novedades y liquidación."
                )

            agregar(
                nombre,
                "Revisión manual",
                "Comprobar fecha de pago y período de cotización; "
                "no se validan automáticamente."
            )

    siguiente = dt.date(
        periodo.year + (periodo.month == 12),
        periodo.month % 12 + 1,
        1
    )

    ultimo = siguiente - dt.timedelta(days=1)
    festivos = holidays.country_holidays(
        "CO",
        years=[siguiente.year]
    )

    fecha = siguiente
    while fecha.weekday() >= 5 or fecha in festivos:
        fecha += dt.timedelta(days=1)

    st.info(
        f"Período esperado: 01/{periodo.month:02d}/{periodo.year} "
        f"a {ultimo:%d/%m/%Y}. "
        f"Fecha de elaboración: {fecha:%d/%m/%Y}."
    )

    if not hallazgos:
        agregar(
            "Expediente",
            "Revisión manual",
            "No se encontraron alertas básicas; "
            "falta verificar integralmente el expediente."
        )

    tabla = pd.DataFrame(hallazgos).drop_duplicates()

    st.subheader("Observaciones preliminares")
    st.dataframe(
        tabla,
        hide_index=True,
        use_container_width=True
    )

    correo = (
        "Cordial saludo,\n\n"
        "De acuerdo con la revisión realizada, "
        "agradezco verificar los siguientes puntos:\n\n"
    )

    for documento, grupo in tabla.groupby(
        "Documento",
        sort=False
    ):
        correo += (
            documento + "\n"
            + "\n".join(
                "- " + x for x in grupo["Observación"]
            )
            + "\n\n"
        )

    correo += (
        "Quedo atenta al envío de la documentación ajustada."
        "\n\nCordialmente,"
    )

    st.text_area(
        "Borrador para Outlook",
        correo,
        height=280
    )

    st.download_button(
        "Descargar observaciones",
        correo,
        "observaciones_demo.txt"
    )

    st.warning(
        "Demostración: no certifica firmas, macros, anonimización, "
        "pago anticipado ni coincidencia completa de obligaciones."
    )
