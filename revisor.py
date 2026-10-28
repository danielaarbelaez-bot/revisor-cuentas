import io
import re
import zipfile
import unicodedata
import datetime as dt
from collections import defaultdict

import pandas as pd
import streamlit as st
import holidays
from docx import Document
from pypdf import PdfReader
from openpyxl import load_workbook

st.title('Revisor de cuentas de cobro | Demostración V2')
st.warning('Solo para pruebas con datos ficticios. No cargue documentos institucionales ni datos personales reales.')

MAX_FILE = 20 * 1024 * 1024
MAX_TOTAL = 80 * 1024 * 1024
MAX_ZIP_FILES = 40


def normalizar(valor):
    texto = unicodedata.normalize('NFKD', str(valor if valor is not None else ''))
    return ''.join(c for c in texto if not unicodedata.combining(c)).upper().strip()


@st.cache_data
def base_demo():
    return pd.DataFrame([
        {'NOMBRE': 'ANA MARIA PRUEBA', 'CEDULA': '999000001', 'CONTRATO': 'DEF-DEMO-001', 'REGIONAL': 'ANTIOQUIA', 'PROYECTO': 'EXIGIBILIDAD', 'HONORARIOS': 5900000},
        {'NOMBRE': 'LAURA PATRICIA EJEMPLO', 'CEDULA': '999000002', 'CONTRATO': 'DEF-DEMO-002', 'REGIONAL': 'QUINDIO', 'PROYECTO': 'VICTIMAS', 'HONORARIOS': 5900000},
    ])


def documentos(archivos):
    salida, total = [], 0
    for archivo in archivos:
        contenido = archivo.getvalue()
        if len(contenido) > MAX_FILE:
            st.error(f'{archivo.name}: supera el límite de 20 MB.')
            continue
        if archivo.name.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(io.BytesIO(contenido)) as z:
                    candidatos = [x for x in z.infolist() if not x.is_dir()]
                    if len(candidatos) > MAX_ZIP_FILES:
                        st.error(f'{archivo.name}: contiene demasiados archivos.')
                        continue
                    for info in candidatos:
                        nombre = info.filename.replace('\\', '/')
                        if ('..' in nombre.split('/') or nombre.startswith('/') or info.flag_bits & 0x1 or
                            not nombre.lower().endswith(('.pdf', '.docx', '.xlsx', '.xlsm'))):
                            continue
                        if info.file_size > MAX_FILE or total + info.file_size > MAX_TOTAL:
                            st.warning('Se omitió un archivo por límites de tamaño.')
                            continue
                        data = z.read(info)
                        total += len(data)
                        salida.append((nombre, data))
            except (zipfile.BadZipFile, RuntimeError, ValueError, OSError):
                st.error(f'No se pudo abrir el ZIP {archivo.name}.')
        elif archivo.name.lower().endswith(('.pdf', '.docx', '.xlsx', '.xlsm')):
            if total + len(contenido) <= MAX_TOTAL:
                salida.append((archivo.name, contenido))
                total += len(contenido)
    return salida


def extraer(nombre, contenido):
    ext = nombre.lower().rsplit('.', 1)[-1]
    datos = {'texto': '', 'ubicaciones': [], 'hojas': {}, 'version_f68': None, 'error': None, 'escaneo': False}
    try:
        if ext == 'pdf':
            pdf = PdfReader(io.BytesIO(contenido))
            for i, pagina in enumerate(pdf.pages):
                texto = pagina.extract_text() or ''
                datos['ubicaciones'].append((f'página {i+1}', texto))
                if len(texto.strip()) < 20:
                    datos['escaneo'] = True
        elif ext == 'docx':
            doc = Document(io.BytesIO(contenido))
            for i, p in enumerate(doc.paragraphs):
                if p.text.strip():
                    datos['ubicaciones'].append((f'párrafo {i+1}', p.text))
            for t, tabla in enumerate(doc.tables):
                for i, fila in enumerate(tabla.rows):
                    for j, celda in enumerate(fila.cells):
                        if celda.text.strip():
                            datos['ubicaciones'].append((f'tabla {t+1}, fila {i+1}, columna {j+1}', celda.text))
            # Word may contain text in headers, footers and embedded images not extracted here.
        elif ext in ('xlsx', 'xlsm'):
            libro = load_workbook(io.BytesIO(contenido), read_only=True, data_only=False)
            for hoja in libro.worksheets:
                valores = {}
                for fila in hoja.iter_rows():
                    for celda in fila:
                        if celda.value is not None and str(celda.value).strip():
                            val = str(celda.value)
                            valores[celda.coordinate] = val
                            datos['ubicaciones'].append((f'hoja {hoja.title}, celda {celda.coordinate}', val))
                datos['hojas'][hoja.title] = valores
            if 'F68' in normalizar(nombre):
                for titulo, celdas in datos['hojas'].items():
                    if 'FORMULARIO PARA PAGOS' in normalizar(titulo):
                        datos['version_f68'] = celdas.get('AA4')
                        break
        datos['texto'] = '\n'.join(texto for _, texto in datos['ubicaciones'])
    except Exception as exc:
        datos['error'] = type(exc).__name__
    return datos


def identificar(analizados, base):
    texto = normalizar('\n'.join(x['texto'] for x in analizados.values()))
    candidatos = []
    for _, fila in base.iterrows():
        señales = []
        if normalizar(fila['CONTRATO']) in texto:
            señales.append('contrato')
        if re.search(r'(?<!\d)' + re.escape(str(fila['CEDULA'])) + r'(?!\d)', texto):
            señales.append('cédula')
        if normalizar(fila['NOMBRE']) in texto:
            señales.append('nombre')
        candidatos.append((len(señales), fila, señales))
    candidatos.sort(key=lambda v: v[0], reverse=True)
    if not candidatos or candidatos[0][0] < 2 or (len(candidatos) > 1 and candidatos[0][0] == candidatos[1][0]):
        return None, 'Identificación no concluyente: se requieren dos coincidencias y un resultado único.'
    return candidatos[0][1], 'Coincidencia preliminar mediante ' + ', '.join(candidatos[0][2]) + '.'


def es_marcador(valor):
    v = normalizar(valor).strip(' .:-_')
    return not v or v in {'N/A', 'NA', 'NO APLICA', 'ANONIMIZADO', 'RESERVADO', 'CONFIDENCIAL', 'XXXXX', 'XXXX', 'XXX', 'NOMBRE', 'APELLIDO', 'CEDULA', 'DOCUMENTO', 'TELEFONO', 'CORREO', 'DIRECCION', 'SIN DATO'} or bool(re.fullmatch(r'[X*#\-\s]{3,}', v))


def riesgos_privacidad(nombre, datos, es_matriz):
    avisos = []
    campos = re.compile(r'\b(NOMBRE(?:S)?|APELLIDO(?:S)?|C[EÉ]DULA|DOCUMENTO(?: DE IDENTIDAD)?|TEL[EÉ]FONO|CELULAR|CORREO(?: ELECTR[OÓ]NICO)?|DIRECCI[OÓ]N)\b', re.I)
    email = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b')
    telefono = re.compile(r'(?<!\d)3\d{9}(?!\d)')
    identificacion = re.compile(r'(?<!\d)\d{7,10}(?!\d)')
    if es_matriz and datos['hojas']:
        for hoja, celdas in datos['hojas'].items():
            encabezados = [(pos, val) for pos, val in celdas.items() if campos.search(normalizar(val))]
            for pos, encabezado in encabezados:
                col = re.match(r'[A-Z]+', pos).group()
                fila_enc = int(re.search(r'\d+', pos).group())
                ejemplos = []
                for celda, valor in celdas.items():
                    if re.match(r'[A-Z]+', celda).group() != col or int(re.search(r'\d+', celda).group()) <= fila_enc:
                        continue
                    if not es_marcador(valor) and not str(valor).startswith('='):
                        ejemplos.append(celda)
                if ejemplos:
                    avisos.append((f'hoja {hoja}, columna {col}', f'La columna «{encabezado[:45]}» contiene posibles identificadores en {len(ejemplos)} celda(s). Revisar anonimización. No se muestran valores.'))
                else:
                    avisos.append((f'hoja {hoja}, columna {col}', f'Existe la columna «{encabezado[:45]}»; no se detectaron valores identificables en celdas posteriores. Comprobar necesidad de conservarla.'))
    # En soportes se alerta de patrones, sin mostrar los datos; las cédulas aisladas son ambiguas.
    if not ('F62' in normalizar(nombre) or 'F68' in normalizar(nombre) or 'PLANILLA' in normalizar(nombre)):
        for ubicacion, texto in datos['ubicaciones']:
            if email.search(texto) or telefono.search(texto):
                avisos.append((ubicacion, 'Posible correo o teléfono identificable. Verificar si corresponde a un tercero y si debe anonimizarse.'))
            elif not es_matriz and identificacion.search(texto) and campos.search(texto):
                avisos.append((ubicacion, 'Posible número de identificación asociado a un campo personal. Revisar anonimización.'))
    # Prevent repetitive alerts from large files.
    return avisos[:35]


st.success('Base ficticia incorporada automáticamente. No necesitas subir un Excel.')
with st.expander('Ver contratistas ficticias para la prueba'):
    st.dataframe(base_demo(), hide_index=True, use_container_width=True)

periodo = st.date_input('Mes de la cuenta', value=dt.date(2026, 10, 1))
archivos = st.file_uploader('Cargar F62, F68, planilla, matrices o ZIP de demostración', type=['pdf', 'docx', 'xlsx', 'xlsm', 'zip'], accept_multiple_files=True)

if st.button('Revisar cuenta automáticamente', type='primary', use_container_width=True):
    if not archivos:
        st.warning('Primero carga documentos de demostración.')
        st.stop()
    contenidos = documentos(archivos)
    if not contenidos:
        st.error('No se encontraron archivos compatibles.')
        st.stop()
    analizados = {nombre: extraer(nombre, contenido) for nombre, contenido in contenidos}
    contratista, mensaje = identificar(analizados, base_demo())
    if contratista is None:
        st.warning(mensaje)
    else:
        st.success('Contratista identificada: ' + str(contratista['NOMBRE']))
        st.caption(f"Contrato: {contratista['CONTRATO']} | Regional: {contratista['REGIONAL']} | Proyecto: {contratista['PROYECTO']}")
        st.caption(mensaje)

    hallazgos = []
    def agregar(documento, estado, observacion, ubicacion='—'):
        hallazgos.append({'Documento': documento, 'Estado': estado, 'Ubicación': ubicacion, 'Observación': observacion})

    nombres = [normalizar(n) for n, _ in contenidos]
    requeridos = {
        'F62': any('F62' in n for n in nombres),
        'F68': any('F68' in n for n in nombres),
        'Planilla PDF': any('PLANILLA' in n and n.endswith('.PDF') for n in nombres),
        'Matriz de casos': any('CASO' in n and n.endswith(('.XLSX', '.XLSM')) for n in nombres),
        'Matriz PYD Excel': any('PYD' in n and n.endswith(('.XLSX', '.XLSM')) for n in nombres),
    }
    for etiqueta, existe in requeridos.items():
        if not existe:
            agregar('Expediente', 'Verificar', f'No se identificó {etiqueta}; revisar si el archivo tiene otro nombre.')

    f62, f68 = [], []
    for nombre, datos in analizados.items():
        clave = normalizar(nombre)
        plano = normalizar(datos['texto'])
        if datos['error']:
            agregar(nombre, 'Revisión manual', 'No fue posible leer el documento completamente; verificar archivo o protección.')
            continue
        if not datos['texto'].strip():
            agregar(nombre, 'Revisión manual', 'No se extrajo texto; podría ser un escaneo o documento protegido.')
        if datos['escaneo']:
            agregar(nombre, 'Revisión manual', 'Algunas páginas carecen de texto extraíble. Las imágenes y firmas no fueron verificadas.')
        if 'F62' in clave:
            f62.append((nombre, datos))
        if 'F68' in clave:
            f68.append((nombre, datos))
            if datos['version_f68'] is None:
                agregar(nombre, 'Revisión manual', 'No se encontró la celda de versión AA4 en la hoja «Formulario para pagos»; comprobar plantilla.')
            elif str(datos['version_f68']).strip() not in {'3', '03', '3.0'}:
                agregar(nombre, 'Corregir', f'La versión indicada en AA4 no corresponde a la versión 3 esperada.', 'Formulario para pagos!AA4')
        if 'F62' in clave or 'F68' in clave:
            agregar(nombre, 'Revisión manual', 'Comprobar firmas y correspondencia entre obligaciones, actividades y soportes; no se certifican automáticamente.')
            if contratista is not None and normalizar(contratista['CONTRATO']) not in plano:
                agregar(nombre, 'Verificar', 'No se encontró el contrato esperado en el texto extraído; comprobar el campo del formato.')
        if 'PLANILLA' in clave and clave.endswith('.PDF'):
            agregar(nombre, 'Revisión manual', 'Validar IBC, novedades, aportes, período cotizado y fecha de pago contra la planilla original. El valor de $2.360.000 no es universal.')
        es_matriz = clave.endswith(('.XLSX', '.XLSM')) and ('CASO' in clave or 'PYD' in clave)
        if es_matriz or ('SOPORTE' in clave or 'EVIDENCIA' in clave or 'ANEXO' in clave):
            for ubicacion, aviso in riesgos_privacidad(nombre, datos, es_matriz):
                agregar(nombre, 'Privacidad: verificar', aviso, ubicacion)

    if len(f62) == 1 and len(f68) == 1:
        nombre62, datos62 = f62[0]
        nombre68, datos68 = f68[0]
        texto62, texto68 = normalizar(datos62['texto']), normalizar(datos68['texto'])
        if contratista is not None:
            contrato = normalizar(contratista['CONTRATO'])
            if contrato in texto62 and contrato in texto68:
                st.caption('Coincidencia preliminar: número de contrato presente en ambos formatos.')
            else:
                agregar('Comparación F62–F68', 'Verificar', 'El número de contrato no pudo confirmarse en ambos formatos.')
        agregar('Comparación F62–F68', 'Revisión manual', 'Pendiente cotejo estructurado de períodos, valores, obligaciones y pagos con las celdas/campos específicos de las plantillas.')
    elif len(f62) > 1 or len(f68) > 1:
        agregar('Comparación F62–F68', 'Verificar', 'Se recibieron varios F62 o F68; seleccionar el par correspondiente antes de comparar.')

    siguiente = dt.date(periodo.year + (periodo.month == 12), periodo.month % 12 + 1, 1)
    ultimo = siguiente - dt.timedelta(days=1)
    festivos = holidays.country_holidays('CO', years=[siguiente.year, siguiente.year + 1])
    fecha = siguiente
    while fecha.weekday() >= 5 or fecha in festivos:
        fecha += dt.timedelta(days=1)
    st.info(f'Período seleccionado: 01/{periodo.month:02d}/{periodo.year} a {ultimo:%d/%m/%Y}. Primer día hábil siguiente: {fecha:%d/%m/%Y} (referencia; confirmar plazo contractual).')

    if not hallazgos:
        agregar('Expediente', 'Revisión manual', 'Sin alertas automáticas; no equivale a validación integral.')
    tabla = pd.DataFrame(hallazgos).drop_duplicates()
    st.subheader('Observaciones preliminares')
    st.dataframe(tabla, hide_index=True, use_container_width=True)
    st.caption('Las alertas de privacidad son indicios, no pruebas concluyentes. No se verifican imágenes, escaneos, firmas ni datos ocultos de Excel.')

    correo = 'Cordial saludo,\n\nDe acuerdo con la revisión preliminar realizada, agradezco verificar los siguientes puntos:\n\n'
    for documento, grupo in tabla.groupby('Documento', sort=False):
        correo += documento + '\n'
        for _, fila in grupo.iterrows():
            ubic = '' if fila['Ubicación'] == '—' else f" ({fila['Ubicación']})"
            correo += f"- [{fila['Estado']}]{ubic} {fila['Observación']}\n"
        correo += '\n'
    correo += 'Quedo atenta al envío de los ajustes correspondientes.\n\nCordialmente,'
    st.text_area('Borrador para Outlook', correo, height=280)
    st.download_button('Descargar observaciones', correo, 'observaciones_demo_v2.txt', mime='text/plain')
    st.warning('Demostración: no certifica anonimización completa, firmas, macros, pagos, ni concordancia integral de obligaciones. No apta para expedientes reales.')
