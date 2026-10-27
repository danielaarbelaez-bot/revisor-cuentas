import streamlit as st, pandas as pd, re, io, zipfile, unicodedata, datetime, holidays
from docx import Document
from pypdf import PdfReader
from rapidfuzz import fuzz
st.set_page_config(page_title='Revisor de cuentas de cobro',layout='wide')
st.title('Revisor de cuentas de cobro | Defensoría del Pueblo')
st.caption('Prototipo local. No cargue documentos con datos personales en un servidor público.')
def norm(v):
 s=unicodedata.normalize('NFKD',str(v or ''))
 return ''.join(c for c in s if not unicodedata.combining(c)).upper().strip()
def textfile(name,b):
 ext=name.lower().split('.')[-1]
 try:
  if ext=='pdf':return '\n'.join(p.extract_text() or '' for p in PdfReader(io.BytesIO(b)).pages)
  if ext=='docx':
   d=Document(io.BytesIO(b));return '\n'.join([p.text for p in d.paragraphs]+[' '.join(c.text for c in r.cells) for t in d.tables for r in t.rows])
  if ext in ('xlsx','xlsm'):
   x=pd.read_excel(io.BytesIO(b),sheet_name=None,header=None,dtype=str)
   return '\n'.join(str(v) for df in x.values() for v in df.fillna('').values.flatten() if str(v).strip())
 except Exception as e:return '[NO SE PUDO EXTRAER: '+str(e)+']'
 return '[FORMATO NO SOPORTADO]'
def unpack(items):
 out=[]
 for f in items:
  b=f.getvalue()
  if f.name.lower().endswith('.zip'):
   try:
    with zipfile.ZipFile(io.BytesIO(b)) as z:
     for n in z.namelist():
      if not n.endswith('/') and not n.startswith('__MACOSX') and not n.startswith('/') and '..' not in n.split('/') and n.lower().endswith(('.pdf','.docx','.xlsx','.xlsm')) and z.getinfo(n).file_size<25_000_000:out.append((n,z.read(n)))
   except Exception as e:st.error(f'ZIP no válido: {e}')
  else:out.append((f.name,b))
 return out
@st.cache_data
def master(b):
 df=pd.read_excel(io.BytesIO(b),sheet_name='FINANCIERA');df.columns=[norm(c) for c in df.columns]
 return df.dropna(subset=['NOMBRE']).copy()
base=st.file_uploader('Base maestra CONTRATISTAS 2026.xlsx',type=['xlsx'],key='base')
period=st.date_input('Mes de la cuenta (seleccione cualquier día del mes)',value=datetime.date(2026,10,1))
files=st.file_uploader('Cargue F62 (Word), F68 (Excel), planilla (PDF), soportes (PDF), matrices (Excel) o ZIP',type=['zip','docx','xlsx','xlsm','pdf'],accept_multiple_files=True)
if base:
 df=master(base.getvalue()); names=df['NOMBRE'].astype(str).tolist(); chosen=st.selectbox('Contratista',names)
 row=df[df['NOMBRE']==chosen].iloc[0]
 st.write('Contrato:',row.get('NO. CONTRATO',''), '| Regional:',row.get('REGIONAL',''), '| Proyecto:',row.get('PROYECTO',''))
else:st.info('Cargue la base maestra para habilitar el cotejo de identidad y contrato.')
if st.button('Analizar documentos',type='primary'):
 if not files:st.warning('Debe adjuntar documentos.');st.stop()
 docs=unpack(files); results=[]; texts={}
 def flag(doc,level,msg):results.append({'Documento':doc,'Prioridad':level,'Observación':msg})
 for n,b in docs:
  t=textfile(n,b);texts[n]=t;lo=norm(n);nt=norm(t)
  if t.startswith('[NO SE PUDO'):flag(n,'Revisión manual','No se pudo leer el contenido; verificar archivo, contraseña o escaneo.')
  if re.search(r'F68',lo) and not re.search(r'\bV\s*3\b|VERSION\s*3|VERSION\s*03',nt+ ' '+lo):flag(n,'Revisión manual','No se pudo confirmar que el F68 corresponda a la versión 3.')
  if re.search(r'F62|F68',lo):
   if 'LISA CRISTINA' not in nt:flag(n,'Revisión manual','Verificar el nombre de la supervisora Dra. Lisa Cristina Gómez.')
   if 'NATALIA POVEDA' in nt or 'YINA NATALIA' in nt:flag(n,'Corregir','Revisar nombre de supervisora: no debe aparecer la Dra. Natalia como supervisora.')
   if not re.search(r'FIRMA|FIRMADO|SUSCRIBE',nt):flag(n,'Revisión manual','No se identificó texto relativo a firma; comprobar visualmente la firma de la contratista.')
  if 'CASO' in lo and n.lower().endswith(('.xlsx','.xlsm')):
   for word in ['NOMBRE','APELLIDO','CEDULA','DOCUMENTO','TELEFONO','CELULAR','CORREO','DIRECCION']:
    if word in nt:flag(n,'Privacidad','Revisar y eliminar columna/casillas de '+word.lower()+' con datos identificables de víctimas.')
  if 'PYD' in lo and not n.lower().endswith(('.xlsx','.xlsm')):flag(n,'Corregir','La matriz PYD debe entregarse en Excel.')
  if re.search(r'F62|F68',lo):
   if base:
    if norm(row.get('NO. CONTRATO','')) not in nt:flag(n,'Corregir','Verificar número de contrato frente a la base maestra.')
    ced=re.sub(r'\D','',str(row.get('CEDULA','')))
    if ced and ced not in re.sub(r'\D','',t):flag(n,'Revisión manual','No se pudo localizar la cédula de la contratista en el documento.')
   if 'OCTUBRE' in nt and period.month!=10:flag(n,'Revisión manual','El documento menciona octubre; comprobar que corresponde al período seleccionado.')
   for k in [1,2,4,5,6,7]:
    if not re.search(r'(?m)(?:^|\n)\s*'+str(k)+r'\s*[.)-]',t):flag(n,'Revisión manual',f'No se pudo identificar claramente la obligación {k}; verificar estructura y actividad.')
  if 'PLANILLA' in lo and not n.lower().endswith('.pdf'):flag(n,'Corregir','La planilla de seguridad social debe enviarse en PDF.')
  if 'PLANILLA' in lo and n.lower().endswith('.pdf'):
   if '2360000' not in re.sub(r'[^0-9]','',t):flag(n,'Revisión manual','No se pudo confirmar IBC de $2.360.000; verificar liquidación, novedades y período de cotización.')
   flag(n,'Revisión manual','Verificar período de cotización, fecha de pago y condición de pago anticipado; no basta el texto extraído para validarlo.')
  if n.lower().endswith('.pdf') and len(t.strip())<80:flag(n,'Revisión manual','PDF posiblemente escaneado: requiere revisión visual y de anonimización.')
 kinds={'F62':any('F62' in norm(n) for n in texts),'F68':any('F68' in norm(n) for n in texts),'planilla PDF':any('PLANILLA' in norm(n) and n.lower().endswith('.pdf') for n in texts),'matriz de casos':any('CASO' in norm(n) for n in texts),'matriz PYD Excel':any('PYD' in norm(n) and n.lower().endswith(('.xlsx','.xlsm')) for n in texts)}
 for k,ok in kinds.items():
  if not ok:flag('Expediente','Faltante',f'No se identificó {k}; verificar si el archivo tiene otro nombre.')
 yy,mm=period.year,period.month
 nxt=datetime.date(yy+int(mm==12),mm%12+1,1);co=holidays.country_holidays('CO',years=[nxt.year]);expected=nxt
 while expected.weekday()>=5 or expected in co:expected+=datetime.timedelta(days=1)
 last=nxt-datetime.timedelta(days=1)
 st.info(f'Período esperado: 01/{mm:02d}/{yy} a {last:%d/%m/%Y}. Fecha de elaboración esperada: {expected:%d/%m/%Y} (primer día hábil colombiano).')
 if mm==10 and yy==2026:st.caption('Para octubre de 2026: 3 de noviembre de 2026.')
 for n,t in texts.items():
  if 'F62' in norm(n) or 'F68' in norm(n):
   if str(expected.day) not in t:flag(n,'Revisión manual',f'Verificar fecha de elaboración {expected:%d/%m/%Y}; no se pudo confirmar automáticamente.')
 f62=' '.join(t for n,t in texts.items() if 'F62' in norm(n));f68=' '.join(t for n,t in texts.items() if 'F68' in norm(n))
 if f62 and f68:
  for k in [1,2,4,5,6,7]:
   a=re.search(r'(?m)^\s*'+str(k)+r'\s*[.)-]\s*(.{20,300})',f62);b=re.search(r'(?m)^\s*'+str(k)+r'\s*[.)-]\s*(.{20,300})',f68)
   if a and b and fuzz.ratio(norm(a.group(1)),norm(b.group(1)))<70:flag('F62 y F68','Corregir',f'Obligación {k}: posible diferencia entre actividades registradas en F62 y F68; cotejar ambos documentos.')
 if not results:flag('Expediente','Revisión manual','Sin hallazgos por reglas automáticas; verificar soportes, firmas, datos personales y correspondencia contractual.')
 report=pd.DataFrame(results).drop_duplicates();st.subheader('Hallazgos');st.dataframe(report,use_container_width=True,hide_index=True)
 st.download_button('Descargar informe CSV',report.to_csv(index=False).encode('utf-8-sig'),'hallazgos.csv','text/csv')
 notes='\n\n'.join(f'{name}:\n'+'\n'.join('- '+x for x in group['Observación']) for name,group in report.groupby('Documento',sort=False))
 email='Cordial saludo,\n\nDe acuerdo con la revisión realizada, agradezco verificar y efectuar los siguientes ajustes:\n\n'+notes+'\n\nQuedo atenta al envío de la documentación ajustada.\n\nCordialmente,'
 st.subheader('Borrador de observaciones');st.text_area('Copiar al correo',email,height=320)
 st.download_button('Descargar observaciones',email,'observaciones.txt')
 st.warning('Este prototipo no certifica anonimización, autenticidad de firmas, funcionamiento de macros, concordancia íntegra de obligaciones ni pago anticipado. Todos requieren controles adicionales.')
