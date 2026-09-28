# -*- coding: utf-8 -*-
"""Actualizador desde el PC administrador hacia SISTEMA SMAP 2026 WEB.

Uso:
  python actualizar_smap_web.py

Configura SMAP_WEB_URL y SMAP_SYNC_TOKEN como variables de entorno, o edita
los valores de abajo antes de usarlo.
"""
import json, os, math, urllib.request
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook

WEB_URL = os.environ.get('SMAP_WEB_URL', 'https://REEMPLAZAR-URL-RENDER.onrender.com')
TOKEN = os.environ.get('SMAP_SYNC_TOKEN', 'CAMBIAR_TOKEN_SMAPPRO')
EXCEL = Path(os.environ.get('SMAP_EXCEL', r'C:\Users\USER\Desktop\INTERVENTORIA VELNEC\SMAP\SMAP_VELNEC_2026_JCA.xlsx'))
SHEET = 'Tabla SMAP-CIV'

def norm(s): return ' '.join(str(s or '').replace('\n',' ').split()).strip().lower()
def blank(v): return v is None or (isinstance(v,float) and math.isnan(v)) or str(v).strip()=='' or str(v).strip().lower() in {'nan','#n/a','none'}
def serial(v):
    if isinstance(v,(datetime,date)): return v.isoformat()
    if isinstance(v,float) and math.isnan(v): return None
    if isinstance(v,float) and v.is_integer(): return int(v)
    return v

def idx(H,*names):
    for n in names:
        n=norm(n)
        for i,h in enumerate(H):
            if h==n: return i
    return -1

wb=load_workbook(EXCEL,read_only=True,data_only=True)
ws=wb[SHEET]
it=ws.iter_rows(min_row=2,values_only=True)
headers=list(next(it)); H=[norm(x) for x in headers]
I={
'consecutivo':idx(H,'Consecutivo'),'mes':idx(H,'Mes'),'anio':idx(H,'Año'),'solicitante':idx(H,'Solicitante'),
'direccion':idx(H,'Dirección'),'lat':idx(H,'Coordenadas Latitud'),'lon':idx(H,'Coordenadas Longitud'),'barrio':idx(H,'Barrio/Vereda'),
'comuna':idx(H,'Comuna'),'nombreComuna':idx(H,'Nombre comuna'),'falla':idx(H,'N° Falla'),'tipoRecorrido':idx(H,'TIPO RECORRIDO'),
'oficio':idx(H,'Oficio de envio'),'fechaEnvio':idx(H,'Fecha de envio CIV'),'estado':idx(H,'Estado de SMAP'),'noSS':idx(H,'No. de SS'),
'tipoSS':idx(H,'Tipo de SS'),'fechaSS':idx(H,'Fecha recibo del numero de SS'),'diasSS':idx(H,'Días Trascurridos entrega numero de SS'),
'cumpleSS':idx(H,'Cumple ANS entrega de numero de SS'),'nivel':idx(H,'Tipo de Nivel'),'fechaEntrega':idx(H,'Fecha entrega de Respuesta'),
'diasRespuesta':idx(H,'Días Trascurridos entrega de Respuesta'),'vencimiento':idx(H,'VENCIMIENTO ENTREGA RESPUESTA'),
'cumpleRespuesta':idx(H,'Cumple ANS entrega de Respuesta'),'fechaRespuesta':idx(H,'Fecha de Respuesta'),'radicadoResp':idx(H,'Radicado Respuesta EPM2'),
'observaciones':idx(H,'Observaciones'),'aval':idx(H,'Radicado AVAL VELNEC'),'fechaAval':idx(H,'Fecha de AVAL'),'aprobDistrito':idx(H,'RAD.APROBACION DISTRITO'),
'fechaAprob':idx(H,'Fecha de APROB DISTRITO'),'ssRevision':idx(H,'FECHA RESPUEST SS REVISAR_EJECUCION'),'diasRevision':idx(H,'Días Trascurridos entrega de SS REVISAR'),
'vencRevision':idx(H,'VENCIMIENTO'),'cumpleRevision':idx(H,'Cumple ANS entrega de SS REVISAR'),'radEjec':idx(H,'Radicado EJECUCION VELNEC'),
'fechaEjec':idx(H,'Fecha de RADI EJECUCION'),'iniciales':idx(H,'Iniciales'),'interventor':idx(H,'INTERVENTOR'),'tecnico':idx(H,'Tecnicos'),'fallaSmap':idx(H,'FALLAS smap')}
exact=next((i for i,h in enumerate(H) if h=='interventor'),-1)
I['interventor']=exact if exact>=0 else 5

data=[]
for r in it:
    if not r or I['consecutivo']<0: continue
    con=r[I['consecutivo']] if I['consecutivo']<len(r) else None
    if blank(con): continue
    d={k:(serial(r[j]) if 0<=j<len(r) else None) for k,j in I.items()}
    d['pendienteActualizacion']=not bool(d.get('anio') or d.get('mes') or d.get('direccion'))
    data.append(d)
wb.close()

payload=json.dumps({'data':data},ensure_ascii=False,separators=(',',':')).encode('utf-8')
req=urllib.request.Request(WEB_URL.rstrip('/')+'/api/sync',data=payload,method='POST',headers={'Content-Type':'application/json','X-SMAP-TOKEN':TOKEN})
with urllib.request.urlopen(req,timeout=60) as resp:
    result=json.loads(resp.read().decode('utf-8'))
print('SISTEMA SMAP 2026')
print('Excel:',EXCEL)
print('Registros enviados:',len(data))
print('Respuesta web:',result)
