SISTEMA SMAP 2026 · VERSION LOCAL

1) Copie esta carpeta a su PC.
2) Verifique config.json: la ruta predeterminada es:
   C:\Users\USER\Desktop\INTERVENTORIA VELNEC\SMAP\SMAP_VELNEC_2026_JCA.xlsx
3) Ejecute INICIAR_SISTEMA_SMAP.bat.
4) Se abrira http://127.0.0.1:8765 en el navegador.
5) El sistema lee directamente la hoja Tabla SMAP-CIV.
6) Revisa cambios del archivo cada 5 segundos y actualiza los datos al guardar el Excel.

REGLA SMAP PENDIENTES:
Los consecutivos ya creados pero sin Año/Mes/Dirección se clasifican como PENDIENTES DE ACTUALIZACION. No son errores, no se cuentan como incumplimientos ANS ni como vencidos.

MAPA:
Usa Latitud/Longitud reales. El mapa base usa Esri World Street Map. Requiere Internet para cargar las teselas, pero los puntos y coordenadas vienen del Excel local.

TECNICOS:
El tablero ya soporta el gráfico por técnico y filtros Año/Mes. En la base actual, la columna Tecnicos no tiene asignaciones por fila; por eso la vista mostrara una alerta informativa hasta que esa columna sea diligenciada.

COLORES Y DISEÑO:
El botón Diseño permite cambiar ancho/alto de cada panel y el orden por arrastre. Los colores se guardan en el navegador.

PARA FUTURO:
Esta arquitectura deja lista la separación entre fuente de datos y tablero para luego migrar a una versión web compartida.


V2.2 - Correcciones: búsqueda por SMAP/No. SS/radicado, resaltado de punto buscado, luminaria LED real en encabezado, leyendas de tortas a la derecha y valores dentro, diseño con orden/ancho/alto manual, puerto local 8876.


CORRECCION V2.7: La gráfica SMAP generadas por interventor toma la asignación real por fila desde la columna SOLICITANTE, donde aparecen valores como Interventor 17, Interventor 25, etc. La columna N°INTERVENTOR es únicamente un listado maestro y no se utiliza para contar.


CORRECCION V2.14: la gráfica SMAP por interventor usa prioritariamente la columna INTERVENTOR del Excel; si el archivo histórico usa Solicitante, la toma como respaldo. La gráfica cuenta las solicitudes por persona y responde a los filtros de Año/Mes/Comuna/Estado/Nivel/Tipo de recorrido. No usa la lista maestra N°INTERVENTOR para contar solicitudes.


V2.16 - Corrección de gráfica SMAP por interventor: usa la columna INTERVENTOR (columna F) de cada solicitud, cuenta todas las solicitudes filtradas y permite análisis por mes/año. Se actualizó la caché inicial para evitar mostrar el listado maestro N°INTERVENTOR.
