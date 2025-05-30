# script-conciliacion-py
Script en Python para conciliar comprobantes entre dos archivos de Excel exportados por sistema interno y AFIP.  

¿Qué hace el script?

Lee y procesa dos archivos Excel:

- 📄 `Facturas en sistema.xlsx` (exportado de un sistema contable)  
- 📄 `Facturas en afip.xlsx` (descargado desde Mis Comprobantes Recibidos)

Limpia y normaliza los datos clave (punto de venta, número de factura e importes), incluso si vienen en formatos distintos.

Compara cada comprobante y lo clasifica automáticamente en una de estas categorías:

- ✅ Coincidencia exacta  
- ⚠️ Coincidencia parcial (mismo número, distinto importe)  
- 🟥 Solo en AFIP  
- 🟦 Solo en el sistema  

Exporta los resultados a un Excel con solapas separadas según su clasificación.

Si falta algún dato clave en AFIP → va a **Solo en el SISTEMA**.  
Si falta algún dato clave en el sistema → va a **Solo en AFIP**.

Incluye una solapa “Mis Comprobantes Recibidos” con los datos tal como figuran en AFIP + una columna "condición" para una comoda revision de los resultados. Las filas están coloreadas automáticamente según el tipo de discrepancia.

---

## Tecnologías utilizadas

- Python  
- pandas  
- openpyxl
- 

> **ACLARACIÓN:** Los datos utilizados son ficticios. No representan información real.


