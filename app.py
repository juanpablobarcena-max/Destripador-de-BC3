import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conversor BC3 a Excel", page_icon="🏗️", layout="wide")

st.title("🏗️ Conversor Nativo de BC3 a Excel")
archivo_bc3 = st.file_uploader("Sube aquí tu archivo .bc3", type=["bc3"])

# --- Función auxiliar para limpiar números del BC3 ---
def parsear_numero(valor):
    try:
        return float(valor.strip().replace(',', '.'))
    except:
        return 0.0

if archivo_bc3 is not None:
    try:
        contenido = archivo_bc3.read().decode('latin-1')
    except:
        contenido = archivo_bc3.read().decode('utf-8')
    
    # 1. DICCIONARIOS PARA ALMACENAR DATOS TEMPORALMENTE
    conceptos = {}
    jerarquia = {}
    mediciones = {}

    lineas = contenido.splitlines()
    
    # 2. PARSEO DEL ARCHIVO LÍNEA A LÍNEA
    for linea in lineas:
        
        # A. Extraer Conceptos (Textos, Unidades, Precios)
        if linea.startswith('~C|'):
            partes = linea.split('|')
            if len(partes) >= 5:
                codigo = partes[1].strip()
                precio_str = partes[4].strip().replace('\\', '')
                conceptos[codigo] = {
                    "Ud": partes[2].strip(),
                    "Descripción": partes[3].strip(),
                    "Precio": parsear_numero(precio_str)
                }
                
        # B. Extraer Jerarquía / Descomposición (Padres e Hijos)
        elif linea.startswith('~D|'):
            partes = linea.split('|')
            if len(partes) >= 3:
                padre = partes[1].strip()
                hijos = []
                # Los hijos vienen separados por '|', y sus detalles por '\'
                for hijo_str in partes[2:]:
                    if hijo_str:
                        datos_hijo = hijo_str.split('\\')
                        if len(datos_hijo) >= 1:
                            hijos.append(datos_hijo[0].strip()) # Guardamos solo el código del hijo
                jerarquia[padre] = hijos
                
        # C. Extraer Mediciones (Cantidades)
        elif linea.startswith('~M|'):
            partes = linea.split('|')
            if len(partes) >= 3:
                codigo = partes[1].strip()
                # Las líneas de medición están a partir del índice 3 normalmente
                lineas_med = partes[3:] if len(partes) > 3 else []
                total_partida = 0.0
                
                for m in lineas_med:
                    if m:
                        # Formato: Comentario \ Uds \ Largo \ Ancho \ Alto
                        datos_m = m.split('\\')
                        factores = []
                        
                        # Multiplicamos las dimensiones (índices 1 al 4) si existen
                        for i in range(1, 5):
                            if len(datos_m) > i and datos_m[i].strip():
                                factores.append(parsear_numero(datos_m[i]))
                        
                        # Si hay números, calculamos el subtotal de la línea
                        if factores:
                            subtotal = 1.0
                            for f in factores:
                                subtotal *= f
                            total_partida += subtotal

                mediciones[codigo] = total_partida

    # 3. RECONSTRUCCIÓN DEL PRESUPUESTO
    datos_extraidos = []
    
    # Buscamos la raíz del proyecto (El código que es padre pero no es hijo de nadie)
    todos_los_hijos = set([h for lista in jerarquia.values() for h in lista])
    raices = [p for p in jerarquia.keys() if p not in todos_los_hijos]
    
    if raices:
        raiz_principal = raices[0]
        capitulos = jerarquia.get(raiz_principal, [])
        
        # Recorremos cada Capítulo
        for cod_cap in capitulos:
            nombre_capitulo = conceptos.get(cod_cap, {}).get("Descripción", cod_cap)
            partidas = jerarquia.get(cod_cap, [])
            
            # Recorremos cada Partida dentro del Capítulo
            for cod_partida in partidas:
                concepto = conceptos.get(cod_partida, {})
                cantidad = mediciones.get(cod_partida, 0.0)
                precio = concepto.get("Precio", 0.0)
                
                datos_extraidos.append({
                    "Capítulo": nombre_capitulo,
                    "Código": cod_partida,
                    "Ud": concepto.get("Ud", ""),
                    "Descripción": concepto.get("Descripción", ""),
                    "Precio (€)": precio,
                    "Cantidad": round(cantidad, 3),
                    "Importe (€)": round(cantidad * precio, 2)
                })

    # 4. RENDERIZADO EN STREAMLIT Y EXCEL
    if datos_extraidos:
        df_bc3 = pd.DataFrame(datos_extraidos)
        
        # Mostramos KPIs rápidos en Streamlit
        st.subheader(f"Total Presupuesto: {df_bc3['Importe (€)'].sum():,.2f} €")
        st.dataframe(df_bc3, use_container_width=True)
        
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df_bc3.to_excel(writer, index=False, sheet_name="Presupuesto")
        
        st.download_button(
            label="⬇️ Descargar Excel Estructurado", 
            data=buffer_excel.getvalue(), 
            file_name="Presupuesto_Estructurado.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No se han encontrado partidas. Revisa el formato de tu BC3.")
