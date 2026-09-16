import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conversor BC3 a Excel", page_icon="🏗️", layout="wide")

st.title("🏗️ Conversor Nativo de BC3 a Excel")
st.write("Sube tu archivo exportado de Presto o Arquímedes (.bc3) y extrae las partidas limpias.")

# 1. Cargador de archivos
archivo_bc3 = st.file_uploader("Sube aquí tu archivo .bc3", type=["bc3"])

if archivo_bc3 is not None:
    # Los archivos BC3 en España suelen usar codificación ANSI/Latin-1
    try:
        contenido = archivo_bc3.read().decode('latin-1')
    except:
        contenido = archivo_bc3.read().decode('utf-8')
    
    st.info("Leyendo la estructura del BC3...")
    
    datos_extraidos = []
    
    # 2. El motor que "traduce" el estándar BC3
    lineas = contenido.splitlines()
    for linea in lineas:
        # En el estándar BC3, los registros de Conceptos empiezan por ~C|
        if linea.startswith('~C|'):
            partes = linea.split('|')
            
            # Asegurarnos de que la línea tiene los datos completos
            if len(partes) >= 5:
                codigo = partes[1].strip()
                unidad = partes[2].strip()
                resumen = partes[3].strip()
                
                # Formatear el precio quitando símbolos raros y convirtiendo a número si es posible
                precio_str = partes[4].strip().replace('\\', '')
                try:
                    precio = float(precio_str)
                except ValueError:
                    precio = 0.0
                
                # Añadir a nuestra lista
                datos_extraidos.append({
                    "Código": codigo,
                    "Ud": unidad,
                    "Descripción": resumen,
                    "Precio (€)": precio
                })

    # 3. Mostrar y exportar
    if datos_extraidos:
        st.success(f"¡Éxito! Se han encontrado {len(datos_extraidos)} registros (Capítulos y Partidas).")
        
        df_bc3 = pd.DataFrame(datos_extraidos)
        
        st.dataframe(df_bc3, use_container_width=True)
        
        # Crear Excel
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df_bc3.to_excel(writer, index=False, sheet_name="Presupuesto")
        
        st.download_button(
            label="⬇️ Descargar Presupuesto en Excel",
            data=buffer_excel.getvalue(),
            file_name="Presupuesto_Extraido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No se encontraron partidas válidas en este archivo BC3.")
