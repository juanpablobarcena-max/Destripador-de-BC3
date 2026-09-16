import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conversor BC3 a Excel", page_icon="🏗️", layout="wide")

st.title("🏗️ Conversor Nativo de BC3 a Excel")
archivo_bc3 = st.file_uploader("Sube aquí tu archivo .bc3", type=["bc3"])

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
    
    conceptos = {}
    jerarquia = {}
    
    lineas = contenido.splitlines()
    
    for linea in lineas:
        # 1. Leer Conceptos (~C)
        if linea.startswith('~C|'):
            partes = linea.split('|')
            if len(partes) >= 5:
                # Limpiamos el código de almohadillas '#' para unificar jerarquías
                codigo = partes[1].strip().replace('#', '')
                precio_str = partes[4].strip().replace('\\', '')
                conceptos[codigo] = {
                    "Ud": partes[2].strip(),
                    "Descripción": partes[3].strip(),
                    "Precio": parsear_numero(precio_str)
                }
                
        # 2. Leer Textos Largos (~T)
        elif linea.startswith('~T|'):
            partes = linea.split('|')
            if len(partes) >= 3:
                codigo = partes[1].strip().replace('#', '')
                texto = partes[2].strip()
                if codigo in conceptos:
                    # Sobrescribimos la descripción corta con la larga si existe
                    conceptos[codigo]["Descripción"] = texto

        # 3. Leer Descomposición / Jerarquía (~D)
        elif linea.startswith('~D|'):
            partes = linea.split('|')
            if len(partes) >= 3:
                padre = partes[1].strip().replace('#', '')
                hijos = []
                # Los hijos vienen separados por '\'
                # Formato FIEBDC: Hijo1 \ Factor1 \ Rendimiento1 \ Hijo2 \ ...
                datos_hijos = "|".join(partes[2:]).split('\\')
                
                # Iteramos de 3 en 3 leyendo: código del hijo, factor, cantidad(rendimiento)
                for i in range(0, len(datos_hijos) - 2, 3):
                    hijo_cod = datos_hijos[i].strip().replace('#', '')
                    if hijo_cod:
                        cantidad = parsear_numero(datos_hijos[i+2])
                        hijos.append((hijo_cod, cantidad))
                        
                jerarquia[padre] = hijos

    # 4. RECONSTRUCCIÓN DEL ÁRBOL (Soporta niveles infinitos)
    datos_extraidos = []
    
    # Encontramos la raíz principal (el nodo que tiene hijos pero no es hijo de nadie)
    todos_los_hijos = set([h[0] for lista in jerarquia.values() for h in lista])
    raices = [p for p in jerarquia.keys() if p not in todos_los_hijos]
    
    if raices:
        raiz_principal = raices[0]
        
        # Función recursiva para navegar por la estructura
        def recorrer_arbol(nodo, ruta_capitulo="", cantidad_acumulada=1.0):
            hijos = jerarquia.get(nodo, [])
            
            if hijos:
                # Es un agrupador (Proyecto, Capítulo o Subcapítulo)
                nombre_nodo = conceptos.get(nodo, {}).get("Descripción", nodo)
                
                if nodo == raiz_principal:
                    nueva_ruta = ""
                else:
                    nueva_ruta = f"{ruta_capitulo} > {nombre_nodo}" if ruta_capitulo else nombre_nodo
                    
                # Llamada recursiva hacia adentro del árbol
                for hijo_cod, cantidad_hijo in hijos:
                    recorrer_arbol(hijo_cod, nueva_ruta, cantidad_acumulada * cantidad_hijo)
            else:
                # Es una Partida u hoja final
                concepto = conceptos.get(nodo, {})
                precio = concepto.get("Precio", 0.0)
                
                datos_extraidos.append({
                    "Capítulo": ruta_capitulo,
                    "Código": nodo,
                    "Ud": concepto.get("Ud", ""),
                    "Descripción": concepto.get("Descripción", ""),
                    "Precio (€)": precio,
                    "Cantidad": round(cantidad_acumulada, 3),
                    "Importe (€)": round(cantidad_acumulada * precio, 2)
                })
        
        # Arrancamos la lectura desde la raíz
        recorrer_arbol(raiz_principal)

    if datos_extraidos:
        df_bc3 = pd.DataFrame(datos_extraidos)
        
        # Mostrar KPIs rápidos
        total_presupuesto = df_bc3['Importe (€)'].sum()
        st.subheader(f"Total Presupuesto: {total_presupuesto:,.2f} €")
        
        st.dataframe(df_bc3, use_container_width=True)
        
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df_bc3.to_excel(writer, index=False, sheet_name="Presupuesto")
        
        st.download_button(
            label="⬇️ Descargar Excel Completo", 
            data=buffer_excel.getvalue(), 
            file_name="Presupuesto_Parseado.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No se ha podido extraer el árbol de partidas. Revisa que sea un BC3 válido.")
