import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Generador de Proyecto BC3", page_icon="🏗️", layout="wide")
st.title("🏗️ Conversor BC3 a Documentos de Proyecto")
st.markdown("Extrae Mediciones, Cuadro de Precios 1, Cuadro de Precios 2 y Resumen.")

archivo_bc3 = st.file_uploader("Sube aquí tu archivo .bc3", type=["bc3"])

def parsear_numero(valor):
    try:
        if not valor.strip(): return None
        return float(valor.strip().replace(',', '.'))
    except:
        return None

if archivo_bc3 is not None:
    try:
        contenido = archivo_bc3.read().decode('latin-1')
    except:
        contenido = archivo_bc3.read().decode('utf-8')
    
    conceptos = {}
    jerarquia = {}
    mediciones = {}
    
    lineas = contenido.splitlines()
    
    # 1. EXTRACCIÓN DE DATOS BC3
    for linea in lineas:
        # A. Conceptos
        if linea.startswith('~C|'):
            partes = linea.split('|')
            if len(partes) >= 5:
                codigo = partes[1].strip().replace('#', '')
                precio_str = partes[4].strip().replace('\\', '')
                # En FIEBDC, el índice 6 guarda el tipo (0:Raiz, 1:Mano Obra, 2:Maquinaria, 3:Material)
                tipo = partes[6].strip() if len(partes) > 6 else "0"
                conceptos[codigo] = {
                    "Ud": partes[2].strip(),
                    "Descripción": partes[3].strip(),
                    "Precio": parsear_numero(precio_str) or 0.0,
                    "Tipo": tipo
                }
                
        # B. Textos largos (Descripciones completas)
        elif linea.startswith('~T|'):
            partes = linea.split('|')
            if len(partes) >= 3:
                codigo = partes[1].strip().replace('#', '')
                if codigo in conceptos:
                    conceptos[codigo]["Descripción"] = partes[2].strip()

        # C. Jerarquía (Descomposición)
        elif linea.startswith('~D|'):
            partes = linea.split('|')
            if len(partes) >= 3:
                padre = partes[1].strip().replace('#', '')
                datos_hijos = "|".join(partes[2:]).split('\\')
                hijos = []
                for i in range(0, len(datos_hijos) - 2, 3):
                    hijo_cod = datos_hijos[i].strip().replace('#', '')
                    if hijo_cod:
                        cantidad = parsear_numero(datos_hijos[i+2]) or 0.0
                        hijos.append((hijo_cod, cantidad))
                jerarquia[padre] = hijos

        # D. Mediciones Detalladas
        elif linea.startswith('~M|'):
            partes = linea.split('|')
            if len(partes) >= 5:
                # El código viene como Capítulo\Partida, cogemos la Partida
                cod_mix = partes[1].split('\\')
                codigo_partida = cod_mix[-1].strip().replace('#', '')
                
                lineas_med = partes[4].strip()
                if lineas_med.startswith('\\'): lineas_med = lineas_med[1:]
                if lineas_med.endswith('\\'): lineas_med = lineas_med[:-1]
                
                tokens = lineas_med.split('\\')
                lineas_detalle = []
                # FIEBDC agrupa las mediciones en bloques de 5 (Comentario, Uds, Largo, Ancho, Alto)
                for i in range(0, len(tokens), 5):
                    chunk = tokens[i:i+5]
                    if len(chunk) == 5:
                        c = chunk[0].strip()
                        n_vals = [parsear_numero(x) for x in chunk[1:5]]
                        
                        # Cálculo de subtotal
                        factores = [x for x in n_vals if x is not None]
                        if factores:
                            subtotal = 1.0
                            for f in factores: subtotal *= f
                        else:
                            subtotal = 0.0
                            
                        if c or subtotal != 0.0:
                            lineas_detalle.append({
                                "Comentario": c, "Uds": chunk[1], 
                                "Largo": chunk[2], "Ancho": chunk[3], "Alto": chunk[4], 
                                "Subtotal": subtotal
                            })
                mediciones[codigo_partida] = lineas_detalle

    # 2. GENERACIÓN DE ESTRUCTURAS PARA EL EXCEL
    todos_los_hijos = set([h[0] for lista in jerarquia.values() for h in lista])
    raices = [p for p in jerarquia.keys() if p not in todos_los_hijos]
    
    if raices:
        raiz_principal = raices[0]
        
        datos_presupuesto = []
        resumen_capitulos = []
        partidas_unicas = set()

        def procesar_arbol(nodo, ruta="", nivel=0, cod_padre=""):
            hijos = jerarquia.get(nodo, [])
            concepto = conceptos.get(nodo, {})
            nombre = concepto.get("Descripción", nodo)
            precio = concepto.get("Precio", 0.0)
            ud = concepto.get("Ud", "")

            # Es un Capítulo o Subcapítulo
            if hijos and nivel > 0:
                datos_presupuesto.append({
                    "Nivel": f"CAPÍTULO" if nivel==1 else "SUBCAPÍTULO",
                    "Código": nodo,
                    "Ud": "", "Descripción": nombre, "Uds": "", "Largo": "", "Ancho": "", "Alto": "", 
                    "Cantidad": "", "Precio (€)": "", "Importe (€)": ""
                })
                
                importe_capitulo = 0.0
                for h_cod, h_cant in hijos:
                    importe_capitulo += procesar_arbol(h_cod, f"{ruta} > {nombre}" if ruta else nombre, nivel+1, nodo)
                
                if nivel == 1:
                    resumen_capitulos.append({"Capítulo": nodo, "Descripción": nombre, "Importe (€)": importe_capitulo})
                
                return importe_capitulo

            # Es la Raíz
            elif hijos and nivel == 0:
                total_proyecto = 0.0
                for h_cod, h_cant in hijos:
                    total_proyecto += procesar_arbol(h_cod, "", nivel+1, nodo)
                return total_proyecto

            # Es una Partida
            else:
                partidas_unicas.add(nodo)
                
                # Calcular total medición
                lineas_med = mediciones.get(nodo, [])
                total_cantidad = sum(m["Subtotal"] for m in lineas_med) if lineas_med else jerarquia.get(cod_padre, [(nodo, 1.0)])[0][1]
                importe = total_cantidad * precio
                
                # Fila de la Partida
                datos_presupuesto.append({
                    "Nivel": "PARTIDA", "Código": nodo, "Ud": ud, "Descripción": nombre, 
                    "Uds": "", "Largo": "", "Ancho": "", "Alto": "", 
                    "Cantidad": round(total_cantidad, 3), "Precio (€)": round(precio, 2), "Importe (€)": round(importe, 2)
                })
                
                # Filas de Medición
                for m in lineas_med:
                    datos_presupuesto.append({
                        "Nivel": "Medición", "Código": "", "Ud": "", "Descripción": f"  ↳ {m['Comentario']}", 
                        "Uds": m["Uds"], "Largo": m["Largo"], "Ancho": m["Ancho"], "Alto": m["Alto"], 
                        "Cantidad": round(m["Subtotal"], 3), "Precio (€)": "", "Importe (€)": ""
                    })
                return importe

        # Procesar todo el árbol
        pem_total = procesar_arbol(raiz_principal)

        # 3. GENERAR CUADROS DE PRECIOS 1 Y 2
        cuadro_precios_1 = []
        cuadro_precios_2 = []
        
        for cod in sorted(list(partidas_unicas)):
            c = conceptos.get(cod, {})
            precio = c.get("Precio", 0.0)
            
            # Cuadro 1
            cuadro_precios_1.append({
                "Código": cod, "Ud": c.get("Ud", ""), "Descripción": c.get("Descripción", ""), "Precio (€)": round(precio, 2)
            })
            
            # Cuadro 2 (Descomposición)
            mo, mq, mt, rest = 0.0, 0.0, 0.0, 0.0
            for hijo_cod, rendimiento in jerarquia.get(cod, []):
                hijo_obj = conceptos.get(hijo_cod, {})
                tipo = hijo_obj.get("Tipo", "0")
                coste = rendimiento * hijo_obj.get("Precio", 0.0)
                
                if tipo == '1': mo += coste
                elif tipo == '2': mq += coste
                elif tipo == '3': mt += coste
                else: rest += coste
                
            cuadro_precios_2.append({
                "Código": cod, "Ud": c.get("Ud", ""), "Descripción": c.get("Descripción", ""), 
                "Mano de Obra (€)": round(mo, 2), "Maquinaria (€)": round(mq, 2), 
                "Materiales (€)": round(mt, 2), "Resto/Auxiliares (€)": round(rest, 2),
                "Precio Total (€)": round(precio, 2)
            })

        # 4. PREPARAR RESUMEN CON KPIs
        df_resumen = pd.DataFrame(resumen_capitulos)
        gg = pem_total * 0.13
        bi = pem_total * 0.06
        base_imponible = pem_total + gg + bi
        iva = base_imponible * 0.21
        total_contrata = base_imponible + iva

        # 5. CREACIÓN DEL EXCEL MULTIPESTAÑA
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            
            # Pestaña 1: Resumen
            df_res_export = pd.concat([df_resumen, pd.DataFrame([
                {"Capítulo": "", "Descripción": "PRESUPUESTO EJECUCIÓN MATERIAL (PEM)", "Importe (€)": round(pem_total, 2)},
                {"Capítulo": "", "Descripción": "13% Gastos Generales (GG)", "Importe (€)": round(gg, 2)},
                {"Capítulo": "", "Descripción": "6% Beneficio Industrial (BI)", "Importe (€)": round(bi, 2)},
                {"Capítulo": "", "Descripción": "PRESUPUESTO BASE LICITACIÓN S/IVA", "Importe (€)": round(base_imponible, 2)},
                {"Capítulo": "", "Descripción": "21% IVA", "Importe (€)": round(iva, 2)},
                {"Capítulo": "", "Descripción": "TOTAL PRESUPUESTO CONTRATA", "Importe (€)": round(total_contrata, 2)}
            ])], ignore_index=True)
            df_res_export.to_excel(writer, index=False, sheet_name="Resumen")

            # Pestaña 2: Mediciones y Presupuesto
            df_mediciones = pd.DataFrame(datos_presupuesto)
            df_mediciones.to_excel(writer, index=False, sheet_name="Presupuesto y Mediciones")
            
            # Pestaña 3: Cuadro de Precios 1
            pd.DataFrame(cuadro_precios_1).to_excel(writer, index=False, sheet_name="Cuadro Precios 1")
            
            # Pestaña 4: Cuadro de Precios 2
            pd.DataFrame(cuadro_precios_2).to_excel(writer, index=False, sheet_name="Cuadro Precios 2")

        # 6. RENDERIZADO EN STREAMLIT
        st.success("✅ Archivo BC3 procesado con éxito.")
        
        # Mostrar KPIs tipo dashboard
        col1, col2, col3 = st.columns(3)
        col1.metric("Presupuesto Ejecución Material (PEM)", f"{pem_total:,.2f} €")
        col2.metric("Base Licitación (PEM + GG + BI)", f"{base_imponible:,.2f} €")
        col3.metric("Total Contrata (Con IVA)", f"{total_contrata:,.2f} €")
        
        st.subheader("📊 Resumen por Capítulos")
        st.dataframe(df_resumen, use_container_width=True)

        st.download_button(
            label="⬇️ Descargar Documentos Proyecto (Excel Multipestaña)", 
            data=buffer_excel.getvalue(), 
            file_name="Documentos_Proyecto.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.warning("⚠️ No se ha podido extraer el árbol. Comprueba el formato BC3.")
