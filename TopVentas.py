import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment

# Configuración de la interfaz
st.set_page_config(page_title="Generador Top 50 Ventas ES", layout="centered")

st.title("🚀 Generador de Top 50 Ventas")
st.markdown(f"""
### Reglas de Negocio Aplicadas:
1. **Prioridad de Orden:** Se respeta estrictamente el orden del fichero de **Ventas**.
2. **Filtros Críticos:** Solo productos con **Stock > 5**, que no empiecen por 'V' y no estén en **Excluidos**.
3. **Búsqueda Inteligente:** * **EAN:** Tarifa (Col F) ➔ Respaldo en EANs (Col B).
    * **PVPR:** Tarifa (Col G) ➔ Respaldo en Feed_España (Col Q).
""")

st.divider()

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)
with col1:
    file_ventas = st.file_uploader("1. VENTAS (Excel)", type=["xlsx"])
    file_stock = st.file_uploader("2. STOCK (CSV)", type=["csv"])
    file_excluidos = st.file_uploader("3. EXCLUIDOS (Excel)", type=["xlsx"])
with col2:
    file_tarifa = st.file_uploader("4. TARIFA (CSV/Excel)", type=["xlsx", "csv"])
    file_feed = st.file_uploader("5. feed_España (Excel)", type=["xlsx"])
    file_ean_extra = st.file_uploader("6. EANs (Excel)", type=["xlsx"])

def normalize_sku(series):
    """Limpia SKUs para asegurar que coincidan (texto, sin .0, 5 dígitos si es número)"""
    s = series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return s.apply(lambda x: x.zfill(5) if x.isdigit() else x.upper())

if st.button("🚀 Generar Reporte con Orden de Ventas", use_container_width=True):
    if not all([file_ventas, file_stock, file_excluidos, file_tarifa, file_feed, file_ean_extra]):
        st.error("⚠️ Debes subir los 6 archivos para procesar.")
    else:
        try:
            with st.spinner('Procesando manteniendo el orden de Ventas...'):
                # 1. CARGA DE BASES DE DATOS
                df_v = pd.read_excel(file_ventas)
                # Guardamos el orden original creando una columna de índice
                df_v['ORDEN_ORIGINAL'] = range(len(df_v))
                sku_v_name = df_v.columns[0]
                df_v['SKU_JOIN'] = normalize_sku(df_v[sku_v_name])

                # Stock (Col G es índice 6)
                df_stk = pd.read_csv(file_stock, sep=None, engine='python')
                df_stk['SKU_JOIN'] = normalize_sku(df_stk.iloc[:, 0])
                col_stk_val = df_stk.columns[6]
                df_stk[col_stk_val] = pd.to_numeric(df_stk[col_stk_val], errors='coerce').fillna(0)
                df_stk_clean = df_stk.drop_duplicates('SKU_JOIN')[['SKU_JOIN', col_stk_val]]

                # Excluidos
                df_exc = pd.read_excel(file_excluidos)
                list_exc = normalize_sku(df_exc.iloc[:, 0]).unique().tolist()

                # Tarifa (SKU: E[4], EAN: F[5], PVPR: G[6])
                df_t = pd.read_csv(file_tarifa, sep=None, engine='python') if file_tarifa.name.endswith('.csv') else pd.read_excel(file_tarifa)
                df_t_clean = df_t.iloc[:, [4, 5, 6]].copy()
                df_t_clean.columns = ['SKU_JOIN', 'EAN_T', 'PVP_T']
                df_t_clean['SKU_JOIN'] = normalize_sku(df_t_clean['SKU_JOIN'])
                df_t_clean = df_t_clean.drop_duplicates('SKU_JOIN')

                # Feed España (MPN: M[12], Price: Q[16])
                df_f = pd.read_excel(file_feed)
                df_f_clean = df_f.iloc[:, [12, 16]].copy()
                df_f_clean.columns = ['SKU_JOIN', 'PVP_F']
                df_f_clean['SKU_JOIN'] = normalize_sku(df_f_clean['SKU_JOIN'])
                df_f_clean = df_f_clean.drop_duplicates('SKU_JOIN')

                # EANs Extra (SKU: A[0], EAN: B[1])
                df_e = pd.read_excel(file_ean_extra)
                df_e_clean = df_e.iloc[:, [0, 1]].copy()
                df_e_clean.columns = ['SKU_JOIN', 'EAN_E']
                df_e_clean['SKU_JOIN'] = normalize_sku(df_e_clean['SKU_JOIN'])
                df_e_clean = df_e_clean.drop_duplicates('SKU_JOIN')

                # --- PROCESO DE FILTRADO MANTENIENDO ORDEN ---
                # Unimos Stock a Ventas para filtrar
                df_res = pd.merge(df_v, df_stk_clean, on='SKU_JOIN', how='left')
                
                # Aplicamos condiciones de descarte
                df_res = df_res[df_res[col_stk_val] > 5] # REGLA: Stock > 5
                df_res = df_res[~df_res['SKU_JOIN'].isin(list_exc)] # REGLA: No Excluidos
                df_res = df_res[~df_res['SKU_JOIN'].str.startswith('V')] # REGLA: No empieza por V
                
                # Ordenamos por el índice original para asegurar que el orden no cambió
                df_res = df_res.sort_values('ORDEN_ORIGINAL')

                # --- ENRIQUECIMIENTO DE DATOS (Cruce de info) ---
                df_res = pd.merge(df_res, df_t_clean, on='SKU_JOIN', how='left')
                df_res = pd.merge(df_res, df_f_clean, on='SKU_JOIN', how='left')
                df_res = pd.merge(df_res, df_e_clean, on='SKU_JOIN', how='left')

                # Lógica de cascada (Rescate)
                df_res['EAN_FINAL'] = df_res['EAN_T'].fillna(df_res['EAN_E'])
                df_res['PVP_FINAL'] = df_res['PVP_T'].fillna(df_res['PVP_F'])

                # Tomamos los 50 primeros que han sobrevivido a los filtros
                top_50 = df_res.head(50).copy()

                # Construcción tabla final
                final_df = pd.DataFrame({
                    "SKU": top_50['SKU_JOIN'],
                    "EAN": pd.to_numeric(top_50['EAN_FINAL'], errors='coerce').fillna(0).astype(int).astype(str).replace("0", ""),
                    "Título del Producto": top_50.iloc[:, 2], # Col C de Ventas
                    "Familia": top_50.iloc[:, 3],           # Col D de Ventas
                    "Subfamilia": top_50.iloc[:, 4],        # Col E de Ventas
                    "PVPR": pd.to_numeric(top_50['PVP_FINAL'], errors='coerce').fillna(0)
                })

                st.success(f"✅ Se han encontrado {len(final_df)} registros que cumplen los criterios en el orden original.")
                st.dataframe(final_df.style.format({"PVPR": "{:.2f} €"}), use_container_width=True)

                # --- EXPORTACIÓN CON DISEÑO ---
                today = datetime.now().strftime('%Y%m%d')
                name_file = f"{today}_50TopVentasES.xlsx"
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Top50')
                    ws = writer.sheets['Top50']
                    
                    # Estilo Encabezado: Azul (1F4E78), Blanco, Negrita, Centrado
                    blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    white_font = Font(color="FFFFFF", bold=True)
                    
                    for cell in ws[1]:
                        cell.fill = blue_fill
                        cell.font = white_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                    # Formato Moneda y Ajuste
                    for row in range(2, len(final_df) + 2):
                        ws[f'F{row}'].number_format = '#,##0.00 €'
                        ws[f'B{row}'].number_format = '0' # EAN
                    
                    # Ajuste de ancho automático
                    for col in ws.columns:
                        ws.column_dimensions[col[0].column_letter].width = 22

                st.download_button(f"📥 Descargar {name_file}", output.getvalue(), file_name=name_file, use_container_width=True)

        except Exception as e:
            st.error(f"Error en el procesado: {e}")