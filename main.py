from fastapi import FastAPI, HTTPException, Query
from pydantic import BaseModel
import pandas as pd
from typing import Optional

# Inicialización de la aplicación FastAPI
app = FastAPI(
    title="API de Gestión de Finanzas",
    description="Software para lectura, consulta y administración de datos financieros desde Excel.",
    version="1.0.0",
)

# Variable global con el nombre del archivo de origen
ARCHIVO_EXCEL = "Finanzas.xlsx"

# Variables globales para almacenar los datos en memoria y los catálogos válidos
df_finanzas = pd.DataFrame()
paises_validos = set()
segmentos_validos = set()
productos_validos = set()


def guardar_excel() -> None:
    """Guarda el DataFrame actual en el archivo Excel manteniendo las columnas sin índice extra."""
    df_para_guardar = df_finanzas.copy()
    if "id" in df_para_guardar.columns:
        df_para_guardar = df_para_guardar.drop(columns=["id"])
    df_para_guardar.to_excel(ARCHIVO_EXCEL, index=False)


def cargar_datos() -> None:
    """
    Lee el archivo Excel al iniciar el servidor, normaliza nombres de columnas y fechas,
    asigna un ID único y extrae los valores válidos iniciales de países, segmentos y productos.
    """
    global df_finanzas, paises_validos, segmentos_validos, productos_validos
    try:
        df = pd.read_excel(ARCHIVO_EXCEL)

        # Limpiar espacios en blanco extra en los nombres de las columnas
        df.columns = df.columns.astype(str).str.strip()

        # Renombrar columnas para estandarizar nombres
        columnas_esperadas = {
            "País": "Pais",
            "Unidades vendidas": "Unidades_vendidas",
            "Precio de manufactura": "Precio_manufactura",
            "Precio de venta": "Precio_venta",
            "Venta brutas": "Venta_brutas",
            "Costo de bienes vendidos": "Costo_bienes",
            "Nombre Mes": "Nombre_Mes",
        }
        df.rename(columns=columnas_esperadas, inplace=True)

        # Asegurar que la columna Fecha se maneje como texto limpio (DD/MM/YYYY)
        if "Fecha" in df.columns:
            df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

        # Asignar un ID único basado en el índice para facilitar búsquedas y eliminaciones
        df.reset_index(drop=True, inplace=True)
        df.insert(0, "id", range(len(df)))

        df_finanzas = df

        # Poblar los conjuntos de valores válidos existentes
        paises_validos = set(df_finanzas["Pais"].dropna().astype(str).str.strip()) - {""}
        segmentos_validos = set(df_finanzas["Segmento"].dropna().astype(str).str.strip()) - {""}
        productos_validos = set(df_finanzas["Producto"].dropna().astype(str).str.strip()) - {""}

    except FileNotFoundError:
        raise RuntimeError(f"Error crítico: No se encontró el archivo '{ARCHIVO_EXCEL}'.")
    except Exception as e:
        raise RuntimeError(f"Error al procesar el archivo Excel: {str(e)}")


# Cargar los datos al arrancar la API
cargar_datos()


# Modelos de datos (Pydantic) para validar las peticiones POST
class NuevaFinanza(BaseModel):
    Segmento: str
    Pais: str
    Producto: str
    Unidades_vendidas: float
    Ventas: float
    Beneficio: float
    Fecha: str  # Formato esperado: DD/MM/YYYY


class NuevoCatalogo(BaseModel):
    tipo: str  # Debe ser: "Pais", "Segmento" o "Producto"
    valor: str  # El nuevo nombre a registrar


# ==========================================
# ENDPOINTS REQUERIDOS POR EL CLIENTE
# ==========================================

# 1) Mostrar todos los países
@app.get("/paises", tags=["Consultas"])
def mostrar_paises():
    """Devuelve la lista completa y ordenada de países registrados en el sistema."""
    return {"paises": sorted(list(paises_validos))}


# 2) Por cada producto, mostrar en qué países se vende
@app.get("/productos/paises", tags=["Consultas"])
def paises_por_producto():
    """Agrupa los datos por producto y lista los países únicos donde se comercializa cada uno."""
    if df_finanzas.empty:
        raise HTTPException(status_code=404, detail="No hay datos financieros cargados.")

    resultado = {}
    agrupado = df_finanzas.groupby("Producto")["Pais"].unique()
    for producto, lista_paises in agrupado.items():
        paises_limpios = [str(p).strip() for p in lista_paises if pd.notna(p) and str(p).strip() != ""]
        resultado[str(producto)] = sorted(list(set(paises_limpios)))

    return {"productos_por_pais": resultado}


# 3) Calcular Beneficio total por País, Segmento o Producto
@app.get("/beneficio-total", tags=["Cálculos"])
def calcular_beneficio_total(
    pais: Optional[str] = Query(None, description="Filtrar por nombre de País"),
    segmento: Optional[str] = Query(None, description="Filtrar por nombre de Segmento"),
    producto: Optional[str] = Query(None, description="Filtrar por nombre de Producto"),
):
    """
    Calcula la suma total de la columna Beneficio.
    Permite filtrar opcionalmente por País, Segmento o Producto.
    """
    if df_finanzas.empty:
        raise HTTPException(status_code=404, detail="No hay datos financieros cargados.")

    df_filtrado = df_finanzas.copy()

    if pais:
        df_filtrado = df_filtrado[
            df_filtrado["Pais"].fillna("").astype(str).str.strip().str.lower() == pais.strip().lower()
        ]
    if segmento:
        df_filtrado = df_filtrado[
            df_filtrado["Segmento"].fillna("").astype(str).str.strip().str.lower() == segmento.strip().lower()
        ]
    if producto:
        df_filtrado = df_filtrado[
            df_filtrado["Producto"].fillna("").astype(str).str.strip().str.lower() == producto.strip().lower()
        ]

    if df_filtrado.empty:
        raise HTTPException(status_code=404, detail="No se encontraron registros con los filtros proporcionados.")

    beneficio_total = float(pd.to_numeric(df_filtrado["Beneficio"], errors="coerce").fillna(0).sum())
    return {
        "filtros_aplicados": {"pais": pais, "segmento": segmento, "producto": producto},
        "registros_encontrados": len(df_filtrado),
        "beneficio_total": round(beneficio_total, 2),
    }


# 4) Mostrar las finanzas por fecha ingresada
@app.get("/finanzas/fecha", tags=["Consultas"])
def finanzas_por_fecha(
    fecha: str = Query(..., description="Fecha a buscar en formato DD/MM/YYYY (ej. 01/01/2014)")
):
    """Filtra y devuelve todos los registros financieros correspondientes a una fecha exacta."""
    if df_finanzas.empty:
        raise HTTPException(status_code=404, detail="No hay datos financieros cargados.")

    fecha_limpia = fecha.strip()
    df_filtrado = df_finanzas[df_finanzas["Fecha"].fillna("").astype(str).str.strip() == fecha_limpia]

    if df_filtrado.empty:
        raise HTTPException(status_code=404, detail=f"No existen registros para la fecha {fecha_limpia}.")

    return {
        "fecha": fecha_limpia,
        "cantidad": len(df_filtrado),
        "finanzas": df_filtrado.fillna("").to_dict(orient="records"),
    }


# 5) Mostrar cuál es la finanza con mayor beneficio y mostrar sus datos
@app.get("/finanzas/mayor-beneficio", tags=["Consultas"])
def finanza_mayor_beneficio():
    """Encuentra el registro con el valor máximo en la columna Beneficio."""
    if df_finanzas.empty:
        raise HTTPException(status_code=404, detail="No hay registros disponibles.")

    beneficios = pd.to_numeric(df_finanzas["Beneficio"], errors="coerce").dropna()
    if beneficios.empty:
        raise HTTPException(status_code=404, detail="No hay valores numéricos de Beneficio disponibles.")

    indice_max = beneficios.idxmax()
    registro_max = df_finanzas.loc[indice_max].fillna("").to_dict()

    return {
        "mensaje": "Registro con mayor beneficio encontrado",
        "datos": registro_max,
    }


# 6) Mostrar cuál fue la finanza que obtuvo mayor ingreso por unidades vendidas
@app.get("/finanzas/mayor-ingreso-unidades", tags=["Consultas"])
def finanza_mayor_ingreso_unidades():
    """
    Calcula la relación entre Ventas y Unidades_vendidas (ingreso unitario generado)
    y devuelve el registro con el mayor valor.
    """
    if df_finanzas.empty:
        raise HTTPException(status_code=404, detail="No hay registros disponibles.")

    col_unidades = "Unidades_vendidas" if "Unidades_vendidas" in df_finanzas.columns else None
    if not col_unidades:
        candidatas = [c for c in df_finanzas.columns if "Unidades" in c]
        if not candidatas:
            raise HTTPException(status_code=400, detail="No se encontró la columna de unidades vendidas.")
        col_unidades = candidatas[0]

    df_temp = df_finanzas.copy()
    df_temp[col_unidades] = pd.to_numeric(df_temp[col_unidades], errors="coerce")
    df_temp["Ventas"] = pd.to_numeric(df_temp["Ventas"], errors="coerce")
    df_temp = df_temp[df_temp[col_unidades] > 0].dropna(subset=["Ventas", col_unidades])

    if df_temp.empty:
        raise HTTPException(status_code=404, detail="No hay registros válidos con unidades vendidas mayores a cero.")

    df_temp["ingreso_por_unidad"] = df_temp["Ventas"] / df_temp[col_unidades]

    indice_max = df_temp["ingreso_por_unidad"].idxmax()
    registro_max = df_temp.loc[indice_max].fillna("").to_dict()
    registro_max["ingreso_por_unidad"] = round(float(registro_max["ingreso_por_unidad"]), 2)

    return {
        "mensaje": "Registro con mayor ingreso por unidad vendida",
        "datos": registro_max,
    }


# 7) Mostrar el top 5 de los países que más ventas han realizado
@app.get("/paises/top-ventas", tags=["Cálculos"])
def top_5_paises_ventas():
    """Agrupa las ventas totales por país y devuelve los 5 mayores."""
    if df_finanzas.empty:
        raise HTTPException(status_code=404, detail="No hay datos disponibles.")

    df_temp = df_finanzas.copy()
    df_temp["Ventas"] = pd.to_numeric(df_temp["Ventas"], errors="coerce").fillna(0)

    top_paises = (
        df_temp.groupby("Pais")["Ventas"]
        .sum()
        .sort_values(ascending=False)
        .head(5)
    )

    ranking = [
        {"puesto": i + 1, "pais": pais, "ventas_totales": round(float(ventas), 2)}
        for i, (pais, ventas) in enumerate(top_paises.items())
    ]

    return {"top_5_paises": ranking}


# 8) Crear una finanza usando países, segmentos y productos ya ingresados
@app.post("/finanzas", status_code=201, tags=["Gestión"])
def crear_finanza(datos: NuevaFinanza):
    """
    Crea un nuevo registro financiero validando previamente que el País,
    Segmento y Producto existan en los catálogos registrados.
    """
    global df_finanzas

    pais_limpio = datos.Pais.strip()
    segmento_limpio = datos.Segmento.strip()
    producto_limpio = datos.Producto.strip()
    fecha_limpia = datos.Fecha.strip()

    if pais_limpio not in paises_validos:
        raise HTTPException(status_code=400, detail=f"El país '{pais_limpio}' no está registrado. Agrégalo primero.")
    if segmento_limpio not in segmentos_validos:
        raise HTTPException(status_code=400, detail=f"El segmento '{segmento_limpio}' no está registrado. Agrégalo primero.")
    if producto_limpio not in productos_validos:
        raise HTTPException(status_code=400, detail=f"El producto '{producto_limpio}' no está registrado. Agrégalo primero.")

    nuevo_id = int(df_finanzas["id"].max() + 1) if not df_finanzas.empty and "id" in df_finanzas.columns else 0

    nueva_fila = {
        "id": nuevo_id,
        "Segmento": segmento_limpio,
        "Pais": pais_limpio,
        "Producto": producto_limpio,
        "Unidades_vendidas": datos.Unidades_vendidas,
        "Ventas": datos.Ventas,
        "Beneficio": datos.Beneficio,
        "Fecha": fecha_limpia,
    }

    df_finanzas = pd.concat([df_finanzas, pd.DataFrame([nueva_fila])], ignore_index=True)
    guardar_excel()

    return {
        "mensaje": "Registro financiero creado exitosamente",
        "finanza_creada": nueva_fila,
    }


# 9) Eliminar finanza
@app.delete("/finanzas/{finanza_id}", tags=["Gestión"])
def eliminar_finanza(finanza_id: int):
    """Elimina un registro financiero existente buscando por su ID."""
    global df_finanzas

    if df_finanzas.empty or "id" not in df_finanzas.columns or finanza_id not in df_finanzas["id"].values:
        raise HTTPException(status_code=404, detail=f"No se encontró ninguna finanza con el ID {finanza_id}.")

    df_finanzas = df_finanzas[df_finanzas["id"] != finanza_id].reset_index(drop=True)
    guardar_excel()

    return {"mensaje": f"Registro financiero con ID {finanza_id} eliminado correctamente."}


# 10) Añadir nuevo País, Segmento o Producto
@app.post("/catalogos", status_code=201, tags=["Gestión"])
def agregar_catalogo(item: NuevoCatalogo):
    """Permite añadir un nuevo País, Segmento o Producto a las listas válidas del sistema."""
    tipo_normalizado = item.tipo.strip().lower()
    valor_limpio = item.valor.strip()

    if not valor_limpio:
        raise HTTPException(status_code=400, detail="El valor no puede estar vacío.")

    if tipo_normalizado in ("pais", "país"):
        if valor_limpio in paises_validos:
            raise HTTPException(status_code=400, detail="Ese país ya existe.")
        paises_validos.add(valor_limpio)
        return {
            "mensaje": f"País '{valor_limpio}' agregado correctamente.",
            "paises_actuales": sorted(list(paises_validos)),
        }

    elif tipo_normalizado == "segmento":
        if valor_limpio in segmentos_validos:
            raise HTTPException(status_code=400, detail="Ese segmento ya existe.")
        segmentos_validos.add(valor_limpio)
        return {
            "mensaje": f"Segmento '{valor_limpio}' agregado correctamente.",
            "segmentos_actuales": sorted(list(segmentos_validos)),
        }

    elif tipo_normalizado == "producto":
        if valor_limpio in productos_validos:
            raise HTTPException(status_code=400, detail="Ese producto ya existe.")
        productos_validos.add(valor_limpio)
        return {
            "mensaje": f"Producto '{valor_limpio}' agregado correctamente.",
            "productos_actuales": sorted(list(productos_validos)),
        }

    else:
        raise HTTPException(status_code=400, detail="Tipo inválido. Debe ser 'Pais', 'Segmento' o 'Producto'.")
        