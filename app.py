"""PCP & Telemetria de Produção — Python 3.11+, Streamlit e Google Sheets.

Execute: python -m streamlit run app.py
O módulo contém dados, regras de negócio, persistência e interface completos.
"""

from __future__ import annotations

import hashlib
import html
import io
import json
import math
import re
import secrets as secure_random
import threading
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from numbers import Real
from typing import Any, Mapping
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials
from openpyxl.styles import Alignment, Font, PatternFill


APP_TITLE = "PCP & Telemetria de Produção"
ONLINE_DEPLOYMENT = False
COLORS = {"ink": "#111318", "muted": "#5A606D", "graphite": "#2D333B",
          "accent": "#C84B31", "ok": "#1B6B5A", "warning": "#B45309",
          "danger": "#B91C1C", "line": "#E2E4E8"}
STATUS = ("Programada", "Em_Processo", "Aguardando_Laudo", "Liberada", "Bloqueada", "Expedida")
QUEUE_STATUS = ("Programada", "Em_Processo")
FAMILIES = {"PA-ALC-325": "Alumina_Padrao", "PA-ALT-014": "Alumina_Padrao",
            "PA-WHT-900": "Brancos_Whitelum", "PA-FRS-75": "FerSiN", "PA-ALH-500": "Quimicos"}
DEFAULT_PASSWORD_HASH = ""
NAVIGATION = ("Torre de Controle PCP", "Motor MRP & Receitas", "Sequenciador de Produção",
              "Simulador de Prazo ATP", "Quarentena & Qualidade")
SEED_CSV = {
    "Estoque_Atual": """SKU,Descricao,Categoria,Unidade,Saldo_Disponivel,Saldo_Quarentena,Estoque_Minimo
MP-BAUX-01,Bauxita Calcinada Grau Refratario,Materia_Prima,ton,145.0,0.0,80.0
MP-ALUM-H02,Hidrato de Alumina Grau Tecnico,Materia_Prima,ton,62.5,15.0,50.0
MP-SIL-03,Silicio Metalico Grau Quimico,Materia_Prima,ton,18.0,0.0,25.0
MP-ADIT-04,Aditivo Dispersante Policarboxilato,Materia_Prima,kg,420.0,0.0,300.0
MP-ALV-05,Agente Branqueador Mineral Grau A,Materia_Prima,kg,190.0,50.0,200.0
EMB-BB-1000,Big Bag Rafia Laminada 1000kg Valvulado,Embalagem,un,180,0,150
EMB-SV-25,Saco Valvulado Kraft Multi-folhado 25kg,Embalagem,un,1400,0,2000
EMB-PAL-PBR,Palete Madeira Tratada PBR 1.00x1.20m,Embalagem,un,210,0,120
PA-ALC-325,Alumina Calcinada Moida Malha 325,Produto_Acabado,ton,38.0,12.0,30.0
PA-ALT-014,Alumina Tabular Grao 0-1mm Alta Pureza,Produto_Acabado,ton,14.5,18.0,25.0
PA-WHT-900,Whitelum Carga Mineral Ultra Branca,Produto_Acabado,ton,22.0,8.5,20.0
PA-FRS-75,FerSiN Nitreto de Ferro Silicio Granulado,Produto_Acabado,ton,9.0,14.0,15.0
PA-ALH-500,Alumina Hidratada Micronizada Retardante,Produto_Acabado,ton,44.0,0.0,35.0
""",
    "Fichas_Tecnicas_BOM": """SKU_Pai,SKU_Insumo,Qtd_Por_Ton,Perda_Tecnica_Pct
PA-ALC-325,MP-ALUM-H02,1.04,3.5
PA-ALC-325,MP-ADIT-04,2.50,1.0
PA-ALC-325,EMB-BB-1000,1.00,0.5
PA-ALC-325,EMB-PAL-PBR,1.00,0.0
PA-ALT-014,MP-BAUX-01,1.02,2.5
PA-ALT-014,EMB-BB-1000,1.00,0.5
PA-ALT-014,EMB-PAL-PBR,1.00,0.0
PA-WHT-900,MP-ALUM-H02,0.98,2.0
PA-WHT-900,MP-ALV-05,12.00,1.5
PA-WHT-900,EMB-SV-25,40.00,1.5
PA-WHT-900,EMB-PAL-PBR,1.00,0.0
PA-FRS-75,MP-SIL-03,0.78,3.0
PA-FRS-75,MP-BAUX-01,0.22,2.0
PA-FRS-75,EMB-BB-1000,1.00,0.5
PA-ALH-500,MP-ALUM-H02,1.00,1.8
PA-ALH-500,MP-ADIT-04,1.80,1.0
PA-ALH-500,EMB-SV-25,40.00,1.5
PA-ALH-500,EMB-PAL-PBR,1.00,0.0
""",
    "Ordens_Producao": """Num_OP,Cliente,SKU_Produto,Qtd_Ton,Familia_Setup,Data_Entrega,Status,Lote
OP-2026-1041,Magnesita Refratarios S.A.,PA-ALT-014,18.0,Alumina_Padrao,2026-11-17,Aguardando_Laudo,LT-2611-088
OP-2026-1042,Gerdau Acominas S.A.,PA-FRS-75,14.0,FerSiN,2026-11-18,Aguardando_Laudo,LT-2611-091
OP-2026-1043,Suvinil Tintas e Vernizes,PA-WHT-900,16.0,Brancos_Whitelum,2026-11-19,Em_Processo,LT-2611-094
OP-2026-1044,Saint-Gobain Ceramicas,PA-ALC-325,28.0,Alumina_Padrao,2026-11-20,Programada,LT-2611-095
OP-2026-1045,Usiminas Siderurgia,PA-FRS-75,22.0,FerSiN,2026-11-21,Programada,LT-2611-096
OP-2026-1046,RHI Magnesita Brasil,PA-ALT-014,30.0,Alumina_Padrao,2026-11-22,Programada,LT-2611-097
OP-2026-1047,Coral AkzoNobel,PA-WHT-900,20.0,Brancos_Whitelum,2026-11-23,Programada,LT-2611-098
OP-2026-1048,ArcelorMittal Tubarao,PA-FRS-75,18.0,FerSiN,2026-11-24,Programada,LT-2611-099
OP-2026-1049,Prysmian Cabos e Sistemas,PA-ALH-500,25.0,Quimicos,2026-11-25,Programada,LT-2611-100
OP-2026-1050,Ceramica Portobello S.A.,PA-ALC-325,20.0,Alumina_Padrao,2026-11-26,Liberada,LT-2611-085
""",
    "Capacidade_Linhas": """Linha_Equipamento,Familia_Compativel,Capacidade_Ton_Hora,Tempo_Setup_Horas
Moinho Pendular Raymond 01,Alumina_Padrao,3.5,4.0
Moinho Classificador Ultra-Fino 02,Brancos_Whitelum,2.2,6.0
Misturador Intensivo e Britagem 03,FerSiN,2.8,5.0
Micronizador a Jato 04,Quimicos,1.8,3.5
""",
}
SCHEMAS = {name: text.splitlines()[0].split(",") for name, text in SEED_CSV.items()}
NUMERIC = {
    "Estoque_Atual": ("Saldo_Disponivel", "Saldo_Quarentena", "Estoque_Minimo"),
    "Fichas_Tecnicas_BOM": ("Qtd_Por_Ton", "Perda_Tecnica_Pct"),
    "Ordens_Producao": ("Qtd_Ton",),
    "Capacidade_Linhas": ("Capacidade_Ton_Hora", "Tempo_Setup_Horas"),
}


class DataError(ValueError):
    """Dados inconsistentes que precisam de correção pelo operador."""


def local_now() -> datetime:
    try:
        return datetime.now(ZoneInfo("America/Sao_Paulo"))
    except ZoneInfoNotFoundError:
        return datetime.now(timezone(timedelta(hours=-3)))


def secret(key: str, default: Any = None) -> Any:
    try:
        return st.secrets.get(key, default)
    except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
        return default


def production_mode() -> bool:
    return ONLINE_DEPLOYMENT or bool(secret("production_mode", False))


def seed_tables() -> dict[str, pd.DataFrame]:
    result = {}
    for name, csv in SEED_CSV.items():
        frame = pd.read_csv(io.StringIO(csv), dtype=str, keep_default_na=False)
        frame.index = pd.RangeIndex(2, len(frame) + 2, name="sheet_row")
        result[name] = frame
    return validate_tables(result)


def parse_number(value: Any) -> float:
    if isinstance(value, str):
        value = value.strip()
        if "," in value:
            value = value.replace(".", "").replace(",", ".")
    return float(value)


def validate_tables(tables: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    clean = {}
    for name, columns in SCHEMAS.items():
        if name not in tables or list(tables[name].columns) != columns:
            raise DataError(f"A aba {name} precisa dos cabeçalhos, nesta ordem: {', '.join(columns)}.")
        frame = tables[name].copy(deep=True)
        for col in columns:
            if col not in NUMERIC[name]:
                frame[col] = frame[col].astype(str).str.strip()
                if frame[col].eq("").any() or frame[col].eq("nan").any():
                    raise DataError(f"A aba {name} possui valor vazio em {col}.")
        for col in NUMERIC[name]:
            try:
                frame[col] = frame[col].map(parse_number).astype(float)
            except (ValueError, TypeError, OverflowError) as exc:
                raise DataError(f"Valor numérico inválido em {name}.{col}.") from exc
            if not frame[col].map(math.isfinite).all() or frame[col].lt(0).any():
                raise DataError(f"{name}.{col} precisa conter números finitos e não negativos.")
        clean[name] = frame
    stock, bom, orders, lines = (clean[n] for n in SCHEMAS)
    for frame, col in ((stock, "SKU"), (orders, "Num_OP"), (orders, "Lote"),
                       (lines, "Linha_Equipamento"), (lines, "Familia_Compativel")):
        if frame[col].duplicated().any():
            raise DataError(f"Identificador duplicado em {col}.")
    if bom.duplicated(["SKU_Pai", "SKU_Insumo"]).any():
        raise DataError("Há componentes repetidos para o mesmo produto na ficha técnica.")
    if stock.empty or lines.empty or bom.empty:
        raise DataError("Estoque, fichas técnicas e capacidade precisam de registros cadastrados.")
    if not stock["Categoria"].isin(["Materia_Prima", "Embalagem", "Produto_Acabado"]).all():
        raise DataError("Categoria de estoque inválida.")
    if not stock["Unidade"].isin(["ton", "kg", "un"]).all():
        raise DataError("As unidades de estoque devem ser ton, kg ou un.")
    finished = stock[stock["Categoria"].eq("Produto_Acabado")]
    if not finished["Unidade"].eq("ton").all():
        raise DataError("Produtos acabados precisam utilizar a unidade ton.")
    if not orders["Status"].isin(STATUS).all():
        raise DataError("Status de OP inválido. Use: " + ", ".join(STATUS))
    if orders["Qtd_Ton"].le(0).any() or bom["Qtd_Por_Ton"].le(0).any():
        raise DataError("Quantidades de OP e consumo por tonelada devem ser maiores que zero.")
    if lines["Capacidade_Ton_Hora"].le(0).any():
        raise DataError("Capacidade das linhas deve ser maior que zero.")
    if not bom["SKU_Insumo"].isin(stock["SKU"]).all():
        raise DataError("Há um insumo da ficha técnica ausente no estoque.")
    if not bom["SKU_Pai"].isin(finished["SKU"]).all():
        raise DataError("Todo SKU pai da ficha técnica deve ser um produto acabado.")
    if bom["SKU_Insumo"].isin(finished["SKU"]).any():
        raise DataError("Esta versão usa BOM de nível único; produtos acabados não podem ser insumos.")
    if not orders["SKU_Produto"].isin(finished["SKU"]).all():
        raise DataError("Há uma OP com produto ausente no estoque de produtos acabados.")
    if not orders["SKU_Produto"].isin(bom["SKU_Pai"]).all():
        raise DataError("Há uma OP sem ficha técnica cadastrada.")
    try:
        orders["Data_Entrega"] = pd.to_datetime(orders["Data_Entrega"], format="%Y-%m-%d", errors="raise").dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError) as exc:
        raise DataError("Data_Entrega deve ser uma data válida no formato AAAA-MM-DD.") from exc
    if not orders["Familia_Setup"].isin(lines["Familia_Compativel"]).all():
        raise DataError("Há uma família de OP sem linha compatível.")
    if orders.groupby("SKU_Produto")["Familia_Setup"].nunique().gt(1).any():
        raise DataError("Cada produto deve estar associado a uma única família de setup.")
    for sku, expected in FAMILIES.items():
        if not orders.loc[orders["SKU_Produto"].eq(sku), "Familia_Setup"].eq(expected).all():
            raise DataError(f"Família incompatível para {sku}: use {expected}.")
    return clean


@st.cache_resource(show_spinner=False)
def cloud_connection(sheet_url: str, credentials_json: str) -> gspread.Spreadsheet:
    info = json.loads(credentials_json)
    info["private_key"] = info.get("private_key", "").replace("\\n", "\n")
    credentials = Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(credentials)
    client.set_timeout(30)
    return client.open_by_url(sheet_url)


def cloud_tables(sheet: gspread.Spreadsheet) -> dict[str, pd.DataFrame]:
    ranges = [f"'{name}'!A1:{gspread.utils.rowcol_to_a1(1, len(cols))[:-1]}" for name, cols in SCHEMAS.items()]
    response = sheet.values_batch_get(ranges, params={"valueRenderOption": "UNFORMATTED_VALUE", "dateTimeRenderOption": "FORMATTED_STRING"})
    result = {}
    for name, value_range in zip(SCHEMAS, response.get("valueRanges", [])):
        values = value_range.get("values", [])
        if not values or values[0] != SCHEMAS[name]:
            raise DataError(f"Cabeçalho da aba {name} ausente ou incorreto. Use o modelo de quatro abas.")
        rows, positions = [], []
        for sheet_row, row in enumerate(values[1:], start=2):
            if not any(str(cell).strip() for cell in row):
                continue
            if len(row) > len(SCHEMAS[name]):
                raise DataError(f"Número de colunas inválido na aba {name}.")
            rows.append(row + [""] * (len(SCHEMAS[name]) - len(row)))
            positions.append(sheet_row)
        result[name] = pd.DataFrame(rows, columns=SCHEMAS[name], index=pd.Index(positions, name="sheet_row"))
    return validate_tables(result)


@st.cache_data(ttl=60, show_spinner=False)
def cached_cloud_tables(sheet_url: str, credentials_json: str) -> dict[str, pd.DataFrame]:
    return cloud_tables(cloud_connection(sheet_url, credentials_json))


@st.cache_resource(show_spinner=False)
def write_lock(connection_id: str) -> Any:
    return threading.RLock()


def cell_update(sheet_id: int, row: int, column: int, value: Any) -> dict:
    value_key = "numberValue" if isinstance(value, (int, float)) else "stringValue"
    return {"updateCells": {"range": {"sheetId": sheet_id, "startRowIndex": row - 1,
            "endRowIndex": row, "startColumnIndex": column - 1, "endColumnIndex": column},
            "rows": [{"values": [{"userEnteredValue": {value_key: value}}]}], "fields": "userEnteredValue"}}


def family_for(sku: str, tables: dict[str, pd.DataFrame]) -> str:
    previous = tables["Ordens_Producao"].loc[lambda x: x["SKU_Produto"].eq(sku), "Familia_Setup"]
    family = previous.iloc[0] if not previous.empty else FAMILIES.get(sku)
    if not family or family not in tables["Capacidade_Linhas"]["Familia_Compativel"].values:
        raise DataError(f"Não há família/linha definida para {sku}. Cadastre uma OP de referência na planilha.")
    return family


def approval_result(tables: dict[str, pd.DataFrame], op: str) -> tuple[dict[str, pd.DataFrame], dict]:
    result = {k: v.copy(deep=True) for k, v in tables.items()}
    orders, stock = result["Ordens_Producao"], result["Estoque_Atual"]
    selected = orders[orders["Num_OP"].eq(op)]
    if selected.empty:
        raise DataError("OP não encontrada. Atualize a tela.")
    row = selected.iloc[0]
    if row["Status"] != "Aguardando_Laudo":
        raise DataError("O lote já foi tratado ou não está aguardando laudo. Nenhum saldo foi alterado.")
    stock_id = stock.index[stock["SKU"].eq(row["SKU_Produto"])][0]
    before = float(stock.at[stock_id, "Saldo_Quarentena"])
    quantity = float(row["Qtd_Ton"])
    stock.at[stock_id, "Saldo_Quarentena"] = max(0.0, before - quantity)
    stock.at[stock_id, "Saldo_Disponivel"] += quantity
    orders.at[selected.index[0], "Status"] = "Liberada"
    return result, {"op_row": int(selected.index[0]), "stock_row": int(stock_id),
                    "quantity": quantity, "shortfall": max(0.0, quantity - before), "sku": row["SKU_Produto"]}


@dataclass
class Repository:
    mode: str
    sheet_url: str = ""
    credentials_json: str = ""

    @classmethod
    def configured(cls) -> "Repository":
        account, url = secret("gcp_service_account"), str(secret("sheet_url", "")).strip()
        if not account or not url:
            if production_mode():
                raise DataError("O acesso online está aguardando a conexão com o Google Sheets. Configure a conta de serviço e a URL; o aplicativo não gravará seus dados em memória no modo de produção.")
            return cls("sandbox")
        return cls("cloud", url, json.dumps(dict(account), sort_keys=True))

    @property
    def connection_id(self) -> str:
        return hashlib.sha256((self.sheet_url + self.credentials_json).encode()).hexdigest()

    def load(self, fresh: bool = False) -> dict[str, pd.DataFrame]:
        if self.mode == "sandbox":
            if "sandbox_tables" not in st.session_state:
                st.session_state["sandbox_tables"] = seed_tables()
            return {k: v.copy(deep=True) for k, v in st.session_state["sandbox_tables"].items()}
        if fresh:
            return cloud_tables(cloud_connection(self.sheet_url, self.credentials_json))
        return cached_cloud_tables(self.sheet_url, self.credentials_json)

    def commit(self, updated: dict[str, pd.DataFrame], changes: list[tuple[str, int, str, Any]]) -> None:
        if self.mode == "sandbox":
            st.session_state["sandbox_tables"] = updated
            st.cache_data.clear()
            return
        sheet = cloud_connection(self.sheet_url, self.credentials_json)
        sheet_ids = {worksheet.title: worksheet.id for worksheet in sheet.worksheets()}
        requests = [cell_update(sheet_ids[name], row, SCHEMAS[name].index(column) + 1, value)
                    for name, row, column, value in changes]
        try:
            sheet.batch_update({"requests": requests})
        finally:
            st.cache_data.clear()

    def record(self, action: str, detail: str) -> None:
        event = {"Data_Hora": local_now().isoformat(timespec="seconds"), "Acao": action,
                 "Detalhe": detail, "Modo": self.mode}
        st.session_state.setdefault("event_log", []).insert(0, event)
        st.session_state["event_log"] = st.session_state["event_log"][:300]
        st.session_state["flash"] = detail
        st.session_state["last_write"] = local_now().strftime("%d/%m/%Y %H:%M:%S")

    def add_order(self, payload: dict[str, Any]) -> None:
        with write_lock(self.connection_id):
            tables = self.load(fresh=True)
            orders = tables["Ordens_Producao"]
            values = {col: payload[col] for col in SCHEMAS["Ordens_Producao"]}
            if orders["Num_OP"].eq(values["Num_OP"]).any() or orders["Lote"].eq(values["Lote"]).any():
                raise DataError("Número de OP ou lote já cadastrado. Gere outro identificador.")
            if values["Status"] != "Programada":
                raise DataError("Novas OPs devem iniciar como Programada.")
            if values["Familia_Setup"] != family_for(str(values["SKU_Produto"]), tables):
                raise DataError("Família de setup incompatível com o produto.")
            new_index = int(max(orders.index, default=1)) + 1
            tables["Ordens_Producao"] = pd.concat([orders, pd.DataFrame([values], index=[new_index])])
            tables = validate_tables(tables)
            if self.mode == "sandbox":
                st.session_state["sandbox_tables"] = tables
                st.cache_data.clear()
            else:
                worksheet = cloud_connection(self.sheet_url, self.credentials_json).worksheet("Ordens_Producao")
                try:
                    worksheet.append_row([values[c] for c in SCHEMAS["Ordens_Producao"]], value_input_option="RAW", table_range="A1:H")
                finally:
                    st.cache_data.clear()
            self.record("NOVA_OP", f"{values['Num_OP']} cadastrada · {fmt(values['Qtd_Ton'])} ton · {values['SKU_Produto']}.")

    def approve(self, op: str) -> None:
        with write_lock(self.connection_id):
            updated, info = approval_result(self.load(fresh=True), op)
            stock = updated["Estoque_Atual"]
            changes = [("Ordens_Producao", info["op_row"], "Status", "Liberada"),
                       ("Estoque_Atual", info["stock_row"], "Saldo_Disponivel", float(stock.at[info["stock_row"], "Saldo_Disponivel"])),
                       ("Estoque_Atual", info["stock_row"], "Saldo_Quarentena", float(stock.at[info["stock_row"], "Saldo_Quarentena"]))]
            self.commit(updated, changes)
            detail = f"{op}: lote liberado; {fmt(info['quantity'])} ton transferidas para disponível."
            if info["shortfall"] > 0:
                detail += f" Divergência de {fmt(info['shortfall'])} ton na quarentena; saldo limitado a zero conforme regra solicitada."
            self.record("LIBERACAO_LOTE", detail)

    def change_status(self, op: str, target: str) -> None:
        with write_lock(self.connection_id):
            tables = self.load(fresh=True)
            orders = tables["Ordens_Producao"]
            selected = orders[orders["Num_OP"].eq(op)]
            if selected.empty:
                raise DataError("OP não encontrada.")
            row = selected.iloc[0]
            allowed = transitions(row["Status"])
            if target not in allowed:
                raise DataError("Transição de status inválida ou OP alterada por outro operador. Atualize a tela.")
            if row["Status"] == "Aguardando_Laudo" and target == "Liberada":
                self.approve(op)
                return
            orders.at[selected.index[0], "Status"] = target
            self.commit(tables, [("Ordens_Producao", int(selected.index[0]), "Status", target)])
            self.record("STATUS_OP", f"{op}: {label(row['Status'])} → {label(target)}.")


def transitions(status: str) -> list[str]:
    return {"Programada": ["Em_Processo", "Bloqueada"],
            "Em_Processo": ["Aguardando_Laudo", "Bloqueada"],
            "Aguardando_Laudo": ["Liberada", "Bloqueada"],
            "Liberada": ["Expedida", "Bloqueada"], "Bloqueada": [], "Expedida": []}.get(status, [])


def calculate_mrp(tables: dict[str, pd.DataFrame], include_in_process: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    statuses = QUEUE_STATUS if include_in_process else ("Programada",)
    orders = tables["Ordens_Producao"].loc[lambda x: x["Status"].isin(statuses)]
    detail = orders.merge(tables["Fichas_Tecnicas_BOM"], left_on="SKU_Produto", right_on="SKU_Pai", how="left", validate="many_to_many")
    if not detail.empty and detail["SKU_Insumo"].isna().any():
        raise DataError("Uma OP não possui BOM; o MRP foi interrompido para evitar subestimar a demanda.")
    detail["Necessidade_Bruta"] = detail["Qtd_Ton"] * detail["Qtd_Por_Ton"] * (1 + detail["Perda_Tecnica_Pct"] / 100)
    demand = detail.groupby("SKU_Insumo", as_index=False)["Necessidade_Bruta"].sum()
    stock = tables["Estoque_Atual"].loc[lambda x: x["Categoria"].ne("Produto_Acabado")]
    result = stock.merge(demand, left_on="SKU", right_on="SKU_Insumo", how="left").drop(columns="SKU_Insumo")
    result["Necessidade_Bruta"] = result["Necessidade_Bruta"].fillna(0.0)
    result["Saldo_Liquido_Projetado"] = result["Saldo_Disponivel"] - result["Necessidade_Bruta"]
    result["Necessidade_Compra_Urgente"] = (-result["Saldo_Liquido_Projetado"]).clip(lower=0)
    result["Compra_Com_Reserva_Minima"] = (result["Estoque_Minimo"] - result["Saldo_Liquido_Projetado"]).clip(lower=0)
    # Frações de embalagem são mantidas no MRP; a requisição arredonda a compra para unidades inteiras.
    result["Compra_Sugerida"] = result.apply(lambda r: math.ceil(r["Compra_Com_Reserva_Minima"] - 1e-9)
                                            if r["Unidade"] == "un" else r["Compra_Com_Reserva_Minima"], axis=1)
    return result.sort_values(["Necessidade_Compra_Urgente", "SKU"], ascending=[False, True]).reset_index(drop=True), detail


def count_setups(sequence: pd.DataFrame, lines: pd.DataFrame) -> tuple[int, float]:
    """Conta apenas trocas após a primeira OP; cobra setup da família que entra."""
    durations = lines.set_index("Familia_Compativel")["Tempo_Setup_Horas"].to_dict()
    families = sequence["Familia_Setup"].tolist()
    changed = [family for previous, family in zip(families, families[1:]) if family != previous]
    return len(changed), sum(float(durations[family]) for family in changed)


def sequence_scenarios(tables: dict[str, pd.DataFrame]) -> tuple[pd.DataFrame, pd.DataFrame]:
    chronological = tables["Ordens_Producao"].loc[lambda x: x["Status"].eq("Programada")].sort_values(["Data_Entrega", "Num_OP"], kind="stable")
    priorities = chronological.groupby("Familia_Setup")["Data_Entrega"].min().sort_values(kind="stable").index.tolist()
    optimized = chronological.copy()
    optimized["_priority"] = optimized["Familia_Setup"].map({f: i for i, f in enumerate(priorities)})
    return chronological, optimized.sort_values(["_priority", "Data_Entrega", "Num_OP"], kind="stable").drop(columns="_priority")


def add_business_days(start: date, days: int) -> date:
    current = start
    while current.weekday() >= 5:
        current += timedelta(days=1)
    for _ in range(max(0, days)):
        current += timedelta(days=1)
        while current.weekday() >= 5:
            current += timedelta(days=1)
    return current


def calculate_atp(tables: dict[str, pd.DataFrame], sku: str, quantity: float, today: date) -> dict:
    if not math.isfinite(quantity) or quantity <= 0:
        raise DataError("Informe uma quantidade finita maior que zero.")
    selected = tables["Estoque_Atual"].loc[lambda x: x["SKU"].eq(sku) & x["Categoria"].eq("Produto_Acabado")]
    if selected.empty:
        raise DataError("Selecione um produto acabado válido.")
    product = selected.iloc[0]
    available = float(product["Saldo_Disponivel"])
    net = max(0.0, quantity - available)
    result = {"sku": sku, "quantity": quantity, "available": available, "net": net,
              "immediate": net == 0, "release": today, "days": 0, "hours": 0.0,
              "queue_tons": 0.0, "queue": pd.DataFrame(), "material_risks": pd.DataFrame()}
    if net == 0:
        return result
    family = family_for(sku, tables)
    line = tables["Capacidade_Linhas"].loc[lambda x: x["Familia_Compativel"].eq(family)].iloc[0]
    queue = tables["Ordens_Producao"].loc[lambda x: x["Status"].isin(QUEUE_STATUS) & x["Familia_Setup"].eq(family)]
    queued_tons = float(queue["Qtd_Ton"].sum())
    hours = (queued_tons + net) / float(line["Capacidade_Ton_Hora"]) + float(line["Tempo_Setup_Horas"])
    days = math.ceil(hours / 16)
    # Estimativa conservadora: D+N dias úteis, sem fração restante do turno atual.
    release = add_business_days(today, days)
    current_mrp, _ = calculate_mrp(tables, include_in_process=True)
    recipe = tables["Fichas_Tecnicas_BOM"].loc[lambda x: x["SKU_Pai"].eq(sku)].copy()
    recipe["Demanda_Novo_Pedido"] = net * recipe["Qtd_Por_Ton"] * (1 + recipe["Perda_Tecnica_Pct"] / 100)
    risks = recipe.merge(current_mrp[["SKU", "Unidade", "Saldo_Liquido_Projetado"]], left_on="SKU_Insumo", right_on="SKU", how="left")
    risks["Saldo_Apos_Pedido"] = risks["Saldo_Liquido_Projetado"] - risks["Demanda_Novo_Pedido"]
    result.update({"family": family, "line": line["Linha_Equipamento"], "capacity": float(line["Capacidade_Ton_Hora"]),
                   "setup": float(line["Tempo_Setup_Horas"]), "queue": queue, "queue_tons": queued_tons,
                   "hours": hours, "days": days, "release": release,
                   "material_risks": risks[risks["Saldo_Apos_Pedido"].lt(-1e-9)]})
    return result


def replenishment_plan(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    stock = tables["Estoque_Atual"].loc[lambda x: x["Categoria"].eq("Produto_Acabado")].copy()
    pending = tables["Ordens_Producao"].loc[lambda x: x["Status"].isin(QUEUE_STATUS)].groupby("SKU_Produto")["Qtd_Ton"].sum()
    stock["Producao_Em_Fila"] = stock["SKU"].map(pending).fillna(0.0)
    stock["Saldo_Projetado"] = stock["Saldo_Disponivel"] + stock["Producao_Em_Fila"]
    stock["Producao_Sugerida_Ton"] = (stock["Estoque_Minimo"] - stock["Saldo_Projetado"]).clip(lower=0)
    return stock


def capacity_plan(tables: dict[str, pd.DataFrame], today: date) -> pd.DataFrame:
    records = []
    queue = tables["Ordens_Producao"].loc[lambda x: x["Status"].isin(QUEUE_STATUS)]
    for _, line in tables["Capacidade_Linhas"].iterrows():
        selected = queue[queue["Familia_Setup"].eq(line["Familia_Compativel"])].sort_values(["Data_Entrega", "Num_OP"])
        elapsed = float(line["Tempo_Setup_Horas"])
        for _, op in selected.iterrows():
            elapsed += float(op["Qtd_Ton"]) / float(line["Capacidade_Ton_Hora"])
            estimated = add_business_days(today, math.ceil(elapsed / 16))
            records.append({"Num_OP": op["Num_OP"], "Linha": line["Linha_Equipamento"], "Familia": op["Familia_Setup"],
                            "Qtd_Ton": op["Qtd_Ton"], "Horas_Acumuladas": elapsed, "Data_Entrega": op["Data_Entrega"],
                            "Liberacao_Estimada": estimated.isoformat(),
                            "Risco": "RISCO DE ATRASO" if estimated > date.fromisoformat(op["Data_Entrega"]) else "NO PRAZO"})
    return pd.DataFrame(records, columns=["Num_OP", "Linha", "Familia", "Qtd_Ton", "Horas_Acumuladas", "Data_Entrega", "Liberacao_Estimada", "Risco"])


def exceptions(tables: dict[str, pd.DataFrame], today: date) -> pd.DataFrame:
    records = []
    orders, stock = tables["Ordens_Producao"], tables["Estoque_Atual"]
    mrp, _ = calculate_mrp(tables, include_in_process=True)
    for _, row in mrp[mrp["Necessidade_Compra_Urgente"].gt(1e-9)].iterrows():
        records.append([1, "MATERIAL", row["SKU"], f"Déficit de {fmt(row['Necessidade_Compra_Urgente'])} {row['Unidade']}", "Emitir requisição no Motor MRP"])
    for _, row in orders[~orders["Status"].isin(["Expedida", "Liberada"])].iterrows():
        deadline = date.fromisoformat(row["Data_Entrega"])
        if deadline <= today:
            records.append([1 if deadline < today else 2, "PRAZO", row["Num_OP"], "Entrega vencida" if deadline < today else "Entrega hoje", "Priorizar PCP e validar o prazo"])
    for _, row in orders[orders["Status"].eq("Aguardando_Laudo")].iterrows():
        records.append([2, "QUALIDADE", row["Lote"], f"{fmt(row['Qtd_Ton'])} ton aguardando laudo", "Analisar laudo e liberar ou bloquear lote"])
    retained = orders[orders["Status"].isin(["Aguardando_Laudo", "Bloqueada"])].groupby("SKU_Produto")["Qtd_Ton"].sum()
    for _, row in stock[stock["Categoria"].eq("Produto_Acabado")].iterrows():
        linked = float(retained.get(row["SKU"], 0))
        if abs(float(row["Saldo_Quarentena"]) - linked) > 1e-6:
            records.append([2, "CONCILIAÇÃO", row["SKU"], f"Quarentena {fmt(row['Saldo_Quarentena'])} ton / OPs retidas {fmt(linked)} ton", "Conferir lotes e apontamentos na planilha"])
    schedule = capacity_plan(tables, today)
    for _, row in schedule[schedule["Risco"].eq("RISCO DE ATRASO")].iterrows():
        records.append([2, "CAPACIDADE", row["Num_OP"], f"Previsão por capacidade: {date.fromisoformat(row['Liberacao_Estimada']).strftime('%d/%m/%Y')}", "Revisar fila e compromisso comercial"])
    return pd.DataFrame(records, columns=["Prioridade", "Tipo", "Referencia", "Ocorrencia", "Acao_Sugerida"]).sort_values(["Prioridade", "Tipo", "Referencia"]).reset_index(drop=True)


def safe_export(frame: pd.DataFrame) -> pd.DataFrame:
    """Neutraliza fórmulas em campos textuais de arquivos baixados."""
    result = frame.copy()
    for col in result.select_dtypes(include=["object", "string"]).columns:
        result[col] = result[col].map(lambda v: "'" + v if isinstance(v, str) and v.lstrip().startswith(("=", "+", "-", "@")) else v)
    return result


def csv_bytes(frame: pd.DataFrame) -> bytes:
    return safe_export(frame).to_csv(index=False, sep=";", decimal=",", lineterminator="\n").encode("utf-8-sig")


def workbook_bytes(frames: Mapping[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in frames.items():
            safe_export(frame).to_excel(writer, sheet_name=name[:31], index=False)
            sheet = writer.sheets[name[:31]]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.fill = PatternFill("solid", fgColor="2D333B")
                cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
                cell.alignment = Alignment(vertical="center")
            for cells in sheet.iter_cols(min_row=2):
                for cell in cells:
                    cell.font = Font(name="Consolas" if isinstance(cell.value, (int, float)) else "Calibri", size=10)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.000"
            for column in sheet.columns:
                width = min(58, max(14, max((len(str(c.value or "")) for c in column), default=10) + 2))
                sheet.column_dimensions[column[0].column_letter].width = width
            sheet.row_dimensions[1].height = 26
    return buffer.getvalue()


CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Plus+Jakarta+Sans:wght@500;600;700&display=swap');
:root{--ink:#111318;--muted:#5A606D;--accent:#C84B31;--line:#E2E4E8;--mono:'JetBrains Mono',monospace;--sans:'Plus Jakarta Sans',sans-serif}
html,body,[class*="css"],.stApp,p,button,label,input,select,textarea{font-family:var(--sans)}
.stApp{background:#F6F6F4;color:var(--ink)}
[data-testid="stHeader"],#MainMenu,footer,[data-testid="stDecoration"],[data-testid="stToolbar"]{display:none!important}
[data-testid="stMainBlockContainer"]{padding:1.2rem 2rem 3rem;max-width:1700px}
[data-testid="stSidebar"]{background:#121417;border-right:1px solid #262A31;min-width:270px;max-width:270px;box-shadow:none!important;transform:none!important}
[data-testid="stSidebarCollapsedControl"]{display:none!important}
[data-testid="stSidebar"] *{color:#E5E7EB}
[data-testid="stSidebar"] [data-testid="stSidebarContent"]{padding-top:0}
[data-testid="stSidebar"] [data-testid="stSidebarHeader"]{height:1rem}
[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"]{display:none}
[data-testid="stSidebar"] [data-baseweb="radio"]{padding:12px 10px;margin:2px 0;border-left:3px solid transparent;align-items:flex-start}
[data-testid="stSidebar"] [data-baseweb="radio"]:has(input:checked){background:#1E2228;border-left-color:#C84B31;border-radius:0 3px 3px 0}
[data-testid="stSidebar"] [data-baseweb="radio"]>div:first-child{display:none}
[data-testid="stSidebar"] [data-baseweb="radio"] p{font-size:12px;font-weight:600;line-height:1.5}
[data-testid="stSidebar"] [data-testid="stRadio"]>label{display:none}
[data-testid="stSidebar"] .stButton button,[data-testid="stSidebar"] .stDownloadButton button{background:#1E2228;border:1px solid #343A44;color:#E5E7EB;font-size:11px}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p{color:#9CA3AF;font-size:10px}
.sidebar-intro{padding:20px 0 28px;border-bottom:1px solid #343A44;margin-bottom:16px}
.sidebar-intro small{font-size:9px;letter-spacing:.16em;color:#9CA3AF}
.sidebar-intro strong{display:block;font-size:21px;line-height:1.45;font-weight:600;margin-top:9px;letter-spacing:-.04em}
.sidebar-index{font-family:var(--mono);font-size:10px;color:#777E8B;margin:18px 0 8px;letter-spacing:.12em}
.context-bar{display:flex;justify-content:space-between;gap:16px;align-items:center;flex-wrap:wrap;border-bottom:1px solid var(--line);padding:4px 0 16px;margin-bottom:30px;font-size:9px;font-weight:700;letter-spacing:.075em}
.context-bar .mode{color:#5A606D;font-family:var(--mono);font-size:9px;letter-spacing:0}
.mode .dot{color:#C84B31;font-size:13px;margin-right:5px}
.page-head{display:flex;justify-content:space-between;gap:20px;align-items:flex-end;margin-bottom:26px}
.eyebrow{font-size:10px;font-weight:700;letter-spacing:.13em;text-transform:uppercase;color:#C84B31;margin-bottom:8px}
.page-head h1{font-size:29px;line-height:1.25;letter-spacing:-.045em;font-weight:600;margin:0 0 9px;padding:0}
.page-head p{font-size:12px;color:var(--muted);margin:0;line-height:1.7}
.page-stamp{font-family:var(--mono);font-size:10px;color:var(--muted);white-space:nowrap;text-align:right}
.kpi-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:0 0 28px}
.kpi{background:white;border:1px solid var(--line);border-radius:4px;border-top:3px solid var(--tone);padding:17px 18px 15px;min-width:0}
.kpi-label{font-size:10px;letter-spacing:.06em;font-weight:600;text-transform:uppercase;color:var(--muted);min-height:26px}
.kpi-value{font-family:var(--mono);font-variant-numeric:tabular-nums;font-size:26px;line-height:1.5;letter-spacing:-.04em;color:var(--ink)}
.kpi-sub{font-size:10px;line-height:1.65;color:var(--muted);margin-top:4px}
.section-label{display:flex;align-items:center;gap:10px;font-size:12px;font-weight:700;margin:6px 0 18px;letter-spacing:-.01em}
.section-label:after{content:'';height:1px;background:var(--line);flex:1}
.section-label small{font-family:var(--mono);font-weight:400;color:var(--muted);font-size:10px}
.mono,input[type="number"],input[type="date"],.stDateInput input,[data-testid="stDataFrame"],.js-plotly-plot text{font-family:var(--mono)!important;font-variant-numeric:tabular-nums}
.badge{display:inline-block;border-radius:3px;padding:2px 7px;font-size:9px;font-weight:600;letter-spacing:.025em;border:1px solid;line-height:1.7}
.badge.ok{color:#1B6B5A;background:#F0F7F4;border-color:#D2E8DF}
.badge.warning{color:#B45309;background:#FFFAF0;border-color:#F0DFC1}
.badge.danger{color:#B91C1C;background:#FFF4F3;border-color:#F0D1CE}
.badge.neutral{color:#5A606D;background:#F6F6F4;border-color:#E2E4E8}
.kanban-head{display:flex;justify-content:space-between;border-top:2px solid var(--tone);padding:12px 0;font-size:10px;text-transform:uppercase;font-weight:700;letter-spacing:.05em}
.kanban-count{font-family:var(--mono);color:var(--muted)}
.order-card{background:#FFF;border:1px solid var(--line);border-radius:4px;padding:14px;margin:0 0 10px;min-height:177px}
.order-card .op{font-size:11px;font-weight:600;margin-bottom:10px;font-family:var(--mono)}
.order-card .client{font-size:11px;font-weight:600;line-height:1.6;min-height:34px;margin-bottom:9px}
.order-card .sku{font-size:10px;color:var(--muted);font-family:var(--mono);line-height:1.6}
.order-card .foot{display:flex;justify-content:space-between;gap:4px;margin-top:14px;align-items:center;flex-wrap:wrap}
.order-card .qty{font-size:12px;font-family:var(--mono);font-weight:600}
.empty-lane{border:1px dashed var(--line);padding:26px 8px;font-size:11px;text-align:center;color:var(--muted);border-radius:4px}
.note{background:#FFF;border:1px solid var(--line);border-left:3px solid #C84B31;border-radius:4px;padding:15px 18px;color:var(--muted);font-size:11px;line-height:1.8;margin:12px 0 20px}
.result-box{background:white;border:1px solid var(--line);border-top:3px solid var(--tone);border-radius:4px;padding:24px;margin:8px 0 20px}
.result-box h2{font-size:19px;letter-spacing:-.03em;padding:6px 0 12px;margin:0}
.result-box p{font-size:12px;color:var(--muted);line-height:1.7;margin:0}
.stButton button,.stDownloadButton button,[data-testid="stFormSubmitButton"] button{border-radius:4px!important;box-shadow:none!important;min-height:38px;font-size:11px;font-weight:600;border:1px solid var(--line)}
button[kind="primary"],button[kind="primaryFormSubmit"]{background:#C84B31!important;border-color:#C84B31!important;color:white!important}
[data-testid="stFormSubmitButton"] button{background:#C84B31!important;border-color:#C84B31!important;color:white!important}
.stButton button:hover,.stDownloadButton button:hover{border-color:#C84B31;color:#C84B31}
[data-testid="stWidgetLabel"] p{font-size:11px;font-weight:600;color:var(--muted)}
[data-baseweb="input"],[data-baseweb="select"]>div,[data-baseweb="textarea"],[data-testid="stExpander"],[data-testid="stForm"]{border-radius:4px!important;box-shadow:none!important}
[data-baseweb="input"],[data-baseweb="select"]>div,[data-baseweb="textarea"]{border:1px solid var(--line);background:white}
[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:4px}
.table-scroll{overflow:auto;background:#FFF;border:1px solid var(--line);border-radius:4px;margin:0 0 16px}
.precision-table{width:100%;border-collapse:collapse;font-size:10px;font-family:var(--sans)}
.precision-table th{position:sticky;top:0;background:#F1F2F1;text-align:left;color:var(--muted);font-size:9px;font-weight:600;letter-spacing:.035em;text-transform:uppercase;padding:12px;white-space:nowrap;border-bottom:1px solid var(--line);z-index:1}
.precision-table td{padding:10px 12px;border-bottom:1px solid #ECEEF2;line-height:1.6;vertical-align:middle;max-width:340px;min-width:65px}
.precision-table tbody tr:last-child td{border-bottom:0}
.precision-table tbody tr:hover{background:#FAFAF8}
.precision-table .numeric{font-family:var(--mono);font-variant-numeric:tabular-nums;text-align:right;white-space:nowrap}
.precision-table .identifier{font-family:var(--mono);font-variant-numeric:tabular-nums;white-space:nowrap;font-size:10px}
.precision-table .deficit{color:#B91C1C;background:#FFF4F3;font-weight:600}
.precision-table .text-wide{min-width:200px}
[data-baseweb="tab-list"]{gap:24px;border-bottom:1px solid var(--line);margin-bottom:12px}
[data-baseweb="tab"]{font-size:11px;font-weight:600}
[data-baseweb="tab-highlight"]{background:#C84B31;height:2px}
[data-testid="stAlert"]{border-radius:4px;box-shadow:none;font-size:12px}
[data-testid="stExpander"] summary p{font-size:12px;font-weight:600}
[data-testid="stBottomBlockContainer"]{display:none}
.footer-note{font-size:9px;letter-spacing:.06em;color:#9298A3;border-top:1px solid var(--line);padding:15px 0;margin-top:28px;line-height:1.8}
.st-key-mobile_logout{display:none}
@media(max-width:1200px){[data-testid="stMainBlockContainer"]{padding:1rem 1.2rem 2rem}.kpi{padding:14px 12px}.kpi-value{font-size:23px}.page-head h1{font-size:25px}}
@media(max-width:850px){
 [data-testid="stAppViewContainer"]{flex-direction:column;overflow-y:auto}
 [data-testid="stSidebar"]{position:relative!important;min-width:100%!important;max-width:100%!important;width:100%!important;height:auto!important;max-height:none;transform:none!important;flex:0 0 auto;overflow:visible}
 [data-testid="stSidebar"] [data-testid="stSidebarContent"]{overflow:visible}
 [data-testid="stSidebar"] [data-testid="stSidebarUserContent"]{padding:10px 18px}
 [data-testid="stSidebar"] [data-testid="stVerticalBlock"]{gap:0!important}
 [data-testid="stSidebar"] [data-testid="stSidebarHeader"], [data-testid="stSidebarCollapsedControl"], [data-testid="stSidebar"] .sidebar-intro, [data-testid="stSidebar"] .sidebar-index, [data-testid="stSidebar"] [data-testid="stCaptionContainer"], [data-testid="stSidebar"] .stButton, [data-testid="stSidebar"] .stDownloadButton{display:none!important}
 [data-testid="stSidebar"] [role="radiogroup"]{flex-direction:row;flex-wrap:wrap;gap:3px}
 [data-testid="stSidebar"] [data-baseweb="radio"]{padding:7px 10px;margin:0;min-width:29%;border-left-width:2px}
 [data-testid="stSidebar"] [data-baseweb="radio"] p{font-size:10px}
 [data-testid="stMain"]{position:relative!important;width:100%!important;overflow:visible}
 .st-key-mobile_logout{display:block}
 .kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.page-head{display:block}.page-stamp{text-align:left;margin-top:12px}.context-bar{font-size:8px}.order-card{min-height:165px}
}
</style>
"""


def esc(value: Any) -> str:
    return html.escape(str(value), quote=True)


def fmt(value: float, digits: int = 1) -> str:
    return f"{float(value):,.{digits}f}".replace(",", "_").replace(".", ",").replace("_", ".")


def label(value: str) -> str:
    return value.replace("_", " ")


def badge(text: str, tone: str = "neutral") -> str:
    return f'<span class="badge {tone}">{esc(text.upper())}</span>'


def deadline_badge(delivery: str, today: date) -> str:
    due = date.fromisoformat(delivery)
    if due < today:
        return badge("Atrasada", "danger")
    if due == today:
        return badge("Entrega hoje", "warning")
    return badge("No prazo", "ok")


def heading(number: str, title: str, subtitle: str) -> None:
    st.markdown(f'<div class="page-head"><div><div class="eyebrow">OPERAÇÕES INDUSTRIAIS / {esc(number)}</div>'
                f'<h1>{esc(title)}</h1><p>{esc(subtitle)}</p></div>'
                f'<div class="page-stamp">{local_now().strftime("%d.%m.%Y")}<br>RIO CLARO · SP</div></div>', unsafe_allow_html=True)


def section(title: str, sub: str = "") -> None:
    st.markdown(f'<div class="section-label">{esc(title)} <small>{esc(sub)}</small></div>', unsafe_allow_html=True)


def kpis(items: list[tuple[str, str, str, str]]) -> None:
    cards = [f'<div class="kpi" style="--tone:{COLORS[tone]}"><div class="kpi-label">{esc(title)}</div>'
             f'<div class="kpi-value">{esc(value)}</div><div class="kpi-sub">{esc(sub)}</div></div>'
             for title, value, sub, tone in items]
    st.markdown('<div class="kpi-grid">' + "".join(cards) + "</div>", unsafe_allow_html=True)


def table(frame: pd.DataFrame, height: int | None = None, deficit_columns: tuple[str, ...] = ()) -> None:
    # HTML garante JetBrains Mono também no conteúdo, sem depender de canvas.
    numeric_cols = set(frame.select_dtypes(include="number").columns)
    header = ''.join(f'<th class="{"numeric" if col in numeric_cols else ""}">{esc(label(str(col)))}</th>' for col in frame.columns)
    rows = []
    for _, record in frame.iterrows():
        cells = []
        for col in frame.columns:
            value = record[col]
            classes = []
            if isinstance(value, Real) and not isinstance(value, bool):
                classes.append("numeric")
                rendered = fmt(float(value), 0 if col in ("Prioridade", "Prioridade_Sugerida") else 3) if pd.notna(value) else "—"
                if col in deficit_columns and float(value) < -1e-9:
                    classes.append("deficit")
            else:
                rendered = esc(value) if pd.notna(value) else "—"
                if col in ("Descricao", "Cliente", "Detalhe", "Ocorrencia", "Acao_Sugerida", "Linha", "Linha_Equipamento"):
                    classes.append("text-wide")
                if any(token in str(col) for token in ("SKU", "Num_OP", "Lote", "Data", "Hora", "Liberacao_Estimada")):
                    classes.append("identifier")
                if col in ("Status", "Risco", "Status_Requisicao"):
                    tone = "ok" if value in ("Liberada", "Expedida", "NO PRAZO") else "danger" if value in ("Bloqueada", "RISCO DE ATRASO") else "warning" if value == "Aguardando_Laudo" else "neutral"
                    rendered = badge(label(str(value)), tone)
            cells.append(f'<td class="{" ".join(classes)}">{rendered}</td>')
        rows.append('<tr>' + ''.join(cells) + '</tr>')
    if not rows:
        rows.append(f'<tr><td colspan="{max(1, len(frame.columns))}">Nenhum registro nesta consulta.</td></tr>')
    st.markdown(f'<div class="table-scroll" style="max-height:{height or 460}px"><table class="precision-table"><thead><tr>{header}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>', unsafe_allow_html=True)


def plot_style(fig: Any, height: int = 350) -> Any:
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font={"family": "JetBrains Mono, monospace", "size": 10, "color": COLORS["muted"]},
                      margin={"l": 0, "r": 12, "t": 16, "b": 0}, height=height,
                      legend={"orientation": "h", "y": 1.15, "x": 0}, bargap=.3)
    fig.update_xaxes(showgrid=False, zeroline=False, title_font_size=10)
    fig.update_yaxes(gridcolor="#ECEEF2", zeroline=False, title_font_size=10)
    return fig


def downloads(frames: dict[str, pd.DataFrame], prefix: str, primary_frame: str | None = None) -> None:
    key = primary_frame or next(iter(frames))
    left, right = st.columns(2)
    left.download_button("Baixar Excel (.xlsx)", workbook_bytes(frames), f"{prefix}.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"xlsx_{prefix}", width="stretch")
    right.download_button("Baixar CSV (.csv)", csv_bytes(frames[key]), f"{prefix}.csv", "text/csv", key=f"csv_{prefix}", width="stretch")


def perform(action: Any, *args: Any) -> None:
    try:
        action(*args)
    except DataError as exc:
        st.error(str(exc))
    except Exception:
        # A resposta de uma escrita pode ser perdida após o commit remoto.
        # Não repetir automaticamente, especialmente append_row.
        st.cache_data.clear()
        st.error("Não foi possível confirmar a gravação. Atualize os dados e confira a OP/estoque antes de repetir a ação. Verifique a conexão e as permissões da conta de serviço.")
    else:
        st.rerun()


def next_order_id(orders: pd.DataFrame, today: date) -> str:
    pattern = re.compile(rf"^OP-{today.year}-(\d+)$")
    numbers = [int(m.group(1)) for value in orders["Num_OP"] if (m := pattern.match(str(value)))]
    return f"OP-{today.year}-{max(numbers, default=1000) + 1:04d}"


def new_order_form(repo: Repository, tables: dict[str, pd.DataFrame], today: date) -> None:
    products = tables["Estoque_Atual"].loc[lambda x: x["Categoria"].eq("Produto_Acabado")]
    descriptions = products.set_index("SKU")["Descricao"].to_dict()
    with st.form("new_order", clear_on_submit=True):
        a, b, c = st.columns([1, 1.6, 1.8])
        num = a.text_input("Número da OP", value=next_order_id(tables["Ordens_Producao"], today), max_chars=40)
        client = b.text_input("Cliente", max_chars=160)
        product = c.selectbox("Produto acabado", products["SKU"].tolist(), format_func=lambda sku: f"{sku} · {descriptions[sku]}")
        d, e, f = st.columns([1, 1, 1])
        quantity = d.number_input("Quantidade (ton)", min_value=.001, value=10.0, step=1.0, format="%.3f")
        delivery = e.date_input("Data de entrega", value=add_business_days(today, 2), format="DD/MM/YYYY")
        batch = f.text_input("Lote único", value=f"LT-{today:%y%m}-{next_order_id(tables['Ordens_Producao'], today).split('-')[-1]}", max_chars=60)
        submitted = st.form_submit_button("Cadastrar OP", type="primary")
    if submitted:
        if not all(v.strip() for v in (num, client, batch)):
            st.error("Preencha número da OP, cliente e lote.")
        else:
            perform(repo.add_order, {"Num_OP": num.strip(), "Cliente": client.strip(), "SKU_Produto": product,
                    "Qtd_Ton": quantity, "Familia_Setup": family_for(product, tables),
                    "Data_Entrega": delivery.isoformat(), "Status": "Programada", "Lote": batch.strip()})


def status_actions(repo: Repository, tables: dict[str, pd.DataFrame]) -> None:
    orders = tables["Ordens_Producao"]
    eligible = orders[orders["Status"].map(lambda value: bool(transitions(value)))]
    if eligible.empty:
        st.info("Todas as OPs estão em status final ou bloqueadas.")
        return
    a, b, c = st.columns([2, 1.2, 1])
    op = a.selectbox("OP existente", eligible["Num_OP"].tolist(), format_func=lambda number: f"{number} · {label(eligible.loc[eligible['Num_OP'].eq(number), 'Status'].iloc[0])}", key="status_op")
    current = eligible.loc[eligible["Num_OP"].eq(op), "Status"].iloc[0]
    target = b.selectbox("Novo status", transitions(current), format_func=label, key=f"status_target_{op}")
    c.write("")
    if c.button("Aplicar alteração", key="apply_status", width="stretch"):
        perform(repo.change_status, op, target)
    st.caption("Fluxo: Programada → Em Processo → Aguardando Laudo → Liberada → Expedida. A liberação transfere estoque. Bloqueio mantém o saldo retido. Reabertura exige conciliação na origem.")


def automation_panel(repo: Repository, tables: dict[str, pd.DataFrame], today: date) -> None:
    alerts = exceptions(tables, today)
    section("Fila automática de exceções", f"{len(alerts)} sinais para ação")
    st.caption("Recalculada a cada atualização. Compras, PCP e Qualidade recebem uma fila única de prioridades na própria tela.")
    if alerts.empty:
        st.success("Nenhuma exceção detectada pelas regras configuradas.")
    else:
        table(alerts)
        downloads({"Excecoes": alerts}, "fila-excecoes")
    with st.expander("Reposição de produto acabado por estoque mínimo"):
        plan = replenishment_plan(tables)
        table(plan[["SKU", "Descricao", "Saldo_Disponivel", "Estoque_Minimo", "Producao_Em_Fila", "Saldo_Projetado", "Producao_Sugerida_Ton"]])
        candidates = plan[plan["Producao_Sugerida_Ton"].gt(1e-9)]
        st.caption("Saldo projetado = disponível + Programada + Em Processo. Quarentena não é oferta garantida. Não há carteira de vendas/reservas nesta base; valide o comprometimento do estoque antes de planejar reposição.")
        if candidates.empty:
            st.info("As OPs em fila já cobrem o estoque mínimo de todos os produtos acabados.")
        else:
            selected = st.selectbox("Produto para repor", candidates["SKU"].tolist(), key="replenish_product")
            quantity = float(candidates.loc[candidates["SKU"].eq(selected), "Producao_Sugerida_Ton"].iloc[0])
            delivery = st.date_input("Entrega interna da reposição", value=add_business_days(today, 5), key="replenish_date", format="DD/MM/YYYY")
            if st.button(f"Criar OP sugerida · {fmt(quantity)} ton", key="replenish_order"):
                current = repo.load(fresh=True)
                fresh_plan = replenishment_plan(current)
                fresh_qty = float(fresh_plan.loc[fresh_plan["SKU"].eq(selected), "Producao_Sugerida_Ton"].iloc[0])
                if fresh_qty <= 1e-9:
                    st.info("A necessidade já foi coberta por outra OP. Atualize os dados.")
                else:
                    op = next_order_id(current["Ordens_Producao"], today)
                    perform(repo.add_order, {"Num_OP": op, "Cliente": "Reposicao interna de estoque", "SKU_Produto": selected,
                            "Qtd_Ton": fresh_qty, "Familia_Setup": family_for(selected, current), "Data_Entrega": delivery.isoformat(),
                            "Status": "Programada", "Lote": f"LT-{today:%y%m}-{op.split('-')[-1]}"})
    with st.expander("Carga das linhas e risco de prazo"):
        schedule = capacity_plan(tables, today)
        if schedule.empty:
            st.info("Não há OPs em fila.")
        else:
            table(schedule)
        st.caption("16 h/dia, segunda a sexta, um setup inicial por linha. OPs em processo têm prioridade na operação real; a projeção usa a fila por data e considera sua tonelagem integral por falta de apontamento parcial. Sem feriados e sem indisponibilidades reais cadastradas.")
    with st.expander("Automações disponíveis e próximos conectores"):
        st.markdown("**Ativas:** MRP com perdas, compra com reserva mínima, alerta de prazo, conciliação de quarentena, carga de linhas, sugestão de reposição, estimativa ATP e geração de relatórios.\n\n**Próxima etapa:** pedidos de compra com fornecedores e lead times, alertas por e-mail/Teams, laudos de laboratório, códigos de barras e telemetria OPC UA/MQTT. Essas integrações precisam das fontes e regras da planta; consulte o roteiro entregue com o projeto.")


def overview(repo: Repository, tables: dict[str, pd.DataFrame], today: date) -> None:
    heading("01", "Torre de controle", "Visibilidade da carteira, dos materiais e das decisões de produção.")
    orders, stock = tables["Ordens_Producao"], tables["Estoque_Atual"]
    active = orders[~orders["Status"].eq("Expedida")]
    queue = orders[orders["Status"].isin(QUEUE_STATUS)]
    retained = orders[orders["Status"].eq("Aguardando_Laudo")]
    rupture = stock[stock["Saldo_Disponivel"].lt(stock["Estoque_Minimo"])]
    kpis([("OPs ativas", str(len(active)), "Carteira total, exceto expedidas", "graphite"),
          ("Tonelagem programada", f"{fmt(queue['Qtd_Ton'].sum())} t", "Programada + em processo", "accent"),
          ("Lotes em quarentena", str(len(retained)), f"{fmt(retained['Qtd_Ton'].sum())} ton aguardando laudo", "warning"),
          ("Itens em ruptura", str(len(rupture)), "Disponível abaixo do estoque mínimo", "danger")])
    day_tab, kanban_tab, auto_tab, stock_tab = st.tabs(["Meu dia", "Carteira de produção", "Automações & exceções", "Estoque & dados"])
    with day_tab:
        alerts = exceptions(tables, today)
        urgent = alerts[alerts["Prioridade"].eq(1)]
        section("Comece pelo que precisa de atenção", f"{len(urgent)} sinais prioritários")
        if urgent.empty:
            st.success("Não há exceções críticas pelas regras atuais. Confira as próximas entregas e avance suas ordens.")
        else:
            table(urgent[["Tipo", "Referencia", "Ocorrencia", "Acao_Sugerida"]])
        due = orders[~orders["Status"].isin(["Expedida", "Liberada"])].sort_values(["Data_Entrega", "Num_OP"])
        section("Próximas entregas", "Carteira ainda não liberada")
        table(due[["Num_OP", "Cliente", "SKU_Produto", "Qtd_Ton", "Data_Entrega", "Status"]].head(8))
        st.markdown('<div class="note">Rotina sugerida: confira os materiais críticos, trate os laudos pendentes e atualize o andamento das OPs. O painel reúne os sinais a cada atualização para você concentrar a revisão em um só lugar.</div>', unsafe_allow_html=True)
    with kanban_tab:
        section("Fluxo de ordens", "01 / 04")
        columns = st.columns(4, gap="small")
        for column, status, tone in zip(columns, STATUS[:4], ["graphite", "accent", "warning", "ok"]):
            selected = orders[orders["Status"].eq(status)].sort_values(["Data_Entrega", "Num_OP"])
            with column:
                st.markdown(f'<div class="kanban-head" style="--tone:{COLORS[tone]}">{esc(label(status))}<span class="kanban-count">{len(selected):02d}</span></div>', unsafe_allow_html=True)
                for _, row in selected.iterrows():
                    st.markdown(f'<div class="order-card"><div class="op">{esc(row["Num_OP"])}</div>'
                                f'<div class="client">{esc(row["Cliente"])}</div><div class="sku">{esc(row["SKU_Produto"])}<br>'
                                f'{esc(row["Lote"])} · {date.fromisoformat(row["Data_Entrega"]).strftime("%d/%m")}</div>'
                                f'<div class="foot"><span class="qty">{fmt(row["Qtd_Ton"])} t</span>{deadline_badge(row["Data_Entrega"], today)}</div></div>', unsafe_allow_html=True)
                if selected.empty:
                    st.markdown('<div class="empty-lane">Sem ordens nesta etapa</div>', unsafe_allow_html=True)
        blocked = orders[orders["Status"].isin(["Bloqueada", "Expedida"])]
        if not blocked.empty:
            with st.expander(f"Bloqueadas e expedidas · {len(blocked)} OPs"):
                table(blocked)
        with st.expander("Ações rápidas · nova OP e atualização de status"):
            section("Cadastrar nova ordem")
            new_order_form(repo, tables, today)
            section("Avançar ordem na produção")
            status_actions(repo, tables)
    with auto_tab:
        automation_panel(repo, tables, today)
    with stock_tab:
        section("Inventário operacional", "Saldos por unidade de medida")
        table(stock, deficit_columns=())
        st.caption("Os saldos são os registros da planilha. Não há baixa automática de insumos nem de expedição; o apontamento real deve ocorrer na origem. MRP não reserva nem consome estoque.")
        downloads(tables, "base-operacional", "Estoque_Atual")
        with st.expander("Fichas técnicas e capacidade de produção"):
            table(tables["Fichas_Tecnicas_BOM"])
            table(tables["Capacidade_Linhas"])


def mrp_page(repo: Repository, tables: dict[str, pd.DataFrame], today: date) -> None:
    heading("02", "Motor MRP & receitas", "Explosão de materiais com perdas técnicas, estoque líquido e compra sugerida.")
    include = st.toggle("Incluir OPs em processo na simulação", value=False, key="mrp_in_process")
    mrp, detail = calculate_mrp(tables, include)
    critical = mrp[mrp["Necessidade_Compra_Urgente"].gt(1e-9)]
    kpis([("Ordens consideradas", str(detail['Num_OP'].nunique()), "Programadas" + (" + em processo" if include else ""), "graphite"),
          ("Materiais com déficit", str(len(critical)), "Necessidade excede o disponível", "danger"),
          ("Itens para recompor mínimo", str(mrp['Compra_Sugerida'].gt(1e-9).sum()), "Compra sugerida considera reserva", "warning"),
          ("Componentes analisados", str(len(mrp)), "Toneladas, kg e unidades separados", "accent")])
    section("Necessidade líquida por insumo")
    columns = ["SKU", "Descricao", "Unidade", "Necessidade_Bruta", "Saldo_Disponivel", "Saldo_Quarentena", "Saldo_Liquido_Projetado", "Necessidade_Compra_Urgente", "Compra_Sugerida"]
    table(mrp[columns], deficit_columns=("Saldo_Liquido_Projetado",))
    st.markdown('<div class="note">Cálculo: quantidade da OP × consumo por tonelada × (1 + perda técnica / 100). A quarentena não compõe a disponibilidade. Compra urgente cobre o déficit; compra sugerida também recompõe o estoque mínimo e arredonda embalagens para unidades inteiras.</div>', unsafe_allow_html=True)
    if not critical.empty:
        section("Materiais críticos", "Necessidade × disponível")
        # Separar unidades evita colocar toneladas, kg e unidades na mesma escala.
        for unit in critical["Unidade"].unique():
            data = critical[critical["Unidade"].eq(unit)].melt(id_vars="SKU", value_vars=["Necessidade_Bruta", "Saldo_Disponivel"], var_name="Serie", value_name="Quantidade")
            fig = px.bar(data, x="Quantidade", y="SKU", color="Serie", barmode="group", orientation="h",
                         color_discrete_map={"Necessidade_Bruta": COLORS["accent"], "Saldo_Disponivel": COLORS["graphite"]},
                         labels={"Quantidade": f"Quantidade ({unit})", "Serie": "", "SKU": ""})
            st.plotly_chart(plot_style(fig, max(200, len(data) * 30 + 70)), use_container_width=True, key=f"mrp_chart_{unit}")
    else:
        st.success("O saldo disponível cobre as OPs consideradas nesta simulação.")
    requisition = mrp[mrp["Compra_Sugerida"].gt(1e-9)].copy()
    requisition["Data_Emissao"] = today.isoformat()
    requisition["Origem"] = "MRP: Programada + Em_Processo" if include else "MRP: Programada"
    requisition["Status_Requisicao"] = "RASCUNHO_PARA_COMPRAS"
    section("Requisição de materiais", "Gerada automaticamente; aprovação de Compras na origem")
    downloads({"Requisicao_Materiais": requisition, "MRP_Analitico": mrp,
               "Explosao_Por_OP": detail}, "requisicao-materiais", "Requisicao_Materiais")
    with st.expander("Rastrear o consumo por OP e ficha técnica"):
        table(detail[["Num_OP", "SKU_Produto", "Qtd_Ton", "SKU_Insumo", "Qtd_Por_Ton", "Perda_Tecnica_Pct", "Necessidade_Bruta"]])


def sequence_page(repo: Repository, tables: dict[str, pd.DataFrame], today: date) -> None:
    heading("03", "Sequenciador de produção", "Comparação por data de entrega e agrupamento por família de setup.")
    original, optimized = sequence_scenarios(tables)
    swaps_a, hours_a = count_setups(original, tables["Capacidade_Linhas"])
    swaps_b, hours_b = count_setups(optimized, tables["Capacidade_Linhas"])
    kpis([("Horas teóricas economizadas", f"{fmt(hours_a - hours_b)} h", "Comparativo de sequência global", "accent"),
          ("Trocas · cronológica", str(swaps_a), f"{fmt(hours_a)} h de setup", "graphite"),
          ("Trocas · por família", str(swaps_b), f"{fmt(hours_b)} h de setup", "ok"),
          ("OPs programadas", str(len(original)), f"{fmt(original['Qtd_Ton'].sum())} ton na comparação", "graphite")])
    st.markdown('<div class="note">O comparativo solicitado modela uma sequência global: cobra a limpeza da família que entra a cada troca, sem setup inicial. Como o cadastro usa uma linha dedicada por família, o ganho acima é teórico e não deve ser somado como economia real da planta. A carga por equipamento está disponível na Torre de Controle. A sequência sugerida não altera compromissos comerciais.</div>', unsafe_allow_html=True)
    if original.empty:
        st.info("Não há OPs programadas para sequenciar.")
        return
    left, right = st.columns(2)
    show_columns = ["Num_OP", "Familia_Setup", "Qtd_Ton", "Data_Entrega"]
    with left:
        section("A / Cronológica")
        table(original[show_columns].reset_index(drop=True))
    with right:
        section("B / Agrupada por família")
        table(optimized[show_columns].reset_index(drop=True))
    comparison = pd.DataFrame({"Cenario": ["Cronológica", "Por família"], "Horas_Setup": [hours_a, hours_b]})
    chart = px.bar(comparison, x="Cenario", y="Horas_Setup", color="Cenario", color_discrete_map={"Cronológica": COLORS["accent"], "Por família": COLORS["graphite"]}, labels={"Cenario": "", "Horas_Setup": "Horas de setup"})
    chart.update_layout(showlegend=False)
    st.plotly_chart(plot_style(chart, 250), use_container_width=True)
    suggestion = optimized[show_columns].copy()
    suggestion.insert(0, "Prioridade_Sugerida", range(1, len(suggestion) + 1))
    downloads({"Sequencia_Sugerida": suggestion, "Sequencia_Cronologica": original[show_columns], "Comparativo": comparison}, "sequenciamento")


def atp_page(repo: Repository, tables: dict[str, pd.DataFrame], today: date) -> None:
    heading("04", "Simulador de prazo ATP", "Uma consulta objetiva para disponibilidade e promessa de entrega ao Comercial.")
    products = tables["Estoque_Atual"].loc[lambda x: x["Categoria"].eq("Produto_Acabado")]
    descriptions = products.set_index("SKU")["Descricao"].to_dict()
    with st.form("atp_form"):
        left, right = st.columns([3, 1])
        sku = left.selectbox("Produto acabado", products["SKU"].tolist(), format_func=lambda code: f"{code} · {descriptions[code]}", key="atp_sku")
        quantity = right.number_input("Quantidade desejada (ton)", min_value=.001, value=50.0, step=1.0, format="%.3f", key="atp_qty")
        submit = st.form_submit_button("Calcular disponibilidade e prazo", type="primary")
    if submit:
        st.session_state["atp_request"] = (sku, quantity)
    if "atp_request" not in st.session_state:
        st.markdown('<div class="note">Selecione o material e a quantidade para consultar o saldo disponível e a fila da linha compatível. Quarentena não está disponível para expedição.</div>', unsafe_allow_html=True)
        return
    sku, quantity = st.session_state["atp_request"]
    if sku not in products["SKU"].values:
        st.info("O produto da última consulta foi removido. Faça uma nova simulação.")
        return
    result = calculate_atp(tables, sku, float(quantity), today)
    st.caption(f"Consulta exibida: {sku} · {fmt(quantity, 3)} ton. O resultado é recalculado quando os dados são atualizados; envie o formulário para alterar a consulta.")
    title = "DISPONIBILIDADE IMEDIATA PARA EXPEDIÇÃO" if result["immediate"] else "PREVISÃO DE LIBERAÇÃO POR CAPACIDADE"
    tone = COLORS["ok"] if result["immediate"] else COLORS["accent"]
    text = "O saldo disponível cobre integralmente a quantidade consultada." if result["immediate"] else f"Liberação estimada em {result['release'].strftime('%d/%m/%Y')} · {result['days']} dias úteis de planejamento."
    st.markdown(f'<div class="result-box" style="--tone:{tone}"><div class="eyebrow">{esc(sku)}</div><h2>{title}</h2><p>{esc(text)}</p></div>', unsafe_allow_html=True)
    kpis([("Pedido consultado", f"{fmt(quantity)} t", descriptions[sku], "graphite"),
          ("Disponível agora", f"{fmt(result['available'])} t", "Quarentena excluída", "ok"),
          ("Líquido a produzir", f"{fmt(result['net'])} t", "Pedido menos saldo disponível", "accent"),
          ("Fila anterior", f"{fmt(result['queue_tons'])} t", "Em processo + programada da família", "warning")])
    if not result["immediate"]:
        section("Memória de cálculo")
        st.markdown(f'<div class="note">Linha: <strong>{esc(result["line"])}</strong><br>'
                    f'<span class="mono">({fmt(result["queue_tons"], 3)} + {fmt(result["net"], 3)}) ton / {fmt(result["capacity"])} ton/h + {fmt(result["setup"])} h = {fmt(result["hours"], 2)} h</span><br>'
                    f'Turno de 16 h/dia, segunda a sexta. Prazo conservador D+N; não considera feriados, manutenção, frete ou tempo de laboratório.</div>', unsafe_allow_html=True)
        if not result["material_risks"].empty:
            st.warning("Prazo condicionado ao abastecimento: há déficit projetado de insumos considerando a fila e o novo pedido. Esta previsão de capacidade não confirma disponibilidade dos materiais.")
            table(result["material_risks"][["SKU_Insumo", "Unidade", "Demanda_Novo_Pedido", "Saldo_Apos_Pedido"]], deficit_columns=("Saldo_Apos_Pedido",))
        if not result["queue"].empty:
            table(result["queue"][["Num_OP", "Cliente", "Qtd_Ton", "Data_Entrega", "Status"]])
    st.caption("Consulta sem reserva: outras OPs ou vendas podem comprometer o mesmo saldo. O modelo atual não contém carteira de pedidos/reservas nem recebimentos de fornecedores.")
    summary = pd.DataFrame([{k: v for k, v in result.items() if k not in ("queue", "material_risks")}])
    downloads({"Consulta_ATP": summary, "Fila": result["queue"], "Riscos_Materiais": result["material_risks"]}, "consulta-atp")


def quality_page(repo: Repository, tables: dict[str, pd.DataFrame], today: date) -> None:
    heading("05", "Quarentena & qualidade", "Liberação consciente de lotes, saldos retidos e conciliação do estoque.")
    orders, stock = tables["Ordens_Producao"], tables["Estoque_Atual"]
    pending = orders[orders["Status"].eq("Aguardando_Laudo")].sort_values(["Data_Entrega", "Num_OP"])
    retained = stock[stock["Saldo_Quarentena"].gt(0)]
    tons_retained = retained.loc[retained["Unidade"].eq("ton"), "Saldo_Quarentena"].sum()
    kpis([("Lotes aguardando laudo", str(len(pending)), f"{fmt(pending['Qtd_Ton'].sum())} ton em OPs retidas", "warning"),
          ("Saldo retido em toneladas", f"{fmt(tons_retained)} t", "MP e PA; kg e unidades separados", "graphite"),
          ("SKUs com retenção", str(len(retained)), "Todas as unidades de medida", "accent"),
          ("Ordens bloqueadas", str(orders['Status'].eq('Bloqueada').sum()), "Sem transferência de saldo", "danger")])
    section("Lotes para decisão", "Ordenados por urgência de entrega")
    if pending.empty:
        st.success("Não há lotes aguardando laudo.")
    stock_indexed = stock.set_index("SKU")
    for _, row in pending.iterrows():
        with st.container(border=True):
            a, b = st.columns([3, 1])
            a.markdown(f'**{esc(row["Num_OP"])} · {esc(row["Cliente"])}**', unsafe_allow_html=True)
            a.markdown(f'<span class="mono">{esc(row["Lote"])} · {esc(row["SKU_Produto"])} · {fmt(row["Qtd_Ton"])} ton</span>', unsafe_allow_html=True)
            b.markdown(deadline_badge(row["Data_Entrega"], today), unsafe_allow_html=True)
            available_quarantine = float(stock_indexed.at[row["SKU_Produto"], "Saldo_Quarentena"])
            if available_quarantine + 1e-9 < float(row["Qtd_Ton"]):
                st.warning(f"Divergência: quarentena do SKU contém {fmt(available_quarantine)} ton para um lote de {fmt(row['Qtd_Ton'])} ton. Pela regra solicitada, a aprovação credita o lote integral e limita a quarentena a zero. Confira o apontamento antes de aprovar.")
            left, right = st.columns(2)
            if left.button("Aprovar laudo e liberar lote", type="primary", key=f"approve_{row['Num_OP']}", width="stretch"):
                perform(repo.approve, row["Num_OP"])
            if right.button("Bloquear lote", key=f"block_{row['Num_OP']}", width="stretch"):
                perform(repo.change_status, row["Num_OP"], "Bloqueada")
    section("Resumo de saldos em quarentena")
    table(retained[["SKU", "Descricao", "Unidade", "Saldo_Quarentena", "Saldo_Disponivel"]])
    st.caption("Bloquear altera o status da OP e mantém o saldo retido. A aprovação não anexa nem valida o conteúdo de um laudo laboratorial: a decisão cabe ao operador responsável. O estoque desta base é agregado por SKU, sem composição por lote.")
    with st.expander("Registro de operações desta sessão", expanded=bool(st.session_state.get("event_log"))):
        events = pd.DataFrame(st.session_state.get("event_log", []))
        if events.empty:
            st.caption("As próximas alterações aparecerão aqui.")
        else:
            table(events)
            downloads({"Operacoes_Sessao": events}, "operacoes-sessao")
        st.caption("Registro em memória desta sessão; não constitui trilha de auditoria permanente. As alterações de OP e estoque persistem no Sheets quando conectado.")


@st.cache_resource(show_spinner=False)
def temporary_cookie_key() -> str:
    """Chave aleatória para testes locais; em produção a chave vem dos secrets."""
    return secure_random.token_urlsafe(48)


@st.cache_resource(show_spinner=False)
def hash_configured_password(password: str) -> str:
    import bcrypt
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt(rounds=12)).decode()


def authenticate() -> Any:
    if not secret("authentication_required", True):
        if production_mode():
            st.error("O modo de produção exige login habilitado.")
            st.stop()
        return None
    import streamlit_authenticator as stauth

    cookie_secret = secret("auth_cookie_key", "")
    if production_mode() and (not isinstance(cookie_secret, str) or len(cookie_secret) < 32):
        st.error("Configure auth_cookie_key com uma chave privada de pelo menos 32 caracteres para manter o acesso entre visitas.")
        st.stop()
    cookie_secret = cookie_secret or temporary_cookie_key()
    username = str(secret("app_username", "usuario")).strip()
    password_hash = secret("app_password_hash", DEFAULT_PASSWORD_HASH)
    # Compatibilidade com instalações anteriores que configuraram app_password.
    password = secret("app_password")
    if password:
        password_hash = hash_configured_password(str(password))
    if not isinstance(password_hash, str) or not re.fullmatch(r"\$2[aby]\$\d{2}\$.{53}", password_hash):
        st.error("Configure o login e o hash da senha nos Secrets privados da hospedagem.")
        st.stop()
    credential_version = hashlib.sha256(str(password).encode()).hexdigest() if password else str(password_hash)
    identity = hashlib.sha256((username.lower() + credential_version).encode()).hexdigest()
    if st.session_state.get("_credential_identity") != identity:
        st.session_state["_credential_identity"] = identity
        st.session_state["_auth_credentials"] = {"usernames": {username.lower(): {"name": username, "password": password_hash}}}
        st.session_state["authentication_status"] = None
    # Trocar a senha invalida automaticamente os cookies emitidos antes da troca.
    signing_key = hashlib.sha256((str(cookie_secret) + identity).encode()).hexdigest()
    remembered = st.session_state.get("remember_browser", True)
    authenticator = stauth.Authenticate(st.session_state["_auth_credentials"],
        cookie_name="pcp_access", cookie_key=signing_key,
        cookie_expiry_days=365 if remembered else 0, auto_hash=False, login_sleep_time=0)
    authenticator.login(location="unrendered")
    if st.session_state.get("authentication_status") is True:
        return authenticator
    left, center, right = st.columns([1, 1.3, 1])
    with center:
        st.markdown('<div style="height:8vh"></div><div class="eyebrow">SEU ESPAÇO DE TRABALHO</div>', unsafe_allow_html=True)
        st.title("Organize a produção.\nGanhe tempo no seu dia.")
        st.caption("Um acesso pessoal para acompanhar prioridades, materiais e prazos.")
        st.checkbox("Manter conectado neste navegador", value=True, key="remember_browser")
        try:
            authenticator.login(location="main", max_login_attempts=5,
                fields={"Form name": "Acesse seu painel", "Username": "Login", "Password": "Senha", "Login": "Entrar"},
                key="personal_login")
        except stauth.utilities.exceptions.LoginError:
            st.error("Limite de tentativas atingido nesta sessão. Feche a página e tente novamente mais tarde.")
        if st.session_state.get("authentication_status") is False:
            st.error("Login ou senha incorretos.")
        st.caption("O acesso fica lembrado por até 1 ano neste navegador. Sair, limpar cookies ou trocar a senha encerra o acesso lembrado.")
    if st.session_state.get("authentication_status") is not True:
        st.stop()
    return authenticator


@st.fragment(run_every="60s")
def workspace(repo: Repository, page: str) -> None:
    try:
        tables = repo.load()
    except DataError as exc:
        st.error(str(exc))
        st.download_button("Baixar modelo com as quatro abas", workbook_bytes(seed_tables()), "modelo-quatro-abas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.stop()
    except Exception:
        st.error("A conexão com o Google Sheets não foi concluída. Confira a URL, as quatro abas, a chave da conta de serviço, a API habilitada e o compartilhamento como Editor. Nenhuma gravação foi enviada.")
        st.download_button("Baixar modelo com as quatro abas", workbook_bytes(seed_tables()), "modelo-quatro-abas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.stop()
    mode = "GOOGLE SHEETS SINCRONIZADO" if repo.mode == "cloud" else "MODO SANDBOX LOCAL"
    st.markdown(f'<div class="context-bar"><span>PCP & TELEMETRIA DE PRODUÇÃO | UNIDADE RIO CLARO (SP)</span>'
                f'<span class="mode"><span class="dot">●</span>{mode}</span></div>', unsafe_allow_html=True)
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))
    pages = dict(zip(NAVIGATION, (overview, mrp_page, sequence_page, atp_page, quality_page)))
    pages[page](repo, tables, local_now().date())
    st.markdown(f'<div class="footer-note">DADOS REAVALIADOS A CADA 60 S COM A SESSÃO ATIVA · LEITURA {local_now().strftime("%H:%M:%S")} BRT<br>'
                'TELEMETRIA OPERACIONAL DERIVADA DE OPs E SALDOS · SENSORES E APONTAMENTOS DE MÁQUINA NÃO CONECTADOS</div>', unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")
    st.markdown(CSS, unsafe_allow_html=True)
    authenticator = authenticate()
    try:
        repo = Repository.configured()
    except DataError as exc:
        st.info(str(exc))
        st.stop()
    with st.sidebar:
        st.markdown('<div class="sidebar-intro"><small>PLANEJAMENTO INDUSTRIAL</small><strong>Controle<br>de produção</strong></div><div class="sidebar-index">01 — 05 / ÁREAS DE TRABALHO</div>', unsafe_allow_html=True)
        page = st.radio("Navegação", NAVIGATION, label_visibility="collapsed", key="navigation")
        st.markdown('<div class="sidebar-index">AMBIENTE OPERACIONAL</div>', unsafe_allow_html=True)
        st.caption("GOOGLE SHEETS / CONECTADO" if repo.mode == "cloud" else "SANDBOX / DADOS EM MEMÓRIA")
        if authenticator is not None:
            st.caption(f"Acesso pessoal · {st.session_state.get('name', 'Pessoal')}")
        if repo.mode == "sandbox":
            st.caption("Alterações pertencem a esta sessão e são perdidas ao encerrá-la. Os dados iniciais seguem a base de novembro/2026.")
        if st.button("Atualizar dados", width="stretch"):
            st.cache_data.clear()
            st.rerun()
        st.download_button("Modelo de integração · 4 abas", workbook_bytes(seed_tables()), "modelo-quatro-abas.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
        if st.session_state.get("last_write"):
            st.caption("Última gravação\n" + st.session_state["last_write"])
        if authenticator is not None:
            authenticator.logout("Sair deste navegador", location="main", key="personal_logout", use_container_width=True)
            if st.session_state.get("authentication_status") is not True:
                st.stop()
        st.markdown('<div class="sidebar-index">16 H / DIA · SEG — SEX</div>', unsafe_allow_html=True)
        st.caption("Rio Claro · SP\nInterface operacional / v1.0")
    if authenticator is not None:
        with st.container(key="mobile_logout"):
            authenticator.logout("Sair deste navegador", location="main", key="mobile_personal_logout")
            if st.session_state.get("authentication_status") is not True:
                st.stop()
    workspace(repo, page)


if __name__ == "__main__":
    main()
