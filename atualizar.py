import xlwings as xw
import pandas as pd
import gspread
import datetime as dt

from oauth2client.service_account import ServiceAccountCredentials
from flask import Flask, redirect, jsonify
from gspread.utils import rowcol_to_a1

app = Flask(__name__)

# ======================================================
# Atualiza o Excel local
# ======================================================
def atualizar_excel():
    caminho_arquivo = r"Z:\\Controladoria\\Reservas\\Reservas.xlsx"

    wb = xw.Book(caminho_arquivo)
    wb.api.RefreshAll()
    wb.app.api.CalculateFullRebuild()
    wb.save()
    wb.close()

    return caminho_arquivo


# ======================================================
# Importa dados para o Google Sheets
# ======================================================
def importar_para_sheets(caminho_arquivo):

    # --------------------------------------------------
    # Lê o Excel
    # --------------------------------------------------
    df = pd.read_excel(caminho_arquivo, dtype=str)

    # --------------------------------------------------
    # Limpeza geral
    # --------------------------------------------------
    df = df.fillna("")

    # --------------------------------------------------
    # Cria NOMECOMPLETO
    # --------------------------------------------------
    if {"NOME", "SOBRENOME"}.issubset(df.columns):
        df["NOMECOMPLETO"] = (
            df["NOME"].str.strip() + " " + df["SOBRENOME"].str.strip()
        )

    # --------------------------------------------------
    # 🔥 FORÇA NUMRESERVA COMO TEXTO ABSOLUTO
    # --------------------------------------------------
    if "NUMRESERVA" in df.columns:
        df["NUMRESERVA"] = (
            "'" + df["NUMRESERVA"].astype(str).str.strip()
        )

    # --------------------------------------------------
    # BLINDAGEM TOTAL DE DATA / HORA
    # --------------------------------------------------
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: ""
            if x == ""
            else x.strftime("%d/%m/%Y %H:%M:%S")
            if isinstance(x, (pd.Timestamp, dt.datetime))
            else x.strftime("%H:%M:%S")
            if isinstance(x, dt.time)
            else x.strftime("%d/%m/%Y")
            if isinstance(x, dt.date)
            else x
        )

    # --------------------------------------------------
    # Autenticação Google Sheets
    # --------------------------------------------------
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        r"Z:\\Controladoria\\Reservas\\solar-nimbus-477516-b0-80bc7621dd62.json",
        scope
    )

    client = gspread.authorize(creds)

    # --------------------------------------------------
    # Abre planilha e aba
    # --------------------------------------------------
    sheet = client.open("PYTHON_TESTE").worksheet("RESERVAS")

    # --------------------------------------------------
    # Envia dados
    # --------------------------------------------------
    values = [df.columns.tolist()] + df.values.tolist()

    sheet.clear()
    sheet.update(values, value_input_option="USER_ENTERED")

    # --------------------------------------------------
    # Formata HORACHEGADA
    # --------------------------------------------------
    try:
        col_index = df.columns.get_loc("HORACHEGADA") + 1
        start_cell = rowcol_to_a1(2, col_index)
        end_cell = rowcol_to_a1(sheet.row_count, col_index)

        sheet.format(f"{start_cell}:{end_cell}", {
            "numberFormat": {
                "type": "DATE_TIME",
                "pattern": "dd/MM/yyyy HH:mm:ss"
            }
        })
    except Exception as e:
        print(f"Erro ao formatar HORACHEGADA: {e}")


# ======================================================
# Rotas Flask
# ======================================================
@app.route("/atualizar", methods=["GET"])
def atualizar():
    caminho = atualizar_excel()
    importar_para_sheets(caminho)

    return redirect(
        "https://docs.google.com/spreadsheets/d/1G_OD22AxbExxl08n-xQTRCY1Oyh3frKDux_a_iIivSo/edit"
    )


@app.route("/atualizar-somente", methods=["GET"])
def atualizar_somente():
    caminho = atualizar_excel()
    importar_para_sheets(caminho)

    return jsonify({
        "status": "ok",
        "mensagem": "Planilha atualizada com sucesso"
    })


# ======================================================
# Main
# ======================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
