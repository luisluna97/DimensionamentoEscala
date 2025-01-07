import streamlit as st
import pandas as pd
import datetime
import re
import io
import zipfile
from io import BytesIO

#################################
# 1) Dicionário "De → Para"
#################################
mapeamento_cargos = {
    "ASG": "ASG LIMPEZA",
    "Aux. Lider": "AUX.LIDER DE RAMPA",
    "Operador": "OPERADOR DE EQUIPAMENTOS",
    "ASA": "AUXILIAR DE RAMPA",
    "SUPERVISOr": "SUPERVISOR DE OPERAÇÕES",
    "ENC": "ENC. DE LIMPEZA",
    "LIDER OP.": "LIDER DE OPERAÇÕES",
    "SUPERVISOR PAX": "SUPERVISOR SERV A PAX",
    "PAX LIDER": "ATENDENTE A PAX LIDER",
    "(BALANCEIRO)": "AUX DE RAMPA (BALANCEIRO)",
    "SUPERVISOR TEC. OPERACIONAL": "SUPERVISOR TEC. OPERACIONAL",
    "LIDER PAX": "ATENDENTE A PASSAGEIRO LIDER",
    "AGENTE DE AVIAÇÃO EXECUTIVA": "AGENTE DE AVIAÇÃO EXECUTIVA",
    "INSPETOR SAFETY E QUALIDADE": "INSPETOR SAFETY E QUALIDADE",
    "PAX": "AGENTE SERV A PAX",
    "AGENTE SERV A PASSAGEIRO": "AGENTE SERV A PASSAGEIRO",
    "AGENTE DE PROTEÇÃO": "AGENTE DE PROTEÇÃO",
    "APAC": "APAC",
    "COORDENADOR DE OPERAÇÕES": "COORDENADOR DE OPERAÇÕES",
    "SUPERV DE CARGAS": "SUPERV DE CARGAS",
    "SUPERVISOR DE LIMPEZA": "SUPERVISOR DE LIMPEZA",
    "GERENTE ATENDENTE A PAX": "GERENTE ATENDENTE A PAX"
}

#################################
# 2) Funções auxiliares
#################################

def hhmm_to_minutes(hhmm):
    """ Converte 'HH:MM' em inteiro de minutos após 00:00. """
    hh, mm = hhmm.split(":")
    return int(hh)*60 + int(mm)

def minutes_to_hhmm(total_minutes):
    """ Converte minutos em 'HH:MM' (24h, modulo 1440). """
    total_minutes = total_minutes % (24*60)
    hh = total_minutes // 60
    mm = total_minutes % 60
    return f"{hh:02d}:{mm:02d}"

def eh_celula_numero(valor):
    """ Verifica se a célula começa com dígito (ex.: '3'). """
    if pd.isna(valor):
        return False
    return bool(re.match(r'^\d+', str(valor).strip()))

def calcula_intervalos(h_inicio_str, carga_h_str):
    """
    Se for 7H => 4h + 1h pausa + 3h.
    Caso contrário => turno único (2H,4H,6H etc.).
    Retorna (Entrada1, Saida1, Entrada2, Saida2).
    """
    match = re.match(r"(\d+)", carga_h_str)
    if not match:
        return ("","","","")

    qtd_horas = int(match.group(1))
    inicio_min = hhmm_to_minutes(h_inicio_str)
    
    if qtd_horas == 7:
        e1 = inicio_min
        s1 = e1 + 4*60
        e2 = s1 + 60
        s2 = e2 + 3*60
        return (
            minutes_to_hhmm(e1),
            minutes_to_hhmm(s1),
            minutes_to_hhmm(e2),
            minutes_to_hhmm(s2)
        )
    else:
        e1 = inicio_min
        s1 = e1 + qtd_horas*60
        return (
            minutes_to_hhmm(e1),
            minutes_to_hhmm(s1),
            "",
            ""
        )

def ignorar_poluicao(funcao, carga):
    """
    Se contiver 'HORAS AQUI' ou 'INTERVALO', consideramos poluição.
    """
    f = str(funcao).upper()
    c = str(carga).upper()
    if "HORAS AQUI" in f or "HORAS AQUI" in c:
        return True
    if "INTERVALO" in f or "INTERVALO" in c:
        return True
    return False

def mapear_cargo(funcao):
    """
    Faz busca parcial em 'funcao' para substituir segundo o dicionário mapeamento_cargos.
    Exemplo:
      - Se funcao="LIDER OP.,6H" => 'LIDER OP.' é detectado e substitui => "LIDER DE OPERAÇÕES"
      - Se "ASA" => "AUXILIAR DE RAMPA"
    Retorna a string mapeada se encontrado; senão, retorna original.
    """
    funcao_up = funcao.upper()
    for de, para in mapeamento_cargos.items():
        # Buscamos se 'de' existe (case-insensitive) dentro de funcao
        if de.upper() in funcao_up:
            return para
    return funcao  # se não achar nada, mantém original

#################################
# 3) Processamento principal
#################################

def processar_planilha(excel_file, periodo) -> dict:
    """
    - 'periodo' é uma string como "06/2025".
    - Retorna { nomeAba: DataFrame } com colunas:
        Base, Periodo, Quantidade, Cargo, CargaHoraria, Entrada1, Saida1, Entrada2, Saida2
    """
    xls = pd.ExcelFile(excel_file)
    abas_ignorar = ["BASE", "TABELAS", "PADRÕES", "ARQUIVO BASE"]
    abas_validas = [aba for aba in xls.sheet_names 
                    if aba.upper() not in [a.upper() for a in abas_ignorar]]

    dict_abas_result = {}

    for nome_aba in abas_validas:
        df = pd.read_excel(excel_file, sheet_name=nome_aba, header=None)
        nrows, ncols = df.shape

        col_inicio = 6
        col_fim    = 800
        if col_fim >= ncols:
            col_fim = ncols - 1

        # Mapeia col->horario (linha 2 => df index=1)
        if nrows < 2:
            continue
        
        col_to_horario = {}
        for c in range(col_inicio, col_fim+1):
            val = ""
            if c < ncols:
                tmp = df.iloc[1,c]
                if not pd.isna(tmp):
                    val = str(tmp).strip()
            col_to_horario[c] = val

        linha_inicio = 86
        if linha_inicio >= nrows:
            continue

        registros = []
        
        for lin in range(linha_inicio, nrows):
            col = col_inicio
            while col <= col_fim:
                if col >= ncols:
                    break

                cell_val = df.iloc[lin, col]
                if eh_celula_numero(cell_val):
                    # Extrair quantidade
                    qtd_match = re.match(r'^(\d+)', str(cell_val).strip())
                    quantidade_str = qtd_match.group(1) if qtd_match else ""

                    # Ler possíveis próximas colunas
                    val_next1 = ""
                    val_next2 = ""
                    if col+1 <= col_fim:
                        v1 = df.iloc[lin, col+1]
                        val_next1 = "" if pd.isna(v1) else str(v1).strip()
                    if col+2 <= col_fim:
                        v2 = df.iloc[lin, col+2]
                        val_next2 = "" if pd.isna(v2) else str(v2).strip()

                    funcao = ""
                    carga  = ""
                    col_consumidas = 1

                    if "," in val_next1:
                        # "Aux. Lider,6H"
                        partes = [p.strip() for p in val_next1.split(",")]
                        if len(partes) == 2:
                            funcao, carga = partes
                            col_consumidas = 2
                        else:
                            funcao = val_next1
                            carga  = val_next2
                            col_consumidas = 3
                    else:
                        # Se val_next1 for "4H" etc.
                        if re.match(r'^\d+H$', val_next1):
                            funcao = ""
                            carga  = val_next1
                            col_consumidas = 2
                        elif re.match(r'^\d+H$', val_next2):
                            funcao = val_next1
                            carga  = val_next2
                            col_consumidas = 3
                        else:
                            funcao = val_next1
                            carga  = ""
                            col_consumidas = 2

                    # Ignorar caso sem função e sem 'xH'
                    if not funcao and not re.match(r'^\d+H$', carga):
                        col += col_consumidas
                        continue

                    # Ignorar poluição
                    if ignorar_poluicao(funcao, carga):
                        col += col_consumidas
                        continue

                    # Horário de início
                    horario_inicio_str = col_to_horario.get(col, "")
                    if not horario_inicio_str:
                        horario_inicio_str = "00:00"

                    # Calcula intervalos
                    if re.match(r'^\d+H$', carga):
                        e1, s1, e2, s2 = calcula_intervalos(horario_inicio_str, carga)
                    else:
                        e1, s1, e2, s2 = ("","","","")

                    # Mapear função
                    funcao_mapeada = mapear_cargo(funcao)

                    # Monta registro
                    reg = {
                        "Base": nome_aba,               # nome da aba
                        "Periodo": periodo,             # p.ex. "06/2025"
                        "Quantidade": quantidade_str,
                        "Cargo": funcao_mapeada,        # já mapeado
                        "CargaHoraria": carga,
                        "Entrada1": e1,
                        "Saida1": s1,
                        "Entrada2": e2,
                        "Saida2": s2
                    }
                    registros.append(reg)

                    # Lógica de pulo
                    fim1 = hhmm_to_minutes(s1) if s1 else 0
                    fim2 = hhmm_to_minutes(s2) if s2 else 0
                    fim  = max(fim1, fim2)
                    ini  = hhmm_to_minutes(e1) if e1 else 0

                    if fim > ini:
                        proxima_coluna = None
                        for ctemp in range(col, col_fim+1):
                            htemp = col_to_horario.get(ctemp,"")
                            if htemp:
                                hm = hhmm_to_minutes(htemp)
                                if hm >= fim:
                                    proxima_coluna = ctemp
                                    break
                        if proxima_coluna is not None:
                            col = proxima_coluna
                        else:
                            col = col_fim + 1
                    else:
                        col += col_consumidas
                else:
                    col += 1

        df_out = pd.DataFrame(registros, columns=[
            "Base","Periodo","Quantidade","Cargo","CargaHoraria",
            "Entrada1","Saida1","Entrada2","Saida2"
        ])
        if not df_out.empty:
            dict_abas_result[nome_aba] = df_out

    return dict_abas_result

#################################
# 4) Funções de download
#################################

def gerar_download_link_para_df(df, filename):
    csv_bytes = df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')
    return st.download_button(
        label=f"Baixar {filename}",
        data=csv_bytes,
        file_name=filename,
        mime="text/csv"
    )

def gerar_download_zip(dict_df):
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for aba, df in dict_df.items():
            csv_str = df.to_csv(index=False, sep=';', encoding='utf-8-sig')
            zf.writestr(f"{aba}.csv", csv_str)
    buffer.seek(0)
    return buffer.getvalue(), "todas_abas.zip"


#################################
# 5) APP STREAMLIT (app.py)
#################################

def main():
    st.title("Conversor de Planilha - Pernoite & Mapeamento de Cargo")

    # Escolha do mês e ano
    col1, col2 = st.columns(2)
    with col1:
        meses = [f"{i:02d}" for i in range(1,13)]
        mes_escolhido = st.selectbox("Mês", meses, index=0)
    with col2:
        anos = [str(y) for y in range(2023, 2031)]
        ano_escolhido = st.selectbox("Ano", anos, index=2)  # ex. default 2025
    
    periodo = f"{mes_escolhido}/{ano_escolhido}"

    st.write("**Selecione o arquivo Excel** com as abas de planejamento.")
    uploaded_file = st.file_uploader("Arquivo Excel", type=["xls","xlsx"])

    if uploaded_file is not None:
        st.write("Processando... aguarde.")
        dict_abas = processar_planilha(uploaded_file, periodo)
        
        if not dict_abas:
            st.warning("Nenhuma aba válida encontrada ou nenhum registro extraído.")
            return
        
        st.success(f"Planilha processada com sucesso! {len(dict_abas)} abas com dados.")

        # Multiselect para escolher quais abas baixar
        abas_list = list(dict_abas.keys())
        abas_selecionadas = st.multiselect(
            "Selecione as abas para download (ou deixe todas)",
            abas_list, default=abas_list
        )

        # Download individual
        st.write("### Download individual:")
        for aba in abas_selecionadas:
            df_temp = dict_abas[aba]
            gerar_download_link_para_df(df_temp, f"{aba}.csv")

        # Download zip (todas escolhidas)
        st.write("### Download ZIP de todas as abas selecionadas:")
        dict_filtrado = {k: dict_abas[k] for k in abas_selecionadas}
        zip_bytes, zip_name = gerar_download_zip(dict_filtrado)
        st.download_button(
            label="Baixar TODAS em ZIP",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/x-zip-compressed"
        )


if __name__ == "__main__":
    main()
