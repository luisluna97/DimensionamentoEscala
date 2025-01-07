import streamlit as st
import pandas as pd
import datetime
import re
import io
import zipfile
from io import BytesIO

#################################
# FUNÇÕES AUXILIARES
#################################

def hhmm_to_minutes(hhmm):
    """ Converte 'HH:MM' em minutos desde 00:00 (0..1440). """
    hh, mm = hhmm.split(":")
    return int(hh)*60 + int(mm)

def minutes_to_hhmm(total_minutes):
    """ Converte minutos (podendo ser >1440) em 'HH:MM' no formato 24h (mod 24). """
    total_minutes = total_minutes % (24*60)
    hh = total_minutes // 60
    mm = total_minutes % 60
    return f"{hh:02d}:{mm:02d}"

def eh_celula_numero(valor):
    """ Verifica se a célula começa com dígito. Ex.: '3', '15'... """
    if pd.isna(valor):
        return False
    return bool(re.match(r'^\d+', str(valor).strip()))

def calcula_intervalos(h_inicio_str, carga_h_str):
    """
    Se for 7H => 4h + 1h pausa + 3h.
    Caso contrário => turno direto (2H,4H,6H, etc.).
    Retorna (Entrada1, Saida1, Entrada2, Saida2).
    """
    match = re.match(r"(\d+)", carga_h_str)  # extrai o número
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

def ignorar_poluicao(celula_funcao, celula_carga):
    """
    Retorna True se detectarmos poluições como:
    - "5 HORAS AQUI"
    - "INTERVALO"
    Assim, podemos ignorar esse bloco.
    """
    texto_funcao = str(celula_funcao).upper()
    texto_carga  = str(celula_carga).upper()

    if "HORAS AQUI" in texto_funcao or "HORAS AQUI" in texto_carga:
        return True
    if "INTERVALO" in texto_funcao or "INTERVALO" in texto_carga:
        return True
    
    return False


#################################
# LÓGICA DE PROCESSAMENTO
#################################

def processar_planilha(excel_file) -> dict:
    """
    Recebe um arquivo Excel (BytesIO ou similar),
    retorna um dicionário { nomeAba: DataFrame } com o resultado.
    """

    # Lê o Excel via pd.ExcelFile
    xls = pd.ExcelFile(excel_file)
    abas_ignorar = ["BASE", "TABELAS", "PADRÕES", "ARQUIVO BASE"]
    abas_validas = [aba for aba in xls.sheet_names if aba.upper() not in [a.upper() for a in abas_ignorar]]

    dict_abas_result = {}

    for nome_aba in abas_validas:
        # Lê a aba inteira (sem header)
        df = pd.read_excel(excel_file, sheet_name=nome_aba, header=None)
        nrows, ncols = df.shape

        # Vamos até a coluna de índice 800 (zero-based), se existir
        col_inicio = 6
        col_fim = 800
        if col_fim >= ncols:
            col_fim = ncols - 1

        # Linha 2 do Excel => df index = 1 => mapeia col->horario
        if nrows < 2:
            # Poucas linhas, nada a fazer
            continue
        
        col_to_horario = {}
        for c in range(col_inicio, col_fim+1):
            if c < ncols:
                val = df.iloc[1,c]
                if pd.isna(val):
                    val = ""
                col_to_horario[c] = str(val).strip()
            else:
                col_to_horario[c] = ""

        # Área de planejamento começa na linha 87 => index 86
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

                    # Ler possíveis próximas células
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
                        # Ex.: "AUX LIDER,6H"
                        partes = [p.strip() for p in val_next1.split(",")]
                        if len(partes) == 2:
                            funcao, carga = partes
                            col_consumidas = 2
                        else:
                            funcao = val_next1
                            carga  = val_next2
                            col_consumidas = 3
                    else:
                        # Se val_next1 for '6H' etc.
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

                    # ---- NOVA LÓGICA: ignorar números “soltos” ----
                    # Se não tiver função e não for uma carga válida (ex '6H'),
                    # significa que é só um número sem nada => ignora
                    if not funcao and not re.match(r'^\d+H$', carga):
                        col += col_consumidas
                        continue

                    # Ignorar poluições ("HORAS AQUI", "INTERVALO")
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

                    # Salva no array
                    registros.append({
                        "Quantidade": quantidade_str,
                        "Cargo": funcao,
                        "CargaHoraria": carga,
                        "Entrada1": e1,
                        "Saida1": s1,
                        "Entrada2": e2,
                        "Saida2": s2
                    })

                    # Lógica de pulo (para evitar duplicar o pernoite)
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

        # Monta DataFrame final
        df_out = pd.DataFrame(registros, columns=[
            "Quantidade", "Cargo", "CargaHoraria",
            "Entrada1", "Saida1", "Entrada2", "Saida2"
        ])
        dict_abas_result[nome_aba] = df_out

    return dict_abas_result


#################################
# FUNÇÕES DE DOWNLOAD (STREAMLIT)
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
# APP STREAMLIT (app.py)
#################################

def main():
    st.title("Conversor de Planilha para CSV - Pernoites")

    st.write("""
    **Instruções**:
    - Faça o *upload* de um arquivo Excel contendo as abas de planejamento.
    - As abas "BASE", "TABELAS", "PADRÕES" e "ARQUIVO BASE" serão ignoradas.
    - O script extrairá informações (quantidade, função, carga horária, horário) e 
      tentará unificar turnos pernoite.
    - Algumas poluições como "5 HORAS AQUI" e "INTERVALO" serão descartadas.
    - Agora, se encontrar um número sozinho sem função/carga, ele também será **ignorado**.
    """)

    uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xls", "xlsx"])

    if uploaded_file is not None:
        st.write("Processando... aguarde.")
        dict_abas = processar_planilha(uploaded_file)
        
        if not dict_abas:
            st.warning("Nenhuma aba válida encontrada ou nenhum registro extraído.")
            return
        
        st.success(f"Planilha processada com sucesso! Foram encontradas {len(dict_abas)} abas válidas.")

        # Permitir escolher quais abas o usuário quer baixar
        abas_list = list(dict_abas.keys())
        abas_selecionadas = st.multiselect(
            "Selecione as abas para download (ou deixe vazio para baixar todas):",
            abas_list, 
            default=abas_list
        )

        # Botões de download individual
        if abas_selecionadas:
            st.write("### Download individual das abas selecionadas:")
            for aba in abas_selecionadas:
                df_temp = dict_abas[aba]
                gerar_download_link_para_df(df_temp, f"{aba}.csv")

        # Botão para baixar todas as abas (ou selecionadas) em ZIP
        st.write("### Download de todas as abas em ZIP:")
        if abas_selecionadas:
            dict_filtrado = {k: dict_abas[k] for k in abas_selecionadas}
        else:
            dict_filtrado = dict_abas

        zip_bytes, zip_name = gerar_download_zip(dict_filtrado)
        st.download_button(
            label="Baixar TODAS em ZIP",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/x-zip-compressed"
        )


if __name__ == "__main__":
    main()
