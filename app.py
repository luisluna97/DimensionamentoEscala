import streamlit as st

def main():
    st.title("Teste com Upload")

    uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xls","xlsx"])

    if uploaded_file is not None:
        st.write("Arquivo recebido com sucesso!")
        # Sem processamento de planilha por enquanto
        st.write("Fim do processamento.")

if __name__ == "__main__":
    main()
