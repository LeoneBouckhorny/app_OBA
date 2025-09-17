import streamlit as st
from docx import Document
from collections import defaultdict
import os

# === Funções auxiliares ===
def formatar_nome(nome, uf=False):
    """
    Formata um nome ou texto:
    - Somente a primeira letra de cada palavra em maiúscula
    - Se uf=True, mantém tudo maiúsculo (para siglas de estado)
    """
    nome = ' '.join(nome.strip().split())  # remove espaços extras
    if uf:
        return nome.upper()  # mantém UF em maiúsculas
    return ' '.join(w.capitalize() for w in nome.split())

def processar_docx(uploaded_file):
    doc = Document(uploaded_file)
    dados = []

    for tabela in doc.tables:
        for i, linha in enumerate(tabela.rows):
            if i == 0:
                continue
            valores = [c.text.strip() for c in linha.cells]
            if len(valores) == 6:
                equipe, funcao, escola, cidade, estado, nome = valores
                dados.append({
                    "Equipe": equipe,
                    "Funcao": funcao.lower(),
                    "Escola": escola,
                    "Cidade": cidade,
                    "Estado": estado,
                    "Nome": nome
                })

    equipes = defaultdict(list)
    for item in dados:
        equipes[item["Equipe"]].append(item)

    novo_doc = Document()
    for equipe, membros in sorted(equipes.items(), key=lambda x: x[0]):
        # Ordenar membros: líder -> acompanhante -> alunos em ordem alfabética
        lider = [m for m in membros if "líder" in m["Funcao"] or "lider" in m["Funcao"]]
        acompanhante = [m for m in membros if "acompanhante" in m["Funcao"]]
        alunos = [m for m in membros if "aluno" in m["Funcao"]]
        alunos_sorted = sorted(alunos, key=lambda m: formatar_nome(m["Nome"]))

        ordem_final = lider + acompanhante + alunos_sorted
        for membro in ordem_final:
            novo_doc.add_paragraph(formatar_nome(membro["Nome"]))

        if membros:
            novo_doc.add_paragraph(f"Equipe: {equipe.split()[-1]}")
            novo_doc.add_paragraph(formatar_nome(membros[0]["Escola"]))
            novo_doc.add_paragraph(f"{formatar_nome(membros[0]['Cidade'])} / {formatar_nome(membros[0]['Estado'], uf=True)}")

        novo_doc.add_paragraph("")  # linha em branco entre equipes
    return novo_doc

# === Interface Streamlit ===
st.title("📋 Organizador de Equipes")
st.write("Faça upload do arquivo `.docx` e baixe o arquivo formatado.")

uploaded_file = st.file_uploader("Envie o arquivo DOCX", type=["docx"])

if uploaded_file:
    nome_base = os.path.splitext(uploaded_file.name)[0]
    novo_nome = f"{nome_base}_FORMATADO.docx"

    novo_doc = processar_docx(uploaded_file)

    # Prévia usando st.code para evitar erros de renderização
    st.subheader("Prévia das primeiras equipes:")
    preview = [p.text for p in novo_doc.paragraphs[:20]]
    st.code("\n".join(preview), language="text")

    # Salvar e disponibilizar para download
    novo_doc.save(novo_nome)
    with open(novo_nome, "rb") as f:
        st.download_button(
            label=f"📥 Baixar {novo_nome}",
            data=f,
            file_name=novo_nome,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
