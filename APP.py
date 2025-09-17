import streamlit as st
from docx import Document
from collections import defaultdict
import os

# === Funções auxiliares ===
def formatar_nome(texto):
    """Coloca a primeira letra de cada palavra em maiúscula, exceto UF."""
    texto = ' '.join(texto.strip().split())
    # Se for sigla de estado, mantém maiúscula
    if len(texto) <= 3 and texto.isupper():
        return texto
    return ' '.join(w.capitalize() for w in texto.split())

def processar_docx(uploaded_file):
    doc = Document(uploaded_file)
    dados = []

    # Ler a tabela do DOCX
    for tabela in doc.tables:
        for i, linha in enumerate(tabela.rows):
            if i == 0:
                continue  # Ignora cabeçalho
            valores = [c.text.strip() for c in linha.cells]
            if len(valores) == 7:
                valido, equipe, funcao, escola, cidade, estado, nome = valores
                dados.append({
                    "Valido": valido,
                    "Equipe": equipe,
                    "Funcao": funcao.lower(),
                    "Escola": escola,
                    "Cidade": cidade,
                    "Estado": estado.upper(),
                    "Nome": nome
                })

    # Agrupar por equipe
    equipes = defaultdict(list)
    for item in dados:
        equipes[item["Equipe"]].append(item)

    # Criar novo documento
    novo_doc = Document()
    for equipe, membros in sorted(equipes.items(), key=lambda x: x[0]):
        lider = [m for m in membros if "líder" in m["Funcao"] or "lider" in m["Funcao"]]
        acompanhante = [m for m in membros if "acompanhante" in m["Funcao"]]
        alunos = [m for m in membros if "aluno" in m["Funcao"]]
        alunos_sorted = sorted(alunos, key=lambda m: formatar_nome(m["Nome"]))

        ordem_final = lider + acompanhante + alunos_sorted

        # Adiciona nomes ao documento
        for membro in ordem_final:
            novo_doc.add_paragraph(formatar_nome(membro["Nome"]))

        # Informação adicional por equipe
        if membros:
            # Pegamos o valor de VALIDO do primeiro membro (assumindo que seja igual para todos)
            novo_doc.add_paragraph(f"Equipe: {equipe.split()[-1]}")
            novo_doc.add_paragraph(f"Lançamentos Válidos: {membros[0]['Valido']} m")
            novo_doc.add_paragraph(formatar_nome(membros[0]["Escola"]))
            novo_doc.add_paragraph(f"{formatar_nome(membros[0]['Cidade'])} / {membros[0]['Estado']}")
        
        novo_doc.add_paragraph("")  # Linha em branco entre equipes

    return novo_doc

# === Interface Streamlit ===
st.title("📋 Organizador de Equipes com Lançamentos Válidos")
st.write("Faça upload do arquivo `.docx` e baixe o arquivo formatado por equipe.")

uploaded_file = st.file_uploader("Envie o arquivo DOCX", type=["docx"])

if uploaded_file:
    nome_base = os.path.splitext(uploaded_file.name)[0]
    novo_nome = f"{nome_base}_FORMATADO.docx"

    novo_doc = processar_docx(uploaded_file)

    # Pré-visualização
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
