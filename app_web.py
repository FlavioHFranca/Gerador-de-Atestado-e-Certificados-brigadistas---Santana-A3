import os
import sys
import io
import re
import gc
import zipfile
import tempfile
import subprocess
import streamlit as st
import pandas as pd
import docx
import pdfplumber
from docxtpl import DocxTemplate

CAMINHO_LOGO = "logo_santanaa3.png"
MODELO_CERTIFICADO = "modeloCertificado.docx"

icone_pagina = CAMINHO_LOGO if os.path.exists(CAMINHO_LOGO) else "📜"

st.set_page_config(
    page_title="Santana A3 - Emissor de Brigada",
    page_icon=icone_pagina,
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton button { width: 100%; height: 3.1em; font-size: 1.05em; font-weight: bold; }
    .stDownloadButton button { width: 100%; height: 3.3em; font-size: 1.1em; font-weight: bold; background-color: #2e7d32; color: white; }
</style>
""", unsafe_allow_html=True)

if os.path.exists(CAMINHO_LOGO):
    col_logo, _ = st.columns([1, 2])
    with col_logo:
        st.image(CAMINHO_LOGO, width=100)

st.title("Emissor de Atestado & Certificados")
st.caption("Faça upload do Atestado de Brigada em DOCX ou PDF para gerar os certificados.")


# --- 1. CONVERSÃO DOCX PARA PDF ---
def converter_docx_para_pdf(caminho_docx, pasta_saida):
    nome_base = os.path.splitext(os.path.basename(caminho_docx))[0]
    caminho_pdf = os.path.join(pasta_saida, f"{nome_base}.pdf")

    if sys.platform == "win32":
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(caminho_docx), ReadOnly=True, ConfirmConversions=False)
            doc.SaveAs2(os.path.abspath(caminho_pdf), FileFormat=17)
            doc.Close(SaveChanges=0)
        finally:
            if word:
                word.Quit()
                del word
            gc.collect()
            pythoncom.CoUninitialize()
    else:
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf",
            "--outdir", pasta_saida, caminho_docx
        ], check=True)

    return caminho_pdf


# --- 2. EXTRATORES (DOCX E PDF) ---
def extrair_metadados_texto(texto_unificado):
    empresa = "MAGAZINE LUIZA S/A"
    match_resp = re.search(r"Responsável pelo uso:\s*(.+)", texto_unificado, re.IGNORECASE)
    if match_resp:
        empresa = match_resp.group(1).strip()

    cnpj = ""
    match_cnpj = re.search(r"CNPJ:\s*([\d\.\/\-]+)", texto_unificado, re.IGNORECASE)
    if match_cnpj:
        cnpj = match_cnpj.group(1).strip()

    legislacao = "CBMPI nº 17/2019"
    match_leg = re.search(r"(NT\s*\d+/\d+\s*[A-Z]+)", texto_unificado, re.IGNORECASE)
    if match_leg:
        raw_leg = match_leg.group(1).strip()
        partes = raw_leg.split()
        legislacao = f"{partes[2]} nº {partes[1]}" if len(partes) >= 3 else raw_leg

    data_extenso = "14 de agosto de 2026"
    match_rodape = re.search(r",\s*(\d{1,2}\s+de\s+[a-zçA-ZÇ]+\s+de\s+\d{4})", texto_unificado, re.IGNORECASE)
    if match_rodape:
        data_extenso = match_rodape.group(1).strip()

    return empresa, cnpj, legislacao, data_extenso


def extrair_dados_docx(arquivo_bytes):
    doc = docx.Document(arquivo_bytes)
    texto_unificado = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    empresa, cnpj, legislacao, data_extenso = extrair_metadados_texto(texto_unificado)

    brigadistas = []
    if doc.tables:
        tabela = doc.tables[0]
        for linha in tabela.rows[1:]:
            cols = [c.text.strip() for c in linha.cells]
            if cols and cols[0]:
                brigadistas.append({
                    "Nome": cols[0],
                    "CPF": cols[1] if len(cols) > 1 else "",
                    "Treinamento": cols[2] if len(cols) > 2 else "Formação",
                    "Nível": cols[3] if len(cols) > 3 else "Básico",
                    "Carga Horária": cols[4] if len(cols) > 4 else "04h"
                })

    return {"empresa": empresa, "cnpj": cnpj, "data": data_extenso, "legislacao": legislacao, "brigadistas": brigadistas}


def extrair_dados_pdf(arquivo_bytes):
    texto_unificado = ""
    brigadistas = []

    with pdfplumber.open(arquivo_bytes) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto_unificado += t + "\n"

            tabelas = pagina.extract_tables()
            for tab in tabelas:
                for linha in tab[1:]:  # Pula o cabeçalho
                    cols = [str(c).strip() for c in linha if c is not None]
                    # Garante que a linha tem nome e CPF plausível
                    if cols and len(cols) >= 2 and cols[0] and not cols[0].lower().startswith("nome"):
                        brigadistas.append({
                            "Nome": cols[0],
                            "CPF": cols[1] if len(cols) > 1 else "",
                            "Treinamento": cols[2] if len(cols) > 2 else "Formação",
                            "Nível": cols[3] if len(cols) > 3 else "Básico",
                            "Carga Horária": cols[4] if len(cols) > 4 else "04h"
                        })

    empresa, cnpj, legislacao, data_extenso = extrair_metadados_texto(texto_unificado)
    return {"empresa": empresa, "cnpj": cnpj, "data": data_extenso, "legislacao": legislacao, "brigadistas": brigadistas}


# --- 3. INTERFACE DE UPLOAD HÍBRIDA ---
arquivo_upload = st.file_uploader(
    "📂 Selecione o Atestado (.docx ou .pdf)",
    type=None, # Sem restrição no seletor nativo do celular
    help="Você pode enviar o arquivo original em Word ou em PDF"
)

if arquivo_upload:
    nome_arq = arquivo_upload.name.lower()
    bytes_arquivo = arquivo_upload.getvalue()

    if not (nome_arq.endswith(".docx") or nome_arq.endswith(".pdf")):
        st.error(f"⚠️ O arquivo '{arquivo_upload.name}' não é suportado. Por favor, envie um arquivo .docx ou .pdf.")
    else:
        # Identifica se lê via DOCX ou PDF
        if nome_arq.endswith(".docx"):
            dados = extrair_dados_docx(io.BytesIO(bytes_arquivo))
            eh_pdf_origem = False
        else:
            dados = extrair_dados_pdf(io.BytesIO(bytes_arquivo))
            eh_pdf_origem = True

        brigadistas = dados["brigadistas"]
        total_alunos = len(brigadistas)

        with st.expander("🏢 Dados Identificados do Atestado", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                empresa_edit = st.text_input("Empresa / Filial", value=dados["empresa"])
                cnpj_edit = st.text_input("CNPJ", value=dados["cnpj"])
            with col2:
                data_edit = st.text_input("Data de Emissão", value=dados["data"])
                leg_edit = st.text_input("Legislação Aplicada", value=dados["legislacao"])

        st.subheader(f"👥 Brigadistas Encontrados ({total_alunos})")
        if total_alunos > 0:
            df_brigadistas = pd.DataFrame(brigadistas)
            df_brigadistas.index = df_brigadistas.index + 1
            st.dataframe(
                df_brigadistas[["Nome", "CPF", "Nível", "Carga Horária"]],
                use_container_width=True,
                height=min(300, (total_alunos + 1) * 38)
            )
        else:
            st.warning("⚠️ Nenhuma linha de brigadista foi identificada na tabela do arquivo.")

        match_filial = re.search(r"Filial\s*(\d+)", empresa_edit, re.IGNORECASE)
        nome_filial = f"Filial_{match_filial.group(1)}" if match_filial else "Filial"

        st.markdown("---")

        if st.button("⚡ Gerar Atestado e Certificados em PDF", type="primary", disabled=(total_alunos == 0)):
            if not os.path.exists(MODELO_CERTIFICADO):
                st.error(f"Erro: O modelo '{MODELO_CERTIFICADO}' não foi encontrado.")
            else:
                with st.spinner("Processando documentos e convertendo para PDF..."):
                    buffer_zip = io.BytesIO()

                    with tempfile.TemporaryDirectory() as temp_dir:
                        pasta_docx = os.path.join(temp_dir, "docx")
                        pasta_pdf = os.path.join(temp_dir, "pdf")
                        os.makedirs(pasta_docx, exist_ok=True)
                        os.makedirs(pasta_pdf, exist_ok=True)

                        total_operacoes = total_alunos + 1
                        barra_progresso = st.progress(0.0)

                        # --- PASSO 1: ATESTADO EM PDF ---
                        nome_pdf_atestado = f"{nome_filial} - Atestado_de_Formacao_de_Brigada.pdf"
                        caminho_atestado_pdf_final = os.path.join(pasta_pdf, nome_pdf_atestado)

                        if eh_pdf_origem:
                            # Se já era PDF, apenas salva diretamente na pasta final
                            with open(caminho_atestado_pdf_final, "wb") as f_out:
                                f_out.write(bytes_arquivo)
                        else:
                            # Se era DOCX, converte para PDF
                            caminho_atestado_docx = os.path.join(pasta_docx, f"{nome_filial} - Atestado.docx")
                            with open(caminho_atestado_docx, "wb") as f_out:
                                f_out.write(bytes_arquivo)
                            converter_docx_para_pdf(caminho_atestado_docx, pasta_pdf)
                            # Renomeia para padronizar o nome no ZIP
                            gerado = os.path.join(pasta_pdf, f"{nome_filial} - Atestado.pdf")
                            if os.path.exists(gerado):
                                os.rename(gerado, caminho_atestado_pdf_final)

                        barra_progresso.progress(1 / total_operacoes)

                        # --- PASSO 2: CERTIFICADOS EM PDF ---
                        for idx, p in enumerate(brigadistas, start=1):
                            doc = DocxTemplate(MODELO_CERTIFICADO)
                            contexto = {
                                "empresa": empresa_edit,
                                "cnpj": cnpj_edit,
                                "data_treinamento": data_edit,
                                "legislacao": leg_edit,
                                "nome": p["Nome"],
                                "cpf": p["CPF"],
                                "tipo_treinamento": p["Treinamento"],
                                "nivel": p["Nível"],
                                "carga_horaria": p["Carga Horária"]
                            }
                            doc.render(contexto)

                            nome_limpo = "".join(c for c in p["Nome"] if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")
                            caminho_cert_docx = os.path.join(pasta_docx, f"{nome_filial} - Certificado_{nome_limpo}.docx")
                            doc.save(caminho_cert_docx)

                            converter_docx_para_pdf(caminho_cert_docx, pasta_pdf)
                            barra_progresso.progress((idx + 1) / total_operacoes)

                        # --- PASSO 3: COMPACTAÇÃO NO ZIP ---
                        with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for arq_pdf in sorted(os.listdir(pasta_pdf)):
                                if arq_pdf.lower().endswith(".pdf"):
                                    caminho_completo = os.path.join(pasta_pdf, arq_pdf)
                                    zip_file.write(caminho_completo, arcname=arq_pdf)

                    st.success(f"✓ Concluído! 1 Atestado + {total_alunos} Certificados gerados em PDF para a {nome_filial.replace('_', ' ')}.")
                    st.download_button(
                        label=f"📥 Baixar Pacote Completo (.zip)",
                        data=buffer_zip.getvalue(),
                        file_name=f"Documentos_Brigada_{nome_filial}.zip",
                        mime="application/zip"
                    )