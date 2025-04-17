import streamlit as st
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from pytz import timezone
from PIL import Image

def main():
    st.set_page_config(layout="centered", page_title="Gerador de Laudo")

    # --- Cores UI ---
    UI_COR_AZUL_SPTC = "#eaeff2"
    UI_COR_CINZA_SPTC = "#6E6E6E"

    # --- Exibir data ---
    data_placeholder = st.empty()

    def atualizar_data():
        try:
            brasilia_tz = timezone('America/Sao_Paulo')
            now = datetime.now(brasilia_tz)
            dias_semana_portugues = [
                "Segunda-feira", "Terça-feira", "Quarta-feira",
                "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"
            ]
            meses_portugues = [
                "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
            ]
            dia_semana = dias_semana_portugues[now.weekday()]
            mes = meses_portugues[now.month - 1]
            data_formatada = f"{dia_semana}, {now.day} de {mes} de {now.year}"
            data_placeholder.markdown(f"""
                <div style='text-align: right; font-size: 0.8em; color: {UI_COR_CINZA_SPTC}; margin-bottom: 15px;'>
                    <span>{data_formatada}</span><br>
                    <span style='font-size: 0.8em;'>(Goiânia-GO)</span>
                </div>
            """, unsafe_allow_html=True)
        except Exception as e:
            fallback_str = datetime.now().strftime("%d/%m/%Y")
            data_placeholder.markdown(f"""
                <div style='text-align: right; font-size: 0.8em; color: #FF5555; margin-bottom: 15px;'>
                    <span>{fallback_str} (Local)</span><br>
                    <span style='font-size: 0.8em;'>Erro Fuso Horário: {e}</span>
                </div>
            """, unsafe_allow_html=True)

    atualizar_data()

    # --- Cabeçalho ---
    col_logo, col_titulo = st.columns([1, 5])
    with col_logo:
        logo_path = "logo_policia_cientifica.png"
        try:
            st.image(logo_path, width=100)
        except FileNotFoundError:
            st.error(f"Erro: Logo '{logo_path}' não encontrado.")
        except Exception as e:
            st.warning(f"Logo não carregado: {e}")

    with col_titulo:
        st.markdown(f"<h1 style='color: {UI_COR_AZUL_SPTC}; margin: 0;'>Gerador de Laudo Pericial</h1>", unsafe_allow_html=True)
        st.markdown(f"<p style='color: {UI_COR_CINZA_SPTC}; font-size: 1em;'>Identificação de Drogas - SPTC/GO</p>", unsafe_allow_html=True)

    st.markdown("---")

# ====================
# Constants
# ====================
TIPOS_MATERIAL_BASE = {
    "v": "vegetal dessecado",
    "po": "pulverizado",
    "pd": "petrificado",
    "r": "resinoso"
}

TIPOS_EMBALAGEM_BASE = {
    "e": "microtubo do tipo eppendorf",
    "z": "embalagem do tipo ziplock",
    "a": "papel alumínio",
    "pl": "plástico",
    "pa": "papel"
}

        CORES_FEMININO_EMBALAGEM = {
            "t": "transparente", "b": "branca", "az": "azul", "am": "amarela",
            "vd": "verde", "vm": "vermelha", "p": "preta", "c": "cinza",
            "m": "marrom", "r": "rosa", "l": "laranja", "violeta": "violeta", "roxa": "roxa"
        }

        QUANTIDADES_EXTENSO = {
            1: "uma", 2: "duas", 3: "três", 4: "quatro", 5: "cinco",
            6: "seis", 7: "sete", 8: "oito", 9: "nove", 10: "dez"
        }

        meses_portugues_dict = {
            "January": "janeiro", "February": "fevereiro", "March": "março",
            "April": "abril", "May": "maio", "June": "junho", "July": "julho",
            "August": "agosto", "September": "setembro", "October": "outubro",
            "November": "novembro", "December": "dezembro"
        }

        # --- Funções Auxiliares ---
        def pluralizar_palavra(palavra, quantidade):
            if quantidade == 1:
                return palavra
            if palavra in ["microtubo do tipo eppendorf", "embalagem do tipo ziplock"]:
                return palavra
            if palavra.endswith('m'):
                return re.sub(r'm$', 'ns', palavra)
            if palavra.endswith('ão'):
                return re.sub(r'ão$', 'ões', palavra)
            elif palavra.endswith(('r', 'z')):
                return palavra + 'es'
            else:
                return palavra + 's'

        # Mapeamento de tipo de material para os números dos itens (2.x)
        material_para_itens = {}

        # Novo mapeamento para rastrear qual número do item está vinculado a qual tipo
        referencias_itens_por_tipo = {}

        def mapear_material_para_itens(itens):
            global material_para_itens, referencias_itens_por_tipo
            material_para_itens = {}
            referencias_itens_por_tipo = {}
            for idx, item in enumerate(itens):
                tipo = item.get('tipo_material')
                numero_item = f"2.{idx + 1}"
                if tipo:
                    if tipo not in material_para_itens:
                        material_para_itens[tipo] = []
                    material_para_itens[tipo].append(numero_item)
                    referencias_itens_por_tipo[numero_item] = tipo

        # Função para obter string de referência ("2.1", "2.2 e 2.3", etc.)
        def formatar_referencia_material(tipo_codigo):
            itens = material_para_itens.get(tipo_codigo, [])
            if not itens:
                return "[sem referência]"
            if len(itens) == 1:
                return itens[0]
            return " e ".join(itens)

        # Essa estrutura será usada ao construir os tópicos 4, 5 e 6
        # Ex: formatar_referencia_material("v") → "2.1 e 2.3"

        def obter_quantidade_extenso(qtd):
            return QUANTIDADES_EXTENSO.get(qtd, str(qtd))

        def add_paragraph(doc, text, bold=False, align='justify', size=12):
            align_map = {
                'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
                'center': WD_ALIGN_PARAGRAPH.CENTER,
                'right': WD_ALIGN_PARAGRAPH.RIGHT
            }
            p = doc.add_paragraph()
            p.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)
            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(size)

        st.subheader("📝 1. Informações do Laudo")
        lacre = st.text_input("Digite o número do lacre da contraprova:")
        numero_laudo = st.text_input("Digite o RG da perícia:")
        st.markdown("---") # Separador visual

        st.subheader("📦 2. MATERIAL RECEBIDO")
        itens_data = []
        num_itens = st.number_input("Quantos itens deseja descrever?", min_value=1, step=1, value=1, key="num_itens")
        for i in range(int(st.session_state.get("num_itens", 1))):
            with st.expander(f"Item {i+1}"):
                col1, col2 = st.columns(2)
                with col1:
                    qtd = st.number_input(f"Quantidade de porções:", min_value=1, step=1, value=1, key=f"qtd_{i}")
                    tipo_mat_code = st.selectbox(f"Tipo de material (v, po, pd, r):", options=list(TIPOS_MATERIAL_BASE.keys()), key=f"tipo_mat_{i}")
                    emb_code = st.selectbox(f"Tipo de embalagem (e, z, a, pl, pa):", options=list(TIPOS_EMBALAGEM_BASE.keys()), key=f"emb_{i}")
                with col2:
                    ref = st.text_input(f"Referência do subitem:", key=f"ref_{i}")
                    pessoa = st.text_input(f"Pessoa relacionada (opcional):", key=f"pessoa_{i}")
                    cor_emb_code = None
                    if emb_code == 'pl' or emb_code == 'pa':
                        cor_emb_code = st.selectbox(f"Cor da embalagem:", options=list(CORES_FEMININO_EMBALAGEM.keys()), key=f"cor_emb_{i}")
                    else:
                        cor_emb_code = None

                itens_data.append({
                    'qtd': qtd,
                    'tipo_mat_code': tipo_mat_code,
                    'emb_code': emb_code,
                    'cor_emb_code': cor_emb_code,
                    'ref': ref,
                    'pessoa': pessoa
                })

        st.subheader("📷 3. Upload da Imagem")
        uploaded_image = st.file_uploader("Selecione uma imagem do material recebido (opcional):", type=["png", "jpg", "jpeg"])
    if st.button("✅ Gerar Laudo"):
        document = Document()

        # Seção 2 - Material Recebido
        add_paragraph(document, "2 MATERIAL RECEBIDO PARA EXAME", bold=True)

        if uploaded_image:
            try:
                document.add_picture(uploaded_image, width=Inches(5.0))
                add_paragraph(document, "Ilustração 1 – Material recebido para exame.", bold=True, align='center', size=10)
            except Exception as e:
                st.error(f"Erro ao processar a imagem: {e}")
                add_paragraph(document, "(Imagem não pôde ser carregada)", bold=True, align='center', size=10)
        else:
            add_paragraph(document, "\nIlustração 1 – Material recebido para exame.", bold=True)

        subitens_cannabis = {}
        subitens_cocaina = {}

        for i, item_info in enumerate(itens_data):
            qtd = item_info['qtd']
            tipo_mat_code = item_info['tipo_mat_code']
            emb_code = item_info['emb_code']
            cor_emb_code = item_info['cor_emb_code']
            ref = item_info['ref']
            pessoa = item_info['pessoa']

            tipo_material = TIPOS_MATERIAL_BASE[tipo_mat_code]
            embalagem = TIPOS_EMBALAGEM_BASE[emb_code]

            if cor_emb_code:
                embalagem += f" de cor {CORES_FEMININO_EMBALAGEM[cor_emb_code]}"

            texto_item = f"2.{i+1} {qtd} ({obter_quantidade_extenso(qtd)}) {pluralizar_palavra('porção', qtd)} de material {tipo_material}, "
            texto_item += f"{'acondicionada em' if qtd == 1 else 'acondicionadas, individualmente, em'} {pluralizar_palavra(embalagem, qtd)}, "
            texto_item += f"referente à amostra do subitem {ref} do laudo de constatação supracitado"
            texto_item += f", relacionada a {pessoa}" if pessoa else ""
            texto_item += "."

            add_paragraph(document, texto_item)

            if tipo_mat_code in ["v", "r"]:
                subitens_cannabis[ref] = f"2.{i+1}"
            elif tipo_mat_code in ["po", "pd"]:
                subitens_cocaina[ref] = f"2.{i+1}"

        # OBJETIVO DOS EXAMES
        add_paragraph(document, "\n3 OBJETIVO DOS EXAMES", bold=True)
        add_paragraph(document, "Visa esclarecer à autoridade requisitante quanto às características do material apresentado, bem como se ele contém substância de uso proscrito no Brasil e capaz de causar dependência física e/ou psíquica. O presente laudo pericial busca demonstrar a materialidade da infração penal apurada.", align='justify')

        # EXAMES
        add_paragraph(document, "\n4 EXAMES", bold=True)
        has_cannabis_item = bool(subitens_cannabis)
        has_cocaina_item = bool(subitens_cocaina)

        if has_cannabis_item:
            add_paragraph(document, "4.1 Exames realizados para pesquisa de Cannabis sativa L. (maconha)")
            add_paragraph(document, "4.1.1 Ensaio químico com Fast blue salt B: teste de cor em reação com solução aquosa de sal de azul sólido B em meio alcalino;")
            add_paragraph(document, "4.1.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas contendo eluentes apropriados e posterior revelação com solução aquosa de azul sólido B.")
        if has_cocaina_item:
            idx = "4.2" if has_cannabis_item else "4.1"
            add_paragraph(document, f"{idx} Exames realizados para pesquisa de cocaína")
            add_paragraph(document, f"{idx}.1 Ensaio químico com teste de tiocianato de cobalto-reação de cor com solução de tiocianato de cobalto em meio ácido;")
            add_paragraph(document, f"{idx}.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas com eluentes apropriados e revelação com solução de iodo platinado.")
        if not has_cannabis_item and not has_cocaina_item:
            add_paragraph(document, "4.1 Exames realizados")
            add_paragraph(document, "4.1.1 Exame macroscópico;")

        # RESULTADOS
        add_paragraph(document, "\n5 RESULTADOS", bold=True)
        if has_cannabis_item:
            subitens = " e ".join(subitens_cannabis.keys())
            label = "no subitem" if len(subitens_cannabis) == 1 else "nos subitens"
            add_paragraph(document, f"5.1 Resultados obtidos para o(s) material(is) descrito(s) {label} {subitens}:")
            add_paragraph(document, "5.1.1 No ensaio com Fast blue salt B, foram obtidas coloração característica para canabinol e tetrahidrocanabinol (princípios ativos da Cannabis sativa L.).")
            add_paragraph(document, "5.1.2 Na CCD, obtiveram-se perfis cromatográficos coincidentes com o material de referência (padrão de Cannabis sativa L.); portanto, a substância tetrahidrocanabinol está presente nos materiais questionados.")
        if has_cocaina_item:
            idx = "5.2" if has_cannabis_item else "5.1"
            subitens = " e ".join(subitens_cocaina.keys())
            label = "no subitem" if len(subitens_cocaina) == 1 else "nos subitens"
            add_paragraph(document, f"{idx} Resultados obtidos para o(s) material(is) descrito(s) {label} {subitens}:")
            add_paragraph(document, f"{idx}.1 No teste de tiocianato de cobalto, foram obtidas coloração característica para cocaína;")
            add_paragraph(document, f"{idx}.2 Na CCD, obteve-se perfis cromatográficos coincidentes com o material de referência (padrão de cocaína); portanto, a substância cocaína está presente nos materiais questionados.")

        # CONCLUSÃO
        add_paragraph(document, "\n6 CONCLUSÃO", bold=True)
        conclusoes = []
        if has_cannabis_item:
            subitens = " e ".join(subitens_cannabis.keys())
            label = "no subitem" if len(subitens_cannabis) == 1 else "nos subitens"
            conclusoes.append(f"no(s) material(is) descrito(s) {label} {subitens}, foi detectada a presença de partes da planta Cannabis sativa L., vulgarmente conhecida por maconha. A Cannabis sativa L. contém princípios ativos chamados canabinóis, dentre os quais se encontra o tetrahidrocanabinol, substância perturbadora do sistema nervoso central. Tanto a Cannabis sativa L. quanto a tetrahidrocanabinol são proscritas no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.")
        if has_cocaina_item:
            subitens = " e ".join(subitens_cocaina.keys())
            conclusoes.append(f"no(s) material(is) descrito(s) no(s) subitem(ns) {subitens}, foi detectada a presença de cocaína, substância alcaloide estimulante do sistema nervoso central. A cocaína é proscrita no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.")

        if conclusoes:
            texto_final = "A partir das análises realizadas, conclui-se que, " + " Outrossim, ".join(conclusoes)
        else:
            texto_final = "A partir das análises realizadas, conclui-se que não foram detectadas substâncias de uso proscrito nos materiais analisados."
        add_paragraph(document, texto_final, align='justify')

        # CUSTÓDIA DO MATERIAL
        add_paragraph(document, "\n7 CUSTÓDIA DO MATERIAL", bold=True)
        add_paragraph(document, "7.1 Contraprova")
        add_paragraph(document, f"7.1.1 A amostra contraprova ficará armazenada neste Instituto, conforme Portaria 0003/2019/SSP  (Lacre nº {lacre}).")

        # REFERÊNCIAS
        add_paragraph(document, "\nREFERÊNCIAS", bold=True)
        referencias = [
            "BRASIL. Ministério da Saúde. Portaria SVS/MS n° 344, de 12 de maio de 1998...",
            "GOIÁS. Secretaria de Estado da Segurança Pública. Portaria nº 0003/2019/SSP...",
            "SWGDRUG: Scientific Working Group for the Analysis of Seized Drugs..."
        ]
        for ref in referencias:
            add_paragraph(document, ref, size=10)

        # Geração do arquivo
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        st.success("✅ Laudo gerado com sucesso!")
        st.download_button(
            label="📄 Baixar Laudo",
            data=buffer,
            file_name=f"Laudo_{numero_laudo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
