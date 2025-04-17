"""
Gerador de Laudo Pericial v2.4 (Streamlit - Logo, Date/Time)

Este script gera laudos periciais para identificação de drogas e substâncias correlatas
usando o Streamlit com a logo da Polícia Científica e exibição de data/hora.

Requerimentos:
    - streamlit
    - python-docx
    - Pillow (PIL)
    - pytz

Uso:
    1. Instale as dependências: pip install streamlit python-docx Pillow pytz
    2. Execute o script: streamlit run gerador_laudo.py
    3. Interaja com a interface web para gerar o laudo.
    4. Baixe o laudo gerado como um arquivo .docx.
"""

import re
from datetime import datetime
import io
import pytz
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import time  # For the clock

# --- Constantes ---
# (Mantenho as constantes conforme sua definição)
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

meses_portugues = {
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

def obter_quantidade_extenso(qtd):
    return QUANTIDADES_EXTENSO.get(qtd, str(qtd))

def adicionar_paragrafo(doc, text, style=None, align=None, color=None, size=None, bold=False, italic=False):
    """Adiciona um parágrafo ao documento docx com formatação flexível."""

    p = doc.add_paragraph(style) if style else doc.add_paragraph()

    if align:
        align_map = {
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT
        }
        p.alignment = align_map.get(align)

    run = p.add_run(text)
    if color:
        run.font.color.rgb = color
    if size:
        run.font.size = Pt(size)
    if bold:
        run.font.bold = bold
    if italic:
        run.font.italic = italic

def inserir_imagem_docx(doc, image_file):
    """Insere uma imagem no documento docx."""

    try:
        if image_file:
            img = Image.open(image_file)
            width, height = img.size
            # Ajuste o tamanho da imagem conforme necessário (máximo 6 polegadas de largura)
            max_width_inches = 6
            width_inches = min(max_width_inches, width / 100)
            height_inches = height / 100 * (width_inches / (width / 100))

            doc.add_picture(image_file, width=Inches(width_inches), height=Inches(height_inches))
    except Exception as e:
        st.error(f"Erro ao inserir imagem no docx: {e}")

def configurar_estilos(doc):
    """Configura os estilos de parágrafo e caractere do documento docx."""

    # Cores da paleta
    COR_FUNDO = docx.shared.RGBColor(0x0D, 0x11, 0x17)  # Preto elegante (GitHub Dark)
    COR_TEXTO_PRINCIPAL = docx.shared.RGBColor(0xEAEAEA)  # Branco suave
    COR_TEXTO_SECUNDARIO = docx.shared.RGBColor(0x88, 0x88, 0x88)  # Cinza claro
    COR_DESTAQUE = docx.shared.RGBColor(0x3B, 0x82, 0xF6)  # Azul vibrante

    # Estilo para o título principal
    titulo_principal_style = doc.styles.add_style('TituloPrincipal', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    titulo_principal_style.font.name = 'Gadugi'
    titulo_principal_style.font.size = Pt(14)
    titulo_principal_style.font.bold = True
    titulo_principal_style.font.color.rgb = COR_TEXTO_PRINCIPAL

    # Estilo para títulos secundários
    titulo_secundario_style = doc.styles.add_style('TituloSecundario', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    titulo_secundario_style.font.name = 'Gadugi'
    titulo_secundario_style.font.size = Pt(12)
    titulo_secundario_style.font.bold = True
    titulo_secundario_style.font.color.rgb = COR_DESTAQUE

    # Estilo para parágrafos normais
    paragrafo_style = doc.styles.add_style('ParagrafoNormal', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    paragrafo_style.font.name = 'Gadugi'
    paragrafo_style.font.size = Pt(12)
    paragrafo_style.font.color.rgb = COR_TEXTO_PRINCIPAL

    # Estilo para texto itálico
    italico_style = doc.styles.add_style('Italico', docx.enum.style.WD_STYLE_TYPE.CHARACTER)
    italico_style.font.italic = True

    # Estilo para ilustrações
    ilustracao_style = doc.styles.add_style('Ilustracao', docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
    ilustracao_style.font.name = 'Gadugi'
    ilustracao_style.font.size = Pt(10)
    ilustracao_style.font.bold = True
    ilustracao_style.font.color.rgb = COR_TEXTO_SECUNDARIO
    ilustracao_style.alignment = WD_ALIGN_PARAGRAPH.CENTER

    return COR_FUNDO  # Retorna a cor de fundo para uso posterior

def aplicar_fundo(doc, cor_fundo):
    """Aplica uma cor de fundo (simulada) ao documento docx."""

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Simula o fundo escuro adicionando um retângulo em cada cabeçalho/rodapé (alternativa - complexo em python-docx)
    # Esta parte foi removida para simplificar. Se precisar muito do fundo, pesquise como adicionar shapes em cabeçalhos/rodapés no python-docx
    for section in doc.sections:
        header = section.header
        header_paragraph = header.paragraphs[0]
        header_paragraph.text = "LAUDO DE PERÍCIA CRIMINAL"
        header_paragraph.style = doc.styles['TituloPrincipal']
        header_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

        footer = section.footer
        footer_paragraph = footer.paragraphs[0]
        footer_paragraph.text = f"Página {doc.sections.index(section) + 1}"
        footer_paragraph.style = doc.styles['TituloPrincipal']
        footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def adicionar_cabecalho_rodape(doc):
    """Adiciona cabeçalho e rodapé ao documento docx."""

    for section in doc.sections:
        header = section.header
        header_paragraph = header.paragraphs[0]
        header_paragraph.text = "LAUDO DE PERÍCIA CRIMINAL"
        header_paragraph.style = doc.styles['TituloPrincipal']
        header_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

        footer = section.footer
        footer_paragraph = footer.paragraphs[0]
        footer_paragraph.text = f"Página {doc.sections.index(section) + 1}"
        footer_paragraph.style = doc.styles['TituloPrincipal']
        footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def adicionar_material_recebido(doc, dados_laudo):
    """Adiciona a seção '2 MATERIAL RECEBIDO PARA EXAME' ao laudo docx."""

    adicionar_paragrafo(doc, "2 MATERIAL RECEBIDO PARA EXAME (Ilustração 1)", style='TituloPrincipal')

    tipos_material_itens_codigo = []
    subitens_cannabis = {}
    subitens_cocaina = {}

    for i, item in enumerate(dados_laudo['itens']):
        qtd_ext = obter_quantidade_extenso(item['qtd'])
        tipo_material = TIPOS_MATERIAL_BASE.get(item['tipo_mat'], item['tipo_mat'])
        embalagem = TIPOS_EMBALAGEM_BASE.get(item['emb'], item['emb'])

        if item['cor_emb']:
            cor = CORES_FEMININO_EMBALAGEM.get(item['cor_emb'], item['cor_emb'])
            embalagem += f" de cor {cor}"

        embalagem = pluralizar_palavra(embalagem, item['qtd'])
        porcao = pluralizar_palavra("porção", item['qtd'])
        acond = "acondicionada em" if item['qtd'] == 1 else "acondicionadas, individualmente, em"
        ref_texto = f", relacionada a {item['pessoa']}" if item['pessoa'] else ""
        final_ponto = "."

        texto = f"2.{i + 1} {item['qtd']} ({qtd_ext}) {porcao} de material {tipo_material}, {acond} {embalagem}, referente à amostra do subitem {item['ref']} do laudo de constatação supracitado{ref_texto}{final_ponto}"
        adicionar_paragrafo(doc, texto, style='ParagrafoNormal')

        tipos_material_itens_codigo.append(item['tipo_mat'])
        if item['tipo_mat'] in ["v", "r"]:
            subitens_cannabis[item['ref']] = f"2.{i + 1}"
        elif item['tipo_mat'] in ["po", "pd"]:
            subitens_cocaina[item['ref']] = f"2.{i + 1}"

    return subitens_cannabis, subitens_cocaina

def adicionar_objetivo_exames(doc):
    """Adiciona a seção '3 OBJETIVO DOS EXAMES'."""

    adicionar_paragrafo(doc, "\n3 OBJETIVO DOS EXAMES", style='TituloPrincipal')
    adicionar_paragrafo(doc, "Visa esclarecer à autoridade requisitante quanto às características do material apresentado, bem como se ele contém substância de uso proscrito no Brasil e capaz de causar dependência física e/ou psíquica. O presente laudo pericial busca demonstrar a materialidade da infração penal apurada.", align='justify', style='ParagrafoNormal')

def adicionar_exames(doc, subitens_cannabis, subitens_cocaina):
    """Adiciona a seção '4 EXAMES'."""

    adicionar_paragrafo(doc, "\n4 EXAMES", style='TituloPrincipal')
    has_cannabis_item = bool(subitens_cannabis)
    has_cocaina_item = bool(subitens_cocaina)

    if has_cannabis_item:
        adicionar_paragrafo(doc, "4.1 Exames realizados para pesquisa de Cannabis sativa L.", style='TituloSecundario')
        adicionar_paragrafo(doc, "4.1.1 Ensaio químico com Fast blue salt B: teste de cor em reação com solução aquosa de sal de azul sólido B em meio alcalino;", style='ParagrafoNormal')
        adicionar_paragrafo(doc, "4.1.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas contendo eluentes apropriados e posterior revelação com solução aquosa de azul sólido B.", style='ParagrafoNormal')
    if has_cocaina_item:
        idx = "4.2" if has_cannabis_item else "4.1"
        adicionar_paragrafo(doc, f"{idx} Exames realizados para pesquisa de cocaína", style='TituloSecundario')
        adicionar_paragrafo(doc, f"{idx}.1 Ensaio químico com teste de tiocianato de cobalto-reação de cor com solução de tiocianato de cobalto em meio ácido;", style='ParagrafoNormal')
        adicionar_paragrafo(doc, f"{idx}.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas com eluentes apropriados e revelação com solução de iodo platinado.", style='ParagrafoNormal')
    if not has_cannabis_item and not has_cocaina_item:
        adicionar_paragrafo(doc, "4.1 Exames realizados", style='TituloSecundario')
        adicionar_paragrafo(doc, "4.1.1 Exame macroscópico;", style='ParagrafoNormal')

def adicionar_resultados(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    """Adiciona a seção '5 RESULTADOS'."""

    adicionar_paragrafo(doc, "\n5 RESULTADOS", style='TituloPrincipal')
    if any(item['tipo_mat'] in ["v", "r"] for item in dados_laudo['itens']):
        subitens_cannabis_str = ", ".join([item['ref'] for item in dados_laudo['itens'] if item['tipo_mat'] in ["v", "r"]])
        label = "no subitem" if len(subitens_cannabis_str.split(", ")) == 1 else "nos subitens"
        adicionar_paragrafo(doc, f"5.1 Resultados obtidos para o(s) material(is) descrito(s) {label} {subitens_cannabis_str}:", style='TituloSecundario')
        adicionar_paragrafo(doc, "5.1.1 No ensaio com Fast blue salt B, foram obtidas coloração característica para canabinol e tetrahidrocanabinol (princípios ativos da Cannabis sativa L.).", style='ParagrafoNormal')
        adicionar_paragrafo(doc, "5.1.2 Na CCD, obtiveram-se perfis cromatográficos coincidentes com o material de referência (padrão de Cannabis sativa L.); portanto, a substância tetrahidrocanabinol está presente nos materiais questionados.", style='ParagrafoNormal')
    if any(item['tipo_mat'] in ["po", "pd"] for item in dados_laudo['itens']):
        subitens_cocaina_str = ", ".join([item['ref'] for item in dados_laudo['itens'] if item['tipo_mat'] in ["po", "pd"]])
        label = "no subitem" if len(subitens_cocaina_str.split(", ")) == 1 else "nos subitens"
        idx = "5.2" if any(item['tipo_mat'] in ["v", "r"] for item in dados_laudo['itens']) else "5.1"
        adicionar_paragrafo(doc, f"\n{idx} Resultados obtidos para o(s) material(is) descrito(s) {label} {subitens_cocaina_str}:", style='TituloSecundario')
        adicionar_paragrafo(doc, f"{idx}.1 No teste de tiocianato de cobalto, foram obtidas coloração característica para cocaína;", style='ParagrafoNormal')
        adicionar_paragrafo(doc, f"{idx}.2 Na CCD, obteve-se perfis cromatográficos coincidentes com o material de referência (padrão de cocaína); portanto, a substância cocaína está presente nos materiais questionados.", style='ParagrafoNormal')
    if not any(item['tipo_mat'] in ["v", "r"] for item in dados_laudo['itens']) and not any(item['tipo_mat'] in ["po", "pd"] for item in dados_laudo['itens']):
        adicionar_paragrafo(doc, "5.1 Resultados obtidos", style='TituloSecundario')
        adicionar_paragrafo(doc, "5.1.1 Exame macroscópico", style='ParagrafoNormal')

def adicionar_conclusao(doc, dados_laudo):
    """Adiciona a seção '6 CONCLUSÃO'."""

    adicionar_paragrafo(doc, "\n6 CONCLUSÃO", style='TituloPrincipal')
    conclusoes = []
    if any(item['tipo_mat'] in ["v", "r"] for item in dados_laudo['itens']):
        subitens_cannabis_str = ", ".join([item['ref'] for item in dados_laudo['itens'] if item['tipo_mat'] in ["v", "r"]])
        label = "no subitem" if len(subitens_cannabis_str.split(", ")) == 1 else "nos subitens"
        conclusoes.append(f"no(s) material(is) descrito(s) {label} {subitens_cannabis_str}, foi detectada a presença de partes da planta Cannabis sativa L., vulgarmente conhecida por maconha. A Cannabis sativa L. contém princípios ativos chamados canabinóis, dentre os quais se encontra o tetrahidrocanabinol, substância perturbadora do sistema nervoso central. Tanto a Cannabis sativa L. quanto a tetrahidrocanabinol são proscritas no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.")
    if any(item['tipo_mat'] in ["po", "pd"] for item in dados_laudo['itens']):
        subitens_cocaina_str = ", ".join([item['ref'] for item in dados_laudo['itens'] if item['tipo_mat'] in ["po", "pd"]])
        conclusoes.append(f"no(s) material(is) descrito(s) no(s) subitem(ns) {subitens_cocaina_str}, foi detectada a presença de cocaína, substância alcaloide estimulante do sistema nervoso central. A cocaína é proscrita no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.")

    if conclusoes:
        texto_final = "A partir das análises realizadas, conclui-se que, " + " Outrossim, ".join(conclusoes)
    else:
        texto_final = "A partir das análises realizadas, conclui-se que não foram detectadas substâncias de uso proscrito nos materiais analisados."
    adicionar_paragrafo(doc, texto_final, align='justify', style='ParagrafoNormal')

def adicionar_custodia_material(doc, lacre):
    """Adiciona a seção '7 CUSTÓDIA DO MATERIAL'."""

    adicionar_paragrafo(doc, "\n7 CUSTÓDIA DO MATERIAL", style='TituloPrincipal')
    adicionar_paragrafo(doc, "7.1 Contraprova", style='TituloSecundario')
    adicionar_paragrafo(doc, f"7.1.1 A amostra contraprova ficará armazenada neste Instituto, conforme Portaria 0003/2019/SSP  (Lacre nº {lacre}).", style='ParagrafoNormal')

def adicionar_referencias(doc, subitens_cannabis, subitens_cocaina):
    """Adiciona a seção 'REFERÊNCIAS'."""

    adicionar_paragrafo(doc, "\nREFERÊNCIAS", style='TituloPrincipal')
    referencias = [
        "BRASIL. Ministério da Saúde. Portaria SVS/MS n° 344, de 12 de maio de 1998.  Aprova o regulamento técnico sobre substâncias e medicamentos sujeitos a controle especial .",
        "Diário Oficial da União: Brasília, DF, p. 37, 19 maio 1998. Alterada pela RDC nº 970 de 19/03/2025 da Anvisa.",
        "GOIÁS. Secretaria de Estado da Segurança Pública. Portaria nº 0003/2019/SSP de 10 de janeiro de 2019.  Regulamenta a apreensão, movimentação, exames, acondicionamento, armazenamento e destruição de drogas no âmbito da Secretaria de Estado da Segurança Pública.",
        "Diário Oficial do Estado de Goiás: n° 22.972, Goiânia, GO, p. 4-5, 15 jan. 2019.",
        "SWGDRUG: Scientific Working Group for the Analysis of Seized Drugs.  Recommendations . Version 8.0 june. 2019.",
        "Disponível em:  http://www.swgdrug.org/Documents/SWGDRUG%20Recommendations%20Version%20 8_FINAL_ForPosting_092919.pdf. Acesso em: 07/10/2019."
    ]
    if subitens_cannabis:
        referencias.append("UNODC (United Nations Office on Drugs and Crime). Laboratory and scientific section.  Recommended Methods for the Identification and Analysis of Cannabis and Cannabis Products . New York: 2012.")
    if subitens_cocaina:
        referencias.append("UNODC (United Nations Office on Drugs and Crime). Laboratory and Scientific Section. Recommended Methods for the Identification and Analysis of Cocaine in Seized Materials. New York: 2012.")
    for ref in referencias:
        adicionar_paragrafo(doc, ref, style='ParagrafoNormal')

def adicionar_data_assinatura(doc):
    """Adiciona a data e a assinatura do perito."""

    brasilia_tz = pytz.timezone('America/Sao_Paulo')
    hoje = datetime.now(brasilia_tz)
    data_formatada = f"Goiânia, {hoje.day} de {meses_portugues[hoje.strftime('%B')]} de {hoje.year}."
    adicionar_paragrafo(doc, data_formatada, align='right', style='ParagrafoNormal')

    adicionar_paragrafo(doc, "\nLaudo assinado digitalmente com dados do assinador à esquerda das páginas", align='left', style='ParagrafoNormal')
    adicionar_paragrafo(doc, "Daniel Chendes Lima", align='center', style='ParagrafoNormal')
    adicionar_paragrafo(doc, "Perito Criminal", align='center', style='ParagrafoNormal')

def aplicar_italico(doc):
    """Aplica estilo itálico a palavras e frases específicas no documento."""

    italics = [
        'Cannabis sativa',
        'Scientific Working Group for the Analysis of Seized Drugs',
        'United Nations Office on Drugs and Crime',
        'Fast blue salt B',
        'eppendorf',
        'ziplock'
    ]

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for phrase in italics:
                if phrase in run.text:
                    inline = paragraph.runs
                    # Evita aplicar itálico a partes indesejadas do texto
                    for i in range(len(inline)):
                        if phrase in inline[i].text:
                           inline[i].style = 'Italico'

def gerar_laudo_docx(dados_laudo, image_file):
    """Gera o laudo completo em formato docx."""

    document = Document()
    cor_fundo = configurar_estilos(document)
    aplicar_fundo(document, cor_fundo)
    adicionar_cabecalho_rodape(document)

    subitens_cannabis, subitens_cocaina = adicionar_material_recebido(document, dados_laudo)

    if image_file:
        inserir_imagem_docx(document, image_file)

    adicionar_objetivo_exames(document)
    adicionar_exames(document, subitens_cannabis, subitens_cocaina)
    adicionar_resultados(document, subitens_cannabis, subitens_cocaina, dados_laudo)
    adicionar_conclusao(document, dados_laudo)
    adicionar_custodia_material(document, dados_laudo['lacre'])
    adicionar_referencias(document, subitens_cannabis, subitens_cocaina)
    adicionar_data_assinatura(document)
    aplicar_italico(document)

    return document

# --- Interface Streamlit ---
def main():
    st.set_page_config(layout="wide")  # Use the full width of the page

    # --- Título com Logo, Data e Hora ---
    col1, col2, col3 = st.columns([1, 3, 1])  # Adjust column widths as needed

    with col1:
        st.image("https://www.policiacientifica.go.gov.br/wp-content/uploads/2021/08/logomarca-branca-menor.png", width=150)  # Adjust width as needed

    with col2:
        st.title("Gerador de Laudo Pericial")

    with col3:
        data_e_hora = col3.empty()  # Create an empty container for dynamic content

    # Atualizar a data e hora dinamicamente
    while True:
        now = datetime.now()
        data_formatada = now.strftime("%A, %d de %B de %Y %H:%M:%S")  # Formato: Dia da semana, dia de mês de Mês de Ano Hora:Minuto:Segundo
        data_e_hora.markdown(f"<div style='text-align: right;'>{data_formatada}</div>", unsafe_allow_html=True)
        time.sleep(1)  # Atualiza a cada segundo

    # --- Resto da Aplicação ---

    dados_laudo = {}
    dados_laudo['itens'] = []

    numero_itens = st.number_input("Número de itens a descrever", min_value=1, value=1, step=1)
    dados_laudo['lacre'] = st.text_input("Número do lacre da contraprova")

    for i in range(numero_itens):
        st.subheader(f"Item {i + 1}")
        item = {}
        item['qtd'] = st.number_input(f"Quantidade de porções do item {i + 1}", min_value=1, value=1, step=1)
        item['tipo_mat'] = st.selectbox(f"Tipo de material do item {i + 1}", options=TIPOS_MATERIAL_BASE.keys(), index=0)
        item['emb'] = st.selectbox(f"Tipo de embalagem do item {i + 1}", options=TIPOS_EMBALAGEM_BASE.keys(), index=0)
        if item['emb'] in ['pl', 'pa']:
            item['cor_emb'] = st.selectbox(f"Cor da embalagem do item {i + 1}", options=CORES_FEMININO_EMBALAGEM.keys(), index=0)
        else:
            item['cor_emb'] = None
        item['ref'] = st.text_input(f"Referência do subitem do item {i + 1}")
        item['pessoa'] = st.text_input(f"Pessoa relacionada ao item {i + 1} (opcional)")
        dados_laudo['itens'].append(item)

    image_file = st.file_uploader("Adicionar Imagem (opcional)", type=["png", "jpg", "jpeg"])

    if st.button("Gerar Laudo"):
        doc = gerar_laudo_docx(dados_laudo, image_file)
        # Salvar o documento em um BytesIO para download
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.subheader("Download do Laudo")
        nome_arquivo = st.text_input("Nome do arquivo para download", value="laudo.docx")
        st.download_button(
            label="Baixar Laudo (DOCX)",
            data=buffer,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success("Laudo gerado e pronto para download!")

if __name__ == "__main__":
    main()
