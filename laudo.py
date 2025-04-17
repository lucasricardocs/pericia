import re
from datetime import datetime
import io
import pytz
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from PIL import Image
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import traceback

# --- Constantes --- (MANTIDO)
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
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}
dias_semana_portugues = {
    0: "Segunda-feira", 1: "Terça-feira", 2: "Quarta-feira", 3: "Quinta-feira",
    4: "Sexta-feira", 5: "Sábado", 6: "Domingo"
}
DOCX_COR_AZUL_SPTC = RGBColor(0, 71, 143)
DOCX_COR_CINZA_SPTC = RGBColor(110, 110, 110)
DOCX_COR_PRETO = RGBColor(0, 0, 0)
TERMOS_ITALICO_ORIGINAL = [
    'Cannabis sativa L.',
    'Cannabis sativa',
    'Scientific Working Group for the Analysis of Seized Drugs',
    'United Nations Office on Drugs and Crime',
    'Fast blue salt B',
    'eppendorf',
    'ziplock',
    'Tetrahidrocanabinol',
    'Portaria nº 344/1998',
    'RDC nº 970, de 19/03/2025'
]

# --- Funções Auxiliares --- (MANTIDO)
def pluralizar_palavra(palavra, quantidade):
    if quantidade == 1:
        return palavra
    if palavra in ["microtubo do tipo eppendorf", "embalagem do tipo ziplock", "papel alumínio"]:
        return palavra
    if palavra.endswith('m') and palavra not in ["alumínio"]:
        return re.sub(r'm$', 'ns', palavra)
    if palavra.endswith('ão'):
        return re.sub(r'ão$', 'ões', palavra)
    elif palavra.endswith(('r', 'z', 's')):
        if palavra.endswith(('r', 'z')):
            return palavra + 'es'
        else:
            return palavra
    elif palavra.endswith('l'):
        return palavra[:-1] + 'is'
    else:
        return palavra + 's'

def obter_quantidade_extenso(qtd):
    return QUANTIDADES_EXTENSO.get(qtd, str(qtd))

def adicionar_paragrafo(doc, text, style=None, align=None, color=None, size=None, bold=False, italic=False):
    p = doc.add_paragraph()
    if style and style in doc.styles:
        try:
            p.style = doc.styles[style]
        except Exception as e:
            print(f"Erro ao aplicar estilo '{style}': {e}. Usando 'Normal'.")
            p.style = doc.styles['Normal']
    elif style:
        print(f"Estilo '{style}' não encontrado. Usando 'Normal'.")
        p.style = doc.styles['Normal']
    if align:
        align_map = {
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY, 'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT, 'left': WD_ALIGN_PARAGRAPH.LEFT
        }
        p.alignment = align_map.get(str(align).lower(), WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run(text)
    if color:
        try:
            if isinstance(color, RGBColor):
                run.font.color.rgb = color
            elif isinstance(color, (tuple, list)) and len(color) == 3:
                run.font.color.rgb = RGBColor(color[0], color[1], color[2])
            else:
                print(f"Formato de cor inválido: {color}")
        except Exception as e:
            print(f"Erro ao aplicar cor: {e}")
    if size:
        try:
            run.font.size = Pt(int(size))
        except ValueError:
            print(f"Tamanho de fonte inválido: {size}")
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True

def inserir_imagem_docx(doc, image_file_uploader):
    try:
        if image_file_uploader:
            img_stream = io.BytesIO(image_file_uploader.getvalue())
            img = Image.open(img_stream)
            width_px, height_px = img.size
            max_width_inches = 6.0
            dpi = img.info.get('dpi', (96, 96))[0]
            if dpi <= 0:
                dpi = 96
            width_inches = width_px / dpi
            if width_inches > max_width_inches:
                display_width_inches = max_width_inches
            else:
                display_width_inches = width_inches
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            img_stream.seek(0)
            run.add_picture(img_stream, width=Inches(display_width_inches))
    except Exception as e:
        st.error(f"Erro ao inserir imagem no docx: {e}")
        print(f"Erro detalhado ao inserir imagem: {e}\n{traceback.format_exc()}")

# --- Funções de Estrutura do Documento DOCX --- (MANTIDO)
def configurar_estilos(doc):
    FONTE_PADRAO = 'Gadugi'
    COR_TEXTO_PRINCIPAL = DOCX_COR_PRETO
    COR_DESTAQUE = DOCX_COR_AZUL_SPTC
    COR_TEXTO_SECUNDARIO = DOCX_COR_CINZA_SPTC

    def get_or_add_style(doc, style_name, style_type):
        if style_name in doc.styles:
            return doc.styles[style_name]
        else:
            try:
                return doc.styles.add_style(style_name, style_type)
            except Exception as e:
                print(f"Falha ao adicionar estilo '{style_name}': {e}. Usando 'Normal' como fallback.")
                return doc.styles['Normal']

    paragrafo_style = doc.styles['Normal']
    paragrafo_style.font.name = FONTE_PADRAO
    paragrafo_style.font.size = Pt(12)
    paragrafo_style.font.color.rgb = COR_TEXTO_PRINCIPAL
    paragrafo_style.paragraph_format.line_spacing = 1.15
    paragrafo_style.paragraph_format.space_before = Pt(0)
    paragrafo_style.paragraph_format.space_after = Pt(8)

    titulo_principal_style = get_or_add_style(doc, 'TituloPrincipal', WD_STYLE_TYPE.PARAGRAPH)
    titulo_principal_style.base_style = doc.styles['Normal']
    titulo_principal_style.font.name = FONTE_PADRAO
    titulo_principal_style.font.size = Pt(14)
    titulo_principal_style.font.bold = True
    titulo_principal_style.font.color.rgb = COR_DESTAQUE
    titulo_principal_style.paragraph_format.space_before = Pt(12)
    titulo_principal_style.paragraph_format.space_after = Pt(6)

    titulo_secundario_style = get_or_add_style(doc, 'TituloSecundario', WD_STYLE_TYPE.PARAGRAPH)
    titulo_secundario_style.base_style = doc.styles['Normal']
    titulo_secundario_style.font.name = FONTE_PADRAO
    titulo_secundario_style.font.size = Pt(12)
    titulo_secundario_style.font.bold = True
    titulo_secundario_style.font.color.rgb = COR_DESTAQUE
    titulo_secundario_style.paragraph_format.space_before = Pt(10)
    titulo_secundario_style.paragraph_format.space_after = Pt(4)

    ilustracao_style = get_or_add_style(doc, 'Ilustracao', WD_STYLE_TYPE.PARAGRAPH)
    ilustracao_style.base_style = doc.styles['Normal']
    ilustracao_style.font.name = FONTE_PADRAO
    ilustracao_style.font.size = Pt(10)
    ilustracao_style.font.italic = True
    ilustracao_style.font.color.rgb = COR_TEXTO_SECUNDARIO
    ilustracao_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ilustracao_style.paragraph_format.space_before = Pt(4)
    ilustracao_style.paragraph_format.space_after = Pt(10)

def configurar_pagina(doc):
    for section in doc.sections:
        section.page_height = Inches(11.69)
        section.page_width = Inches(8.27)
        section.top_margin = Inches(1.18)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.79)

def adicionar_cabecalho_rodape(doc):
    FONTE_CABECALHO_RODAPE = 'Gadugi'
    TAMANHO_CABECALHO_RODAPE = Pt(10)
    section = doc.sections[0]

    header = section.header
    if header.paragraphs:
        for para in header.paragraphs:
            p_element = para._element
            if p_element.getparent() is not None:
                p_element.getparent().remove(p_element)

    header_paragraph = header.add_paragraph()
    run_header_left = header_paragraph.add_run("POLÍCIA CIENTÍFICA DE GOIÁS")
    run_header_left.font.name = FONTE_CABECALHO_RODAPE
    run_header_left.font.size = TAMANHO_CABECALHO_RODAPE
    run_header_left.font.bold = True
    header_paragraph.add_run("\t\t")
    run_header_right = header_paragraph.add_run("LAUDO DE PERÍCIA CRIMINAL")
    run_header_right.font.name = FONTE_CABECALHO_RODAPE
    run_header_right.font.size = TAMANHO_CABECALHO_RODAPE
    run_header_right.font.bold = False
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    footer = section.footer
    if footer.paragraphs:
        for para in footer.paragraphs:
            p_element = para._element
            if p_element.getparent() is not None:
                p_element.getparent().remove(p_element)
    page_num_paragraph = footer.add_paragraph()
    page_num_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_page = page_num_paragraph.add_run("Página ")
    run_page.font.name = FONTE_CABECALHO_RODAPE
    run_page.font.size = TAMANHO_CABECALHO_RODAPE
    fld_char_begin = OxmlElement('w:fldChar')
    fld_char_begin.set(qn('w:fldCharType'), 'begin')
    run_page._r.append(fld_char_begin)
    instr_text = OxmlElement('w:instrText')
    instr_text.set(qn('xml:space'), 'preserve')
    instr_text.text = 'PAGE \* MERGEFORMAT'
    run_page._r.append(instr_text)
    fld_char_sep = OxmlElement('w:fldChar')
    fld_char_sep.set(qn('w:fldCharType'), 'separate')
    run_page._r.append(fld_char_sep)
    fld_char_end = OxmlElement('w:fldChar')
    fld_char_end.set(qn('w:fldCharType'), 'end')
    run_page._r.append(fld_char_end)

    run_num_pages = page_num_paragraph.add_run(" de ")
    run_num_pages.font.name = FONTE_CABECALHO_RODAPE
    run_num_pages.font.size = TAMANHO_CABECALHO_RODAPE
    fld_char_begin_np = OxmlElement('w:fldChar')
    fld_char_begin_np.set(qn('w:fldCharType'), 'begin')
    run_num_pages._r.append(fld_char_begin_np)
    instr_text_np = OxmlElement('w:instrText')
    instr_text_np.set(qn('xml:space'), 'preserve')
    instr_text_np.text = 'NUMPAGES \* MERGEFORMAT'
    run_num_pages._r.append(instr_text_np)
    fld_char_sep_np = OxmlElement('w:fldChar')
    fld_char_sep_np.set(qn('w:fldCharType'), 'separate')
    run_num_pages._r.append(fld_char_sep_np)
    fld_char_end_np = OxmlElement('w:fldChar')
    fld_char_end_np.set(qn('w:fldCharType'), 'end')
    run_num_pages._r.append(fld_char_end_np)

# --- Funções das Seções do Laudo --- (MANTIDO)
def adicionar_material_recebido(doc, dados_laudo):
    adicionar_paragrafo(doc, "1 MATERIAL RECEBIDO PARA EXAME", style='TituloPrincipal')

    imagem_carregada = dados_laudo.get('imagem')
    if imagem_carregada:
        inserir_imagem_docx(doc, imagem_carregada)
        adicionar_paragrafo(doc, "Ilustração 1: Material(is) recebido(s).", style='Ilustracao')

    subitens_cannabis = {}
    subitens_cocaina = {}

    if not dados_laudo.get('itens'):
        adicionar_paragrafo(doc, "Nenhum item de material foi descrito para exame.", style='Normal')
        return subitens_cannabis, subitens_cocaina

    for i, item in enumerate(dados_laudo['itens']):
        qtd = item.get('qtd', 1)
        qtd_ext = obter_quantidade_extenso(qtd)
        tipo_mat_cod = item.get('tipo_mat', '')
        tipo_material = TIPOS_MATERIAL_BASE.get(tipo_mat_cod, f"tipo '{tipo_mat_cod}'")
        emb_cod = item.get('emb', '')
        embalagem = TIPOS_EMBALAGEM_BASE.get(emb_cod, f"embalagem '{emb_cod}'")
        cor_emb_cod = item.get('cor_emb')
        desc_cor = ""
        if cor_emb_cod and emb_cod in ['pl', 'pa', 'e', 'z']:
            cor = CORES_FEMININO_EMBALAGEM.get(cor_emb_cod, cor_emb_cod)
            desc_cor = f" de cor {cor}"

        embalagem_base_plural = pluralizar_palavra(embalagem, qtd)
        embalagem_final = f"{embalagem_base_plural}{desc_cor}"
        porcao = pluralizar_palavra("porção", qtd)
        acond = "acondicionada em" if qtd == 1 else "acondicionadas, individualmente, em"
        ref_texto = f", relacionada a {item['pessoa']}" if item.get('pessoa') else ""
        subitem_ref = item.get('ref', '')
        subitem_texto = f", referente à amostra do subitem {subitem_ref} do laudo de constatação supracitado" if subitem_ref else ""
        item_num_str = f"1.{i + 1}"
        final_ponto = "."
        texto = (f"{item_num_str} {qtd} ({qtd_ext}) {porcao} de material {tipo_material}, {acond} {embalagem_final}{subitem_texto}{ref_texto}{final_ponto}")
        adicionar_paragrafo(doc, texto, style='Normal', align='justify')

        chave_mapeamento = subitem_ref if subitem_ref else f"Item_{item_num_str}"
        item_num_referencia = item_num_str
        if tipo_mat_cod in ["v", "r"]:
            subitens_cannabis[chave_mapeamento] = item_num_referencia
        elif tipo_mat_cod in ["po", "pd"]:
            subitens_cocaina[chave_mapeamento] = item_num_referencia

    return subitens_cannabis, subitens_cocaina

def adicionar_objetivo_exames(doc):
    adicionar_paragrafo(doc, "2 OBJETIVO DOS EXAMES", style='TituloPrincipal')
    texto = (
        "Visa esclarecer à autoridade requisitante quanto às características do material apresentado, "
        "bem como se ele contém substância de uso proscrito no Brasil e capaz de causar dependência física e/ou psíquica. "
        "O presente laudo pericial busca demonstrar a materialidade da infração penal apurada.")
    adicionar_paragrafo(doc, texto, align='justify', style='Normal')

def adicionar_exames(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    adicionar_paragrafo(doc, "3 EXAMES", style='TituloPrincipal')

    has_cannabis_item = bool(subitens_cannabis)
    has_cocaina_item = bool(subitens_cocaina)

    idx_subitem = 1
    if has_cannabis_item:
        adicionar_paragrafo(doc, f"3.{idx_subitem} Exames realizados para pesquisa de Cannabis sativa L. (maconha)",
                            style='TituloSecundario')
        adicionar_paragrafo(doc,
                            f"3.{idx_subitem}.1 Ensaio químico com Fast blue salt B: teste de cor em reação com solução aquosa de sal de azul sólido B em meio alcalino;",
                            style='Normal', align='justify')
        adicionar_paragrafo(doc,
                            f"3.{idx_subitem}.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas contendo eluentes apropriados e posterior revelação com solução aquosa de azul sólido B.",
                            style='Normal', align='justify')
        idx_subitem += 1

    if has_cocaina_item:
        adicionar_paragrafo(doc, f"3.{idx_subitem} Exames realizados para pesquisa de cocaína",
                            style='TituloSecundario')
        adicionar_paragrafo(doc,
                            f"3.{idx_subitem}.1 Ensaio químico com teste de tiocianato de cobalto-reação de cor com solução de tiocianato de cobalto em meio ácido;",
                            style='Normal', align='justify')
        adicionar_paragrafo(doc,
                            f"3.{idx_subitem}.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas com eluentes apropriados e revelação com solução de iodo platinado.",
                            style='Normal', align='justify')
        idx_subitem += 1

    if not has_cannabis_item and not has_cocaina_item and dados_laudo.get('itens'):
        adicionar_paragrafo(doc, f"3.{idx_subitem} Exames realizados", style='TituloSecundario')
        adicionar_paragrafo(doc, f"3.{idx_subitem}.1 Exame macroscópico;", style='Normal', align='justify')
        idx_subitem += 1

    if idx_subitem == 1:
        adicionar_paragrafo(doc, "Nenhum exame específico a relatar com base nos materiais descritos.",
                            style='Normal')

def adicionar_resultados(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    adicionar_paragrafo(doc, "4 RESULTADOS", style='TituloPrincipal')

    has_cannabis_item = bool(subitens_cannabis)
    has_cocaina_item = bool(subitens_cocaina)
    idx_subitem = 1

    if has_cannabis_item:
        itens_referencia = sorted(list(subitens_cannabis.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}"
        adicionar_paragrafo(doc, f"4.{idx_subitem} Resultados obtidos para o(s) material(is) descrito(s) {label}:",
                            style='TituloSecundario')
        adicionar_paragrafo(doc,
                            f"4.{idx_subitem}.1 No ensaio com Fast blue salt B, foram obtidas coloração característica para canabinol e tetrahidrocanabinol (princípios ativos da Cannabis sativa L.).",
                            style='Normal', align='justify')
        adicionar_paragrafo(doc,
                            f"4.{idx_subitem}.2 Na CCD, obtiveram-se perfis cromatográficos coincidentes com o material de referência (padrão de Cannabis sativa L.); portanto, a substância tetrahidrocanabinol está presente nos materiais questionados.",
                            style='Normal', align='justify')
        idx_subitem += 1

    if has_cocaina_item:
        itens_referencia = sorted(list(subitens_cocaina.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}"
        adicionar_paragrafo(doc, f"4.{idx_subitem} Resultados obtidos para o(s) material(is) descrito(s) {label}:",
                            style='TituloSecundario')
        adicionar_paragrafo(doc, f"4.{idx_subitem}.1 No teste de tiocianato de cobalto, foram obtidas coloração característica para cocaína;",
                            style='Normal', align='justify')
        adicionar_paragrafo(doc,
                            f"4.{idx_subitem}.2 Na CCD, obteve-se perfis cromatográficos coincidentes com o material de referência (padrão de cocaína); portanto, a substância cocaína está presente nos materiais questionados.",
                            style='Normal', align='justify')
        idx_subitem += 1

    if idx_subitem == 1:
        if dados_laudo.get('itens'):
            adicionar_paragrafo(doc,
                                "Não foram obtidos resultados positivos para Cannabis ou Cocaína nos testes realizados para os materiais descritos.",
                                style='Normal', align='justify')
        else:
            adicionar_paragrafo(doc, "Nenhum material foi submetido a exame, portanto, não há resultados a relatar.",
                                style='Normal', align='justify')

def adicionar_conclusao(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    adicionar_paragrafo(doc, "5 CONCLUSÃO", style='TituloPrincipal')

    conclusoes = []
    if subitens_cannabis:
        itens_referencia = sorted(list(subitens_cannabis.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}"
        conclusoes.append(
            f"no(s) material(is) descrito(s) {label}, foi detectada a presença de partes "
            f"da planta Cannabis sativa L., vulgarmente conhecida por maconha. "
            f"A Cannabis sativa L. contém princípios ativos chamados canabinóis, dentre os quais se encontra o tetrahidrocanabinol, substância perturbadora do sistema nervoso central. "
            f"Tanto a Cannabis sativa L. quanto a tetrahidrocanabinol são proscritas no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.")
    if subitens_cocaina:
        itens_referencia = sorted(list(subitens_cocaina.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}"
        conclusoes.append(
            f"no(s) material(is) descrito(s) {label}, foi detectada a presença de cocaína, substância alcaloide estimulante do sistema nervoso central. A cocaína é proscrita no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.")
    if conclusoes:
        texto_final = "A partir das análises realizadas, conclui-se que, " + " Outrossim, ".join(conclusoes)
    elif dados_laudo.get('itens'):
        texto_final = "A partir das análises realizadas, conclui-se que não foram detectadas substâncias de uso proscrito nos materiais analisados."
    else:
        texto_final = "Não houve material submetido a exame, portanto, não há conclusões a apresentar."
    adicionar_paragrafo(doc, texto_final, align='justify', style='Normal')

def adicionar_custodia_material(doc, dados_laudo):
    adicionar_paragrafo(doc, "6 CUSTÓDIA DO MATERIAL", style='TituloPrincipal')
    adicionar_paragrafo(doc, "6.1 Contraprova", style='TituloSecundario')

    lacre = dados_laudo.get('lacre', '_______')
    texto_contraprova = (
        f"A amostra contraprova ficará armazenada neste Instituto, conforme Portaria 0003/2019/SSP "
        f"(Lacre nº {lacre}).")
    adicionar_paragrafo(doc, texto_contraprova, style='Normal', align='justify')

def adicionar_referencias(doc, subitens_cannabis, subitens_cocaina):
    adicionar_paragrafo(doc, "REFERÊNCIAS", style='TituloPrincipal')
    tamanho_ref = 10
    referencias_base = [
        "BRASIL. Ministério da Saúde. Portaria SVS/MS n° 344, de 12 de maio de 1998. Aprova o regulamento técnico sobre substâncias e medicamentos sujeitos a controle especial. Diário Oficial da União: Brasília, DF, p. 37, 19 maio 1998. Alterada pela RDC nº 970, de 19/03/2025.",
        "GOIÁS. Secretaria de Estado da Segurança Pública. Portaria nº 0003/2019/SSP de 10 de janeiro de 2019. Regulamenta a apreensão, movimentação, exames, acondicionamento, armazenamento e destruição de drogas no âmbito da Secretaria de Estado da Segurança Pública. Diário Oficial do Estado de Goiás: n° 22.972, Goiânia, GO, p. 4-5, 15 jan. 2019.",
        "SWGDRUG: Scientific Working Group for the Analysis of Seized Drugs. Recommendations. Version 8.0 june. 2019. Disponível em: http://www.swgdrug.org/Documents/SWGDRUG%20Recommendations%20Version%208_FINAL_ForPosting_092919.pdf. Acesso em: 07/10/2019."
    ]
    for ref in referencias_base:
        adicionar_paragrafo(doc, ref, style='Normal', align='justify', size=tamanho_ref)
    if subitens_cannabis:
        adicionar_paragrafo(doc,
                            "UNODC (United Nations Office on Drugs and Crime). Laboratory and scientific section. Recommended Methods for the Identification and Analysis of Cannabis and Cannabis Products. New York: 2012.",
                            style='Normal', align='justify', size=tamanho_ref)
    if subitens_cocaina:
        adicionar_paragrafo(doc,
                            "UNODC (United Nations Office on Drugs and Crime). Laboratory and Scientific Section. Recommended Methods for the Identification and Analysis of Cocaine in Seized Materials. New York: 2012.",
                            style='Normal', align='justify', size=tamanho_ref)

def adicionar_encerramento_assinatura(doc):
    try:
        brasilia_tz = pytz.timezone('America/Sao_Paulo')
        hoje = datetime.now(brasilia_tz)
    except Exception:
        hoje = datetime.now()
    mes_atual = meses_portugues.get(hoje.month, f"Mês {hoje.month}")
    data_formatada = f"Goiânia, {hoje.day} de {mes_atual} de {hoje.year}."
    doc.add_paragraph()
    adicionar_paragrafo(doc, data_formatada, align='right', style='Normal')
    doc.add_paragraph();
    doc.add_if full_text[idx:].startswith(phrase):
                # Verifica se é uma palavra completa (simplificado)
                ends_correctly = (idx + len(phrase) == len(full_text)) or (not full_text[idx + len(phrase)].isalnum())
                starts_correctly = (idx == 0) or (not full_text[idx-1].isalnum())
                if ends_correctly and starts_correctly:
                    run = paragraph.add_run(phrase)
                    run.font.name = 'Gadugi'
                    run.font.size = tamanho_fonte
                    run.italic = True # Aplica itálico
                    idx += len(phrase)
                    match_found = True
                    break  # Sai do loop de termos e continua varrendo o texto
            # Se nenhum termo em itálico foi encontrado começando em 'idx'
            if not match_found:
                run = paragraph.add_run(full_text[idx])
                run.font.name = 'Gadugi'
                run.font.size = tamanho_fonte
                run.italic = False # Garante que não seja itálico por padrão
                idx += 1
    # Se o parágrafo ficou vazio após o processo (pouco provável), restaura o texto original
    if not paragraph.text and full_text:
        paragraph.text = full_text

# --- Função Principal de Geração do DOCX ---

def gerar_laudo_docx(dados_laudo):
    """Gera o laudo completo em formato docx."""
    document = Document()

    configurar_estilos(document) # Configura estilos COM fonte Gadugi e cores SPTC
    configurar_pagina(document)
    adicionar_cabecalho_rodape(document)

    # Adiciona Seções na Ordem Correta usando as funções modificadas
    subitens_cannabis, subitens_cocaina = adicionar_material_recebido(document, dados_laudo)
    adicionar_objetivo_exames(document)
    adicionar_exames(document, subitens_cannabis, subitens_cocaina, dados_laudo)
    adicionar_resultados(document, subitens_cannabis, subitens_cocaina, dados_laudo)
    adicionar_conclusao(document, subitens_cannabis, subitens_cocaina, dados_laudo)
    adicionar_custodia_material(document, dados_laudo) # Passa dados_laudo para pegar o lacre
    adicionar_referencias(document, subitens_cannabis, subitens_cocaina)
    adicionar_encerramento_assinatura(document)

    # Aplica fonte Gadugi e itálico usando o método do código Colab original
    aplicar_italico_fonte_original(document)

    return document

# --- Interface Streamlit ---

def main():
    st.set_page_config(layout="centered", page_title="Gerador de Laudo")

    # --- Cores UI ---
    UI_COR_AZUL_SPTC = "#eaeff2"
    UI_COR_CINZA_SPTC = "#6E6E6E"

    # --- MOVIDO: Data/Calendário (Acima da logo/título) ---
    data_placeholder = st.empty()
    def atualizar_data():
        try:
            brasilia_tz = pytz.timezone('America/Sao_Paulo')
            now = datetime.now(brasilia_tz)
            dia_semana = dias_semana_portugues.get(now.weekday(), '')
            mes = meses_portugues.get(now.month, '')
            data_formatada = f"{dia_semana}, {now.day} de {mes} de {now.year}"
            # Adiciona um pouco de margem inferior para separar da linha seguinte
            data_placeholder.markdown(f"""
    <div style="text-align: right; font-size: 0.5em; color: {UI_COR_CINZA_SPTC}; line-height: 1.2; margin-bottom: 15px;">
        <span>{data_formatada}</span><br>
        <span style="font-size: 0.8em;">(Goiânia-GO)</span>
    </div>""", unsafe_allow_html=True)
        except Exception as e:
            now = datetime.now()
            fallback_str = now.strftime("%d/%m/%Y")
            data_placeholder.markdown(f"""
    <div style="text-align: right; font-size: 0.9em; color: #FF5555; line-height: 1.2; margin-bottom: 15px;">
        <span>{fallback_str} (Local)</span><br>
        <span style="font-size: 0.8em;">Erro Fuso Horário: {e}</span>
    </div>""", unsafe_allow_html=True)
    atualizar_data() # Chama a função para exibir a data

    # --- Cabeçalho com Logo e Título --- (Data foi movida para cima)
    # Ajuste as proporções se necessário, removendo a coluna da data
    col_logo, col_titulo = st.columns([1, 5]) # Ex: Proporção 1 para logo, 5 para título
    with col_logo: # Coluna da Logo
        logo_path = "logo_policia_cientifica.png"
        try:
            # Reduz a largura da imagem da logo
            st.image(logo_path, width=100) # <<-- LARGURA REDUZIDA AQUI (Ajuste 100, 110, 120...)
        except FileNotFoundError:
            st.error(f"Erro: Logo '{logo_path}' não encontrado.")
            st.info("Coloque 'logo_policia_cientifica.png' na mesma pasta do script.")
        except Exception as e:
            st.warning(f"Logo não carregado: {e}")
    with col_titulo: # Coluna do Título
        # Adicionado margin para tentar alinhar melhor com logo menor
        st.markdown(f'<h1 style="color: {UI_COR_AZUL_SPTC}; margin-top: 0px; margin-bottom: 0px;">Gerador de Laudo Pericial</h1>', unsafe_allow_html=True)
        st.markdown(f'<p style="color: {UI_COR_CINZA_SPTC}; font-size: 1em;">Identificação de Drogas - SPTC/GO</p>', unsafe_allow_html=True)
    st.markdown("---") # Separador visual

    # --- Inicialização do Estado da Sessão (Adicionado lacre e rg_pericia) ---
    if 'dados_laudo' not in st.session_state:
        st.session_state.dados_laudo = {
            'rg_pericia': '', # Adicionado
            'lacre': '', # Adicionado
            'itens': [],
            'imagem': None
        }
    # Garante que as chaves existem mesmo se o estado já foi inicializado antes
    if 'rg_pericia' not in st.session_state.dados_laudo:
        st.session_state.dados_laudo['rg_pericia'] = ''
    if 'lacre' not in st.session_state.dados_laudo:
        st.session_state.dados_laudo['lacre'] = ''
    if 'itens' not in st.session_state.dados_laudo:
        st.session_state.dados_laudo['itens'] = []
    if 'imagem' not in st.session_state.dados_laudo:
        st.session_state.dados_laudo['imagem'] = None

    # --- Campos de Entrada ---

    # RG da Perícia (Adicionado)
    st.session_state.dados_laudo['rg_pericia'] = st.text_input("RG da Perícia", st.session_state.dados_laudo.get('rg_pericia', ''))

    # Lacre (Adicionado)
    st.session_state.dados_laudo['lacre'] = st.text_input("Número do Lacre", st.session_state.dados_laudo.get('lacre', ''))

    # Upload de Imagem (Mantido)
    st.subheader("Imagem do Material")
    st.session_state.dados_laudo['imagem'] = st.file_uploader("Carregar imagem do material", type=["png", "jpg", "jpeg"])

    # Seção para adicionar itens de material (Mantido com ajustes)
    st.subheader("Itens de Material para Exame")

    # Botão para adicionar novo item
    if st.button("Adicionar Item de Material"):
        st.session_state.dados_laudo['itens'].append({}) # Adiciona um dicionário vazio

    # Loop para exibir os campos de cada item de material
    for i, item in enumerate(st.session_state.dados_laudo['itens']):
        st.markdown(f"**Item {i + 1}**")
        col1, col2, col3 = st.columns(3)
        with col1:
            item['qtd'] = st.number_input("Quantidade", value=item.get('qtd', 1), min_value=1, key=f"qtd_{i}")
        with col2:
            item['tipo_mat'] = st.selectbox("Tipo de Material", options=TIPOS_MATERIAL_BASE.keys(),
                                             index=list(TIPOS_MATERIAL_BASE.keys()).index(item.get('tipo_mat', 'v')),
                                             key=f"tipo_mat_{i}")
        with col3:
            item['emb'] = st.selectbox("Embalagem", options=TIPOS_EMBALAGEM_BASE.keys(),
                                          index=list(TIPOS_EMBALAGEM_BASE.keys()).index(item.get('emb', 'z')),
                                          key=f"emb_{i}")

        # Cor da Embalagem (Condicional, como antes)
        if item['emb'] in ['pl', 'pa', 'e', 'z']:
            item['cor_emb'] = st.selectbox("Cor da Embalagem", options=CORES_FEMININO_EMBALAGEM.keys(),
                                              index=list(CORES_FEMININO_EMBALAGEM.keys()).index(item.get('cor_emb', 't')),
                                              key=f"cor_emb_{i}")
        else:
            # Garante que 'cor_emb' existe para evitar erros futuros
            item['cor_emb'] = item.get('cor_emb', '')

        col4, col5 = st.columns(2)
        with col4:
            item['pessoa'] = st.text_input("Pessoa/Objeto Relacionado", value=item.get('pessoa', ''), key=f"pessoa_{i}")
        with col5:
            item['ref'] = st.text_input("Referência (Subitem Laudo Constatação)", value=item.get('ref', ''), key=f"ref_{i}")

        # Botão para remover o item (Mantido)
        if st.button("Remover Item", key=f"remover_{i}"):
            del st.session_state.dados_laudo['itens'][i]
            st.experimental_rerun() # Atualiza a interface para refletir a remoção

    st.markdown("---")

    # Botão de Gerar Laudo (Mantido com ajustes para usar st.download_button)
    # Garante que há itens antes de tentar gerar o laudo
    if st.button("📊 Gerar Laudo (.docx)"):
        # Validação simples: Verifica se RG da Perícia foi preenchido
        rg_pericia = st.session_state.dados_laudo.get('rg_pericia', '').strip()
        if not rg_pericia:
            st.warning("⚠️ Por favor, informe o RG da Perícia para gerar o nome do arquivo.")
        else:
            with st.spinner("Gerando documento... Por favor, aguarde."):
                try:
                    document = gerar_laudo_docx(st.session_state.dados_laudo)
                    doc_io = io.BytesIO()
                    document.save(doc_io)
                    doc_io.seek(0)

                    # Usa o RG da Perícia para o nome do arquivo
                    file_name = f"{rg_pericia}.docx"

                    st.download_button(
                        label=f"✅ Download Laudo ({file_name})", data=doc_io,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_button"
                    )
                    st.success("Laudo gerado com sucesso! Clique no botão acima para baixar.")
                except Exception as e:
                    st.error(f"❌ Ocorreu um erro ao gerar o laudo:")
                    st.exception(e)
                    print(f"Erro detalhado na geração do DOCX: {e}\n{traceback.format_exc()}")

if __name__ == "__main__":
    main()
