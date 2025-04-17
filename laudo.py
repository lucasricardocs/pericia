# -*- coding: utf-8 -*-
"""
Gerador de Laudo Pericial v3.1 (Streamlit + Lógica Colab + Cores SPTC - Layout Ajustado)

Combina a interface Streamlit e formatação DOCX avançada com a lógica
de geração de texto e entradas (lacre, RG) do script original do Colab.
Usa a fonte 'Gadugi' e o método de itálico do script Colab.
Layout do cabeçalho ajustado conforme feedback.

Requerimentos:
    - streamlit
    - python-docx
    - Pillow (PIL)
    - pytz

Uso:
    1. Instale as dependências: pip install streamlit python-docx Pillow pytz
    2. Salve este código como 'gerador_laudo_combinado_v3_1.py' (ou outro nome)
    3. Salve a imagem do logo como 'logo_policia_cientifica.png' no mesmo diretório.
    4. Execute o script: streamlit run gerador_laudo_combinado_v3_1.py
    5. Interaja com a interface web para inserir dados e gerar o laudo.
    6. Baixe o laudo gerado como um arquivo .docx (nomeado com o RG da Perícia).
"""

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
# Importações necessárias para campos de página
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import traceback

# --- Constantes ---
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

# Cores Institucionais SPTC/GO (para uso no DOCX)
DOCX_COR_AZUL_SPTC = RGBColor(0, 71, 143)
DOCX_COR_CINZA_SPTC = RGBColor(110, 110, 110)
DOCX_COR_PRETO = RGBColor(0, 0, 0)

# Lista de termos para itálico (do código original Colab)
TERMOS_ITALICO_ORIGINAL = [
    'Cannabis sativa L.', # Adicionado L. para consistência
    'Cannabis sativa',
    'Scientific Working Group for the Analysis of Seized Drugs',
    'United Nations Office on Drugs and Crime',
    'Fast blue salt B', # Usado na seção de Exames do código Colab
    'eppendorf',
    'ziplock',
    'Tetrahidrocanabinol', # Mencionado na conclusão Colab
    'Portaria nº 344/1998', # Itálico não usual, mas presente implicitamente na formatação Colab
    'RDC nº 970, de 19/03/2025' # Idem
    # Adicionar outros termos se necessário
]

# --- Funções Auxiliares (Pluralização, Extenso, Parágrafo, Imagem) ---
def pluralizar_palavra(palavra, quantidade):
    """Pluraliza palavras em português (com algumas regras básicas)."""
    if quantidade == 1:
        return palavra
    # Casos especiais que não pluralizam ou têm forma fixa
    if palavra in ["microtubo do tipo eppendorf", "embalagem do tipo ziplock", "papel alumínio"]:
        return palavra
    if palavra.endswith('m') and palavra not in ["alumínio"]: # Evita 'alumínions'
        return re.sub(r'm$', 'ns', palavra) # Ex: item -> itens
    if palavra.endswith('ão'):
        return re.sub(r'ão$', 'ões', palavra) # Ex: porção -> porções
    elif palavra.endswith(('r', 'z', 's')):
        # Termina em 'r' ou 'z': adiciona 'es'
        if palavra.endswith(('r', 'z')):
             return palavra + 'es' # Ex: cor -> cores
        # Termina em 's': geralmente não muda (mas depende da sílaba tônica, simplificado aqui)
        else:
             return palavra # Ex: mês -> meses (precisaria de acentuação), mas lápis -> lápis
    elif palavra.endswith('l'):
         # Troca 'l' por 'is'
        return palavra[:-1] + 'is' # Ex: papel -> papéis, vegetal -> vegetais
    else:
        # Regra geral: adiciona 's'
        return palavra + 's'

def obter_quantidade_extenso(qtd):
    """Retorna a quantidade por extenso (1-10) ou o número como string."""
    return QUANTIDADES_EXTENSO.get(qtd, str(qtd))

def adicionar_paragrafo(doc, text, style=None, align=None, color=None, size=None, bold=False, italic=False):
    """Adiciona um parágrafo ao documento docx com formatação flexível."""
    p = doc.add_paragraph()
    # Aplica estilo de parágrafo
    if style and style in doc.styles:
        try:
            p.style = doc.styles[style]
        except Exception as e:
            print(f"Erro ao aplicar estilo '{style}': {e}. Usando 'Normal'.")
            p.style = doc.styles['Normal']
    elif style: # Se o estilo for passado mas não existir, usar Normal
        print(f"Estilo '{style}' não encontrado. Usando 'Normal'.")
        p.style = doc.styles['Normal']

    # Aplica alinhamento
    if align:
        align_map = {
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY, 'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT, 'left': WD_ALIGN_PARAGRAPH.LEFT
        }
        # Garante que a chave é string e minúscula
        p.alignment = align_map.get(str(align).lower(), WD_ALIGN_PARAGRAPH.LEFT)

    # Adiciona o texto e aplica formatação de caractere
    run = p.add_run(text)
    if color:
        try:
            if isinstance(color, RGBColor): run.font.color.rgb = color
            elif isinstance(color, (tuple, list)) and len(color) == 3: run.font.color.rgb = RGBColor(color[0], color[1], color[2])
            else: print(f"Formato de cor inválido: {color}")
        except Exception as e: print(f"Erro ao aplicar cor: {e}")
    if size:
        try: run.font.size = Pt(int(size))
        except ValueError: print(f"Tamanho de fonte inválido: {size}")
    if bold: run.font.bold = True
    if italic: run.font.italic = True

def inserir_imagem_docx(doc, image_file_uploader):
    """Insere uma imagem vinda do st.file_uploader no documento docx, centralizada."""
    try:
        if image_file_uploader:
            img_stream = io.BytesIO(image_file_uploader.getvalue())
            img = Image.open(img_stream)
            width_px, height_px = img.size
            max_width_inches = 6.0 # Largura máxima A4 menos margens
            dpi = img.info.get('dpi', (96, 96))[0] # Tenta obter DPI, padrão 96
            if dpi <= 0: dpi = 96 # Evita divisão por zero

            width_inches = width_px / dpi

            # Ajusta o tamanho para caber na página se for muito grande
            if width_inches > max_width_inches:
                display_width_inches = max_width_inches
            else:
                display_width_inches = width_inches

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            img_stream.seek(0) # Volta ao início do stream após ler com PIL
            run.add_picture(img_stream, width=Inches(display_width_inches))
    except Exception as e:
        st.error(f"Erro ao inserir imagem no docx: {e}")
        print(f"Erro detalhado ao inserir imagem: {e}\n{traceback.format_exc()}")

# --- Funções de Estrutura do Documento DOCX ---

def configurar_estilos(doc):
    """Configura os estilos de parágrafo e caractere do documento docx
       usando a fonte 'Gadugi' e cores institucionais da SPTC/GO."""

    FONTE_PADRAO = 'Gadugi' # Alterado para Gadugi
    COR_TEXTO_PRINCIPAL = DOCX_COR_PRETO
    COR_DESTAQUE = DOCX_COR_AZUL_SPTC
    COR_TEXTO_SECUNDARIO = DOCX_COR_CINZA_SPTC

    def get_or_add_style(doc, style_name, style_type):
        """Tenta obter um estilo, se não existir, tenta adicioná-lo."""
        if style_name in doc.styles:
            return doc.styles[style_name]
        else:
            try:
                return doc.styles.add_style(style_name, style_type)
            except Exception as e:
                print(f"Falha ao adicionar estilo '{style_name}': {e}. Usando 'Normal' como fallback.")
                return doc.styles['Normal'] # Retorna um estilo padrão seguro

    # Estilo Normal (Base)
    paragrafo_style = doc.styles['Normal']
    paragrafo_style.font.name = FONTE_PADRAO # Gadugi
    paragrafo_style.font.size = Pt(12)
    paragrafo_style.font.color.rgb = COR_TEXTO_PRINCIPAL
    paragrafo_style.paragraph_format.line_spacing = 1.15
    paragrafo_style.paragraph_format.space_before = Pt(0)
    paragrafo_style.paragraph_format.space_after = Pt(8)

    # Estilo para Títulos Principais (Seções)
    titulo_principal_style = get_or_add_style(doc, 'TituloPrincipal', WD_STYLE_TYPE.PARAGRAPH)
    titulo_principal_style.base_style = doc.styles['Normal']
    titulo_principal_style.font.name = FONTE_PADRAO # Gadugi
    titulo_principal_style.font.size = Pt(14)
    titulo_principal_style.font.bold = True
    titulo_principal_style.font.color.rgb = COR_DESTAQUE # Azul SPTC
    titulo_principal_style.paragraph_format.space_before = Pt(12)
    titulo_principal_style.paragraph_format.space_after = Pt(6)

    # Estilo para Títulos Secundários (Subseções)
    titulo_secundario_style = get_or_add_style(doc, 'TituloSecundario', WD_STYLE_TYPE.PARAGRAPH)
    titulo_secundario_style.base_style = doc.styles['Normal']
    titulo_secundario_style.font.name = FONTE_PADRAO # Gadugi
    titulo_secundario_style.font.size = Pt(12)
    titulo_secundario_style.font.bold = True
    titulo_secundario_style.font.color.rgb = COR_DESTAQUE # Azul SPTC
    titulo_secundario_style.paragraph_format.space_before = Pt(10)
    titulo_secundario_style.paragraph_format.space_after = Pt(4)

    # Estilo para Legendas de Ilustrações
    ilustracao_style = get_or_add_style(doc, 'Ilustracao', WD_STYLE_TYPE.PARAGRAPH)
    ilustracao_style.base_style = doc.styles['Normal']
    ilustracao_style.font.name = FONTE_PADRAO # Gadugi
    ilustracao_style.font.size = Pt(10) # Tamanho menor para legenda
    ilustracao_style.font.italic = True
    ilustracao_style.font.color.rgb = COR_TEXTO_SECUNDARIO # Cinza SPTC
    ilustracao_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ilustracao_style.paragraph_format.space_before = Pt(4)
    ilustracao_style.paragraph_format.space_after = Pt(10)

def configurar_pagina(doc):
    """Configura margens da página (padrão ABNT)."""
    for section in doc.sections:
        section.page_height = Inches(11.69) # A4 Altura
        section.page_width = Inches(8.27)  # A4 Largura
        section.top_margin = Inches(1.18)  # 3 cm
        section.bottom_margin = Inches(0.79) # 2 cm
        section.left_margin = Inches(1.18)   # 3 cm
        section.right_margin = Inches(0.79)  # 2 cm

def adicionar_cabecalho_rodape(doc):
    """Adiciona cabeçalho e rodapé padrão ao documento docx."""
    FONTE_CABECALHO_RODAPE = 'Gadugi' # Usar Gadugi aqui também
    TAMANHO_CABECALHO_RODAPE = Pt(10)

    section = doc.sections[0] # Assume que há pelo menos uma seção

    # --- Cabeçalho ---
    header = section.header
    # Limpa cabeçalho existente para evitar duplicação
    if header.paragraphs:
        for para in header.paragraphs:
            p_element = para._element
            if p_element.getparent() is not None:
                p_element.getparent().remove(p_element)

    # Adiciona novo cabeçalho
    header_paragraph = header.add_paragraph()
    run_header_left = header_paragraph.add_run("POLÍCIA CIENTÍFICA DE GOIÁS")
    run_header_left.font.name = FONTE_CABECALHO_RODAPE
    run_header_left.font.size = TAMANHO_CABECALHO_RODAPE
    run_header_left.font.bold = True
    header_paragraph.add_run("\t\t") # Usar tabulação para espaçar
    run_header_right = header_paragraph.add_run("LAUDO DE PERÍCIA CRIMINAL")
    run_header_right.font.name = FONTE_CABECALHO_RODAPE
    run_header_right.font.size = TAMANHO_CABECALHO_RODAPE
    run_header_right.font.bold = False
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT # Alinhado à direita fica melhor

    # --- Rodapé (Numeração de Página) ---
    footer = section.footer
    # Limpa rodapé existente
    if footer.paragraphs:
        for para in footer.paragraphs:
             p_element = para._element
             if p_element.getparent() is not None:
                 p_element.getparent().remove(p_element)
    # Adiciona parágrafo para numeração
    page_num_paragraph = footer.add_paragraph()
    page_num_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Adiciona "Página X"
    run_page = page_num_paragraph.add_run("Página ")
    run_page.font.name = FONTE_CABECALHO_RODAPE
    run_page.font.size = TAMANHO_CABECALHO_RODAPE
    # Campo PAGE
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

    # Adiciona " de Y"
    run_num_pages = page_num_paragraph.add_run(" de ")
    run_num_pages.font.name = FONTE_CABECALHO_RODAPE
    run_num_pages.font.size = TAMANHO_CABECALHO_RODAPE
    # Campo NUMPAGES
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

# --- Funções das Seções do Laudo (Numeração e Conteúdo Ajustados) ---

def adicionar_material_recebido(doc, dados_laudo):
    """Adiciona a seção '1 MATERIAL RECEBIDO PARA EXAME' ao laudo docx."""
    # Numeração corrigida para 1.
    adicionar_paragrafo(doc, "1 MATERIAL RECEBIDO PARA EXAME", style='TituloPrincipal')
    # Texto introdutório pode ser adicionado aqui se desejado. Ex:
    # adicionar_paragrafo(doc, "O material foi recebido neste Instituto devidamente acondicionado e lacrado.", align='justify', style='Normal')

    imagem_carregada = dados_laudo.get('imagem')
    if imagem_carregada:
        inserir_imagem_docx(doc, imagem_carregada)
        # Adiciona legenda à imagem (usará a cor Cinza SPTC e fonte Gadugi definida no estilo 'Ilustracao')
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
        acond = "acondicionada em" if qtd == 1 else "acondicionadas, individualmente, em" # Ajustado ", individualmente,"
        ref_texto = f", relacionada a {item['pessoa']}" if item.get('pessoa') else ""
        subitem_ref = item.get('ref', '')
        # Texto adaptado do código Colab original
        subitem_texto = f", referente à amostra do subitem {subitem_ref} do laudo de constatação supracitado" if subitem_ref else ""
        item_num_str = f"1.{i + 1}" # Numeração corrigida para 1.x
        final_ponto = "."
        texto = (f"{item_num_str} {qtd} ({qtd_ext}) {porcao} de material {tipo_material}, {acond} {embalagem_final}{subitem_texto}{ref_texto}{final_ponto}")
        adicionar_paragrafo(doc, texto, style='Normal', align='justify')

        # Mapeamento para Exames/Resultados/Conclusão
        chave_mapeamento = subitem_ref if subitem_ref else f"Item_{item_num_str}" # Mantém fallback se ref vazia
        item_num_referencia = item_num_str # Usar a referência 1.x para os textos
        if tipo_mat_cod in ["v", "r"]:
            subitens_cannabis[chave_mapeamento] = item_num_referencia
        elif tipo_mat_cod in ["po", "pd"]:
             subitens_cocaina[chave_mapeamento] = item_num_referencia

    return subitens_cannabis, subitens_cocaina

def adicionar_objetivo_exames(doc):
    """Adiciona a seção '2 OBJETIVO DOS EXAMES' (Texto do Colab)."""
    # Numeração corrigida para 2.
    adicionar_paragrafo(doc, "2 OBJETIVO DOS EXAMES", style='TituloPrincipal')
    # Texto do código Colab original
    texto = ("Visa esclarecer à autoridade requisitante quanto às características do material apresentado, "
             "bem como se ele contém substância de uso proscrito no Brasil e capaz de causar dependência física e/ou psíquica. "
             "O presente laudo pericial busca demonstrar a materialidade da infração penal apurada.")
    adicionar_paragrafo(doc, texto, align='justify', style='Normal')

def adicionar_exames(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    """Adiciona a seção '3 EXAMES' (Texto e lógica do Colab)."""
    # Numeração corrigida para 3.
    adicionar_paragrafo(doc, "3 EXAMES", style='TituloPrincipal')

    has_cannabis_item = bool(subitens_cannabis)
    has_cocaina_item = bool(subitens_cocaina)

    # Adota a estrutura de subitens do código Colab
    idx_subitem = 1
    if has_cannabis_item:
        adicionar_paragrafo(doc, f"3.{idx_subitem} Exames realizados para pesquisa de Cannabis sativa L. (maconha)", style='TituloSecundario')
        adicionar_paragrafo(doc, f"3.{idx_subitem}.1 Ensaio químico com Fast blue salt B: teste de cor em reação com solução aquosa de sal de azul sólido B em meio alcalino;", style='Normal', align='justify')
        adicionar_paragrafo(doc, f"3.{idx_subitem}.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas contendo eluentes apropriados e posterior revelação com solução aquosa de azul sólido B.", style='Normal', align='justify')
        idx_subitem += 1

    if has_cocaina_item:
        adicionar_paragrafo(doc, f"3.{idx_subitem} Exames realizados para pesquisa de cocaína", style='TituloSecundario')
        adicionar_paragrafo(doc, f"3.{idx_subitem}.1 Ensaio químico com teste de tiocianato de cobalto-reação de cor com solução de tiocianato de cobalto em meio ácido;", style='Normal', align='justify')
        adicionar_paragrafo(doc, f"3.{idx_subitem}.2 Cromatografia em Camada Delgada (CCD), comparativa com substância padrão, em sistemas com eluentes apropriados e revelação com solução de iodo platinado.", style='Normal', align='justify')
        idx_subitem += 1

    # Se nenhum dos dois foi detectado mas há itens, adiciona exame macroscópico
    if not has_cannabis_item and not has_cocaina_item and dados_laudo.get('itens'):
        adicionar_paragrafo(doc, f"3.{idx_subitem} Exames realizados", style='TituloSecundario')
        adicionar_paragrafo(doc, f"3.{idx_subitem}.1 Exame macroscópico;", style='Normal', align='justify')
        idx_subitem += 1

    if idx_subitem == 1: # Se nenhum item foi adicionado
         adicionar_paragrafo(doc, "Nenhum exame específico a relatar com base nos materiais descritos.", style='Normal')

def adicionar_resultados(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    """Adiciona a seção '4 RESULTADOS' (Texto e lógica do Colab)."""
    # Numeração corrigida para 4.
    adicionar_paragrafo(doc, "4 RESULTADOS", style='TituloPrincipal')

    has_cannabis_item = bool(subitens_cannabis)
    has_cocaina_item = bool(subitens_cocaina)
    idx_subitem = 1

    if has_cannabis_item:
        # Obtém os números dos itens (1.x) associados a Cannabis
        itens_referencia = sorted(list(subitens_cannabis.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}"
        adicionar_paragrafo(doc, f"4.{idx_subitem} Resultados obtidos para o(s) material(is) descrito(s) {label}:", style='TituloSecundario')
        adicionar_paragrafo(doc, f"4.{idx_subitem}.1 No ensaio com Fast blue salt B, foram obtidas coloração característica para canabinol e tetrahidrocanabinol (princípios ativos da Cannabis sativa L.).", style='Normal', align='justify')
        adicionar_paragrafo(doc, f"4.{idx_subitem}.2 Na CCD, obtiveram-se perfis cromatográficos coincidentes com o material de referência (padrão de Cannabis sativa L.); portanto, a substância tetrahidrocanabinol está presente nos materiais questionados.", style='Normal', align='justify')
        idx_subitem += 1

    if has_cocaina_item:
        itens_referencia = sorted(list(subitens_cocaina.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}"
        adicionar_paragrafo(doc, f"4.{idx_subitem} Resultados obtidos para o(s) material(is) descrito(s) {label}:", style='TituloSecundario')
        adicionar_paragrafo(doc, f"4.{idx_subitem}.1 No teste de tiocianato de cobalto, foram obtidas coloração característica para cocaína;", style='Normal', align='justify')
        adicionar_paragrafo(doc, f"4.{idx_subitem}.2 Na CCD, obteve-se perfis cromatográficos coincidentes com o material de referência (padrão de cocaína); portanto, a substância cocaína está presente nos materiais questionados.", style='Normal', align='justify')
        idx_subitem += 1

    if idx_subitem == 1: # Se nenhum resultado foi adicionado
        if dados_laudo.get('itens'):
            adicionar_paragrafo(doc, "Não foram obtidos resultados positivos para Cannabis ou Cocaína nos testes realizados para os materiais descritos.", style='Normal', align='justify')
        else:
            adicionar_paragrafo(doc, "Nenhum material foi submetido a exame, portanto, não há resultados a relatar.", style='Normal', align='justify')


def adicionar_conclusao(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    """Adiciona a seção '5 CONCLUSÃO' (Texto e lógica do Colab)."""
    # Numeração corrigida para 5.
    adicionar_paragrafo(doc, "5 CONCLUSÃO", style='TituloPrincipal')

    conclusoes = []
    if subitens_cannabis:
        itens_referencia = sorted(list(subitens_cannabis.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}"
        conclusoes.append(f"no(s) material(is) descrito(s) {label}, foi detectada a presença de partes "
                           f"da planta Cannabis sativa L., vulgarmente conhecida por maconha. "
                           f"A Cannabis sativa L. contém princípios ativos chamados canabinóis, dentre os quais se encontra o tetrahidrocanabinol, substância perturbadora do sistema nervoso central. "
                           f"Tanto a Cannabis sativa L. quanto a tetrahidrocanabinol são proscritas no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.") # Data da RDC do código Colab

    if subitens_cocaina:
        itens_referencia = sorted(list(subitens_cocaina.values()))
        refs_str = " e ".join(itens_referencia)
        label = f"no item {refs_str}" if len(itens_referencia) == 1 else f"nos itens {refs_str}" # Ajuste na descrição (era 'no(s) subitem(ns)')
        conclusoes.append(f"no(s) material(is) descrito(s) {label}, foi detectada a presença de cocaína, substância alcaloide estimulante do sistema nervoso central. A cocaína é proscrita no país, com fulcro na Portaria nº 344/1998, atualizada por meio da RDC nº 970, de 19/03/2025, da Anvisa.") # Data da RDC do código Colab

    if conclusoes:
        # Junta as conclusões com "Outrossim," como no código Colab
        texto_final = "A partir das análises realizadas, conclui-se que, " + " Outrossim, ".join(conclusoes)
    elif dados_laudo.get('itens'): # Se houve itens mas sem resultado positivo
        texto_final = "A partir das análises realizadas, conclui-se que não foram detectadas substâncias de uso proscrito nos materiais analisados."
    else: # Se não houve itens
        texto_final = "Não houve material submetido a exame, portanto, não há conclusões a apresentar."

    adicionar_paragrafo(doc, texto_final, align='justify', style='Normal')

def adicionar_custodia_material(doc, dados_laudo):
    """Adiciona a seção '6 CUSTÓDIA DO MATERIAL' (Texto do Colab, com Lacre do input)."""
    # Numeração corrigida para 6.
    adicionar_paragrafo(doc, "6 CUSTÓDIA DO MATERIAL", style='TituloPrincipal')
    adicionar_paragrafo(doc, "6.1 Contraprova", style='TituloSecundario') # Usar TituloSecundario para subitem

    # Pega o lacre do estado da sessão (que veio do input do Streamlit)
    lacre = dados_laudo.get('lacre', '_______') # Usa placeholder se não informado

    # Texto adaptado do código Colab
    texto_contraprova = (f"A amostra contraprova ficará armazenada neste Instituto, conforme Portaria 0003/2019/SSP "
                         f"(Lacre nº {lacre}).")
    adicionar_paragrafo(doc, texto_contraprova, style='Normal', align='justify')

def adicionar_referencias(doc, subitens_cannabis, subitens_cocaina):
    """Adiciona a seção 'REFERÊNCIAS' (Texto e lógica do Colab)."""
    adicionar_paragrafo(doc, "REFERÊNCIAS", style='TituloPrincipal')
    # Tamanho da fonte menor para referências
    tamanho_ref = 10

    referencias_base = [
        "BRASIL. Ministério da Saúde. Portaria SVS/MS n° 344, de 12 de maio de 1998. Aprova o regulamento técnico sobre substâncias e medicamentos sujeitos a controle especial. Diário Oficial da União: Brasília, DF, p. 37, 19 maio 1998. Alterada pela RDC nº 970, de 19/03/2025.", # Data da RDC do Colab
        "GOIÁS. Secretaria de Estado da Segurança Pública. Portaria nº 0003/2019/SSP de 10 de janeiro de 2019. Regulamenta a apreensão, movimentação, exames, acondicionamento, armazenamento e destruição de drogas no âmbito da Secretaria de Estado da Segurança Pública. Diário Oficial do Estado de Goiás: n° 22.972, Goiânia, GO, p. 4-5, 15 jan. 2019.",
        "SWGDRUG: Scientific Working Group for the Analysis of Seized Drugs. Recommendations. Version 8.0 june. 2019. Disponível em: http://www.swgdrug.org/Documents/SWGDRUG%20Recommendations%20Version%208_FINAL_ForPosting_092919.pdf. Acesso em: 07/10/2019." # Data de acesso fixa do código Colab
    ]

    for ref in referencias_base:
        adicionar_paragrafo(doc, ref, style='Normal', align='justify', size=tamanho_ref)

    if subitens_cannabis:
        adicionar_paragrafo(doc, "UNODC (United Nations Office on Drugs and Crime). Laboratory and scientific section. Recommended Methods for the Identification and Analysis of Cannabis and Cannabis Products. New York: 2012.", style='Normal', align='justify', size=tamanho_ref) # Ano ajustado para 2012 como no Colab v2
    if subitens_cocaina:
        adicionar_paragrafo(doc, "UNODC (United Nations Office on Drugs and Crime). Laboratory and Scientific Section. Recommended Methods for the Identification and Analysis of Cocaine in Seized Materials. New York: 2012.", style='Normal', align='justify', size=tamanho_ref)

def adicionar_encerramento_assinatura(doc):
    """Adiciona a frase de encerramento, data, local e a assinatura do perito (formato Colab)."""
    # Frase de encerramento pode ser omitida ou adaptada se preferir o "É o laudo."
    # adicionar_paragrafo(doc, "\nÉ o laudo. Nada mais havendo a lavrar, encerra-se o presente.", style='Normal', align='justify')

    try:
        brasilia_tz = pytz.timezone('America/Sao_Paulo')
        hoje = datetime.now(brasilia_tz)
    except Exception:
        hoje = datetime.now() # Fallback
    mes_atual = meses_portugues.get(hoje.month, f"Mês {hoje.month}")
    # Formato da data e local do código Colab
    data_formatada = f"Goiânia, {hoje.day} de {mes_atual} de {hoje.year}."

    doc.add_paragraph() # Espaço
    adicionar_paragrafo(doc, data_formatada, align='right', style='Normal') # Alinhado à direita como no Colab

    doc.add_paragraph(); doc.add_paragraph() # Mais espaço

    # Assinatura - Usando o formato/texto do Colab
    adicionar_paragrafo(doc, "Laudo assinado digitalmente com dados do assinador à esquerda das páginas", align='left', style='Normal', size=9, italic=True) # Nota sobre assinatura digital
    adicionar_paragrafo(doc, "________________________________________", align='center', style='Normal')
    adicionar_paragrafo(doc, "Daniel Chendes Lima", align='center', style='Normal', bold=True) # Nome do Perito do Colab
    adicionar_paragrafo(doc, "Perito Criminal", align='center', style='Normal') # Cargo do Colab
    # Adicionar Matrícula se desejar/tiver
    # adicionar_paragrafo(doc, "Matrícula nº XXXXXXX", align='center', style='Normal')

def aplicar_italico_fonte_original(doc):
    """Aplica fonte Gadugi e itálico a termos específicos, como no código Colab original."""
    termos_para_italico = TERMOS_ITALICO_ORIGINAL

    for paragraph in doc.paragraphs:
        # Verifica se o parágrafo é a legenda da ilustração para usar tamanho 10
        is_ilustracao = "Ilustração 1:" in paragraph.text and paragraph.style.name == 'Ilustracao'
        tamanho_fonte = Pt(10) if is_ilustracao else Pt(12)

        full_text = paragraph.text
        if not full_text: continue # Pula parágrafos vazios

        # Limpa o parágrafo preservando a formatação original (alinhamento, estilo)
        original_alignment = paragraph.alignment
        original_style = paragraph.style
        paragraph.clear()
        paragraph.alignment = original_alignment
        paragraph.style = original_style

        idx = 0
        while idx < len(full_text):
            match_found = False
            # Procura pelo termo mais longo primeiro para evitar correspondências parciais
            # Ordena por comprimento descendente
            termos_ordenados = sorted(termos_para_italico, key=len, reverse=True)
            for phrase in termos_ordenados:
                # Verifica se o termo começa na posição atual
                # Adiciona espaço/início de string antes e espaço/fim de string depois para evitar subpalavras (simplificado)
                # Melhor seria usar regex com word boundaries, mas mantendo a lógica simples do Colab:
                if full_text[idx:].startswith(phrase):
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
                        break # Sai do loop de termos e continua varrendo o texto

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
            'lacre': '',      # Adicionado
            'itens': [],
            'imagem': None
        }
    # Garante que as chaves existem mesmo se o estado já foi inicializado antes
    if 'rg_pericia' not in st.session_state.dados_laudo: st.session_state.dados_laudo['rg_pericia'] = ''
    if 'lacre' not in st.session_state.dados_laudo: st.session_state.dados_laudo['lacre'] = ''
    if 'itens' not in st.session_state.dados_laudo: st.session_state.dados_laudo['itens'] = []
    if 'imagem' not in st.session_state.dados_laudo: st.session_state.dados_laudo['imagem'] = None
    if not isinstance(st.session_state.dados_laudo.get('itens'), list): st.session_state.dados_laudo['itens'] = []


    # --- Inputs Gerais (RG Perícia e Lacre) ---
    st.header("Informações Gerais")
    col_geral1, col_geral2 = st.columns(2)
    with col_geral1:
        st.session_state.dados_laudo['rg_pericia'] = st.text_input(
            "RG da Perícia (para nome do arquivo)",
            value=st.session_state.dados_laudo['rg_pericia'],
            key="rg_pericia_input",
            help="Ex: 2025_04_12345. Será usado para nomear o arquivo .docx."
        )
    with col_geral2:
        st.session_state.dados_laudo['lacre'] = st.text_input(
            "Número do Lacre da Contraprova",
            value=st.session_state.dados_laudo['lacre'],
            key="lacre_input",
            help="Informe o número do lacre que será usado na contraprova."
        )

    st.markdown("---")

    # --- Coleta de Dados para o Laudo (Itens) ---
    st.header("1 MATERIAL RECEBIDO PARA EXAME")

numero_itens = st.number_input(
    "Número de tipos diferentes de material/acondicionamento a descrever",
    min_value=0,
    value=max(0, len(st.session_state.dados_laudo.get('itens', []))),
    step=1,
    key="num_itens_input"
)

# --- Mantida a lógica de adição/remoção de itens ---
current_num_itens = len(st.session_state.dados_laudo['itens'])
if numero_itens > current_num_itens:
    for _ in range(numero_itens - current_num_itens):
        st.session_state.dados_laudo['itens'].append({
            'qtd': 1, 'tipo_mat': list(TIPOS_MATERIAL_BASE.keys())[0],
            'emb': list(TIPOS_EMBALAGEM_BASE.keys())[0], 'cor_emb': None,
            'ref': '', 'pessoa': ''
        })
elif numero_itens < current_num_itens:
    st.session_state.dados_laudo['itens'] = st.session_state.dados_laudo['itens'][:numero_itens]

# --- Interface simplificada mantendo os expanders ---
if numero_itens > 0:
    for i in range(numero_itens):
        with st.expander(f"Item 1.{i + 1}", expanded=True):
            item = st.session_state.dados_laudo['itens'][i]
            
            # Linha 1
            col1, col2 = st.columns([1, 3])
            with col1:
                item['qtd'] = st.number_input(
                    "Quantidade", 
                    min_value=1,
                    value=item['qtd'],
                    key=f"qtd_{i}"
                )
            
            with col2:
                item['tipo_mat'] = st.selectbox(
                    "Tipo de material",
                    options=list(TIPOS_MATERIAL_BASE.keys()),
                    index=list(TIPOS_MATERIAL_BASE.keys()).index(item['tipo_mat']),
                    key=f"mat_{i}"
                )

            # Linha 2
            col3, col4 = st.columns([3, 2])
            with col3:
                item['emb'] = st.selectbox(
                    "Embalagem",
                    options=list(TIPOS_EMBALAGEM_BASE.keys()),
                    index=list(TIPOS_EMBALAGEM_BASE.keys()).index(item['emb']),
                    key=f"emb_{i}"
                )
            
            with col4:
                if item['emb'] in ['pl', 'pa', 'e', 'z']:
                    item['cor_emb'] = st.selectbox(
                        "Cor",
                        options=[None] + list(CORES_FEMININO_EMBALAGEM.keys()),
                        index=0 if item['cor_emb'] is None else list(CORES_FEMININO_EMBALAGEM.keys()).index(item['cor_emb']) + 1,
                        key=f"cor_{i}"
                    )
                else:
                    st.info("Sem cor específica")

            # Linha 3
            item['ref'] = st.text_input(
                "Referência do subitem",
                value=item['ref'],
                key=f"ref_{i}"
            )

            item['pessoa'] = st.text_input(
                "Pessoa relacionada (opcional)",
                value=item['pessoa'],
                key=f"pessoa_{i}"
            )
            
    st.markdown("---")

    # --- Upload de Imagem ---
    st.header("Ilustração (Opcional)")
    uploaded_image = st.file_uploader(
        "Carregar imagem do(s) material(is) recebido(s)",
        type=["png", "jpg", "jpeg", "bmp", "gif"],
        key="image_uploader",
        help="Faça o upload de uma imagem. Será incluída na Seção 1."
        )
    # Atualiza estado da imagem
    if uploaded_image is not None:
        st.session_state.dados_laudo['imagem'] = uploaded_image
    # Detecta se o usuário removeu a imagem
    elif 'image_uploader' in st.session_state and st.session_state.image_uploader is None:
         st.session_state.dados_laudo['imagem'] = None


    # --- Botão de Geração e Download ---
    st.markdown("---")
    st.header("Gerar e Baixar Laudo")

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
