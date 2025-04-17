# -*- coding: utf-8 -*-
"""
Gerador de Laudo Pericial v2.7 (Streamlit - Foco nos Itens + Cores SPTC)

Este script gera laudos periciais para identificação de drogas, focando
diretamente na descrição dos itens recebidos. A seção de informações
gerais foi removida. As cores da interface e do DOCX foram ajustadas
para seguir a identidade visual da SPTC/GO.

Requerimentos:
    - streamlit
    - python-docx
    - Pillow (PIL)
    - pytz

Uso:
    1. Instale as dependências: pip install streamlit python-docx Pillow pytz
    2. Salve este código como 'gerador_laudo_itens.py'
    3. Salve a imagem do logo como 'logo_policia_cientifica.png' no mesmo diretório.
    4. Execute o script: streamlit run gerador_laudo_itens.py
    5. Interaja com a interface web para descrever os itens e gerar o laudo.
    6. Baixe o laudo gerado como um arquivo .docx.
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
# Azul SPTC: #00478F -> RGB(0, 71, 143)
# Cinza SPTC: #6E6E6E -> RGB(110, 110, 110)
# Preto: #000000 -> RGB(0, 0, 0)
# Branco: #FFFFFF -> RGB(255, 255, 255)
DOCX_COR_AZUL_SPTC = RGBColor(0, 71, 143)
DOCX_COR_CINZA_SPTC = RGBColor(110, 110, 110)
DOCX_COR_PRETO = RGBColor(0, 0, 0)


# --- Funções Auxiliares (Pluralização, Extenso, Parágrafo, Imagem) ---
# (Mantidas as mesmas funções auxiliares da versão anterior)
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
       usando as cores institucionais da SPTC/GO."""

    # Usa as cores institucionais definidas globalmente
    COR_TEXTO_PRINCIPAL = DOCX_COR_PRETO        # Preto para corpo do texto
    COR_DESTAQUE = DOCX_COR_AZUL_SPTC           # Azul SPTC para Títulos
    COR_TEXTO_SECUNDARIO = DOCX_COR_CINZA_SPTC  # Cinza SPTC para Legendas/Secundário

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

    # Estilo Normal (Base) - Cor do texto principal (Preto)
    paragrafo_style = doc.styles['Normal']
    paragrafo_style.font.name = 'Calibri'
    paragrafo_style.font.size = Pt(12)
    paragrafo_style.font.color.rgb = COR_TEXTO_PRINCIPAL # Preto
    paragrafo_style.paragraph_format.line_spacing = 1.15
    paragrafo_style.paragraph_format.space_before = Pt(0)
    paragrafo_style.paragraph_format.space_after = Pt(8)

    # Estilo para Títulos Principais (Seções) - Cor de destaque (Azul SPTC)
    titulo_principal_style = get_or_add_style(doc, 'TituloPrincipal', WD_STYLE_TYPE.PARAGRAPH)
    titulo_principal_style.base_style = doc.styles['Normal']
    titulo_principal_style.font.name = 'Calibri'
    titulo_principal_style.font.size = Pt(14)
    titulo_principal_style.font.bold = True
    titulo_principal_style.font.color.rgb = COR_DESTAQUE # Azul SPTC
    titulo_principal_style.paragraph_format.space_before = Pt(12)
    titulo_principal_style.paragraph_format.space_after = Pt(6)

    # Estilo para Títulos Secundários (Subseções) - Cor de destaque (Azul SPTC)
    titulo_secundario_style = get_or_add_style(doc, 'TituloSecundario', WD_STYLE_TYPE.PARAGRAPH)
    titulo_secundario_style.base_style = doc.styles['Normal']
    titulo_secundario_style.font.name = 'Calibri'
    titulo_secundario_style.font.size = Pt(12)
    titulo_secundario_style.font.bold = True
    titulo_secundario_style.font.color.rgb = COR_DESTAQUE # Azul SPTC
    titulo_secundario_style.paragraph_format.space_before = Pt(10)
    titulo_secundario_style.paragraph_format.space_after = Pt(4)

    # Estilo de caractere para Itálico (se não existir)
    if 'Italico' not in doc.styles:
        try:
            italico_style = doc.styles.add_style('Italico', WD_STYLE_TYPE.CHARACTER)
            italico_style.font.italic = True
            italico_style.base_style = doc.styles['Default Paragraph Font']
        except Exception as e:
            print(f"Não foi possível criar estilo 'Italico': {e}")
    elif doc.styles['Italico'].type == WD_STYLE_TYPE.CHARACTER:
        doc.styles['Italico'].font.italic = True

    # Estilo para Legendas de Ilustrações - Cor de texto secundário (Cinza SPTC)
    ilustracao_style = get_or_add_style(doc, 'Ilustracao', WD_STYLE_TYPE.PARAGRAPH)
    ilustracao_style.base_style = doc.styles['Normal']
    ilustracao_style.font.name = 'Calibri'
    ilustracao_style.font.size = Pt(10)
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
    section = doc.sections[0] # Assume que há pelo menos uma seção

    # --- Cabeçalho ---
    header = section.header
    # Limpa cabeçalho existente para evitar duplicação
    if header.paragraphs:
        for para in header.paragraphs:
            p_element = para._element
            p_element.getparent().remove(p_element)
    # Adiciona novo cabeçalho
    header_paragraph = header.add_paragraph()
    run_header_left = header_paragraph.add_run("POLÍCIA CIENTÍFICA DE GOIÁS")
    run_header_left.font.name = 'Calibri'
    run_header_left.font.size = Pt(10)
    run_header_left.font.bold = True
    header_paragraph.add_run("\t\t") # Usar tabulação para espaçar
    run_header_right = header_paragraph.add_run("LAUDO DE PERÍCIA CRIMINAL")
    run_header_right.font.name = 'Calibri'
    run_header_right.font.size = Pt(10)
    run_header_right.font.bold = False
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # --- Rodapé (Numeração de Página) ---
    footer = section.footer
    # Limpa rodapé existente
    if footer.paragraphs:
        for para in footer.paragraphs:
            p_element = para._element
            p_element.getparent().remove(p_element)
    # Adiciona parágrafo para numeração
    page_num_paragraph = footer.add_paragraph()
    page_num_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Adiciona "Página X"
    run_page = page_num_paragraph.add_run("Página ")
    run_page.font.name = 'Calibri'
    run_page.font.size = Pt(10)
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
    run_num_pages.font.name = 'Calibri'
    run_num_pages.font.size = Pt(10)
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

# --- Seção de Preâmbulo REMOVIDA ---

# --- Funções das Seções do Laudo (Numeração Ajustada) ---
# (As funções adicionar_material_recebido, adicionar_objetivo_exames,
# adicionar_exames, adicionar_resultados, adicionar_conclusao,
# adicionar_custodia_material, adicionar_referencias,
# adicionar_encerramento_assinatura, aplicar_italico_especifico
# são mantidas como na versão anterior, pois já usam os estilos
# configurados em configurar_estilos, que agora têm as cores corretas)

def adicionar_material_recebido(doc, dados_laudo):
    """Adiciona a seção '1 MATERIAL RECEBIDO PARA EXAME' ao laudo docx."""
    adicionar_paragrafo(doc, "1 MATERIAL RECEBIDO PARA EXAME", style='TituloPrincipal')
    adicionar_paragrafo(doc, "O material foi recebido neste Instituto devidamente acondicionado e lacrado.", align='justify', style='Normal')

    imagem_carregada = dados_laudo.get('imagem')
    if imagem_carregada:
        inserir_imagem_docx(doc, imagem_carregada)
        # Adiciona legenda à imagem (usará a cor Cinza SPTC definida no estilo 'Ilustracao')
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
        acond = "acondicionada em" if qtd == 1 else "acondicionadas individualmente em"
        ref_texto = f", relacionada a {item['pessoa']}" if item.get('pessoa') else ""
        subitem_ref = item.get('ref', '')
        subitem_texto = f", referente(s) à(s) amostra(s) do(s) subitem(ns) {subitem_ref} do laudo de constatação (se aplicável)" if subitem_ref else ""
        item_num_str = f"1.{i + 1}"
        texto = (f"{item_num_str} – {qtd} ({qtd_ext}) {porcao} de material {tipo_material}, "
                 f"{acond} {embalagem_final}{subitem_texto}{ref_texto}.")
        adicionar_paragrafo(doc, texto, style='Normal', align='justify')

        chave_mapeamento = subitem_ref if subitem_ref else f"Item_{item_num_str}"
        if tipo_mat_cod in ["v", "r"]:
             subitens_cannabis[chave_mapeamento] = item_num_str
        elif tipo_mat_cod in ["po", "pd"]:
             subitens_cocaina[chave_mapeamento] = item_num_str

    return subitens_cannabis, subitens_cocaina

def adicionar_objetivo_exames(doc):
    """Adiciona a seção '2 OBJETIVO DOS EXAMES'."""
    adicionar_paragrafo(doc, "2 OBJETIVO DOS EXAMES", style='TituloPrincipal') # Usará Azul SPTC
    texto = ("O objetivo dos exames é identificar a natureza do material apresentado, verificando "
             "a presença de substâncias entorpecentes ou psicotrópicas capazes de causar dependência "
             "física ou psíquica, cujo uso e/ou comercialização são proscritos em todo o território "
             "nacional, conforme legislação vigente (Portaria SVS/MS nº 344/1998 e suas atualizações).")
    adicionar_paragrafo(doc, texto, align='justify', style='Normal') # Usará Preto

def adicionar_exames(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    """Adiciona a seção '3 EXAMES'."""
    adicionar_paragrafo(doc, "3 EXAMES", style='TituloPrincipal') # Usará Azul SPTC
    adicionar_paragrafo(doc, "Os materiais recebidos foram submetidos aos seguintes exames e testes:", style='Normal', align='justify') # Usará Preto

    has_cannabis_item = bool(subitens_cannabis)
    has_cocaina_item = bool(subitens_cocaina)
    itens_outros = [item for item in dados_laudo.get('itens', []) if item.get('tipo_mat') not in ["v", "r", "po", "pd"]]
    has_outros_item = bool(itens_outros)
    idx_counter = 1

    if dados_laudo.get('itens'):
        adicionar_paragrafo(doc, f"3.{idx_counter} Exame macroscópico:", style='TituloSecundario') # Usará Azul SPTC
        adicionar_paragrafo(doc, "Observação das características gerais do material, como aspecto físico (pó, erva, pedra, etc.), coloração, odor e acondicionamento.", style='Normal', align='justify') # Usará Preto
        idx_counter += 1

    if has_cannabis_item:
        # Adiciona parágrafo com run específico para itálico
        p_cannabis = doc.add_paragraph()
        p_cannabis.style = doc.styles['TituloSecundario'] # Aplica estilo Azul SPTC
        p_cannabis.add_run(f"3.{idx_counter} Testes para ")
        run_italic_cannabis = p_cannabis.add_run("Cannabis sativa")
        run_italic_cannabis.italic = True
        p_cannabis.add_run(" L.:")

        adicionar_paragrafo(doc, "   a) Reação Duquenois-Levine modificado;", style='Normal') # Preto
        adicionar_paragrafo(doc, "   b) Reação Fast Blue B Salt;", style='Normal') # Preto
        adicionar_paragrafo(doc, "   c) Cromatografia em Camada Delgada (CCD) comparativa com padrão de referência.", style='Normal') # Preto
        idx_counter += 1

    if has_cocaina_item:
        adicionar_paragrafo(doc, f"3.{idx_counter} Testes para cocaína:", style='TituloSecundario') # Azul SPTC
        adicionar_paragrafo(doc, "   a) Reação Tiocianato de Cobalto;", style='Normal') # Preto
        adicionar_paragrafo(doc, "   b) Cromatografia em Camada Delgada (CCD) comparativa com padrão de referência.", style='Normal') # Preto
        idx_counter += 1

    if has_outros_item:
        nums_itens_outros = [f"1.{i+1}" for i, item in enumerate(dados_laudo['itens']) if item.get('tipo_mat') not in ["v", "r", "po", "pd"]]
        desc_itens_str = ", ".join(sorted(nums_itens_outros))
        label_desc = "no item" if len(nums_itens_outros) == 1 else "nos itens"
        adicionar_paragrafo(doc, f"3.{idx_counter} Testes para outras substâncias (material {label_desc} {desc_itens_str}):", style='TituloSecundario') # Azul SPTC
        adicionar_paragrafo(doc, "Realização de testes preliminares de coloração e/ou CCD apropriados para investigação de outras substâncias controladas (Ex: anfetaminas, opiáceos), conforme características observadas no exame macroscópico.", style='Normal', align='justify') # Preto
        idx_counter += 1

    if idx_counter == 1 and not dados_laudo.get('itens'):
         adicionar_paragrafo(doc, "Nenhum material descrito para submissão a exames.", style='Normal') # Preto

def adicionar_resultados(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    """Adiciona a seção '4 RESULTADOS'."""
    adicionar_paragrafo(doc, "4 RESULTADOS", style='TituloPrincipal') # Azul SPTC
    idx_counter = 1

    if subitens_cannabis:
        desc_itens_nums = sorted(list(subitens_cannabis.values()))
        desc_itens_str = ", ".join(desc_itens_nums)
        label_desc = "no item" if len(desc_itens_nums) == 1 else "nos itens"
        adicionar_paragrafo(doc, f"4.{idx_counter} Para o(s) material(is) descrito(s) {label_desc} {desc_itens_str}:", style='TituloSecundario') # Azul SPTC

        # Parágrafo com itálico
        p_macro_c = doc.add_paragraph(style='Normal') # Preto
        p_macro_c.add_run("   a) Exame macroscópico: Material com características compatíveis com ")
        run_italic_mc = p_macro_c.add_run("Cannabis sativa")
        run_italic_mc.italic = True
        p_macro_c.add_run(" L. (odor característico, aspecto de erva picada e prensada ou fragmentos resinosos).")
        p_macro_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        adicionar_paragrafo(doc, "   b) Testes químicos (Duquenois-Levine e Fast Blue B Salt): Resultados positivos para a presença de canabinoides.", style='Normal', align='justify') # Preto
        adicionar_paragrafo(doc, "   c) CCD: Resultado compatível com padrão de referência para Tetrahidrocanabinol (THC).", style='Normal', align='justify') # Preto
        idx_counter += 1

    if subitens_cocaina:
        desc_itens_nums = sorted(list(subitens_cocaina.values()))
        desc_itens_str = ", ".join(desc_itens_nums)
        label_desc = "no item" if len(desc_itens_nums) == 1 else "nos itens"
        adicionar_paragrafo(doc, f"4.{idx_counter} Para o(s) material(is) descrito(s) {label_desc} {desc_itens_str}:", style='TituloSecundario') # Azul SPTC
        adicionar_paragrafo(doc, "   a) Exame macroscópico: Material pulverulento de coloração esbranquiçada ou amarelada, ou material petrificado (\"crack\"), com odor característico.", style='Normal', align='justify') # Preto
        adicionar_paragrafo(doc, "   b) Teste químico (Tiocianato de Cobalto): Resultado positivo para a presença de cocaína.", style='Normal', align='justify') # Preto
        adicionar_paragrafo(doc, "   c) CCD: Resultado compatível com padrão de referência para Cocaína.", style='Normal', align='justify') # Preto
        idx_counter += 1

    itens_outros = [item for i, item in enumerate(dados_laudo.get('itens', [])) if item.get('tipo_mat') not in ["v", "r", "po", "pd"]]
    if itens_outros:
        nums_itens_outros = sorted([f"1.{i+1}" for i, item in enumerate(dados_laudo['itens']) if item.get('tipo_mat') not in ["v", "r", "po", "pd"]])
        desc_itens_str = ", ".join(nums_itens_outros)
        label_desc = "no item" if len(nums_itens_outros) == 1 else "nos itens"
        adicionar_paragrafo(doc, f"4.{idx_counter} Para o(s) material(is) descrito(s) {label_desc} {desc_itens_str}:", style='TituloSecundario') # Azul SPTC
        adicionar_paragrafo(doc, "   a) Exame macroscópico: [Descrever características observadas para estes itens, ex: comprimidos, pó de outra cor, etc.].", style='Normal', align='justify') # Preto
        adicionar_paragrafo(doc, "   b) Demais testes: [Descrever resultados dos testes aplicados, ex: 'Resultados negativos para as principais substâncias testadas', ou 'Resultado positivo para [outra substância]'].", style='Normal', align='justify') # Preto
        idx_counter += 1

    if idx_counter == 1 and not dados_laudo.get('itens'):
        adicionar_paragrafo(doc, "Nenhum material foi submetido a exame, portanto, não há resultados a relatar.", style='Normal', align='justify') # Preto
    elif idx_counter == 1:
         adicionar_paragrafo(doc, "Resultados para os itens descritos não puderam ser classificados como Cannabis ou Cocaína com base nos testes padrões aqui listados.", style='Normal', align='justify') # Preto


def adicionar_conclusao(doc, subitens_cannabis, subitens_cocaina, dados_laudo):
    """Adiciona a seção '5 CONCLUSÃO'."""
    adicionar_paragrafo(doc, "5 CONCLUSÃO", style='TituloPrincipal') # Azul SPTC

    conclusoes = []
    ref_legal = ("substância(s) de uso proscrito no Brasil, conforme a Portaria SVS/MS nº 344/1998 e suas atualizações")

    if subitens_cannabis:
        desc_itens_nums = sorted(list(subitens_cannabis.values()))
        desc_str = ", ".join(desc_itens_nums)
        label_desc = "no material descrito no item" if len(desc_itens_nums) == 1 else "nos materiais descritos nos itens"
        # Conclusão com run para itálico
        concl_cannabis_text = f"{label_desc} {desc_str}, foi detectada a presença de Tetrahidrocanabinol (THC), princípio ativo da Cannabis sativa L. (maconha), {ref_legal}"
        # Adicionar parágrafo e runs manualmente se precisar de itálico aqui, ou usar aplicar_italico_especifico no final
        conclusoes.append(concl_cannabis_text) # Adiciona texto normal por enquanto

    if subitens_cocaina:
        desc_itens_nums = sorted(list(subitens_cocaina.values()))
        desc_str = ", ".join(desc_itens_nums)
        label_desc = "no material descrito no item" if len(desc_itens_nums) == 1 else "nos materiais descritos nos itens"
        conclusoes.append(f"{label_desc} {desc_str}, foi detectada a presença de Cocaína, {ref_legal}")

    itens_outros = [item for i, item in enumerate(dados_laudo.get('itens', [])) if item.get('tipo_mat') not in ["v", "r", "po", "pd"]]
    if itens_outros:
        nums_itens_outros = sorted([f"1.{i+1}" for i, item in enumerate(dados_laudo['itens']) if item.get('tipo_mat') not in ["v", "r", "po", "pd"]])
        desc_str = ", ".join(nums_itens_outros)
        label_desc = "no material descrito no item" if len(nums_itens_outros) == 1 else "nos materiais descritos nos itens"
        conclusoes.append(f"{label_desc} {desc_str}, [concluir sobre a presença de outras substâncias controladas ou indicar resultado negativo para as substâncias pesquisadas]")

    if conclusoes:
        texto_final = "Face ao exposto e com base nos resultados obtidos nos exames realizados, conclui-se que "
        if len(conclusoes) > 1:
            texto_final += "; ".join(conclusoes[:-1]) + "; e " + conclusoes[-1] + "."
        else:
            texto_final += conclusoes[0] + "."
    elif dados_laudo.get('itens'):
        # Conclusão negativa com itálico
        texto_final = ("Face ao exposto e com base nos resultados obtidos nos exames realizados, conclui-se que "
                       "não foram detectadas as substâncias Cannabis sativa L. (maconha) ou Cocaína nos materiais examinados.")
    else:
        texto_final = "Não houve material submetido a exame, portanto, não há conclusões a apresentar."

    # Adiciona o parágrafo de conclusão (itálico será aplicado depois)
    adicionar_paragrafo(doc, texto_final, align='justify', style='Normal') # Preto

def adicionar_custodia_material(doc, dados_laudo):
    """Adiciona a seção '6 CUSTÓDIA DO MATERIAL'."""
    adicionar_paragrafo(doc, "6 CUSTÓDIA DO MATERIAL", style='TituloPrincipal') # Azul SPTC
    adicionar_paragrafo(doc, "6.1 Contraprova:", style='TituloSecundario') # Azul SPTC

    lacre_placeholder = '_____________'
    texto_contraprova = (f"A(s) amostra(s) para eventual contraprova foi(foram) devidamente acondicionada(s) "
                         f"e lacrada(s) novamente com o lacre nº {lacre_placeholder}, encontrando-se à disposição "
                         "da autoridade competente ou da justiça, arquivada(s) neste Instituto.")
    adicionar_paragrafo(doc, texto_contraprova, style='Normal', align='justify') # Preto

def adicionar_referencias(doc, subitens_cannabis, subitens_cocaina):
    """Adiciona a seção 'REFERÊNCIAS'."""
    adicionar_paragrafo(doc, "REFERÊNCIAS", style='TituloPrincipal') # Azul SPTC
    adicionar_paragrafo(doc, "BRASIL. Ministério da Saúde. Agência Nacional de Vigilância Sanitária. Portaria SVS/MS nº 344, de 12 de maio de 1998. Aprova o Regulamento Técnico sobre substâncias e medicamentos sujeitos a controle especial. Diário Oficial da União, Brasília, DF, 15 maio 1998. (e suas atualizações).", style='Normal', align='justify', size=10) # Preto, menor
    adicionar_paragrafo(doc, "GOIÁS. Secretaria de Estado da Segurança Pública. Superintendência de Polícia Técnico-Científica. Procedimento Operacional Padrão – Química Forense (POP-QUIM).", style='Normal', align='justify', size=10) # Preto, menor
    # A data de acesso deve ser atualizada ou removida se não for dinâmica
    hoje_ref = datetime.now().strftime('%d/%m/%Y')
    adicionar_paragrafo(doc, f"SCIENTIFIC WORKING GROUP FOR THE ANALYSIS OF SEIZED DRUGS (SWGDRUG). Recommendations. Version 8.0. Disponível em: <www.swgdrug.org>. Acesso em: {hoje_ref}.", style='Normal', align='justify', size=10) # Preto, menor

    if subitens_cannabis:
        adicionar_paragrafo(doc, "UNITED NATIONS OFFICE ON DRUGS AND CRIME (UNODC). Recommended methods for the identification and analysis of cannabis and cannabis products. Manual for Use by National Drug Analysis Laboratories. New York: UN, 2009.", style='Normal', align='justify', size=10) # Preto, menor
    if subitens_cocaina:
        adicionar_paragrafo(doc, "UNITED NATIONS OFFICE ON DRUGS AND CRIME (UNODC). Recommended methods for the identification and analysis of cocaine in seized materials. Manual for Use by National Drug Analysis Laboratories. New York: UN, 2012.", style='Normal', align='justify', size=10) # Preto, menor

def adicionar_encerramento_assinatura(doc):
    """Adiciona a frase de encerramento, data, local e a assinatura do perito."""
    adicionar_paragrafo(doc, "\nÉ o laudo. Nada mais havendo a lavrar, encerra-se o presente.", style='Normal', align='justify') # Preto

    try:
        brasilia_tz = pytz.timezone('America/Sao_Paulo')
        hoje = datetime.now(brasilia_tz)
    except Exception:
        hoje = datetime.now()
    data_formatada = f"Goiânia, {hoje.day} de {meses_portugues.get(hoje.month, 'MêsInválido')} de {hoje.year}."

    doc.add_paragraph()
    adicionar_paragrafo(doc, data_formatada, align='center', style='Normal') # Preto
    doc.add_paragraph(); doc.add_paragraph()

    adicionar_paragrafo(doc, "________________________________________", align='center', style='Normal') # Preto
    adicionar_paragrafo(doc, "NOME DO PERITO CRIMINAL", align='center', style='Normal', bold=True) # Preto
    adicionar_paragrafo(doc, "Perito Criminal - SPTC/GO", align='center', style='Normal') # Preto
    adicionar_paragrafo(doc, "Matrícula nº XXXXXXX", align='center', style='Normal') # Preto

def aplicar_italico_especifico(doc):
    """Aplica estilo itálico a termos científicos e latinos específicos no documento."""
    termos_italico = ['Cannabis sativa', 'Cannabis sativa L.', 'Tetrahidrocanabinol', 'THC']
    expressoes_latinas = ['et al.', 'i.e.', 'e.g.', 'supra', 'infra', 'in vitro', 'in vivo', 'a priori', 'a posteriori']
    termos_completos = termos_italico + expressoes_latinas
    regex_pattern = r"(?:^|\W)(" + "|".join(re.escape(termo) for termo in termos_completos) + r")($|\W)"

    for paragraph in doc.paragraphs:
        if not any(termo in paragraph.text for termo in termos_completos):
            continue

        # Preserva runs existentes se houver múltiplas formatações no parágrafo
        runs_originais = list(paragraph.runs)
        texto_original_completo = paragraph.text # Pega o texto completo antes de limpar

        # Salva formatação do parágrafo
        original_alignment = paragraph.alignment
        original_style = paragraph.style
        paragraph.clear()
        paragraph.alignment = original_alignment
        paragraph.style = original_style

        last_index = 0
        for match in re.finditer(regex_pattern, texto_original_completo):
            start, end = match.span(1)
            termo_encontrado = match.group(1)

            if start > last_index:
                 # Adiciona texto normal antes do termo, tentando preservar formatação original
                 # (Simplificação: assume formatação uniforme do parágrafo)
                 run_normal = paragraph.add_run(texto_original_completo[last_index:start])
                 # TODO: Idealmente, copiar formatação do run original correspondente

            run_italic = paragraph.add_run(termo_encontrado)
            run_italic.italic = True
            # TODO: Idealmente, copiar outra formatação (bold, size, etc.) do run original

            last_index = end

        if last_index < len(texto_original_completo):
            run_normal = paragraph.add_run(texto_original_completo[last_index:])
            # TODO: Copiar formatação

        # Se o parágrafo ficou vazio (talvez erro na lógica?), restaura o texto original
        if not paragraph.text and texto_original_completo:
            paragraph.text = texto_original_completo


# --- Função Principal de Geração do DOCX ---

def gerar_laudo_docx(dados_laudo):
    """Gera o laudo completo em formato docx (foco nos itens)."""
    document = Document()
    configurar_estilos(document) # Configura estilos COM as cores SPTC
    configurar_pagina(document)
    adicionar_cabecalho_rodape(document)

    # Adiciona Seções na Ordem Correta
    subitens_cannabis, subitens_cocaina = adicionar_material_recebido(document, dados_laudo)
    adicionar_objetivo_exames(document)
    adicionar_exames(document, subitens_cannabis, subitens_cocaina, dados_laudo)
    adicionar_resultados(document, subitens_cannabis, subitens_cocaina, dados_laudo)
    adicionar_conclusao(document, subitens_cannabis, subitens_cocaina, dados_laudo)
    adicionar_custodia_material(document, dados_laudo)
    adicionar_referencias(document, subitens_cannabis, subitens_cocaina)
    adicionar_encerramento_assinatura(document)

    # Aplica itálico (depois de todo o texto ser adicionado)
    aplicar_italico_especifico(document)

    return document

# --- Interface Streamlit ---
def main():
    st.set_page_config(layout="wide", page_title="Gerador de Laudo - Itens")

    # --- Cabeçalho com Logo, Título, Data --- (Hora removida, Cores e Logo ajustados)
    col1, col2, col3 = st.columns([1, 4, 2]) # Ajuste proporção se necessário

    # Define cores institucionais (para uso na interface Streamlit)
    UI_COR_AZUL_SPTC = "#00478F"
    UI_COR_CINZA_SPTC = "#6E6E6E"

    with col1:
        logo_path = "logo_policia_cientifica.png" # Caminho para o logo local
        try:
            # Tenta carregar o logo local e aumenta o tamanho
            st.image(logo_path, width=150) # Logo maior
        except FileNotFoundError:
            # Erro se o arquivo local NÃO for encontrado (removido fallback de URL)
            st.error(f"Erro: Arquivo do logo '{logo_path}' não encontrado no diretório.")
            st.info("Certifique-se de que o arquivo do logo está na mesma pasta que o script ou no repositório GitHub.")
        except Exception as e:
            st.warning(f"Logo não pôde ser carregado: {e}")

    with col2:
        # Usa markdown para aplicar a cor azul institucional ao título
        st.markdown(f'<h1 style="color: {UI_COR_AZUL_SPTC};">Gerador de Laudo Pericial</h1>', unsafe_allow_html=True)
        # Usa markdown para aplicar a cor cinza institucional ao caption
        st.markdown(f'<p style="color: {UI_COR_CINZA_SPTC}; font-size: 0.9em;">Identificação de Drogas - Foco nos Itens - SPTC/GO</p>', unsafe_allow_html=True)

    with col3:
        data_placeholder = st.empty() # Placeholder apenas para a data agora

        # Função para atualizar data (sem hora)
        def atualizar_data():
            try:
                brasilia_tz = pytz.timezone('America/Sao_Paulo')
                now = datetime.now(brasilia_tz)
                dia_semana = dias_semana_portugues.get(now.weekday(), '')
                mes = meses_portugues.get(now.month, '')
                # Formata apenas a data
                data_formatada = f"{dia_semana}, {now.day} de {mes} de {now.year}"

                # Usa HTML/Markdown para formatação, aplicando a cor cinza institucional
                data_placeholder.markdown(f"""
                <div style="text-align: right; font-size: 0.9em; color: {UI_COR_CINZA_SPTC}; line-height: 1.2; margin-top: 10px;">
                    <span>{data_formatada}</span>
                    <br>
                    <span style="font-size: 0.8em;">(Goiânia-GO)</span>
                </div>
                """, unsafe_allow_html=True) # Adicionado referência local/fuso horário
            except Exception as e:
                now = datetime.now()
                fallback_str = now.strftime("%d/%m/%Y") # Formato de data fallback
                data_placeholder.markdown(f"""
                <div style="text-align: right; font-size: 0.9em; color: #FF5555; line-height: 1.2; margin-top: 10px;">
                    <span>{fallback_str} (Local)</span><br>
                    <span style="font-size: 0.8em;">Erro Fuso Horário: {e}</span>
                </div>
                """, unsafe_allow_html=True)

        atualizar_data() # Atualiza na carga inicial

    st.markdown("---") # Divisor visual

    # --- REMOVIDA A SEÇÃO DE INFORMAÇÕES GERAIS ---

    # --- Inicialização do Estado da Sessão (Ajustada) ---
    if 'dados_laudo' not in st.session_state:
        st.session_state.dados_laudo = {
            'itens': [],
            'imagem': None
        }
    if 'itens' not in st.session_state.dados_laudo:
        st.session_state.dados_laudo['itens'] = []
    if 'imagem' not in st.session_state.dados_laudo:
        st.session_state.dados_laudo['imagem'] = None
    if not isinstance(st.session_state.dados_laudo.get('itens'), list):
         st.session_state.dados_laudo['itens'] = []

    # --- Coleta de Dados para o Laudo (Foco nos Itens) ---
    st.header("MATERIAL RECEBIDO PARA EXAME")

    numero_itens = st.number_input(
        "Número de tipos diferentes de material/acondicionamento a descrever",
        min_value=0,
        value=max(0, len(st.session_state.dados_laudo.get('itens', []))),
        step=1,
        key="num_itens_input",
        help="Informe quantos grupos distintos de material (com mesma embalagem, cor, etc.) você recebeu. Ex: 5 eppendorfs azuis contendo pó = 1 item; 3 porções em plástico transparente = 1 item."
    )

    # --- Lógica para adicionar/remover itens no estado da sessão ---
    if not isinstance(st.session_state.dados_laudo.get('itens'), list):
        st.session_state.dados_laudo['itens'] = []
    current_num_itens_in_state = len(st.session_state.dados_laudo['itens'])
    if numero_itens > current_num_itens_in_state:
        for _ in range(numero_itens - current_num_itens_in_state):
            st.session_state.dados_laudo['itens'].append({
                'qtd': 1, 'tipo_mat': list(TIPOS_MATERIAL_BASE.keys())[0],
                'emb': list(TIPOS_EMBALAGEM_BASE.keys())[0], 'cor_emb': None,
                'ref': '', 'pessoa': ''
            })
    elif numero_itens < current_num_itens_in_state:
        st.session_state.dados_laudo['itens'] = st.session_state.dados_laudo['itens'][:numero_itens]

    # --- Loop para exibir campos de cada item ---
    if numero_itens > 0:
        st.markdown("---")
        for i in range(numero_itens):
            with st.expander(f"Detalhes do Item {i + 1}", expanded=True):
                item_key_prefix = f"item_{i}_"
                cols_item1 = st.columns([1, 3, 3])
                with cols_item1[0]:
                    if not isinstance(st.session_state.dados_laudo['itens'][i].get('qtd'), int):
                        st.session_state.dados_laudo['itens'][i]['qtd'] = 1
                    st.session_state.dados_laudo['itens'][i]['qtd'] = st.number_input(
                        "Qtd", min_value=1,
                        value=st.session_state.dados_laudo['itens'][i]['qtd'],
                        step=1, key=item_key_prefix + "qtd",
                        help="Número de unidades deste item (ex: 5 eppendorfs)")
                with cols_item1[1]:
                    if st.session_state.dados_laudo['itens'][i].get('tipo_mat') not in TIPOS_MATERIAL_BASE:
                         st.session_state.dados_laudo['itens'][i]['tipo_mat'] = list(TIPOS_MATERIAL_BASE.keys())[0]
                    st.session_state.dados_laudo['itens'][i]['tipo_mat'] = st.selectbox(
                        "Material", options=list(TIPOS_MATERIAL_BASE.keys()),
                        format_func=lambda x: f"{x.upper()} ({TIPOS_MATERIAL_BASE.get(x, '?')})",
                        index=list(TIPOS_MATERIAL_BASE.keys()).index(st.session_state.dados_laudo['itens'][i]['tipo_mat']),
                        key=item_key_prefix + "tipo_mat",
                        help="Selecione o aspecto principal do material.")
                with cols_item1[2]:
                    if st.session_state.dados_laudo['itens'][i].get('emb') not in TIPOS_EMBALAGEM_BASE:
                         st.session_state.dados_laudo['itens'][i]['emb'] = list(TIPOS_EMBALAGEM_BASE.keys())[0]
                    st.session_state.dados_laudo['itens'][i]['emb'] = st.selectbox(
                        "Embalagem", options=list(TIPOS_EMBALAGEM_BASE.keys()),
                        format_func=lambda x: f"{x.upper()} ({TIPOS_EMBALAGEM_BASE.get(x, '?')})",
                        index=list(TIPOS_EMBALAGEM_BASE.keys()).index(st.session_state.dados_laudo['itens'][i]['emb']),
                        key=item_key_prefix + "emb",
                        help="Selecione o tipo de acondicionamento primário.")

                cols_item2 = st.columns([2, 2, 3])
                with cols_item2[0]:
                    embalagem_selecionada = st.session_state.dados_laudo['itens'][i]['emb']
                    if embalagem_selecionada in ['pl', 'pa', 'e', 'z']:
                        if 'cor_emb' not in st.session_state.dados_laudo['itens'][i]:
                             st.session_state.dados_laudo['itens'][i]['cor_emb'] = None
                        opcoes_cor = {None: " - Selecione - "}
                        opcoes_cor.update({k: v.capitalize() for k, v in CORES_FEMININO_EMBALAGEM.items()})
                        current_cor_key = st.session_state.dados_laudo['itens'][i]['cor_emb']
                        try: cor_index = list(opcoes_cor.keys()).index(current_cor_key)
                        except ValueError: cor_index = 0
                        st.session_state.dados_laudo['itens'][i]['cor_emb'] = st.selectbox(
                            "Cor Embalagem", options=list(opcoes_cor.keys()),
                            format_func=lambda x: opcoes_cor[x], index=cor_index,
                            key=item_key_prefix + "cor_emb",
                            help="Selecione a cor da embalagem, se houver e for relevante."
                        )
                    else:
                        st.text_input("Cor Embalagem", value="N/A", key=item_key_prefix + "cor_emb_disabled", disabled=True, help="Cor não aplicável para este tipo de embalagem.")
                        st.session_state.dados_laudo['itens'][i]['cor_emb'] = None
                with cols_item2[1]:
                    if 'ref' not in st.session_state.dados_laudo['itens'][i]: st.session_state.dados_laudo['itens'][i]['ref'] = ''
                    st.session_state.dados_laudo['itens'][i]['ref'] = st.text_input(
                        "Ref. Constatação", value=st.session_state.dados_laudo['itens'][i]['ref'],
                        key=item_key_prefix + "ref",
                        help="Informe o número do subitem correspondente no Laudo de Constatação, se houver (ex: 1.1, 2.3).")
                with cols_item2[2]:
                    if 'pessoa' not in st.session_state.dados_laudo['itens'][i]: st.session_state.dados_laudo['itens'][i]['pessoa'] = ''
                    st.session_state.dados_laudo['itens'][i]['pessoa'] = st.text_input(
                        "Pessoa Relacionada", value=st.session_state.dados_laudo['itens'][i]['pessoa'],
                        key=item_key_prefix + "pessoa",
                        help="(Opcional) Nome da pessoa a quem este material estava associado, se informado.")
                st.markdown("---", unsafe_allow_html=False)

    st.markdown("---")

    # --- Upload de Imagem ---
    st.header("Ilustração (Opcional)")
    uploaded_image = st.file_uploader(
        "Carregar imagem do(s) material(is) recebido(s)",
        type=["png", "jpg", "jpeg", "bmp", "gif"],
        key="image_uploader",
        help="Faça o upload de uma imagem que mostre os materiais recebidos. Será incluída na Seção 1."
        )
    if uploaded_image is not None:
        st.session_state.dados_laudo['imagem'] = uploaded_image
    else:
        if 'image_uploader' in st.session_state and st.session_state.image_uploader is None:
             st.session_state.dados_laudo['imagem'] = None

    # --- Botão de Geração e Download ---
    st.markdown("---")
    st.header("Gerar e Baixar Laudo")

    if st.button("📊 Gerar Laudo (.docx)"):
        with st.spinner("Gerando documento... Por favor, aguarde."):
            try:
                document = gerar_laudo_docx(st.session_state.dados_laudo)
                doc_io = io.BytesIO()
                document.save(doc_io)
                doc_io.seek(0)
                now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_name = f"Laudo_Drogas_{now_str}.docx"
                st.download_button(
                    label="✅ Download do Laudo Concluído!", data=doc_io,
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
