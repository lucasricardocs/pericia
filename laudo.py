# --- Interface Streamlit ---
    import streamlit as st
    import re
    from datetime import datetime
    import docx
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

        # --- MOVIDO: Data/Calendário (Acima da logo/título) ---
        data_placeholder = st.empty()
        def atualizar_data():
            try:
                brasilia_tz = timezone('America/Sao_Paulo')
                now = datetime.now(brasilia_tz)
                dias_semana_portugues = {
                    0: "Segunda-feira", 1: "Terça-feira", 2: "Quarta-feira",
                    3: "Quinta-feira", 4: "Sexta-feira", 5: "Sábado", 6: "Domingo"
                }
                meses_portugues = {
                    1: "janeiro", 2: "fevereiro", 3: "março",
                    4: "abril", 5: "maio", 6: "junho", 7: "julho",
                    8: "agosto", 9: "setembro", 10: "outubro",
                    11: "novembro", 12: "dezembro"
                }
                dia_semana = dias_semana_portugues.get(now.weekday(), '')
                mes = meses_portugues.get(now.month, '')
                data_formatada = f"{dia_semana}, {now.day} de {mes} de {now.year}"
                # Adiciona um pouco de margem inferior para separar da linha seguinte
                data_placeholder.markdown(f"""
        <div style="text-align: right; font-size: 0.8em; color: {UI_COR_CINZA_SPTC}; line-height: 1.2; margin-bottom: 15px;">
            <span>{data_formatada}</span><br>
            <span style="font-size: 0.8em;">(Goiânia-GO)</span>
        </div>""", unsafe_allow_html=True)
            except Exception as e:
                now = datetime.now()
                fallback_str = now.strftime("%d/%m/%Y")
                data_placeholder.markdown(f"""
        <div style="text-align: right; font-size: 0.8em; color: #FF5555; line-height: 1.2; margin-bottom: 15px;">
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

            # 2 MATERIAL RECEBIDO PARA EXAME (Ilustração 1)
            add_paragraph(document, "2 MATERIAL RECEBIDO PARA EXAME", bold=True)

            if uploaded_image is not None:
                try:
                    image = Image.open(uploaded_image)
                    # Adiciona a imagem ao documento, ajustando a largura (em polegadas)
                    document.add_picture(uploaded_image, width=Inches(5.0))
                    add_paragraph(document, "Ilustração 1 – Material recebido para exame.", bold=True, align='center', size=10)
                except Exception as e:
                    st.error(f"Erro ao processar a imagem: {e}")
                    add_paragraph(document, "(Imagem do material recebido não pôde ser carregada)", bold=True, align='center', size=10)
            else:
                add_paragraph(document, "\nIlustração 1 – Material recebido para exame.", bold=True)

            tipos_material_itens_codigo = []
            subitens_cannabis = {}
            subitens_cocaina = {}

            itens_data_final = []
            for i in range(int(st.session_state.get("num_itens", 1))):
                qtd = st.session_state.get(f"qtd_{i}")
                tipo_mat_code = st.session_state.get(f"tipo_mat_{i}")
                emb_code = st.session_state.get(f"emb_{i}")
                cor_emb_code = st.session_state.get(f"cor_emb_{i}")
                ref = st.session_state.get(f"ref_{i}")
                pessoa = st.session_state.get(f"pessoa_{i}")

                itens_data_final.append({
                    'qtd': qtd,
                    'tipo_mat_code': tipo_mat_code,
                    'emb_code': emb_code,
                    'cor_emb_code': cor_emb_code,
                    'ref': ref,
                    'pessoa': pessoa
                })

            for i, item_info in enumerate(itens_data_final):
                qtd = item_info['qtd']
                qtd_ext = obter_quantidade_extenso(qtd)
                tipo_mat_code = item_info['tipo_mat_code']
                emb_code = item_info['emb_code']
                cor_emb_code = item_info['cor_emb_code']
                ref = item_info['ref']
                pessoa = item_info['pessoa']

                tipo_material = TIPOS_MATERIAL_BASE.get(tipo_mat_code, tipo_mat_code)
                embalagem = TIPOS_EMBALAGEM_BASE.get(emb_code, emb_code)

                if cor_emb_code:
                    cor = CORES_FEMININO_EMBALAGEM.get(cor_emb_code, cor_emb_code)
                    embalagem += f" de cor {cor}"

                embalagem = pluralizar_palavra(embalagem, qtd)
                porcao = pluralizar_palavra("porção", qtd)
                acond = "acondicionada em" if qtd == 1 else "acondicionadas, individualmente, em"
                ref_texto = f", relacionada a {pessoa}" if pessoa else ""
                final_ponto = "."

                texto = f"2.{i+1} {qtd} ({qtd_ext}) {porcao} de material {tipo_material}, {acond} {embalagem}, referente à amostra do subitem {ref} do laudo de constatação supracitado{ref_texto}{final_ponto}"
                add_paragraph(document, texto)

                tipos_material_itens_codigo.append(tipo_mat_code)
                if tipo_mat_code in ["v", "r"]:
                    subitens_cannabis[ref] = f"2.{i+1}"
                elif tipo_mat_code in ["po", "pd"]:
                    subitens_cocaina[ref] = f"2.{i+1}"

            # 3 OBJETIVO DOS EXAMES
            add_paragraph(document, "\n3 OBJETIVO DOS EXAMES", bold=True)
            add_paragraph(document, "Visa esclarecer à autoridade requisitante quanto às características do material apresentado, bem como se ele contém substância de uso proscrito no Brasil e capaz de causar dependência física e/ou psíquica. O presente laudo pericial busca demonstrar a materialidade da infração penal apurada.", align='justify')

            # 4 EXAMES
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

            # 5 RESULTADOS
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

            # 6 CONCLUSÃO
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

            # 7 CUSTÓDIA DO MATERIAL
            add_paragraph(document, "\n7 CUSTÓDIA DO MATERIAL", bold=True)
            add_paragraph(document, "7.1 Contraprova")
            add_paragraph(document, f"7.1.1 A amostra contraprova ficará armazenada neste Instituto, conforme Portaria 0003/2019/SSP  (Lacre nº {lacre}).")

            # REFERÊNCIAS
            add_paragraph(document, "\nREFERÊNCIAS", bold=True)
            referencias = [
                "BRASIL. Ministério da Saúde. Portaria SVS/MS n° 344, de 12 de maio de 1998. Aprova o regulamento técnico sobre substâncias e medicamentos sujeitos a controle especial. Diário Oficial da União: Brasília, DF, p. 37, 19 maio 1998. Alterada pela RDC nº 970, de 19/03/2025.",
                "GOIÁS. Secretaria de Estado da Segurança Pública. Portaria nº 0003/2019/SSP de 10 de janeiro de 2019. Regulamenta a apreensão, movimentação, exames, acondicionamento, armazenamento e destruição de drogas no âmbito da Secretaria de Estado da Segurança Pública. Diário Oficial do Estado de Goiás: n° 22.972, Goiânia, GO, p. 4-5, 15 jan. 2019.",
                "SWGDRUG: Scientific Working Group for the Analysis of Seized Drugs. Recommendations. Version 8.0 june. 2019. Disponível em: http://www.swgdrug.org/Documents/SWGDRUG%20Recommendations%20Version%208_FINAL_ForPosting_092919.pdf. Acesso em: 07/10/2019."
            ]
            if has_cannabis_item:
                referencias.append("UNODC (United Nations Office on Drugs and Crime). Laboratory and scientific section. Recommended Methods for the Identification and Analysis of Cannabis and Cannabis Products. New York: 2012.")
            if has_cocaina_item:
                referencias.append("UNODC (United Nations Office on Drugs and Crime). Laboratory and Scientific Section. Recommended Methods for the Identification and Analysis of Cocaine in Seized Materials. New York: 2012.")
            for ref in referencias:
                add_paragraph(document, ref)

            brasilia_tz = timezone('America/Sao_Paulo')
            hoje = datetime.now(brasilia_tz)
            data_formatada = f"Goiânia, {hoje.day} de {meses_portugues_dict[hoje.strftime('%B')]} de {hoje.year}."
            add_paragraph(document, data_formatada, align='right')

            add_paragraph(document, "\nLaudo assinado digitalmente com dados do assinador à esquerda das páginas", align='left')
            add_paragraph(document, "Daniel Chendes Lima", align='center')
            add_paragraph(document, "Perito Criminal", align='center')

            # Aplicar fonte Gadugi a todo o documento e itálico apenas nas expressões específicas
            italics = [
                'Cannabis sativa',
                'Scientific Working Group for the Analysis of Seized Drugs',
                'United Nations Office on Drugs and Crime',
                'Fast blue salt B',
                'eppendorf',
                'ziplock'
            ]

            for paragraph in document.paragraphs:
                full_text = paragraph.text
                is_ilustracao = "Ilustração 1 – Material recebido para exame." in full_text
                paragraph.clear()
                idx = 0
                while idx < len(full_text):
                    match_found = False
                    for phrase in italics:
                        if full_text[idx:].startswith(phrase):
                            run = paragraph.add_run(phrase)
                            run.font.name = 'Gadugi'
                            run.font.size = Pt(10) if is_ilustracao else Pt(12)
                            run.italic = True
                            idx += len(phrase)
                            match_found = True
                            break
                    if not match_found:
                        run = paragraph.add_run(full_text[idx])
                        run.font.name = 'Gadugi'
                        run.font.size = Pt(10) if is_ilustracao else Pt(12)
                        idx += 1

            file_stream = io.BytesIO()
            document.save(file_stream)
            file_stream.seek(0)

            st.download_button(
                label="✅ Gerar Laudo",
                data=file_stream,
                file_name=f"{numero_laudo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Laudo gerado com sucesso!")

    if __name__ == "__main__":
        main()
