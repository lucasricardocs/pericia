
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
