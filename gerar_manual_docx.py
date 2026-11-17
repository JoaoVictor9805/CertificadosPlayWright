# -*- coding: utf-8 -*-
"""
Script para geração do Manual do Usuário em formato Word (.docx) via python-docx.
Mantém o mesmo padrão visual e estrutural do manual original em PDF, atualizado
com as funcionalidades implementadas (Coluna Nome da Empresa, consulta do próprio
certificado em Meu Perfil, enquadramento da Caixa Postal, Matriz/Filiais e print
único quando a Caixa Postal estiver vazia).
"""

from pathlib import Path
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor


COR_AZUL_PRIMARIO = RGBColor(0x1F, 0x4E, 0x78)   # #1F4E78
COR_AZUL_SECUNDARIO = RGBColor(0x2B, 0x5B, 0x84)  # #2B5B84
COR_SUBTITULO = RGBColor(0x5A, 0x6A, 0x80)        # #5A6A80
COR_TEXTO = RGBColor(0x2D, 0x37, 0x48)            # #2D3748
COR_NOTA_TEXTO = RGBColor(0x4A, 0x55, 0x68)       # #4A5568
COR_RODAPE = RGBColor(0x71, 0x80, 0x96)           # #718096
COR_BRANCO = RGBColor(0xFF, 0xFF, 0xFF)


def _set_cell_shading(cell, fill_hex: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def _set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_cell_borders(cell, **edges):
    """
    edges: dict com chaves top/bottom/left/right -> tuple(sz, color_hex) ou None para 'none'.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        if edge in edges:
            spec = edges[edge]
            el = OxmlElement(f"w:{edge}")
            if spec is None:
                el.set(qn("w:val"), "none")
            else:
                sz, color = spec
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(sz))
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), color)
            tcBorders.append(el)
    tcPr.append(tcBorders)


def _set_paragraph_bottom_border(paragraph, color_hex: str, sz: int = 8):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _set_paragraph_top_border(paragraph, color_hex: str, sz: int = 6):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), str(sz))
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), color_hex)
    pBdr.append(top)
    pPr.append(pBdr)


def _add_field(run, field_name: str):
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> %s </w:instrText>' % (nsdecls("w"), field_name))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w"))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def _add_run(p, text: str, *, bold=False, italic=False, size_pt=10, color=COR_TEXTO, font_name="Calibri"):
    r = p.add_run(text)
    r.font.name = font_name
    r.font.size = Pt(size_pt)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return r


def _add_heading_h2(doc: Document, texto: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    _add_run(p, texto, bold=True, size_pt=13, color=COR_AZUL_PRIMARIO)
    _set_paragraph_bottom_border(p, "E2E8F0", sz=6)
    return p


def _add_heading_h3(doc: Document, texto: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    _add_run(p, texto, bold=True, size_pt=11, color=COR_AZUL_SECUNDARIO)
    return p


def _append_formatted_runs(p, parts, size_pt=10, default_color=COR_TEXTO):
    """Adiciona trechos de texto simples ou formatados (texto, bold, italic, color) ao parágrafo."""
    if isinstance(parts, str):
        _add_run(p, parts, size_pt=size_pt, color=default_color)
        return
    for item in parts:
        if isinstance(item, str):
            _add_run(p, item, size_pt=size_pt, color=default_color)
        else:
            txt = item[0]
            b = item[1] if len(item) > 1 else False
            it = item[2] if len(item) > 2 else False
            col = item[3] if len(item) > 3 else default_color
            _add_run(p, txt, bold=b, italic=it, size_pt=size_pt, color=col)


def _add_body_p(doc: Document, parts, space_after=6, justify=True):
    """Adiciona parágrafo de corpo com suporte a trechos formatados."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _append_formatted_runs(p, parts, size_pt=10, default_color=COR_TEXTO)
    return p


def _add_bullet(doc: Document, parts, space_after=3):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.22
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _append_formatted_runs(p, parts, size_pt=10, default_color=COR_TEXTO)
    return p


def _add_action_badge(doc: Document, nome_arquivo: str):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    c_icon, c_badge = tbl.rows[0].cells

    c_icon.width = Cm(0.8)
    _set_cell_borders(c_icon, top=None, bottom=None, left=None, right=None)
    _set_cell_margins(c_icon, top=40, bottom=40, left=0, right=40)
    p_icon = c_icon.paragraphs[0]
    p_icon.paragraph_format.space_before = Pt(0)
    p_icon.paragraph_format.space_after = Pt(0)
    _add_run(p_icon, "👉", size_pt=11)

    c_badge.width = Cm(5.8)
    _set_cell_shading(c_badge, "EBF3FA")
    _set_cell_borders(
        c_badge,
        top=(6, "B0C8DE"),
        bottom=(6, "B0C8DE"),
        left=(6, "B0C8DE"),
        right=(6, "B0C8DE"),
    )
    _set_cell_margins(c_badge, top=60, bottom=60, left=140, right=140)
    p_badge = c_badge.paragraphs[0]
    p_badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_badge.paragraph_format.space_before = Pt(0)
    p_badge.paragraph_format.space_after = Pt(0)
    _add_run(p_badge, nome_arquivo, bold=True, size_pt=10, color=COR_AZUL_PRIMARIO)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(4)


def _add_inline_note(doc: Document, titulo: str, corpo_parts):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    cell.width = Cm(17.4)
    _set_cell_shading(cell, "F8FAFC")
    _set_cell_borders(
        cell,
        top=None,
        bottom=None,
        right=None,
        left=(24, "1F4E78"),
    )
    _set_cell_margins(cell, top=100, bottom=100, left=170, right=150)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.25
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _add_run(p, f"{titulo} ", bold=True, size_pt=9.8, color=COR_AZUL_PRIMARIO)
    _append_formatted_runs(p, corpo_parts, size_pt=9.8, default_color=COR_NOTA_TEXTO)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(4)


def _add_option_box(doc: Document, titulo: str, opcoes):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    cell.width = Cm(17.4)
    _set_cell_shading(cell, "F8FAFC")
    _set_cell_borders(
        cell,
        top=(6, "CBD5E0"),
        bottom=(6, "CBD5E0"),
        left=(6, "CBD5E0"),
        right=(6, "CBD5E0"),
    )
    _set_cell_margins(cell, top=110, bottom=110, left=160, right=160)

    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(4)
    _add_run(p0, titulo, bold=True, size_pt=9.8, color=COR_AZUL_PRIMARIO)

    for idx, opt_parts in enumerate(opcoes):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2 if idx < len(opcoes) - 1 else 0)
        p.paragraph_format.line_spacing = 1.22
        _append_formatted_runs(p, opt_parts, size_pt=9.8, default_color=COR_TEXTO)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(4)


def _configure_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    _set_paragraph_top_border(p, "E2E8F0", sz=6)

    # Tab stop alinhado à direita (largura útil: 21.0 - 1.8 - 1.8 = 17.4 cm)
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(17.4), WD_ALIGN_PARAGRAPH.RIGHT)

    _add_run(p, "Robô Societário – DET / Gov.br\t", size_pt=8.5, color=COR_RODAPE)
    _add_run(p, "Página\u00A0", size_pt=8.5, color=COR_RODAPE)
    r_pg = p.add_run()
    r_pg.font.name = "Calibri"
    r_pg.font.size = Pt(8.5)
    r_pg.font.color.rgb = COR_RODAPE
    _add_field(r_pg, "PAGE")

    _add_run(p, "\u00A0de\u00A0", size_pt=8.5, color=COR_RODAPE)
    r_num = p.add_run()
    r_num.font.name = "Calibri"
    r_num.font.size = Pt(8.5)
    r_num.font.color.rgb = COR_RODAPE
    _add_field(r_num, "NUMPAGES")


def gerar_docx(caminho_saida: Path):
    caminho_saida = Path(caminho_saida).resolve()
    doc = Document()

    # Configuração A4 com as mesmas margens do PDF (14mm sup/inf, 18mm esq/dir)
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.footer_distance = Cm(0.8)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = COR_TEXTO

    _configure_footer(section)

    # ==========================================================================
    # CABEÇALHO DO DOCUMENTO
    # ==========================================================================
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(4)
    _add_run(p_title, "ROBÔ SOCIETÁRIO – DET / GOV.BR", bold=True, size_pt=20, color=COR_AZUL_PRIMARIO)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(10)
    _add_run(p_sub, "Manual de Operação e Guia Prático para o Usuário Final", size_pt=11, color=COR_SUBTITULO)
    _set_paragraph_bottom_border(p_sub, "CBD5E0", sz=10)

    # ==========================================================================
    # SEÇÃO 1
    # ==========================================================================
    _add_heading_h2(doc, "1. Apresentação e Objetivo do Sistema")
    _add_body_p(
        doc,
        "O Robô Societário automatiza as consultas em lote no portal do Domicílio Eletrônico Trabalhista "
        "(DET / Gov.br), eliminando o trabalho manual de acessar cada empresa individualmente.",
    )
    _add_body_p(doc, "O sistema realiza as seguintes funções de forma autônoma:")
    _add_bullet(
        doc,
        "Acesso via Gov.br utilizando certificado digital — seja com o próprio certificado da empresa "
        "titular (Meu Perfil) ou com procuração eletrônica vinculada ao CNPJ consultado (Procurador);",
    )
    _add_bullet(
        doc,
        "Verificação da Caixa Postal em busca de mensagens e avisos pendentes ou já lidos "
        "(incluindo alternância automática entre Matriz e Filiais quando disponível);",
    )
    _add_bullet(
        doc,
        "Localização de novas notificações e download automático dos termos e anexos oficiais (PDFs ou ZIPs);",
    )
    _add_bullet(
        doc,
        "Registro de capturas de tela (prints) de todas as etapas para auditoria e evidência jurídica, "
        "organizadas por empresa;",
    )
    _add_bullet(
        doc,
        "Emissão de relatório geral consolidado em Excel com o status detalhado de cada empresa consultada.",
    )

    # ==========================================================================
    # SEÇÃO 2 - PASSOS 1 E 2
    # ==========================================================================
    _add_heading_h2(doc, "2. Guia Passo a Passo de Utilização")

    _add_heading_h3(doc, "Passo 1: Instalação dos Certificados Digitais na Máquina")
    _add_body_p(
        doc,
        "Antes de iniciar as consultas, certifique-se de que os certificados digitais (modelo A1) necessários "
        "estejam devidamente instalados no seu perfil de usuário ativo no Windows. O robô reconhece e utiliza "
        "os certificados já presentes na máquina para realizar as autenticações de forma transparente.",
    )

    _add_heading_h3(doc, "Passo 2: Preparação Inicial dos Certificados (Atualizar_Certificados.bat)")
    _add_body_p(
        doc,
        [
            ("Os Passos 2 e 3 formam a ", False),
            ("configuração inicial do sistema", True),
            (" e são realizados apenas ", False),
            ("antes da primeira execução", True),
            (" (você só precisará repetir este Passo 2 se instalar ou renovar algum certificado digital no computador). Dê dois cliques no arquivo utilitário:", False),
        ],
    )
    _add_action_badge(doc, "Atualizar_Certificados.bat")
    _add_body_p(
        doc,
        [
            ("Esse assistente lê os certificados digitais instalados no Windows e atualiza automaticamente a aba ", False),
            ("Certificados_Disponiveis", True),
            (" dentro da planilha ", False),
            ("PlanilhasBase/CNPJs.xlsx", True),
            (" (que é a planilha de configurações do robô). Se o arquivo ainda não existir na pasta, ele será criado automaticamente.", False),
        ],
    )
    _add_inline_note(
        doc,
        "💡 Antes de clicar:",
        [
            ("Certifique-se de que a planilha ", False),
            ("CNPJs.xlsx", False, True),
            (" esteja fechada no Excel para que o assistente consiga gravar os certificados sem bloqueio do Windows.", False),
        ],
    )

    # ==========================================================================
    # PASSO 3 (Configuração da Planilha com as 3 colunas)
    # ==========================================================================
    _add_heading_h3(doc, "Passo 3: Configuração da Planilha de Consultas (PlanilhasBase/CNPJs.xlsx)")
    _add_body_p(
        doc,
        [
            ("A planilha ", False),
            ("CNPJs.xlsx", True),
            (" (dentro da pasta ", False),
            ("PlanilhasBase", True),
            (") é a ", False),
            ("planilha de configurações do Robô Societário", True),
            (". Tudo o que você configurar aqui fica salvo para as próximas execuções — ou seja, ", False),
            ("todo esse preenchimento é feito somente antes da primeira vez que você rodar o robô", True),
            (", sendo necessário editar novamente apenas quando quiser incluir, alterar ou remover alguma empresa.", False),
        ],
    )
    _add_body_p(doc, "Abra o arquivo PlanilhasBase/CNPJs.xlsx no Excel e siga esta ordem:")
    _add_bullet(
        doc,
        [
            ("1º — Consulte a aba 'Certificados_Disponiveis': ", True),
            ("Nela está a lista pronta dos certificados da sua máquina (gerada no Passo 2). Você usará essa aba apenas para copiar o nome exato de cada certificado.", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("2º — Preencha a aba 'CNPJs_consultados': ", True),
            ("É nesta aba que você cadastra as empresas que o robô deve consultar (uma empresa por linha), preenchendo as 3 colunas abaixo:", False),
        ],
    )

    # Tabela das 3 colunas
    tbl = doc.add_table(rows=2, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    larguras = [Cm(4.5), Cm(6.7), Cm(6.2)]
    cabecalhos = [
        "Coluna A: CNPJ",
        "Coluna B: Certificado",
        "Coluna C: Nome da Empresa",
    ]
    descricoes = [
        [
            ("Digite os 14 dígitos do CNPJ da empresa que será consultada.\n", False),
            ("Exemplo: ", True),
            ("39407892000185", False, True),
        ],
        [
            ("Copie (Ctrl+C) na aba ", False),
            ("Certificados_Disponiveis", False, True),
            (" (coluna ", False),
            ("Nome Completo (certmgr)", False, True),
            (") e cole (Ctrl+V) o certificado que tem acesso àquela empresa.", False),
        ],
        [
            ("Digite o nome da empresa (opcional, porém recomendado). Ele aparece no relatório final e dá nome à pasta dos prints; se ficar vazio, a pasta usa o CNPJ.", False),
        ],
    ]

    for idx, titulo_col in enumerate(cabecalhos):
        cell = tbl.rows[0].cells[idx]
        cell.width = larguras[idx]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_shading(cell, "1F4E78")
        _set_cell_borders(
            cell,
            top=(6, "1F4E78"),
            bottom=(6, "1F4E78"),
            left=(6, "1F4E78"),
            right=(6, "1F4E78"),
        )
        _set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        _add_run(p, titulo_col, bold=True, size_pt=10, color=COR_BRANCO)

    for idx, desc_parts in enumerate(descricoes):
        cell = tbl.rows[1].cells[idx]
        cell.width = larguras[idx]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _set_cell_shading(cell, "FFFFFF")
        _set_cell_borders(
            cell,
            top=(6, "CBD5E0"),
            bottom=(6, "CBD5E0"),
            left=(6, "CBD5E0"),
            right=(6, "CBD5E0"),
        )
        _set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.2
        _append_formatted_runs(p, desc_parts, size_pt=9.5, default_color=COR_TEXTO)

    spacer_tbl = doc.add_paragraph()
    spacer_tbl.paragraph_format.space_before = Pt(0)
    spacer_tbl.paragraph_format.space_after = Pt(4)

    _add_inline_note(
        doc,
        "💡 Finalizando a Configuração:",
        [
            ("Depois de preencher todas as linhas na aba ", False),
            ("CNPJs_consultados", False, True),
            (" (sem deixar o certificado escrito diferente da aba ", False),
            ("Certificados_Disponiveis", False, True),
            ("), ", False),
            ("salve a planilha e feche o Excel", True),
            (". A partir daqui, o robô já está configurado e pronto para uso — nas próximas vezes, basta abrir direto o ", False),
            ("RoboSocietario.exe", True),
            (" no Passo 4.", False),
        ],
    )
    _add_inline_note(
        doc,
        "💡 Empresa usando o Próprio Certificado (sem procuração):",
        [
            ("Se você cadastrar na Coluna A o mesmo CNPJ do certificado informado na Coluna B (por exemplo, consultando a própria empresa dona do certificado), o robô reconhece isso sozinho, coloca essa empresa como a ", False),
            ("primeira da fila", True),
            (" daquele certificado e acessa diretamente o perfil próprio (", False),
            ("Meu Perfil", False, True),
            ("), seguindo normalmente como Procurador para as demais empresas ligadas ao mesmo certificado.", False),
        ],
    )

    # ==========================================================================
    # PASSO 4
    # ==========================================================================
    _add_heading_h3(doc, "Passo 4: Executando a Automação no Dia a Dia (RoboSocietario.exe)")
    _add_body_p(
        doc,
        [
            ("Com a planilha de configurações (", False),
            ("PlanilhasBase/CNPJs.xlsx", False, True),
            (") salva e fechada, você já pode iniciar as consultas sempre que precisar dando dois cliques no aplicativo principal:", False),
        ],
    )
    _add_action_badge(doc, "RoboSocietario.exe")
    _add_body_p(
        doc,
        "Ao abrir o programa, uma janela preta de acompanhamento (Console) mostrará em tempo real o andamento "
        "da fila, o CNPJ em consulta e cada etapa concluída. O robô abrirá o Google Chrome sozinho e fará todo "
        "o processo sem necessidade de intervenção manual.",
    )
    _add_inline_note(
        doc,
        "💡 Cuidados Durante a Execução:",
        [
            ("Enquanto o robô estiver abrindo o Chrome e confirmando o certificado digital na janela do Windows, ", False),
            ("evite mexer no mouse ou no teclado", True),
            (" para não tirar o foco da tela. Também não feche o navegador aberto pelo robô nem a janela do Console antes da mensagem final de conclusão.", False),
        ],
    )
    _add_body_p(
        doc,
        [
            ("Se houver pendências da execução anterior: ", True),
            ("antes de iniciar as consultas, o robô perguntará na tela o que você deseja fazer:", False),
        ],
    )
    _add_option_box(
        doc,
        "Opções de Execução Disponíveis:",
        [
            [
                ("• ", False),
                ("[Opção 1 - ENTER]: ", True),
                ("Consultar todas as empresas cadastradas na planilha de configurações (", False),
                ("PlanilhasBase/CNPJs.xlsx", False, True),
                (").", False),
            ],
            [
                ("• ", False),
                ("[Opção 2]: ", True),
                ("Reprocessar somente as empresas que ficaram pendentes na última tentativa (", False),
                ("PlanilhasBase/CNPJs_Pendentes.xlsx", False, True),
                (").", False),
            ],
        ],
    )
    _add_body_p(
        doc,
        [
            ("Proteção automática contra falhas do portal: ", True),
            ("como sites governamentais podem oscilar, você não precisa refazer manualmente o que falhar na primeira tentativa:", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("1. Repetição automática (Segunda passagem): ", True),
            ("Se alguma empresa apresentar instabilidade durante a execução, o próprio robô tenta consultá-la novamente ao final da rodada.", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("2. Registro de pendências (CNPJs_Pendentes.xlsx): ", True),
            ("Somente se o erro persistir (ou se o Gov.br exigir validação em duas etapas no celular), a empresa é salva em ", False),
            ("PlanilhasBase/CNPJs_Pendentes.xlsx", False, True),
            (" com CNPJ, Certificado e Nome da Empresa. Quando todas forem resolvidas em uma execução futura, esse arquivo de pendências é apagado automaticamente. ", False),
            ("Não é necessário fazer nenhuma modificação manual nesse arquivo", True),
            (" — seu intuito é exclusivamente orientar uma nova execução focada apenas no que ficou pendente da execução anterior.", False),
        ],
    )

    # ==========================================================================
    # PASSO 5
    # ==========================================================================
    _add_heading_h3(doc, "Passo 5: Conferência dos Resultados e Evidências (Pasta Resultados)")
    _add_body_p(
        doc,
        [
            ("Quando o robô terminar todas as empresas, ele exibirá o resumo final no Console e pedirá para você pressionar ENTER para fechar a janela. Tudo o que foi gerado fica salvo dentro da pasta ", False),
            ("Resultados", True),
            (", dividido em dois locais:", False),
        ],
    )

    _add_body_p(
        doc,
        [
            ("1. Pasta 'Resultados/Relatorios' (Resumo Geral em Excel):\n", True),
            ("É o primeiro lugar que você deve abrir ao final da rotina. Cada execução gera uma planilha com data e hora (exemplo: ", False),
            ("Relatorio_Execucao_30-10-2026_09-00-00.xlsx", False, True),
            (") mostrando CNPJ, ", False),
            ("Nome da Empresa", True),
            (", Certificado utilizado, detalhes da 1ª e 2ª tentativa, quantidade de arquivos salvos e o ", False),
            ("Status Final", True),
            (" colorido para leitura rápida:", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Verde (SUCESSO): ", True),
            ("Empresa consultada e evidências salvas normalmente.", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Amarelo (Necessita confirmação de identidade): ", True),
            ("O login Gov.br do certificado pediu código/aprovação no celular (2FA).", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Vermelho (FALHA / NAO_EXECUTADO / SEM_PROCURACAO): ", True),
            ("Erro persistente, certificado não encontrado na máquina ou ausência de procuração no DET.", False),
        ],
    )

    _add_body_p(
        doc,
        [
            ("2. Pasta 'Resultados/CapturasDeTela-Downloads' (Prints e Anexos por Empresa):\n", True),
            ("Guarda todas as provas visuais e documentos organizados em pastas seguindo a ordem: ", False),
            ("Data e Hora da Execução → Certificado → Nome da Empresa", True),
            (" (ou o CNPJ, caso a Coluna C tenha ficado em branco). Dentro da pasta de cada empresa você encontra:", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("CaixaPostal: ", True),
            ("prints da Caixa Postal enquadrados direto no título ", False),
            ("Caixa Postal.", False, True),
            (" Quando houver mensagens, são gerados 2 prints (todas e apenas não lidas); se a empresa possuir seleção entre ", False),
            ("Matriz e Filiais", True),
            (", o robô captura os prints dos dois perfis (", False),
            ("matriz ", False, True),
            ("e ", False),
            ("filiais", False, True),
            (").", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Notificacao: ", True),
            ("prints das notificações encontradas com data e hora visíveis (ou comprovando que não há notificações).", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("ArquivosBaixados: ", True),
            ("termos de notificação e anexos oficiais (PDFs ou ZIPs) baixados automaticamente, a partir da data do dia da execução.", False),
        ],
    )

    _add_inline_note(
        doc,
        "💡 O que fazer quando o Status ficar Amarelo (2FA)?",
        [
            ("Quando o Gov.br exigir verificação em duas etapas no celular, o robô salva automaticamente um print comprobatório na pasta do certificado, pula aquele grupo.", False),
        ],
    )

    _add_body_p(
        doc,
        [
            ("Rotina Automática Agendada (Sugerido: Sexta-feira às 08:00):\n", True, False, RGBColor(0x1F, 0x49, 0x7D)),
            ("O Robô Societário também pode operar de forma 100% autônoma toda sexta-feira às 08:00 pelo Agendador de Tarefas do Windows, sem precisar de cliques nem confirmações:", False),
        ],
        justify=False,
    )
    _add_bullet(
        doc,
        [
            ("1. Até quinta-feira: ", True),
            ("basta manter a planilha de configurações (", False),
            ("PlanilhasBase/CNPJs.xlsx", False, True),
            (") atualizada caso entre ou saia alguma empresa.", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("2. Na sexta-feira às 08:00: ", True),
            ("o robô inicia sozinho, executa todas as consultas, salva os prints/downloads e encerra ao final sem pedir ENTER.", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("3. Conferência: ", True),
            ("basta abrir a pasta ", False),
            ("Resultados/Relatorios", False, True),
            (" para visualizar o relatório consolidado da semana.", False),
        ],
    )

    # ==========================================================================
    # SEÇÃO 3 - AGENDADOR DE TAREFAS
    # ==========================================================================
    _add_heading_h2(doc, "3. Como Configurar o Agendamento Automático (Agendador de Tarefas do Windows)")
    _add_body_p(
        doc,
        [
            ("Para que o Robô Societário rode sozinho no dia e horário programados (por exemplo, toda sexta-feira às 08:00), siga o passo a passo abaixo no ", False),
            ("Agendador de Tarefas (Task Scheduler)", True),
            (" do Windows:", False),
        ],
    )

    _add_heading_h3(doc, "1. Abrir e criar a tarefa")
    _add_bullet(
        doc,
        [
            ("Pressione as teclas ", False),
            ("Windows + R", True),
            (", digite ", False),
            ("taskschd.msc", False, True),
            (" e pressione ", False),
            ("ENTER", True),
            (" (ou pesquise por ", False),
            ("Agendador de Tarefas", True),
            (" no Menu Iniciar).", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("No painel da direita, clique em ", False),
            ("Criar Tarefa...", True),
            (" ", False),
            ('(use "Criar Tarefa", e não "Criar Tarefa Básica", para ter acesso a todas as configurações necessárias).', False, True),
        ],
    )

    _add_heading_h3(doc, '2. Aba "Geral"')
    _add_bullet(
        doc,
        [
            ("Em ", False),
            ("Nome", True),
            (", digite um identificador, como: ", False),
            ("RoboSocietario - Consulta DET", False, True),
            (".", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Em ", False),
            ("Opções de segurança", True),
            (", deixe marcada a opção ", False),
            ("Executar somente quando o usuário estiver conectado", True),
            (".", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Marque ", False),
            ("Executar com privilégios mais altos", True),
            (" caso o usuário da máquina tenha restrições administrativas.", False),
        ],
    )
    _add_inline_note(
        doc,
        "⚠ Atenção:",
        [
            ("Não utilize a opção ", False),
            ('"Executar estando o usuário conectado ou não"', False, True),
            (", pois o robô precisa da interface gráfica aberta para interagir com o Google Chrome e com a janela de seleção do certificado digital do Windows.", False),
        ],
    )

    _add_heading_h3(doc, '3. Aba "Disparadores" (Dia e Horário)')
    _add_bullet(
        doc,
        [
            ("Clique em ", False),
            ("Novo...", True),
            (" e selecione ", False),
            ("Semanalmente", True),
            (".", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Defina ", False),
            ("Iniciar", True),
            (" com a data atual e o horário desejado (exemplo: ", False),
            ("08:00:00", False, True),
            (").", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Marque o dia da semana desejado (exemplo: ", False),
            ("Sexta-feira", True),
            ("), confirme que ", False),
            ("Habilitado", True),
            (" está marcado e clique em ", False),
            ("OK", True),
            (".", False),
        ],
    )

    _add_heading_h3(doc, '4. Aba "Ações" (Programa e Modo Silencioso)')
    _add_bullet(
        doc,
        [
            ("Clique em ", False),
            ("Novo...", True),
            (" e mantenha a ação ", False),
            ("Iniciar um programa", True),
            (".", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Em ", False),
            ("Programa/script", True),
            (", clique em ", False),
            ("Procurar...", True),
            (" e selecione o arquivo ", False),
            ("RoboSocietario.exe", False, True),
            (" dentro da pasta do sistema.", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Em ", False),
            ("Adicione argumentos (opcional)", True),
            (", digite exatamente: ", False),
            ("--agendado", True),
            (" ", False),
            ("(esse parâmetro desativa perguntas interativas e faz o robô encerrar sozinho ao terminar, sem aguardar você pressionar ENTER).", False, True),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Em ", False),
            ("Iniciar em (opcional)", True),
            (", informe o caminho da pasta onde o ", False),
            ("RoboSocietario.exe", False, True),
            (" está salvo (exemplo: ", False),
            (r"C:\RoboSocietario", False, True),
            ("), sem aspas, e clique em ", False),
            ("OK", True),
            (".", False),
        ],
    )

    _add_heading_h3(doc, '5. Abas "Condições" e "Configurações"')
    _add_bullet(
        doc,
        [
            ("Na aba ", False),
            ("Condições", True),
            (": caso use um notebook, desmarque ", False),
            ('"Iniciar a tarefa somente se o computador estiver conectado à energia CA"', False, True),
            (" se quiser permitir a execução mesmo fora da tomada.", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Na aba ", False),
            ("Configurações", True),
            (": mantenha marcada a opção ", False),
            ('"Permitir que a tarefa seja executada sob demanda"', False, True),
            (" e ajuste ", False),
            ('"Interromper a tarefa se ela for executada por mais de"', False, True),
            (" para um tempo compatível com a quantidade de empresas da sua planilha (exemplo: ", False),
            ("8 horas", False, True),
            (", ou desmarque para não interromper um lote grande).", False),
        ],
    )
    _add_bullet(
        doc,
        [
            ("Clique em ", False),
            ("OK", True),
            (" para salvar a tarefa.", False),
        ],
    )

    _add_heading_h3(doc, "6. Testar o agendamento")
    _add_bullet(
        doc,
        [
            ("Na lista de tarefas do Agendador, clique com o botão direito sobre ", False),
            ("RoboSocietario - Consulta DET", False, True),
            (" e escolha ", False),
            ("Executar", True),
            (".", False),
        ],
    )
    _add_bullet(
        doc,
        "Acompanhe para confirmar que o Chrome abre, realiza as consultas, grava os resultados na pasta Resultados e fecha a janela automaticamente ao final.",
    )

    tbl_lemb = doc.add_table(rows=1, cols=1)
    tbl_lemb.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_lemb = tbl_lemb.rows[0].cells[0]
    cell_lemb.width = Cm(17.4)
    _set_cell_shading(cell_lemb, "F8FAFC")
    _set_cell_borders(cell_lemb, top=None, bottom=None, right=None, left=(24, "1F4E78"))
    _set_cell_margins(cell_lemb, top=100, bottom=100, left=170, right=150)

    p0_lemb = cell_lemb.paragraphs[0]
    p0_lemb.paragraph_format.space_before = Pt(0)
    p0_lemb.paragraph_format.space_after = Pt(4)
    p0_lemb.paragraph_format.line_spacing = 1.25
    p0_lemb.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _add_run(p0_lemb, "💡 Lembretes Importantes para a Execução Agendada: ", bold=True, size_pt=9.8, color=COR_AZUL_PRIMARIO)
    _add_run(p0_lemb, "Para que a rotina agendada funcione sem falhas:", size_pt=9.8, color=COR_NOTA_TEXTO)

    lembretes = [
        [
            ("• No horário programado, o computador deve estar ", False),
            ("ligado e com a sessão do usuário desbloqueada", True),
            (" (com a tela bloqueada o Windows não exibe os diálogos de certificado para o robô clicar);", False),
        ],
        [
            ("• A planilha ", False),
            ("PlanilhasBase/CNPJs.xlsx", False, True),
            (" deve estar ", False),
            ("fechada", True),
            (" no Excel;", False),
        ],
        [
            ("• Caso existam itens em ", False),
            ("CNPJs_Pendentes.xlsx", False, True),
            (" de uma execução anterior, no modo ", False),
            ("--agendado", False, True),
            (" o robô processará automaticamente os pendentes se ", False),
            ("CNPJs.xlsx", False, True),
            (" não existir, ou executará a base completa (", False),
            ("CNPJs.xlsx", False, True),
            (") se ambas estiverem presentes.", False),
        ],
    ]
    for idx, parts in enumerate(lembretes):
        p_l = cell_lemb.add_paragraph()
        p_l.paragraph_format.space_before = Pt(0)
        p_l.paragraph_format.space_after = Pt(2 if idx < len(lembretes) - 1 else 0)
        p_l.paragraph_format.line_spacing = 1.22
        p_l.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _append_formatted_runs(p_l, parts, size_pt=9.8, default_color=COR_NOTA_TEXTO)

    doc.save(str(caminho_saida))
    print(f"[SUCESSO] Manual em DOCX gerado: {caminho_saida}")


if __name__ == "__main__":
    destino = Path(__file__).resolve().parent / "Manual_do_Usuario_RoboSocietario.docx"
    gerar_docx(destino)
