# -*- coding: utf-8 -*-
"""
Script para geração do Manual do Usuário em formato PDF via Playwright.
Gera um documento com diagramação idêntica ao modelo institucional, com 3 páginas perfeitas.
"""

from pathlib import Path
from playwright.sync_api import sync_playwright

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
    @page {
        size: A4;
        margin: 18mm 20mm 18mm 20mm;
    }
    * {
        box-sizing: border-box;
    }
    body {
        font-family: "Calibri", "Segoe UI", Arial, sans-serif;
        color: #262626;
        line-height: 1.45;
        font-size: 11pt;
        margin: 0;
        padding: 0;
    }
    .page {
        page-break-after: always;
        height: 100%;
        display: flex;
        flex-direction: column;
    }
    .page:last-child {
        page-break-after: avoid;
    }
    .title-container {
        text-align: center;
        margin-top: 10px;
        margin-bottom: 22px;
    }
    h1.main-title {
        color: #1F4E78;
        font-size: 20pt;
        font-weight: bold;
        letter-spacing: 0.5px;
        margin: 0 0 10px 0;
    }
    hr.title-hr {
        border: 0;
        border-top: 1px solid #D9D9D9;
        margin: 0 auto;
        width: 100%;
    }
    h2 {
        color: #1F4E78;
        font-size: 13pt;
        font-weight: bold;
        margin-top: 16px;
        margin-bottom: 8px;
    }
    h3 {
        color: #2F5597;
        font-size: 11.5pt;
        font-weight: bold;
        margin-top: 14px;
        margin-bottom: 6px;
    }
    p {
        margin-top: 0;
        margin-bottom: 8px;
        text-align: justify;
    }
    ul {
        margin-top: 4px;
        margin-bottom: 12px;
        padding-left: 22px;
    }
    li {
        margin-bottom: 4px;
    }
    .file-action {
        font-size: 11pt;
        margin: 8px 0 12px 18px;
        color: #111;
    }
    .callout {
        background-color: #F2F5F9;
        border-left: 4px solid #1F4E78;
        padding: 10px 14px;
        margin: 12px 0 14px 0;
        border-radius: 2px;
        font-size: 10.5pt;
        line-height: 1.4;
    }
    .callout strong {
        color: #1F4E78;
    }
    table.data-table {
        width: 100%;
        border-collapse: collapse;
        margin: 10px 0 6px 0;
    }
    table.data-table th {
        background-color: #1F4E78;
        color: #FFFFFF;
        font-weight: bold;
        text-align: center;
        padding: 8px 10px;
        font-size: 10.5pt;
        border: 1px solid #1F4E78;
    }
    table.data-table td {
        border: 1px solid #B0C4DE;
        padding: 10px 12px;
        font-size: 10pt;
        vertical-align: top;
        background-color: #FFFFFF;
    }
    .footnote {
        font-size: 9.5pt;
        color: #555555;
        margin-top: 4px;
        margin-bottom: 10px;
    }
    .console-box {
        background-color: #111111;
        color: #F0F0F0;
        font-family: "Consolas", "Courier New", monospace;
        font-size: 10pt;
        padding: 14px 18px;
        border-radius: 4px;
        margin: 10px 0 14px 0;
        line-height: 1.4;
        box-shadow: inset 0 0 4px rgba(0,0,0,0.5);
    }
    .console-line {
        white-space: pre-wrap;
    }
</style>
</head>
<body>

<!-- PÁGINA 1 -->
<div class="page">
    <div class="title-container">
        <h1 class="main-title">ROBÔ SOCIETÁRIO – DET / GOV.BR</h1>
        <hr class="title-hr">
    </div>

    <h2>1. Apresentação e Objetivo do Sistema</h2>
    <p>O Robô Societário foi desenvolvido para automatizar as consultas em lote no portal DET (Domicílio Eletrônico Trabalhista) do Ministério do Trabalho e Emprego (Gov.br).</p>
    <p>A principais funções da ferramenta são:</p>
    <ul>
        <li>Verificação de notificações e avisos pendentes de leitura ou já recebidos;</li>
        <li>Download automático de termos de notificação e anexos oficiais em formato PDF ou ZIP;</li>
        <li>Captura de telas (prints) de todas as etapas;</li>
        <li>Emissão de relatório em Excel contendo o status de cada CNPJ consultado.</li>
    </ul>

    <h2>2. Guia Passo a Passo de Utilização</h2>
    <h3>Passo 1: Instalação dos Certificados Digitais na Máquina</h3>
    <p>Instale os certificados digitais A1 no perfil de usuário ativo na máquina Windows.<br>
    (Isso pode ser verificado pelo comando Windows + R no teclado -&gt; <strong>certmgr.msc</strong> -&gt; Pessoal -&gt; Certificados).</p>

    <h3>Passo 2: Atualização dos Certificados no Excel (2 Cliques)</h3>
    <p>Na primeira vez que você for executar o robô ou sempre que você instalar um novo certificado na máquina, é interessante dar 2 cliques no arquivo:</p>
    
    <div class="file-action">👉 <strong>Atualizar_Certificados.bat</strong></div>

    <p>Esse arquivo é um “script” que fará a varredura do Windows e atualizará a aba 'Certificados_Disponiveis' da planilha <strong>CNPJs.xlsx</strong> (localizada dentro da pasta <strong>PlanilhasBase</strong>) com os nomes, números de documento e datas de validade de cada certificado instalado na máquina. Isso evita de ter que pegar os nomes dos certificados manualmente e traz exatamente o nome que o robô espera na aba “Nome Completo (certmgr)”.</p>

    <p>Depois de executar, clique “Enter” ou apenas feche a janela “preta” (console) que abriu durante a execução.</p>

    <div class="callout">
        <strong>📌 AVISO DE ARQUIVO FECHADO:</strong> Certifique-se de que a planilha CNPJs.xlsx (dentro da pasta <strong>PlanilhasBase</strong>) não esteja aberta no Excel durante esse clique. O Windows impede alterações em arquivos que já estejam abertos por outro programa.
    </div>
</div>

<!-- PÁGINA 2 -->
<div class="page">
    <h3>Passo 3: Preenchimento da Planilha CNPJs.xlsx</h3>
    <p>Abra o arquivo <strong>CNPJs.xlsx</strong> localizado dentro da pasta <strong>PlanilhasBase</strong> no Microsoft Excel. A planilha possui duas abas planejadas para facilitar o trabalho:</p>

    <p><strong>1. Aba 'Certificados_Disponiveis' (Apenas Consulta):</strong><br>
    Exibe todos os certificados instalados na sua máquina, é o resultado da execução do Passo 2.</p>

    <p><strong>2. Aba ‘CNPJs_consultados’ (Preenchimento das Consultas):</strong><br>
    É aqui onde você informa quais empresas serão consultadas pelo robô, nesse padrão:</p>

    <table class="data-table">
        <thead>
            <tr>
                <th style="width: 48%;">CNPJ</th>
                <th style="width: 52%;">Certificado (ao qual o CNPJ pertence)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Digite os 14 dígitos do CNPJ da empresa a ser consultada (ex: 39407892000185).</td>
                <td>Digite o Nome Completo (certmgr) do certificado procurador.</td>
            </tr>
        </tbody>
    </table>

    <p class="footnote">** A planilha estará preenchida com alguns CNPJs na primeira vez que abri-la. Eles servem apenas como exemplo e podem ser deletados a qualquer momento desejado.</p>

    <div class="callout">
        <strong>📌 FACILIDADE DE PREENCHIMENTO:</strong> Para evitar preencher o nome do certificado com valores que o robô não aceite, é interessante copiar (ctrl + c) os nomes presentes na aba ‘Certificados_Disponiveis’ e colar diretamente na coluna de certificado da aba ‘CNPJs_consultados’.
    </div>

    <p><em>Importante: Após preencher e conferir a lista, salve o arquivo e feche o Excel antes de prosseguir.</em></p>

    <h3>Passo 4: Executando a Automação (RoboSocietario.exe)</h3>
    <p>Para iniciar as consultas, dê dois cliques diretamente no executável:</p>

    <div class="file-action">👉 <strong>RoboSocietario.exe</strong></div>

    <p>Na primeira execução, o sistema abrirá uma janela preta (Console), inicializará o Google Chrome e fará todo o processo principal automaticamente (abrir site do DET, acessar certificados, CNPJs, tirar prints…).</p>

    <p>Devido a instabilidades no site do DET, alguns CNPJs podem falhar. Pensando nisso, caso algum CNPJ fique pendente, ele será registrado em uma nova planilha dentro de <strong>PlanilhasBase</strong>: “<strong>CNPJs_pendentes.xlsx</strong>”.</p>
</div>

<!-- PÁGINA 3 -->
<div class="page">
    <p>Se isso acontecer, ao dar dois cliques em RoboSocietario.exe, ele abrirá essa tela:</p>

    <div class="console-box">
        <div class="console-line">[AVISO] Identificada planilha com itens pendentes da execução anterior:</div>
        <div class="console-line">   [1] Executar base completa (PlanilhasBase/CNPJs.xlsx)</div>
        <div class="console-line">   [2] Reprocessar apenas pendentes anteriores (PlanilhasBase/CNPJs_Pendentes.xlsx)</div>
        <br>
        <div class="console-line">Digite a opção [Pressione ENTER para opção 1]:</div>
    </div>

    <p>Basta escolher se deseja executar todo o processo novamente, ou apenas o que ficou pendente.</p>
    <ul>
        <li><strong>Limpeza Automática:</strong> Quando todos os CNPJs forem concluídos com êxito, a planilha de pendências é excluída automaticamente.</li>
    </ul>

    <h3>Passo 5: Conferência dos Resultados e Evidências</h3>
    <p>Ao finalizar a execução, o robô exibirá uma mensagem de conclusão e aguardará você pressionar ENTER. Você poderá conferir os resultados nos seguintes locais:</p>

    <p><strong>1. Pasta 'Resultados/Relatorios':</strong><br>
    Contém a planilha Relatorio_Execucao_&lt;Data_Hora&gt;.xlsx com o consolidado de todas as consultas, indicando se houve notificações, se havia anexos, quantidade de mensagens e links diretos para os prints.</p>

    <p><strong>2. Pasta 'Resultados/CapturasDeTela-Downloads':</strong><br>
    Organizada por data e hora da execução, contendo uma subpasta dedicada para cada CNPJ com todos os prints comprobatórios e arquivos de anexos (PDFs/ZIPs) baixados.</p>

    <h2>3. Guia de Funcionamento agendado (Sexta-feira, 8:00)</h2>
    <p><strong>1.</strong> O robô será executado toda sexta feira às 8:00 da manhã, automaticamente tirando prints e fazendo download dos arquivos depois da planilha “CNPJs.xlsx” em <strong>PlanilhasBase</strong> ser configurada.</p>
    <p><strong>2.</strong> A possibilidade de realizar a execução manualmente existe caso o site apresente instabilidade durante o horário da execução e não consiga finalizar com sucesso a consulta aos CNPJs. Dessa forma, o responsável pode executar o processo para todos os CNPJs ou somente para os CNPJs que ficaram pendentes a qualquer momento.</p>
</div>

</body>
</html>
"""

def gerar_pdf(caminho_saida: Path):
    caminho_saida = Path(caminho_saida).resolve()
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.set_content(HTML_TEMPLATE, wait_until="networkidle")
        page.pdf(
            path=str(caminho_saida),
            format="A4",
            print_background=True,
            margin={
                "top": "15mm",
                "bottom": "15mm",
                "left": "18mm",
                "right": "18mm",
            },
        )
        browser.close()
    print(f"[SUCESSO] Manual em PDF gerado: {caminho_saida}")

if __name__ == "__main__":
    destino = Path(__file__).resolve().parent / "Manual_do_Usuario_RoboSocietario.pdf"
    gerar_pdf(destino)
