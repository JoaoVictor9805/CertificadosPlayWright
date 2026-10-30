# -*- coding: utf-8 -*-
"""
Script para geração do Manual do Usuário em formato PDF via Playwright.
Gera um documento objetivo, amigável e direto, focado exclusivamente na
apresentação do projeto e no seu passo a passo de utilização em 3 páginas A4.
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
        margin: 14mm 18mm 14mm 18mm;
    }
    * {
        box-sizing: border-box;
    }
    body {
        font-family: "Calibri", "Segoe UI", Arial, sans-serif;
        color: #2D3748;
        line-height: 1.38;
        font-size: 10pt;
        margin: 0;
        padding: 0;
    }
    .page {
        page-break-after: always;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
    }
    .page:last-child {
        page-break-after: avoid;
    }
    .title-container {
        text-align: center;
        margin-top: 4px;
        margin-bottom: 16px;
    }
    h1.main-title {
        color: #1F4E78;
        font-size: 20pt;
        font-weight: bold;
        letter-spacing: 0.5px;
        margin: 0 0 6px 0;
    }
    .subtitle {
        color: #5A6A80;
        font-size: 11pt;
        font-weight: 500;
        margin: 0 0 10px 0;
    }
    hr.title-hr {
        border: 0;
        border-top: 1.5px solid #CBD5E0;
        margin: 0 auto;
        width: 100%;
    }
    h2 {
        color: #1F4E78;
        font-size: 13pt;
        font-weight: bold;
        margin-top: 14px;
        margin-bottom: 8px;
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 4px;
    }
    h3 {
        color: #2B5B84;
        font-size: 11pt;
        font-weight: bold;
        margin-top: 12px;
        margin-bottom: 4px;
    }
    p {
        margin-top: 0;
        margin-bottom: 7px;
        text-align: justify;
    }
    ul {
        margin-top: 3px;
        margin-bottom: 8px;
        padding-left: 20px;
    }
    li {
        margin-bottom: 3px;
    }
    .action-badge {
        display: inline-block;
        background-color: #EBF3FA;
        color: #1F4E78;
        font-weight: bold;
        font-size: 10pt;
        padding: 3px 10px;
        border-radius: 4px;
        border: 1px solid #B0C8DE;
        margin: 4px 0 6px 0;
    }
    .inline-note {
        background-color: #F8FAFC;
        border-left: 3.5px solid #1F4E78;
        padding: 7px 12px;
        margin: 5px 0 7px 0;
        border-radius: 0 3px 3px 0;
        font-size: 9.8pt;
        color: #4A5568;
        line-height: 1.4;
    }
    .inline-note strong {
        color: #1F4E78;
    }
    table.data-table {
        width: 100%;
        border-collapse: collapse;
        margin: 8px 0 8px 0;
    }
    table.data-table th {
        background-color: #1F4E78;
        color: #FFFFFF;
        font-weight: bold;
        text-align: center;
        padding: 7px 10px;
        font-size: 10pt;
        border: 1px solid #1F4E78;
    }
    table.data-table td {
        border: 1px solid #CBD5E0;
        padding: 8px 10px;
        font-size: 9.5pt;
        vertical-align: top;
        background-color: #FFFFFF;
    }
    .option-box {
        background-color: #F8FAFC;
        border: 1px solid #CBD5E0;
        border-radius: 4px;
        padding: 8px 12px;
        margin: 5px 0 7px 0;
        font-size: 9.8pt;
        line-height: 1.4;
    }
    .option-box .opt-title {
        font-weight: bold;
        color: #1F4E78;
        margin-bottom: 4px;
    }
    .page-footer {
        margin-top: auto;
        padding-top: 6px;
        border-top: 1px solid #E2E8F0;
        display: flex;
        justify-content: space-between;
        font-size: 8.5pt;
        color: #718096;
    }
</style>
</head>
<body>

<!-- PÁGINA 1 -->
<div class="page">
    <div class="title-container">
        <h1 class="main-title">ROBÔ SOCIETÁRIO – DET / GOV.BR</h1>
        <div class="subtitle">Manual de Operação e Guia Prático para o Usuário Final</div>
        <hr class="title-hr">
    </div>

    <h2>1. Apresentação e Objetivo do Sistema</h2> as consultas em lote no portal do Domicílio Eletrônico Trabalhista (DET / Gov.br), eliminando o trabalho manual de acessar cada empresa individualmente.</p>
    
    <p>O sistema realiza as seguintes funções de forma autônoma:</p>
    <ul>
        <liAcesso via Gov.br utilizando certificado digital com procuração eletrônica vinculada ao CNPJ da empresa consultada;</li>
        <li>Verificação da Caixa Postal em busca de mensagens e avisos pendentes ou já lidos;</li>
        <li>Localização de novas notificações e download automático dos termos e anexos oficiais (PDFs ou ZIPs);</li>
        <li>Registro de capturas de tela (prints) de todas as etapas para auditoria e evidência jurídica;</li>
        <li>Emissão de relatório geral consolidado em Excel com o status detalhado de cada empresa consultada.</li>
    </ul>

    <h2>2. Guia Passo a Passo de Utilização</h2>
    
    <h3>Passo 1: Instalação dos Certificados Digitais na Máquina</h3>
    <p>Antes de iniciar as consultas, certifique-se de que os certificados digitais (modelo A1) necessários estejam devidamente instalados no seu perfil de usuário ativo no Windows. O robô reconhece e utiliza os certificados já presentes na máquina para realizar as autenticações de forma transparente.</p>

    <h3>Passo 2: Atualização dos Certificados na Planilha (Atualizar_Certificados.bat)</h3>
    <p>Na primeira vez que você for usar o sistema, ou sempre que instalar um novo certificado digital no computador, basta dar dois cliques no arquivo utilitário:</p>
    
    <div>👉 <span class="action-badge">Atualizar_Certificados.bat</span></div>

    <p>Em poucos segundos, esse assistente fará a varredura do Windows e preencherá automaticamente a 2ª aba (<strong>Certificados_Disponiveis</strong>) da planilha <strong>CNPJs.xlsx</strong> (localizada dentro da pasta <strong>PlanilhasBase</strong>). Ele cataloga todos os certificados ativos, trazendo a Razão Social, o número do documento e a data de validade de cada um.</p>

    <div class="inline-note">
        <strong>💡 Observação de Uso:</strong> Para que o assistente consiga salvar as informações sem impedimentos do Windows, certifique-se de que o arquivo <em>CNPJs.xlsx</em> não esteja aberto no Excel durante esse clique rápido.
    </div>

    <div class="page-footer">
        <span>Robô Societário – DET / Gov.br</span>
        <span>Página 1 de 3</span>
    </div>
</div>

<!-- PÁGINA 2 -->
<div class="page">
    <h3>Passo 3: Preenchimento da Planilha de Consultas (CNPJs.xlsx)</h3>
    <p>Agora que os certificados foram coletados na etapa anterior, abra a planilha <strong>CNPJs.xlsx</strong> localizada dentro da pasta <strong>PlanilhasBase</strong> no Excel. O preenchimento conecta diretamente o que foi gerado no Passo 2 com as empresas que você quer consultar:</p>
    
    <ul>
        <li><strong>Aba 'Certificados_Disponiveis' (Resultado do Passo 2):</strong> Mostra a lista de todos os certificados encontrados na sua máquina. Você usará essa aba como fonte de consulta para saber exatamente como o certificado está identificado.</li>
        <li><strong>Aba 'CNPJs_consultados' (Onde você preenche as empresas):</strong> É aqui onde você informa quais empresas serão consultadas pelo robô, seguindo este formato simples:</li>
    </ul>

    <table class="data-table">
        <thead>
            <tr>
                <th style="width: 44%;">Coluna A: CNPJ</th>
                <th style="width: 56%;">Coluna B: Certificado</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Digite os 14 dígitos do CNPJ da empresa que deseja consultar (exemplo: <code>39407892000185</code>).</td>
                <td>Copie e cole o nome exatamente como aparece na aba <em>Certificados_Disponiveis</em>.</td>
            </tr>
        </tbody>
    </table>

    <div class="inline-note">
        <strong>💡 Conexão entre as Abas:</strong> Para que o robô localize o certificado correto sem erros de digitação, vá até a aba <em>Certificados_Disponiveis</em>, copie (Ctrl+C) o nome do certificado desejado e cole (Ctrl+V) na Coluna B da aba <em>CNPJs_consultados</em>. Após preencher e conferir a lista, salve o arquivo e feche o Excel.
    </div>

    <h3>Passo 4: Executando a Automação (RoboSocietario.exe)</h3>
    <p>Com a planilha preenchida e fechada, basta dar dois cliques no aplicativo principal para iniciar as consultas:</p>

    <div>👉 <span class="action-badge">RoboSocietario.exe</span></div>

    <p>O robô inicializará o Google Chrome automaticamente e executará as consultas empresa por empresa. Uma janela de texto preta (Console) acompanhará a execução em tempo real, informando o percentual de avanço, o CNPJ em atendimento e cada ação realizada.</p>

    <div class="inline-note">
        <strong>💡 Movimentação do Mouse:</strong> Durante os primeiros instantes em que a janela do Windows surge na tela para confirmar o certificado digital, evite mexer no mouse ou no teclado para que o robô realize a seleção física com total precisão.
    </div>

    <p><strong>Gerenciamento de Instabilidades do Portal:</strong> Os portais do governo frequentemente passam por oscilações temporárias de conexão. Pensando nisso, o robô possui mecanismos de proteção automáticos:</p>
    <ul>
        <li><strong>Planilha de pendências (CNPJs_Pendentes.xlsx):</strong> Se alguma empresa não responder (ou se exigir validação por celular / 2FA) durante a execução do robô, ela é registrada automaticamente em <em>PlanilhasBase/CNPJs_Pendentes.xlsx</em>.</li>
        <li><strong>Segunda passagem:</strong> Se algum CNPJ for registrado como pendente, o robô automaticamente tentará analisar ele novamente. Caso obtenha sucesso nessa análise, o CNPJ sai da planilha CNPJs_Pendentes e é registrado normalmente.</li>
    </ul>

    <p>Caso existam pendências de uma execução anterior, ao abrir o <strong>RoboSocietario.exe</strong> você poderá escolher de forma simples:</p>

    <div class="option-box">
        <div class="opt-title">Opções de Execução Disponíveis:</div>
        <div>• <strong>[Opção 1 - ENTER]:</strong> Executar a base completa cadastrada na planilha principal (<em>PlanilhasBase/CNPJs.xlsx</em>).</div>
        <div>• <strong>[Opção 2]:</strong> Reprocessar apenas as empresas que ficaram pendentes (<em>PlanilhasBase/CNPJs_Pendentes.xlsx</em>).</div>
    </div>

    <p>Assim que todas as empresas pendentes forem concluídas com sucesso, a planilha de pendências é excluída automaticamente pelo robô.</p>

    <div class="page-footer">
        <span>Robô Societário – DET / Gov.br</span>
        <span>Página 2 de 3</span>
    </div>
</div>

<!-- PÁGINA 3 -->
<div class="page">
    <h3>Passo 5: Conferência dos Resultados e Evidências</h3>
    <p>Ao finalizar o processamento de todos os itens, o robô exibirá uma mensagem de conclusão na tela e aguardará você pressionar ENTER para fechar a janela. Todos os resultados e evidências ficam organizados dentro da pasta <strong>Resultados</strong>:</p>

    <p><strong>1. Pasta 'Resultados/Relatorios':</strong><br>
    Contém a planilha consolidada em Excel gerada a cada execução, identificada com data e hora (exemplo: <em>Relatorio_Execucao_30-10-2026_09-00-00.xlsx</em>). Essa planilha lista cada CNPJ consultado, o certificado utilizado, o status e detalhes da 1ª e 2ª tentativa, a quantidade de arquivos guardados (ou o motivo em caso de erro), o Status Final destacado visualmente por cor (verde para Sucesso, amarelo para Confirmação de Identidade e vermelho para Falhas) e a data/hora exata da consulta.</p>

    <p><strong>2. Pasta 'Resultados/CapturasDeTela-Downloads':</strong><br>
    Armazena o dossiê completo de auditoria organizado cronologicamente por data e hora da execução, agrupado por certificado digital e com uma subpasta individual para cada CNPJ consultado. Dentro de cada pasta ficam salvos:</p>
    <ul>
        <li>Capturas de tela (prints) em alta resolução da Caixa Postal (antes e após a aplicação de filtros de não lidas);</li>
        <li>Capturas de tela das notificações oficiais com data e hora visíveis;</li>
        <li>Arquivos de termos de notificação e documentos anexos baixados em formato oficial (PDFs ou ZIPs).</li>
    </ul>

    <div class="inline-note">
        <strong>💡 Verificação em Duas Etapas (2FA):</strong> Caso a conta Gov.br de uma empresa solicite confirmação de identidade no aplicativo do celular, o robô salva um print da tela como comprovante e guarda a empresa na planilha de pendências, prosseguindo com as demais empresas da lista sem travar a rotina.
    </div>

    <p><strong>Rotina Automática Agendada (Sexta-feira às 08:00):</strong><br>
    O Robô Societário também será configurado para operar de forma 100% autônoma toda sexta-feira às 08:00 através do Agendador de Tarefas do Windows. Nessa rotina agendada, o funcionamento é totalmente silencioso:</p>
    <ul>
        <li>A equipe só precisa manter a planilha <em>PlanilhasBase/CNPJs.xlsx</em> preenchida previamente com as empresas e certificados da semana;</li>
        <li>Pontualmente às 08:00, o robô inicia sozinho, realiza todas as consultas em lote, grava os downloads e capturas, e encerra o processo sem exigir nenhuma confirmação manual;</li>
        <li>Basta abrir a pasta <em>Resultados/Relatorios</em> para conferir o relatório consolidado pronto com os resultados da semana.</li>
    </ul>

    <div class="page-footer">
        <span>Robô Societário – DET / Gov.br</span>
        <span>Página 3 de 3</span>
    </div>
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
                "top": "0mm",
                "bottom": "0mm",
                "left": "0mm",
                "right": "0mm",
            },
        )
        browser.close()
    print(f"[SUCESSO] Manual em PDF gerado: {caminho_saida}")

if __name__ == "__main__":
    destino = Path(__file__).resolve().parent / "Manual_do_Usuario_RoboSocietario.pdf"
    gerar_pdf(destino)
