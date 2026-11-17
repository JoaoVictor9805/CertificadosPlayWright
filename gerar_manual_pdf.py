# -*- coding: utf-8 -*-
"""
Script para geração do Manual do Usuário em formato PDF via Playwright.
Reflete exatamente o conteúdo e o padrão visual de Manual_do_Usuario_RoboSocietario.docx
(incluindo Passo 1 a Passo 5 e Seção 3 de Agendamento no Task Scheduler).
"""

from pathlib import Path
from playwright.sync_api import sync_playwright

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
    @page {
        size: A4;
        margin: 16mm 18mm 20mm 18mm;
    }
    * {
        box-sizing: border-box;
    }
    body {
        font-family: "Calibri", "Segoe UI", Arial, sans-serif;
        color: #2D3748;
        line-height: 1.5;
        font-size: 10pt;
        margin: 0;
        padding: 0;
    }
    .title-container {
        text-align: center;
        margin-top: 2px;
        margin-bottom: 18px;
    }
    h1.main-title {
        color: #1F4E78;
        font-size: 20pt;
        font-weight: bold;
        letter-spacing: 0.4px;
        margin: 0 0 6px 0;
    }
    .subtitle {
        color: #5A6A80;
        font-size: 11pt;
        font-weight: 500;
        margin: 0 0 12px 0;
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
        margin-top: 22px;
        margin-bottom: 10px;
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 5px;
        break-after: avoid;
        page-break-after: avoid;
    }
    h2:first-of-type {
        margin-top: 8px;
    }
    h3 {
        color: #2B5B84;
        font-size: 11pt;
        font-weight: bold;
        margin-top: 16px;
        margin-bottom: 7px;
        break-after: avoid;
        page-break-after: avoid;
    }
    p {
        margin-top: 0;
        margin-bottom: 9px;
        text-align: justify;
    }
    p.sub-heading-p {
        text-align: left;
        margin-top: 13px;
        margin-bottom: 7px;
        line-height: 1.48;
    }
    p.sub-heading-p strong.blue-label {
        color: #1F497D;
    }
    ul {
        margin-top: 4px;
        margin-bottom: 11px;
        padding-left: 20px;
    }
    li {
        margin-bottom: 6px;
        text-align: justify;
    }
    .action-row {
        margin: 10px 0 12px 0;
        break-inside: avoid;
        page-break-inside: avoid;
    }
    .action-badge {
        display: inline-block;
        background-color: #EBF3FA;
        color: #1F4E78;
        font-weight: bold;
        font-size: 10pt;
        padding: 5px 16px;
        border-radius: 3px;
        border: 1px solid #B0C8DE;
        margin-left: 4px;
    }
    .inline-note {
        background-color: #F8FAFC;
        border-left: 3.5px solid #1F4E78;
        padding: 10px 14px;
        margin: 11px 0 13px 0;
        border-radius: 0 3px 3px 0;
        font-size: 9.8pt;
        color: #4A5568;
        line-height: 1.48;
        text-align: justify;
        break-inside: avoid;
        page-break-inside: avoid;
    }
    .inline-note strong.note-title {
        color: #1F4E78;
    }
    .inline-note ul.note-list {
        margin: 6px 0 0 0;
        padding-left: 0;
        list-style: none;
    }
    .inline-note ul.note-list li {
        margin-bottom: 5px;
    }
    .inline-note ul.note-list li:last-child {
        margin-bottom: 0;
    }
    table.data-table {
        width: 100%;
        border-collapse: collapse;
        margin: 10px 0 13px 0;
        break-inside: avoid;
        page-break-inside: avoid;
    }
    table.data-table th {
        background-color: #1F4E78;
        color: #FFFFFF;
        font-weight: bold;
        text-align: center;
        padding: 8px 10px;
        font-size: 10pt;
        border: 1px solid #1F4E78;
    }
    table.data-table td {
        border: 1px solid #CBD5E0;
        padding: 10px 12px;
        font-size: 9.5pt;
        line-height: 1.45;
        vertical-align: top;
        background-color: #FFFFFF;
    }
    .option-box {
        background-color: #F8FAFC;
        border: 1px solid #CBD5E0;
        border-radius: 3px;
        padding: 10px 14px;
        margin: 10px 0 13px 0;
        font-size: 9.8pt;
        line-height: 1.48;
        break-inside: avoid;
        page-break-inside: avoid;
    }
    .option-box > div {
        margin-bottom: 4px;
    }
    .option-box > div:last-child {
        margin-bottom: 0;
    }
    .option-box .opt-title {
        font-weight: bold;
        color: #1F4E78;
        margin-bottom: 4px;
    }
</style>
</head>
<body>

<div class="title-container">
    <h1 class="main-title">ROBÔ SOCIETÁRIO – DET / GOV.BR</h1>
    <div class="subtitle">Manual de Operação e Guia Prático para o Usuário Final</div>
    <hr class="title-hr">
</div>

<h2>1. Apresentação e Objetivo do Sistema</h2>
<p>O Robô Societário automatiza as consultas em lote no portal do Domicílio Eletrônico Trabalhista (DET / Gov.br), eliminando o trabalho manual de acessar cada empresa individualmente.</p>

<p>O sistema realiza as seguintes funções de forma autônoma:</p>
<ul>
    <li>Acesso via Gov.br utilizando certificado digital — seja com o próprio certificado da empresa titular (Meu Perfil) ou com procuração eletrônica vinculada ao CNPJ consultado (Procurador);</li>
    <li>Verificação da Caixa Postal em busca de mensagens e avisos pendentes ou já lidos (incluindo alternância automática entre Matriz e Filiais quando disponível);</li>
    <li>Localização de novas notificações e download automático dos termos e anexos oficiais (PDFs ou ZIPs);</li>
    <li>Registro de capturas de tela (prints) de todas as etapas para auditoria e evidência jurídica, organizadas por empresa;</li>
    <li>Emissão de relatório geral consolidado em Excel com o status detalhado de cada empresa consultada.</li>
</ul>

<h2>2. Guia Passo a Passo de Utilização</h2>

<h3>Passo 1: Instalação dos Certificados Digitais na Máquina</h3>
<p>Antes de iniciar as consultas, certifique-se de que os certificados digitais (modelo A1) necessários estejam devidamente instalados no seu perfil de usuário ativo no Windows. O robô reconhece e utiliza os certificados já presentes na máquina para realizar as autenticações de forma transparente.</p>

<h3>Passo 2: Preparação Inicial dos Certificados (Atualizar_Certificados.bat)</h3>
<p>Os Passos 2 e 3 formam a <strong>configuração inicial do sistema</strong> e são realizados apenas <strong>antes da primeira execução</strong> (você só precisará repetir este Passo 2 se instalar ou renovar algum certificado digital no computador). Dê dois cliques no arquivo utilitário:</p>

<div class="action-row">👉 <span class="action-badge">Atualizar_Certificados.bat</span></div>

<p>Esse assistente lê os certificados digitais instalados no Windows e atualiza automaticamente a aba <strong>Certificados_Disponiveis</strong> dentro da planilha <strong>PlanilhasBase/CNPJs.xlsx</strong> (que é a planilha de configurações do robô). Se o arquivo ainda não existir na pasta, ele será criado automaticamente.</p>

<div class="inline-note">
    <strong class="note-title">💡 Antes de clicar:</strong> Certifique-se de que a planilha <em>CNPJs.xlsx</em> esteja fechada no Excel para que o assistente consiga gravar os certificados sem bloqueio do Windows.
</div>

<h3>Passo 3: Configuração da Planilha de Consultas (PlanilhasBase/CNPJs.xlsx)</h3>
<p>A planilha <strong>CNPJs.xlsx</strong> (dentro da pasta <strong>PlanilhasBase</strong>) é a <strong>planilha de configurações do Robô Societário</strong>. Tudo o que você configurar aqui fica salvo para as próximas execuções — ou seja, <strong>todo esse preenchimento é feito somente antes da primeira vez que você rodar o robô</strong>, sendo necessário editar novamente apenas quando quiser incluir, alterar ou remover alguma empresa.</p>

<p>Abra o arquivo PlanilhasBase/CNPJs.xlsx no Excel e siga esta ordem:</p>
<ul>
    <li><strong>1º — Consulte a aba 'Certificados_Disponiveis':</strong> Nela está a lista pronta dos certificados da sua máquina (gerada no Passo 2). Você usará essa aba apenas para copiar o nome exato de cada certificado.</li>
    <li><strong>2º — Preencha a aba 'CNPJs_consultados':</strong> É nesta aba que você cadastra as empresas que o robô deve consultar (uma empresa por linha), preenchendo as 3 colunas abaixo:</li>
</ul>

<table class="data-table">
    <thead>
        <tr>
            <th style="width: 26%;">Coluna A: CNPJ</th>
            <th style="width: 38.5%;">Coluna B: Certificado</th>
            <th style="width: 35.5%;">Coluna C: Nome da Empresa</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>Digite os 14 dígitos do CNPJ da empresa que será consultada.<br><strong>Exemplo:</strong> <em>39407892000185</em></td>
            <td>Copie (Ctrl+C) na aba <em>Certificados_Disponiveis</em> (coluna <em>Nome Completo (certmgr)</em>) e cole (Ctrl+V) o certificado que tem acesso àquela empresa.</td>
            <td>Digite o nome da empresa (opcional, porém recomendado). Ele aparece no relatório final e dá nome à pasta dos prints; se ficar vazio, a pasta usa o CNPJ.</td>
        </tr>
    </tbody>
</table>

<div class="inline-note">
    <strong class="note-title">💡 Finalizando a Configuração:</strong> Depois de preencher todas as linhas na aba <em>CNPJs_consultados</em> (sem deixar o certificado escrito diferente da aba <em>Certificados_Disponiveis</em>), <strong>salve a planilha e feche o Excel</strong>. A partir daqui, o robô já está configurado e pronto para uso — nas próximas vezes, basta abrir direto o <strong>RoboSocietario.exe</strong> no Passo 4.
</div>

<div class="inline-note">
    <strong class="note-title">💡 Empresa usando o Próprio Certificado (sem procuração):</strong> Se você cadastrar na Coluna A o mesmo CNPJ do certificado informado na Coluna B (por exemplo, consultando a própria empresa dona do certificado), o robô reconhece isso sozinho, coloca essa empresa como a <strong>primeira da fila</strong> daquele certificado e acessa diretamente o perfil próprio (<em>Meu Perfil</em>), seguindo normalmente como Procurador para as demais empresas ligadas ao mesmo certificado.
</div>

<h3>Passo 4: Executando a Automação no Dia a Dia (RoboSocietario.exe)</h3>
<p>Com a planilha de configurações (<em>PlanilhasBase/CNPJs.xlsx</em>) salva e fechada, você já pode iniciar as consultas sempre que precisar dando dois cliques no aplicativo principal:</p>

<div class="action-row">👉 <span class="action-badge">RoboSocietario.exe</span></div>

<p>Ao abrir o programa, uma janela preta de acompanhamento (Console) mostrará em tempo real o andamento da fila, o CNPJ em consulta e cada etapa concluída. O robô abrirá o Google Chrome sozinho e fará todo o processo sem necessidade de intervenção manual.</p>

<div class="inline-note">
    <strong class="note-title">💡 Cuidados Durante a Execução:</strong> Enquanto o robô estiver abrindo o Chrome e confirmando o certificado digital na janela do Windows, <strong>evite mexer no mouse ou no teclado</strong> para não tirar o foco da tela. Também não feche o navegador aberto pelo robô nem a janela do Console antes da mensagem final de conclusão.
</div>

<p><strong>Se houver pendências da execução anterior:</strong> antes de iniciar as consultas, o robô perguntará na tela o que você deseja fazer:</p>

<div class="option-box">
    <div class="opt-title">Opções de Execução Disponíveis:</div>
    <div>• <strong>[Opção 1 - ENTER]:</strong> Consultar todas as empresas cadastradas na planilha de configurações (<em>PlanilhasBase/CNPJs.xlsx</em>).</div>
    <div>• <strong>[Opção 2]:</strong> Reprocessar somente as empresas que ficaram pendentes na última tentativa (<em>PlanilhasBase/CNPJs_Pendentes.xlsx</em>).</div>
</div>

<p><strong>Proteção automática contra falhas do portal:</strong> como sites governamentais podem oscilar, você não precisa refazer manualmente o que falhar na primeira tentativa:</p>
<ul>
    <li><strong>1. Repetição automática (Segunda passagem):</strong> Se alguma empresa apresentar instabilidade durante a execução, o próprio robô tenta consultá-la novamente ao final da rodada.</li>
    <li><strong>2. Registro de pendências (CNPJs_Pendentes.xlsx):</strong> Somente se o erro persistir (ou se o Gov.br exigir validação em duas etapas no celular), a empresa é salva em <em>PlanilhasBase/CNPJs_Pendentes.xlsx</em> com CNPJ, Certificado e Nome da Empresa. Quando todas forem resolvidas em uma execução futura, esse arquivo de pendências é apagado automaticamente. <strong>Não é necessário fazer nenhuma modificação manual nesse arquivo</strong> — seu intuito é exclusivamente orientar uma nova execução focada apenas no que ficou pendente da execução anterior.</li>
</ul>

<h3>Passo 5: Conferência dos Resultados e Evidências (Pasta Resultados)</h3>
<p>Quando o robô terminar todas as empresas, ele exibirá o resumo final no Console e pedirá para você pressionar ENTER para fechar a janela. Tudo o que foi gerado fica salvo dentro da pasta <strong>Resultados</strong>, dividido em dois locais:</p>

<p class="sub-heading-p"><strong>1. Pasta 'Resultados/Relatorios' (Resumo Geral em Excel):</strong><br>
É o primeiro lugar que você deve abrir ao final da rotina. Cada execução gera uma planilha com data e hora (exemplo: <em>Relatorio_Execucao_30-10-2026_09-00-00.xlsx</em>) mostrando CNPJ, <strong>Nome da Empresa</strong>, Certificado utilizado, detalhes da 1ª e 2ª tentativa, quantidade de arquivos salvos e o <strong>Status Final</strong> colorido para leitura rápida:</p>
<ul>
    <li><strong>Verde (SUCESSO):</strong> Empresa consultada e evidências salvas normalmente.</li>
    <li><strong>Amarelo (Necessita confirmação de identidade):</strong> O login Gov.br do certificado pediu código/aprovação no celular (2FA).</li>
    <li><strong>Vermelho (FALHA / NAO_EXECUTADO / SEM_PROCURACAO):</strong> Erro persistente, certificado não encontrado na máquina ou ausência de procuração no DET.</li>
</ul>

<p class="sub-heading-p"><strong>2. Pasta 'Resultados/CapturasDeTela-Downloads' (Prints e Anexos por Empresa):</strong><br>
Guarda todas as provas visuais e documentos organizados em pastas seguindo a ordem: <strong>Data e Hora da Execução → Certificado → Nome da Empresa</strong> (ou o CNPJ, caso a Coluna C tenha ficado em branco). Dentro da pasta de cada empresa você encontra:</p>
<ul>
    <li><strong>CaixaPostal:</strong> prints da Caixa Postal enquadrados direto no título <em>Caixa Postal.</em> Quando houver mensagens, são gerados 2 prints (todas e apenas não lidas); se a empresa possuir seleção entre <strong>Matriz e Filiais</strong>, o robô captura os prints dos dois perfis (<em>matriz</em> e <em>filiais</em>).</li>
    <li><strong>Notificacao:</strong> prints das notificações encontradas com data e hora visíveis (ou comprovando que não há notificações).</li>
    <li><strong>ArquivosBaixados:</strong> termos de notificação e anexos oficiais (PDFs ou ZIPs) baixados automaticamente, a partir da data do dia da execução.</li>
</ul>

<div class="inline-note">
    <strong class="note-title">💡 O que fazer quando o Status ficar Amarelo (2FA)?</strong> Quando o Gov.br exigir verificação em duas etapas no celular, o robô salva automaticamente um print comprobatório na pasta do certificado, pula aquele grupo.
</div>

<p class="sub-heading-p"><strong class="blue-label">Rotina Automática Agendada (Sugerido: Sexta-feira às 08:00):</strong><br>
O Robô Societário também pode operar de forma 100% autônoma toda sexta-feira às 08:00 pelo Agendador de Tarefas do Windows, sem precisar de cliques nem confirmações:</p>
<ul>
    <li><strong>1. Até quinta-feira:</strong> basta manter a planilha de configurações (<em>PlanilhasBase/CNPJs.xlsx</em>) atualizada caso entre ou saia alguma empresa.</li>
    <li><strong>2. Na sexta-feira às 08:00:</strong> o robô inicia sozinho, executa todas as consultas, salva os prints/downloads e encerra ao final sem pedir ENTER.</li>
    <li><strong>3. Conferência:</strong> basta abrir a pasta <em>Resultados/Relatorios</em> para visualizar o relatório consolidado da semana.</li>
</ul>

<h2>3. Como Configurar o Agendamento Automático (Agendador de Tarefas do Windows)</h2>
<p>Para que o Robô Societário rode sozinho no dia e horário programados (por exemplo, toda sexta-feira às 08:00), siga o passo a passo abaixo no <strong>Agendador de Tarefas (Task Scheduler)</strong> do Windows:</p>

<h3>1. Abrir e criar a tarefa</h3>
<ul>
    <li>Pressione as teclas <strong>Windows + R</strong>, digite <em>taskschd.msc</em> e pressione <strong>ENTER</strong> (ou pesquise por <strong>Agendador de Tarefas</strong> no Menu Iniciar).</li>
    <li>No painel da direita, clique em <strong>Criar Tarefa...</strong> <em>(use "Criar Tarefa", e não "Criar Tarefa Básica", para ter acesso a todas as configurações necessárias)</em>.</li>
</ul>

<h3>2. Aba "Geral"</h3>
<ul>
    <li>Em <strong>Nome</strong>, digite um identificador, como: <em>RoboSocietario - Consulta DET</em>.</li>
    <li>Em <strong>Opções de segurança</strong>, deixe marcada a opção <strong>Executar somente quando o usuário estiver conectado</strong>.</li>
    <li>Marque <strong>Executar com privilégios mais altos</strong> caso o usuário da máquina tenha restrições administrativas.</li>
</ul>

<div class="inline-note">
    <strong class="note-title">⚠ Atenção:</strong> Não utilize a opção <em>"Executar estando o usuário conectado ou não"</em>, pois o robô precisa da interface gráfica aberta para interagir com o Google Chrome e com a janela de seleção do certificado digital do Windows.
</div>

<h3>3. Aba "Disparadores" (Dia e Horário)</h3>
<ul>
    <li>Clique em <strong>Novo...</strong> e selecione <strong>Semanalmente</strong>.</li>
    <li>Defina <strong>Iniciar</strong> com a data atual e o horário desejado (exemplo: <em>08:00:00</em>).</li>
    <li>Marque o dia da semana desejado (exemplo: <strong>Sexta-feira</strong>), confirme que <strong>Habilitado</strong> está marcado e clique em <strong>OK</strong>.</li>
</ul>

<h3>4. Aba "Ações" (Programa e Modo Silencioso)</h3>
<ul>
    <li>Clique em <strong>Novo...</strong> e mantenha a ação <strong>Iniciar um programa</strong>.</li>
    <li>Em <strong>Programa/script</strong>, clique em <strong>Procurar...</strong> e selecione o arquivo <em>RoboSocietario.exe</em> dentro da pasta do sistema.</li>
    <li>Em <strong>Adicione argumentos (opcional)</strong>, digite exatamente: <strong>--agendado</strong> <em>(esse parâmetro desativa perguntas interativas e faz o robô encerrar sozinho ao terminar, sem aguardar você pressionar ENTER)</em>.</li>
    <li>Em <strong>Iniciar em (opcional)</strong>, informe o caminho da pasta onde o <em>RoboSocietario.exe</em> está salvo (exemplo: <em>C:\RoboSocietario</em>), sem aspas, e clique em <strong>OK</strong>.</li>
</ul>

<h3>5. Abas "Condições" e "Configurações"</h3>
<ul>
    <li>Na aba <strong>Condições</strong>: caso use um notebook, desmarque <em>"Iniciar a tarefa somente se o computador estiver conectado à energia CA"</em> se quiser permitir a execução mesmo fora da tomada.</li>
    <li>Na aba <strong>Configurações</strong>: mantenha marcada a opção <em>"Permitir que a tarefa seja executada sob demanda"</em> e ajuste <em>"Interromper a tarefa se ela for executada por mais de"</em> para um tempo compatível com a quantidade de empresas da sua planilha (exemplo: <em>8 horas</em>, ou desmarque para não interromper um lote grande).</li>
    <li>Clique em <strong>OK</strong> para salvar a tarefa.</li>
</ul>

<h3>6. Testar o agendamento</h3>
<ul>
    <li>Na lista de tarefas do Agendador, clique com o botão direito sobre <em>RoboSocietario - Consulta DET</em> e escolha <strong>Executar</strong>.</li>
    <li>Acompanhe para confirmar que o Chrome abre, realiza as consultas, grava os resultados na pasta Resultados e fecha a janela automaticamente ao final.</li>
</ul>

<div class="inline-note">
    <strong class="note-title">💡 Lembretes Importantes para a Execução Agendada:</strong> Para que a rotina agendada funcione sem falhas:
    <ul class="note-list">
        <li>• No horário programado, o computador deve estar <strong>ligado e com a sessão do usuário desbloqueada</strong> (com a tela bloqueada o Windows não exibe os diálogos de certificado para o robô clicar);</li>
        <li>• A planilha <em>PlanilhasBase/CNPJs.xlsx</em> deve estar <strong>fechada</strong> no Excel;</li>
        <li>• Caso existam itens em <em>CNPJs_Pendentes.xlsx</em> de uma execução anterior, no modo <em>--agendado</em> o robô processará automaticamente os pendentes se <em>CNPJs.xlsx</em> não existir, ou executará a base completa (<em>CNPJs.xlsx</em>) se ambas estiverem presentes.</li>
    </ul>
</div>

</body>
</html>
"""

FOOTER_TEMPLATE = """
<div style="width: 100%; margin: 0 18mm; padding-top: 4px; border-top: 1px solid #E2E8F0;
            font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 8.5pt; color: #718096;
            display: flex; justify-content: space-between; align-items: center;">
    <span>Robô Societário – DET / Gov.br</span>
    <span>Página <span class="pageNumber"></span> de <span class="totalPages"></span></span>
</div>
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
            display_header_footer=True,
            header_template="<div></div>",
            footer_template=FOOTER_TEMPLATE,
            margin={
                "top": "16mm",
                "bottom": "20mm",
                "left": "18mm",
                "right": "18mm",
            },
        )
        browser.close()
    print(f"[SUCESSO] Manual em PDF gerado: {caminho_saida}")


if __name__ == "__main__":
    destino = Path(__file__).resolve().parent / "Manual_do_Usuario_RoboSocietario.pdf"
    gerar_pdf(destino)
