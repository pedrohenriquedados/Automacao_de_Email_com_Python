📊 Automação de Relatório Diário de Vendas por Loja
Projeto em Python que automatiza a geração e o envio de relatórios diários de desempenho de vendas para os gerentes de cada loja, via e-mail (Outlook), com indicadores comparados a metas e cores de sinalização (verde/vermelho).

🎯 Objetivo
Eliminar o trabalho manual de consolidar planilhas de vendas, calcular indicadores e enviar relatórios individuais para cada loja todos os dias. O script faz tudo isso automaticamente: lê as bases de dados, calcula os indicadores do dia e do ano, gera um backup em Excel por loja e dispara um e-mail personalizado para cada gerente.

⚙️ Funcionalidades
Consolidação de dados: leitura das bases de e-mails, lojas e vendas (Excel/CSV) com pandas.
Tratamento da base de vendas: merge com a base de lojas para trazer o nome de cada loja.
Separação por loja: criação de um dicionário com um DataFrame individual para cada loja.
Backup automático: as planilhas de cada loja são salvas em pastas organizadas por loja, com o nome do arquivo contendo mês, dia e nome da loja.
Cálculo de indicadores (diário e anual):
Faturamento
Diversidade de produtos vendidos
Ticket médio
Comparação com metas: cada indicador é comparado à meta definida e sinalizado em verde (dentro/acima da meta) ou vermelho (abaixo da meta).
Envio automático de e-mail via Outlook (win32com):
Corpo do e-mail em HTML com tabelas de indicadores do dia e do ano.
Cópia (CC) configurável.
Anexo automático da planilha de backup da loja.
🛠️ Tecnologias utilizadas
Python
pandas — leitura e tratamento dos dados
pathlib — manipulação de pastas e arquivos
pywin32 (win32com) — integração com o Outlook para envio dos e-mails
🔄 Como funciona (fluxo do projeto)
Importação das bases: leitura dos arquivos Emails.xlsx, Lojas.csv e Vendas.xlsx.
Tratamento: a base de vendas recebe o nome da loja via merge com a base de lojas.
Segmentação: as vendas são separadas em um dicionário, uma chave por loja.
Definição do dia de referência: a data mais recente da base de vendas é usada como "dia indicador".
Backup: cada planilha de loja é salva em Backup Arquivos Lojas/<Loja>/<mês>_<dia>_<loja>.xlsx.
Cálculo dos indicadores (dia e ano): faturamento, diversidade de produtos e ticket médio, comparados às metas pré-definidas.
Montagem do e-mail: corpo em HTML com tabelas coloridas (verde/vermelho) para os indicadores.
Envio: o e-mail é enviado automaticamente para o gerente de cada loja, com a planilha de backup em anexo.
📁 Estrutura de pastas esperada
Bases de Dados/
├── Emails.xlsx
├── Lojas.csv
└── Vendas.xlsx

Backup Arquivos Lojas/
└── <Nome da Loja>/
    └── <mês>_<dia>_<Nome da Loja>.xlsx
✅ Pré-requisitos
Python 3.x instalado
Bibliotecas: pandas, pywin32
Microsoft Outlook instalado e configurado na máquina (necessário para o envio dos e-mails)
Instalação das dependências:

pip install pandas pywin32
▶️ Como executar
Configure as bases de dados na pasta Bases de Dados.
Ajuste os caminhos dos arquivos e as metas de desempenho (meta_faturamento_dia, meta_faturamento_ano, etc.) conforme a necessidade do seu negócio.
Execute o notebook/script. Ele irá:
Gerar os backups das planilhas por loja.
Calcular os indicadores.
Enviar automaticamente um e-mail para cada gerente com o resumo do desempenho.
📧 Exemplo do e-mail enviado
O e-mail contém duas tabelas — indicadores do dia e indicadores do ano — mostrando:

Indicador	Valor	Meta	Cenário
Faturamento	R$ X	R$ Y	🟢/🔴
Diversidade de Produtos	X	Y	🟢/🔴
Ticket Médio	R$ X	R$ Y	🟢/🔴
Além disso, a planilha detalhada da loja é enviada em anexo.

💡 Possíveis melhorias futuras
Adicionar logs de execução e tratamento de erros mais robusto.
Migrar o envio de e-mail para uma solução multiplataforma (ex: smtplib), já que win32com depende do Windows/Outlook.
Parametrizar as metas em um arquivo de configuração externo.
Adicionar agendamento automático (ex: Windows Task Scheduler ou cron).
