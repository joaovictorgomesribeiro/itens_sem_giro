📊 Automação de Relatórios SAP Analytics Cloud (SAC)
Este projeto automatiza a extração, tratamento e distribuição de relatórios mensais que estão sem giro a mais de 90 dias provenientes do SAP Analytics Cloud. O objetivo é substituir o processo manual de gerar 24 relatórios individuais por um fluxo único e automatizado.

🚀 O Problema
Anteriormente, eu realizava 24 extrações manuais do SAC, seguidas de formatação individual em Excel. Esse processo consumia tempo considerável (Aproximadamente 2h) e era suscetível a erros de formatação e filtros manuais.

🛠️ A Solução
O projeto utiliza Python (Pandas e Openpyxl) para processar um único arquivo consolidado e realizar as seguintes etapas:

Data Cleaning: Remove cabeçalhos inúteis e limpa linhas de metadados do SAP.

Data Transformation: Renomeia colunas para nomes amigáveis e converte tipos de dados (strings para numéricos).

Formatação Estética: * Aplica estilos de cabeçalho (Azul Escuro, Fonte Branca, Negrito).

Formata números com separadores de milhar e remove casas decimais.

Ajusta automaticamente a largura das colunas e fixa a largura de colunas críticas (Ex: Descrição).

Cálculos Automáticos: Insere uma linha de Resumo no topo de cada arquivo com a Soma do estoque/SKUs e a Média de dias sem giro.

Segregação: Filtra e gera 24 arquivos .xlsx individuais nomeados por filial.

Interface Web: Implementação de uma UI via Streamlit para que o usuário faça o upload do arquivo bruto e baixe os relatórios prontos em um arquivo .zip.

📦 Tecnologias Utilizadas
Python

Pandas: Manipulação e filtragem de dados.

Openpyxl: Engine para formatação avançada e inserção de fórmulas de Excel.

Streamlit: Framework para criação da interface web do projeto.

📈 Resultados
Redução de Tempo: O processo que levava horas agora é concluído em segundos.

Padronização: Todos os 24 relatórios seguem rigorosamente o mesmo padrão visual e de cálculo.

Facilidade de Uso: Interface simples que não exige instalação de software na máquina local.
