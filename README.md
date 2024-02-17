# **SOBRE O REPOSITÓRIO:**:

Este é um repositório com projetos de automatização de processos via web-automation que desenvolvi.

## **PROJETO "01-leme\...":**

### **O QUE ESSE CÓDIGO FAZ?**

Esse código faz o input dos dados de mensuração dos indicadores de iniciativas do Sebrae Minas no Leme (sistema interno da empresa). 

Há um notebook comentado em  português no caminho: ".\01-leme\00-scripts" e o executável (.py) está em ".\01-leme\."

OBS: o código não é replicável fora da rede do Sebrae, então posteriormente irei postar um vídeo com ele rodando para que o processo fique mais claro.

### **METODOLOGIA**

A partir das bibliotecas selenium, webdriver e pandas, criamos um robô que executa as ações de clique e inputs no site do Leme com base em dados coletados pelas Unidades de Inteligência e Gestão Estratégica do Sebrae e também extraídos de dashboards disponíveis no Data Sebrae. Esses dados são alimentados em uma planilha modelo disponível em ".\01-dados\planilha-modelo.xlsx" e o robô utiliza essa planilha para fazer os inputs. 

Essa metodologia também é conhecida como Web Automation. O código está bem robusto com diversos recursos para agir diante de erros que podem acontecer durante a execução, sejam eles por conta de queda do servidor ou por questões relacionadas à má conexão de rede e ao próprio HTML do site. Porém, pode ser que novas alterações sejam necessárias com o tempo para a evolução e a robustez do robô. 


***Developers: Marcilio Duarte e Victor Hugo - Unidade de Inteligência Estratégica do Sebrae Minas.***
