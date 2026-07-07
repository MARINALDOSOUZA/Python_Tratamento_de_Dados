# SEEG - Analisador Dinâmico de Dados (GEE Estados)

> **Criado por:** 5507761

Este projeto é uma aplicação de terminal desenvolvida em Python para download, normalização, filtragem e análise estatística da base de dados de emissões de gases de efeito estufa (GEE) do SEEG (Sistema de Estimativas de Emissões e Remoções de Gases de Efeito Estufa). 

O sistema foi desenhado com foco em **resiliência**, **portabilidade** e **experiência do usuário (UX) no terminal**, operando de forma 100% autônoma em qualquer ambiente Windows.

---

## 🚀 Principais Funcionalidades

* **Gestão Inteligente de Diretórios:** Utiliza a API do Registro do Windows (`winreg`) para localizar a Área de Trabalho real do usuário (ignorando conflitos comuns com o OneDrive) e cria automaticamente o ambiente de trabalho (`Python_df`).
* **Download e Versionamento Automático:** Verifica a existência da base local. Se não existir, baixa da fonte oficial. Se existir, oferece a opção de atualizar, criando um backup automático (com *timestamp*) da base antiga antes de baixar a nova.
* **Normalização e Cache Otimizado (CSV):** Lê o arquivo `.xlsx` original de forma dinâmica (imune a erros de digitação nas abas), filtra linhas indesejadas (ex: `Bunker == 'Sim'`) mantendo as colunas originais intactas, e gera um arquivo de cache `.csv` para acelerar drasticamente as leituras futuras.
* **Filtros Interativos e Avançados:** Permite ao usuário navegar dinamicamente pelas colunas do DataFrame, listar valores únicos e aplicar filtros múltiplos empilhados, visualizando o resultado em tempo real.
* **Motor de Estatísticas Limpo:** Calcula estatísticas descritivas complexas (por ano, estado e atividade). O motor contorna o padrão visual do Pandas, removendo a notação científica e formatando a saída para leitura humana amigável (duas casas decimais).

---

## 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* **Pandas:** Para manipulação, limpeza e análise estatística pesada.
* **Requests:** Para integração web e download do arquivo Excel original.
* **Pathlib & OS:** Para manipulação moderna e multiplataforma de caminhos de arquivos.
* **Winreg:** Para interação direta com o sistema operacional Windows, garantindo portabilidade absoluta de diretórios.
* **Openpyxl:** Motor de leitura de planilhas Excel (dependência do Pandas).

---

## ⚙️ Pré-requisitos e Instalação

Certifique-se de ter o Python instalado na sua máquina. Para instalar as bibliotecas necessárias, utilize o arquivo `requirements.txt`.

1. Clone este repositório:
   ```bash
   git clone [https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git](https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git)
