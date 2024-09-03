# Automação em Python

Este repositório contém scripts em Python para automação de tarefas, incluindo a criação de planilhas Excel, a abertura de páginas HTML em um navegador específico e a gestão de tarefas.

## Scripts Disponíveis

### 1. `Planilha.py`

Este script cria uma planilha Excel com informações sobre disciplinas e notas.

- **Funcionalidades**:
  - Cria uma planilha com as colunas: 'Disciplina', 'Ap1', 'Ap2', 'Ext', 'Média Final'.
  - Adiciona dados predefinidos para várias disciplinas.
  - Ajusta a largura das colunas com base no conteúdo.
  - Aplica formatação numérica às células de notas.
  - Salva o arquivo como `Notas.xlsx`.
  - Abre o arquivo gerado automaticamente após a criação.

- **Requisitos**:
  - Biblioteca `openpyxl`. Instale com `pip install openpyxl`.

- **Como Executar**:

  python Planilha.py

### 2. `Tarefas.py`

Este script abre um arquivo HTML no navegador Opera GX.

- **Funcionalidades**:
  - Abre a página HTML especificada usando o navegador Opera GX.
  
- **Requisitos**:
  - O navegador Opera GX deve estar instalado no caminho especificado.

- **Como Executar**:
  1. Certifique-se de que o navegador Opera GX está instalado no caminho especificado.
  2. Ajuste o caminho para o arquivo HTML no script para refletir o local correto no seu sistema.
  3. Execute o script com o comando:

     python Tarefa.py
