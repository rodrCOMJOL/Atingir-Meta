---

# Automação de Processos no Excel com `pyautogui`

Este script automatiza o procedimento de "Atingir Metas" em uma planilha Excel utilizando a biblioteca `pyautogui`. Ele navega pela interface do Excel, preenchendo células específicas com valores extraídos de uma planilha, e interage com a funcionalidade de "Atingir Metas" do Excel. 

## Funcionalidades

- **Automação do Excel**: Automatiza a inserção de valores nas células especificadas na planilha.
- **Interrupção Segura**: Monitora continuamente a tecla `Esc` para interromper a execução de forma segura, exibindo uma mensagem informativa ao usuário.
- **Verificação de Metas**: Verifica o estado da operação "Atingir Metas" usando imagens como referência, aguardando até que a operação esteja concluída antes de continuar.
- **Notificações ao Usuário**: Exibe mensagens ao usuário quando a execução é finalizada ou interrompida.

## Estrutura do Código

1. **Importações**:
   - `pyautogui`: Para automação de cliques e digitação na interface do Excel.
   - `time`: Para pausas controladas durante a execução.
   - `pandas`: Para manipulação da planilha Excel.
   - `tkinter`: Para criação de janelas pop-up que informam o usuário.
   - `keyboard`: Para monitorar a tecla `Esc` e interromper a execução se necessário.

2. **Função `wait(mensagem_imagem)`**:
   - Verifica se uma imagem específica está presente na tela, o que indica o progresso ou conclusão da ação.
   - Retorna `True` se a imagem for encontrada, permitindo que o script continue.

3. **Função `check_esc()`**:
   - Verifica continuamente se a tecla `Esc` foi pressionada.
   - Se pressionada, exibe uma mensagem informativa e encerra a execução do script.

4. **Loop Principal**:
   - Lê uma planilha Excel usando `pandas`.
   - Itera sobre as linhas da planilha, pulando linhas com valores nulos em colunas específicas.
   - Automatiza a inserção de dados no Excel, utilizando coordenadas de células específicas.

5. **Finalização**:
   - Exibe uma mensagem informativa ao usuário quando o script é concluído com sucesso.

## Como Usar

1. **Pré-requisitos**:
   - Python 3.x
   - Instalação das bibliotecas necessárias: `pyautogui`, `pandas`, `tkinter`, `keyboard`, `openpyxl`.

2. **Execução**:
   - Coloque as imagens de referência (`wait1.jpg` e `wait2.jpg`) no diretório do script.
   - Certifique-se de que a planilha Excel (`planilha.xlsx`) está no diretório correto.
   - Execute o script e siga as instruções na interface do Excel.

3. **Interrupção**:
   - Pressione `Esc` a qualquer momento para interromper o processo.

## Notas

- O script utiliza imagens para detectar quando o Excel finaliza certas operações, por isso, certifique-se de que as imagens de referência estão corretas e de que o Excel está visível na tela durante a execução.
- A planilha Excel deve ter dados válidos nas colunas especificadas pelo código (colunas de índice 5 e 52).

---
