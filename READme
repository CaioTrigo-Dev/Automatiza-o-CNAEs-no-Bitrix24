# 🧩 Automação de Inserção de CNAEs no Bitrix24

## 📋 Descrição

Este script foi desenvolvido para automatizar o processo de **inserção de dados de CNAEs (Classificação Nacional de Atividades Econômicas)** no sistema **Bitrix24**, utilizando Python e automação de interface gráfica com o **PyAutoGUI**.

A automação acessa a página do cliente no Bitrix24, lê os dados de uma planilha Excel com as informações dos CNAEs fornecidos pelo **IBGE**, e realiza a inserção automática dessas informações diretamente na plataforma.

---

## ⚙️ Tecnologias Utilizadas

* **Python 3.10+**
* **PyAutoGUI** — Automação de teclado e mouse
* **Pandas** — Leitura e manipulação da planilha Excel
* **Pyperclip** — Copiar texto para área de transferência (para evitar erros de digitação automática)
* **Time** — Controle de pausas e intervalos durante a automação

---

## 📁 Estrutura Esperada dos Arquivos

```
📂 Projeto
 ├── Automatizacao.py
 ├── adiciona.png           # Captura de tela do botão "Adicionar" do Bitrix24
 ├── Tabela.xlsx            # Planilha com dados de CNAEs
```

### Estrutura da planilha `Tabela.xlsx`

A planilha deve conter as seguintes colunas:

| CNAE      | DESCRIÇÃO          |
| --------- | ------------------ |
| 0111-3/01 | Cultivo de cereais |
| 0112-1/02 | Produção de soja   |
| ...       | ...                |

---

## 🚀 Como Utilizar

1. **Instale as dependências**

   ```bash
   pip install pyautogui pandas pyperclip
   ```

2. **Ajuste os caminhos dos arquivos no script**
   No início do arquivo `Automatizacao.py`, modifique:

   ```python
   caminho = r'C:\Users\<SEU_USUARIO>\Downloads\adicionar.png'
   df = pd.read_excel(r'C:\Users\<SEU_USUARIO>\Downloads\Tabela.xlsx')
   ```

3. **Atualize a URL do Bitrix24**

   ```python
   url = 'https://promobrindes.bitrix24.com.br/page/inteligencia_dados/...'
   ```

4. **Execute o script**

   ```bash
   python Automatizacao.py
   ```

5. O script abrirá o navegador automaticamente, acessará a página do Bitrix24 e começará a inserir os CNAEs da planilha.

---

## 🧠 Funcionamento do Script

1. **Abertura do navegador e acesso ao Bitrix24**
   O script abre o Google Chrome e navega até o link configurado.

2. **Localização do botão "Adicionar"**
   Utiliza `pyautogui.locateCenterOnScreen()` para encontrar o botão na tela por meio da imagem `adicionar.png`.

3. **Iteração sobre os dados da planilha**
   Para cada linha do arquivo Excel, o script:

   * Clica no botão “Adicionar”
   * Copia o texto do CNAE e da descrição
   * Cola as informações no campo adequado
   * Confirma e passa para o próximo item

4. **Tratamento de exceções**
   Caso o botão “Adicionar” não seja encontrado, o script exibe a mensagem:

   ```
   Botão 'Adicionar' não encontrado na tela
   ```

   e continua a execução.

---

## 🧩 Personalização

* Você pode alterar os **intervalos de pausa** com:

  ```python
  pyautogui.PAUSE = 0.5
  ```

  Ajuste para valores maiores caso o sistema do Bitrix24 demore para responder.

* Caso o layout da página mude, será necessário **atualizar a imagem `adicionar.png`** com uma nova captura de tela do botão.

---

## ⚠️ Cuidados

* Não mova o mouse nem use o teclado durante a execução.
* Evite que outras janelas fiquem sobre o navegador.
* Verifique se o Bitrix24 está logado antes de rodar o script.
* Mantenha a resolução da tela constante (a detecção de imagem é sensível a isso).

---

## 🧑‍💻 Autor

**Caio Cesar**
Automação de Processos | Python | Bitrix24 Integrations
