# 📄 Documentação do Código VBA – Automação SAP

## 🛠️ Módulo 1: Criação de Notas (IW51)

📅 **Data:** Junho de 2025

### 🧠 Descrição Geral

Este código VBA foi desenvolvido para **automatizar a abertura de notas (IW51) no SAP**, preenchendo automaticamente os campos com dados extraídos de uma planilha Excel.  
Ele lê as informações linha a linha de uma aba especificada e interage com a interface do SAP GUI via SAP Scripting API, preenchendo os campos necessários para criação das notificações.

---

### 🔧 Pré-requisitos

- SAP GUI instalado e com o Scripting ativado.
- Permissão para executar scripts no SAP.
- Estar logado no SAP e na transação IW51.
- Planilha Excel com os dados corretamente preenchidos.

---

### 📑 Estrutura da Planilha Excel

| Coluna | Descrição                 | Obrigatório? |
|--------|---------------------------|--------------|
| A      | Status (pós-execução)     | ❌           |
| B      | Tipo de Nota              | ✔️           |
| C      | Título                    | ✔️           |
| D      | Número do Equipamento     | ✔️           |
| E      | Número de Série           | ✔️           |
| F      | Cliente                   | ❌           |
| G      | Prioridade                | ✔️           |
| H      | Notificador               | ✔️           |
| I      | Horímetro                 | ✔️           |
| J      | Origem (Repair Coding)    | ✔️           |
| K      | Descrição da Nota         | ✔️           |
| L      | Centro Localizador        | ✔️           |
| M      | Área Operacional          | ✔️           |
| N      | Embarcação                | ✔️           |
| O      | Centro de Custo           | ✔️           |
| P      | Job Code                  | ❌           |
| Q      | Comp Code                 | ❌           |
| R      | Natureza da Demanda       | ✔️           |
| S      | Pessoa de Contato         | ✔️ (se usado)|

---

### 🔍 Fluxo do Código

1. **Acesso à Planilha**  
   - Define a aba de onde os dados serão lidos.

2. **Busca da Última Linha**  
   - Identifica até onde há dados na coluna A.

3. **Conexão com o SAP**  
   - Inicia a comunicação com o SAP GUI via Scripting.

4. **Validação de Dados**  
   - Verifica se os campos obrigatórios estão preenchidos.
   - Exibe erro e encerra se houver campos obrigatórios vazios.

5. **Preenchimento dos Dados no SAP**  
   - Acessa a transação IW51.
   - Preenche os campos com base nos dados da planilha.

6. **Controle de Erros**  
   - Verifica se o SAP está na tela correta (menu IW51).

7. **Execução Linha a Linha**  
   - Processa apenas as linhas com coluna A vazia (ainda não criadas).

---

### ⚙️ Funcionalidades Principais

- Automação da criação de notificações IW51 no SAP.
- Validação de dados antes da execução.
- Alertas para campos obrigatórios ausentes.
- Maximização da janela SAP para melhor controle visual.

---

## 🛠️ Módulo 2: Consulta de Dados do Equipamento (IH08)

📅 **Data:** Junho de 2025

### 🧠 Descrição Geral

Este código VBA automatiza a consulta de informações de equipamentos no SAP, utilizando o **número de série** como referência.  
Os dados consultados (número do equipamento, nome e número do cliente) são preenchidos automaticamente em uma aba do Excel com base nos números de série informados.

---

### 🔧 Pré-requisitos

- SAP GUI instalado com SAP Scripting ativado.
- Permissão para execução de scripts no SAP.
- Estar logado no SAP na transação de consulta (IQ03, IQ09 ou equivalente).
- Planilha com aba chamada **"Equipamento"**, contendo na coluna A os números de série.

---

### 📑 Estrutura da Planilha Excel

| Coluna | Descrição               | Obrigatório? |
|--------|-------------------------|--------------|
| A      | Número de Série         | ✔️           |
| B      | Número do Equipamento   | ❌ (output)  |
| C      | Nome do Cliente         | ❌ (output)  |
| D      | Número do Cliente       | ❌ (output)  |

---

### 🔍 Fluxo do Código

1. **Acesso à Planilha**  
   - Aba "Equipamento".

2. **Busca da Última Linha**  
   - Verifica até onde há dados na coluna A.

3. **Conexão com o SAP**  
   - Estabelece conexão via SAP Scripting API.

4. **Validação de Dados**  
   - Executa somente se colunas B, C e D estiverem vazias.

5. **Execução no SAP**  
   - Acessa transação.
   - Preenche o número de série.
   - Executa a busca e navega até as abas de dados/parceiros.
   - Extrai:
     - Número do Equipamento
     - Nome do Cliente
     - Número do Cliente

6. **Retorno para o Excel**  
   - Preenche os dados extraídos nas colunas B, C e D.

7. **Encerramento e Finalização**  
   - Limpa campo de busca no SAP.
   - Volta para tela inicial para continuar o loop.
   - Ao final, exibe: `"Processo concluído!"`

---

### ⚙️ Funcionalidades Principais

- Automação da consulta de dados por número de série.
- Execução apenas para registros incompletos.
- Maximiza janela do SAP para evitar erros visuais.
- Otimiza tempo e reduz erros manuais.

---

### 📬 Contato

Caso tenha dúvidas ou sugestões, entre em contato comigo:  
**Darlan Dos Santos Monteiro**  
📧 darlanmonteirott@gmail.com  
📱 +55 (22) 99229-7791  


---

### 🧩 Observações Finais

> ⚠️ Estes códigos foram desenvolvidos para uso interno em ambientes corporativos com acesso ao SAP GUI e habilitação de scripting. Recomenda-se realizar testes em ambiente de homologação antes de utilizar em produção.
