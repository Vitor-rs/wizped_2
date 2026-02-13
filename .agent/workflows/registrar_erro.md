---
description: Workflow para registrar erros relatados pelo usuário
---

Quando o usuário relatar erros no projeto (ex: erros de compilação, bugs de lógica, falhas na Ribbon), siga este procedimento para registrar e documentar a solução.

0. **Consultar Log Existente** (Para problemas Complexos/"Chatos")
    * **Antes de iniciar a investigação**, abra o arquivo `docs/error_log.md`.
    * Verifique se o erro já foi registrado anteriormente ou se há erros similares que possam dar contexto.

1. **Identificar o Erro**
    * Verificar a mensagem de erro exata e o contexto (ex: qual botão foi clicado, qual macro falhou).
    * Classificar: Erro de Compilação, Erro de Runtime, Erro de Lógica.

2. **Investigar a Causa**
    * Analisar os arquivos relevantes (VBA, Python, XML).
    * Verificar duplicidade de nomes, sintaxe inválida, referências quebradas.

3. **Aplicar a Solução**
    * Realizar as correções necessárias nos arquivos do projeto.
    * Verificar se a correção resolve o problema relatado.

4. **Registrar no Log de Erros (`docs/error_log.md`)**
    * Abrir (ou criar) o arquivo `docs/error_log.md` na raiz do projeto.
    * Adicionar uma nova entrada com a data atual e título do erro.
    * Seguir o template:

        ```markdown
        ## [YYYY-MM-DD] [Título Resumido do Erro]
        - **Sintoma:** O que o usuário relatou ou o erro exibido.
        - **Causa:** Explicação técnica da origem do problema.
        - **Solução:** Passos realizados para corrigir.
        - **Arquivos Afetados:** Lista de arquivos modificados.
        ```

5. **Notificar o Usuário**
    * Informar ao usuário que o erro foi corrigido e registrado.
    * Fornecer instruções para atualizar o projeto se necessário.
