---
description: Rule: Always log reported errors
---

# Regra: Registrar Erros Relatados

Sempre que o usuário relatar um erro significativo (compilação, runtime, lógica, UX) no projeto:

0. **CONSULTAR O LOG EXISTENTE:** Se o problema for complexo ou recorrente ("chato"), verifique o arquivo `docs/error_log.md` ANTES de investigar, para entender o histórico.

1. **Investigar e Corrigir** o erro.
2. **Registrar** o erro e a **SOLUÇÃO DETALHADA** no arquivo `docs/error_log.md`.
    * A solução deve explicar O QUE foi feito e POR QUE, facilitando consultas futuras.
3. **Seguir o workflow** definido em `.agent/workflows/registrar_erro.md`.

Esta regra é obrigatória para manter o histórico de problemas e soluções do projeto.
