# Registro de Erros e Soluções (Wizped Office)

Este arquivo documenta erros significativos relatados pelo usuário e suas respectivas causas e soluções.

## [2026-02-12] Erro na Ribbon: "Cannot run the macro"

- **Sintoma:** Ao clicar no botão "Gerenciar Alunos" da Ribbon, o Excel exibe a mensagem "Cannot run the macro... The macro may not be available in this workbook".
- **Causa:** Erro de compilação no Projeto VBA ("Ambiguous Name detected"). O formulário `frmAlunos` não pôde ser instanciado porque continha procedimentos duplicados:
    1. Duplicidade na declaração do controle `lblData` em `VBA_01_CriarFormulario.bas`.
    2. Duplicidade na rotina `private Sub txtNome_Change()` em `VBA_02_FormLogica.bas` (uma versão original e outra adicionada para o recurso Active State).
- **Solução:**
    1. Removida a linha duplicada de `lblData` em `VBA_01`.
    2. Removida a rotina `txtNome_Change` duplicada em `VBA_02` e unificada a lógica de `mFormModificado = True` com `VerificarBloqueioAtivo`.
    3. Atualizada a rotina `txtID_Change` para incluir o flag de modificação.
- **Arquivos Afetados:**
  - `vba\VBA_01_CriarFormulario.bas`
  - `vba\VBA_02_FormLogica.bas`
