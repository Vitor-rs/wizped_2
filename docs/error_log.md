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

## [2026-02-12] Erro de Compilação - Macro 'OnGerenciarAlunos' não encontrada
- **Sintoma:** O usuário relatou 'Não é possível executar a macro OnGerenciarAlunos'.
- **Causa:** Erro de sintaxe (duplo 'Next r') introduzido em 'vba/VBA_02_FormLogica.bas'.
- **Solução:** Removida a linha duplicada.
- **Arquivos Afetados:** 'vba/VBA_02_FormLogica.bas'


## [2026-02-12-FixV2] Refatoração de Função de Busca (Resolução Definitiva)
- **Sintoma:** O erro 'Cannot run the macro' persistiu mesmo após correção de sintaxe.
- **Causa Provável:** Ambiguidade de nomes ou conflito entre funções privadas duplicadas (VBA_02 e VBA_04) ou problema de compilação em escopo privado.
- **Solução:** Removida a cópia privada de BuscarIDExperienciaPorNome de VBA_02. Renomeada a função em VBA_04 para GlobalBuscarIDExperiencia e tornada Pública para uso global.
- **Arquivos Afetados:** ba/VBA_02_FormLogica.bas, ba/VBA_04_ModImportarCadastro.bas


## [2026-02-12-FixV3] Erro de Sintaxe ao Copiar/Colar (Attribute VB_Name)
- **Sintoma:** Linhas como 'Attribute VB_Name = ...' aparecendo em vermelho no editor VBA. Mensagem 'Compile Error: Expected: Statement'.
- **Causa:** O usuário está copiando o conteúdo dos arquivos .bas e colando no editor, em vez de importar os arquivos. As linhas de metadados Attribute não são válidas dentro do corpo do módulo.
- **Solução:** Comentadas as linhas Attribute VB_Name em todos os módulos (VBA_01, VBA_03, VBA_04, VBA_05).
- **Arquivos Afetados:** Todos os módulos .bas exceto VBA_02 (que já não tinha).

