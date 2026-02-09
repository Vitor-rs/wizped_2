# Wizped Office

Gerenciador de Fichas de Frequência e Alunos para escolas, integrado com Excel e Sponte Web.

## Estrutura do Projeto

* **`gerador_fichas_freq.xlsm`**: A ferramenta principal (Excel com Macros).
* **`src/`**: Scripts Python para importação de dados e automação.
* **`vba/`**: Código fonte VBA exportado para versionamento.
* **`ribbon/`**: XML da interface customizada do Excel.
* **`docs/`**: Documentação técnica e esquemas de banco de dados.
* **`assets/`**: Imagens e recursos visuais.

## Instalação e Configuração

Este projeto utiliza `uv` para gerenciamento de dependências Python.

1. **Instalar UV** (se ainda não tiver):

    ```powershell
    pip install uv
    ```

2. **Sincronizar dependências**:
    Na pasta raiz do projeto:

    ```powershell
    uv sync
    ```

## Uso

### Excel (Frontend)

Abra o arquivo `gerador_fichas_freq.xlsm`. A guia "Wizped" na faixa de opções (Ribbon) contém os botões principais:
* **Gerenciar Alunos**: Abre o formulário de cadastro.
* **Importar Cadastro**: Lê o PDF do Sponte Web e atualiza o banco de dados.

### Python (Backend)

O script `src/wizped_import.py` é chamado automaticamente pelo VBA, mas pode ser testado manualmente:

```powershell
uv run src/wizped_import.py "caminho/para/relatorio.pdf"
```

## Versionamento

Os arquivos VBA devem ser exportados periodicamente para a pasta `vba/` para garantir que as alterações no código sejam rastreadas pelo Git.
