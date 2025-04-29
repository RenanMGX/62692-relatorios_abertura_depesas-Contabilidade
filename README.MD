# Relatórios de Abertura de Despesas - Contabilidade

## Descrição

Este projeto tem como objetivo atualizar e processar diversos relatórios de despesas, como Despesas Administrativas, Comerciais e Outras Despesas. O script principal (`main.py`) gerencia a execução do fluxo, lendo argumentos de um arquivo JSON e executando funções que extraem, processam e criam tabelas para cada tipo de despesa.

## Estrutura do Projeto

- **main.py**  
  Contém as classes principais:
  - **JsonArgs**: Responsável por ler os argumentos do arquivo `json/args.json`. Converte a data contida no JSON para o objeto `datetime` e, opcionalmente, remove o arquivo após a leitura.
  - **Execute**: Gerencia a execução do processo.
    - `start()`:  
      - Limpa mensagens anteriores e inicia o processo de atualização.
      - Lê os argumentos. Se não encontrar os argumentos, interrompe a execução.
      - Para cada tipo de despesa, baixa o relatório usando a classe `ExtrairRelatorio`, processa os dados com `AlimentarBase` e cria tabelas com `Tabela`.
      - Registra mensagens informativas e erros durante o processo, utilizando as classes `Informativo` e `Logs`.
    - `test()`: Testa a leitura dos argumentos sem deletar o arquivo JSON.

- **Entities/**  
  Contém módulos complementares:
  - `tratar_dados.py`: Possivelmente responsáveis pelas classes `AlimentarBase` e `Tabela`, que processam e organizam os dados dos relatórios.
  - `extrair_relatorio.py`: Contém a classe `ExtrairRelatorio`, utilizada para extrair os relatórios via o método `fbl3n`.
  - `dependencies/`: Inclui módulos auxiliares para argumentos, funções gerais, logs e exibição de mensagens informativas.

## Como Utilizar

1. **Preparar os Argumentos**  
   Crie um arquivo JSON chamado `args.json` dentro da pasta `json` (no mesmo diretório de `main.py`). Este arquivo deve conter:
   - `files_path`: Um dicionário com os caminhos dos arquivos de destino para:
     - `desp_adm` (Despesas Administrativas)
     - `desp_comercial` (Despesas Comerciais)
     - `outras_despesas` (Outras Despesas)
   - `date`: Uma data no formato ISO (ex.: `"2025-04-29T00:00:00"`).

2. **Executar o Script**  
   Abra o terminal na raiz do projeto e execute um dos comandos abaixo:
   - Para iniciar o processo de atualização:
     ```bash
     python main.py start
     ```
   - Para testar a leitura dos argumentos sem removê-los:
     ```bash
     python main.py test
     ```

3. **Fluxo de Execução**  
   Ao executar `python main.py start`, o fluxo é o seguinte:
   - O script inicia limpando as mensagens anteriores.
   - Chama o método `JsonArgs.get(delete_after=True)` para ler os argumentos do JSON.
   - Verifica se os caminhos dos arquivos para cada tipo de despesa estão informados.
   - Para cada relatório:
     - Utiliza `ExtrairRelatorio.fbl3n` para baixar o relatório correspondente.
     - Processa o relatório com `AlimentarBase` e cria a tabela apropriada com `Tabela`.
     - Registra mensagens de sucesso ou de erro utilizando `Informativo` e `Logs`.
   - Ao final, é registrada uma mensagem informando a conclusão do processo.

## Requisitos

- Python 3.7 ou superior.
- Dependências definidas nos módulos dentro do diretório `Entities/dependencies`.

## Contribuições

Contribuições são bem-vindas. Caso deseje melhorar ou corrigir alguma funcionalidade, sinta-se à vontade para abrir um _issue_ ou enviar um _pull request_.

## Licença

Este projeto está licenciado sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.