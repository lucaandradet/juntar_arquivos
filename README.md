# Juntar Arquivos

Ferramenta em Python com interface gráfica que junta todos os arquivos de uma pasta que tenham a mesma extensão em um único arquivo consolidado, mantendo **apenas um cabeçalho**.

Pensada para o dia a dia de quem recebe dados divididos em vários arquivos (extrações mensais, partes de um export, relatórios por região etc.) e precisa de tudo em um arquivo só, sem abrir e copiar um por um.

## Funcionalidades

- **Seleção de pasta por janela:** nada de linha de comando, basta escolher a pasta.
- **Arquivo gerado com o nome da pasta:** a pasta `Vendas_2026` com arquivos `.csv` gera `Vendas_2026/Vendas_2026.csv`.
- **Um único cabeçalho:** o cabeçalho do primeiro arquivo é mantido e os dos demais são descartados.
- **Escolha de extensão:** se a pasta tiver arquivos de mais de um tipo (ex.: `.csv` e `.txt`), o programa pergunta qual deles juntar.
- **Janela de progresso:** mostra a quantidade de arquivos, o arquivo em processamento ("Arquivo 3 de 10") e a porcentagem do andamento.
- **Tela de carregamento** na abertura do programa.
- **Vários lotes seguidos:** ao final, pergunta se deseja juntar outra pasta, já abrindo a seleção na última pasta usada.
- **Avisos automáticos:** informa se algum arquivo tiver cabeçalho diferente do primeiro ou estiver vazio.
- **Números formatados** no padrão brasileiro (ex.: `23.461.002 linhas`).
- **Janelas centralizadas** na tela.

## Formatos suportados

| Tipo | Extensões | Como é feito |
|---|---|---|
| Texto | `.csv`, `.txt`, `.tsv` e outros arquivos de texto | Concatenação linha a linha, sem alterar delimitadores nem o conteúdo dos dados |
| Excel | `.xlsx`, `.xlsm`, `.xls` | Leitura com pandas, com todas as colunas tratadas como texto para preservar zeros à esquerda |

Detalhes do tratamento de arquivos texto:

- Aceita arquivos em **UTF-8** e **Latin-1 (ANSI)**, inclusive misturados na mesma pasta.
- O arquivo final é gravado em **UTF-8 com BOM**, para o Excel abrir os acentos corretamente.
- Linhas em branco são ignoradas.
- Quebras de linha do Windows (`\r\n`) e do Linux (`\n`) são tratadas.
- Arquivos temporários do Excel (`~$...`) e o próprio arquivo consolidado de execuções anteriores não entram na junção.

## Requisitos

- Python 3.9 ou superior
- Tkinter (já vem com o Python no Windows)
- Apenas para arquivos Excel:

```bash
pip install pandas openpyxl
```

Para juntar somente arquivos texto não é preciso instalar nada.

## Como usar

1. Baixe o arquivo `juntar_arquivos.py`.
2. Execute com dois cliques ou pelo terminal:

   ```bash
   python juntar_arquivos.py
   ```

3. Aguarde a tela de carregamento e selecione a pasta com os arquivos.
4. Se houver mais de uma extensão na pasta, escolha qual deseja juntar.
5. Acompanhe o andamento na janela de progresso.
6. Ao final, veja o resumo com a quantidade de arquivos e linhas e o caminho do arquivo gerado.
7. Responda se deseja juntar outra pasta ou encerrar o programa.

## Observações

- Os arquivos devem ter **as mesmas colunas, na mesma ordem**. Se algum cabeçalho for diferente, o programa junta mesmo assim e mostra um aviso no final.
- Se o arquivo consolidado já existir na pasta, ele é **sobrescrito**.
- A janela de progresso não pode ser fechada durante a execução, para não gerar um arquivo incompleto.
