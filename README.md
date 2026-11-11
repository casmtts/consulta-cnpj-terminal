# Consulta automatizada de situação cadastral de CNPJ (terminal)

Script em Python, sem dependências externas, que consulta em lote a situação cadastral de CNPJs na Receita Federal e gera uma tabela, um CSV e um relatório `.txt`.

## O problema

Quem precisa verificar a situação de muitas empresas, como fornecedores, clientes, parceiros, prestadores de serviço ou processos de *due diligence*, costuma fazer isso **manualmente**:

1. abrir o site de consulta de CNPJ;
2. digitar um CNPJ, resolver o CAPTCHA e esperar a resposta;
3. copiar situação, data, razão social e município para uma planilha;
4. repetir tudo para o próximo CNPJ.

Com dezenas ou centenas de empresas, isso leva horas, é repetitivo, sujeito a erro de digitação (um dígito trocado consulta a empresa errada) e não deixa registro organizado do que foi verificado, nem quando.

## A solução

Este script faz esse trabalho sozinho:

- lê os CNPJs de um arquivo `.txt`, `.csv` ou `.xlsx` (ou aceita CNPJs digitados direto na linha de comando);
- valida os dígitos verificadores **antes** de consultar, então entradas inválidas (CPF, números incompletos, erros de digitação) são apontadas com o motivo, sem gastar consulta;
- completa com zeros à esquerda CNPJs que o Excel teve o zero inicial removido;
- consulta a situação cadastral pela **BrasilAPI** e, se falhar, usa a **ReceitaWS** como alternativa;
- mostra o andamento e uma tabela no terminal;
- salva a tabela completa em **CSV** (pronta para abrir no Excel) e um **relatório `.txt`** agrupado por situação cadastral, com totais, detalhes e lista de erros e pendências.

Ctrl+C interrompe a consulta e salva o que já foi consultado.

## Requisitos

- Python 3.8 ou superior (testado com 3.9)
- Conexão com a internet
- Apenas para ler arquivos `.xlsx`: `pip3 install openpyxl` (ou salve a planilha como CSV e não precisa instalar nada)

Não há dependências obrigatórias: o script usa só a biblioteca padrão do Python.

## Como usar

Lista em arquivo de texto (um CNPJ por linha, ou separados por espaço, vírgula, ponto e vírgula ou tab):

```bash
python3 consulta_cnpj_terminal.py base.txt
```

Planilha CSV, salvando a tabela e o relatório:

```bash
python3 consulta_cnpj_terminal.py base.csv -o resultado.csv -r relatorio.txt
```

CNPJs digitados direto:

```bash
python3 consulta_cnpj_terminal.py 33.000.167/0001-01 11.222.333/0001-81
```

Somente as empresas com situação ATIVA na tabela, com 2 segundos entre consultas, sem imprimir a tabela no fim:

```bash
python3 consulta_cnpj_terminal.py base.txt --filtro ATIVA --intervalo 2 --sem-tabela -r relatorio.txt
```

### Opções

| Opção | Descrição |
| --- | --- |
| `entradas` | arquivo `.txt`, `.csv` ou `.xlsx` com CNPJs, e/ou CNPJs digitados direto |
| `-i`, `--intervalo` | segundos entre consultas (padrão `1.5`) |
| `-o`, `--csv` | salva a tabela completa em CSV, com CNPJ sem máscara |
| `-r`, `--relatorio` | salva um relatório `.txt` com resumo e situação cadastral |
| `-f`, `--filtro` | mostra na tabela só as linhas que contêm este texto |
| `-c`, `--coluna` | nome da coluna com os documentos em planilhas e CSV (padrão: coluna com `cnpj` ou `cpf` no cabeçalho) |
| `--sem-tabela` | não imprime a tabela no fim, só o resumo |

## Exemplo de saída

Exemplo ilustrativo, com empresas fictícias:

```
[1/3] 11.222.333/0001-81 -> ATIVA | EMPRESA EXEMPLO LTDA (BrasilAPI)
[2/3] 33.000.167/0001-01 -> BAIXADA | COMERCIO MODELO ME (BrasilAPI)
[3/3] 12.345.678/0001-99 -> ERRO: dígitos verificadores não conferem

Resumo: total 3 | consultados 2 | ativas 2 | não ativas 0 | erros 1 | pendentes 0
```

O relatório `.txt` tem esta estrutura:

```
RELATÓRIO DE CONSULTA DE SITUAÇÃO CADASTRAL (RECEITA FEDERAL) - CNPJ
RESUMO
  Total na lista: ...
  Consultados com sucesso: ...
  Com erro: ...
SITUAÇÃO CADASTRAL
  ATIVA: ...
  BAIXADA: ...
DETALHES - ATIVA (...)
  CNPJ | data da situação | UF | razão social
ERROS (...)
  CNPJ | motivo
```

## Formato da entrada

- **`.txt`**: CNPJs separados por quebra de linha, espaço, vírgula, ponto e vírgula, tab ou `|`.
- **`.csv` / `.xlsx`**: se houver uma coluna com `cnpj` ou `cpf` no cabeçalho, só ela é lida. Senão, o script lê todas as células. Use `--coluna` para escolher a coluna pelo nome.
- Máscara (`00.000.000/0000-00`) é aceita ou dispensada.
- Só CNPJ é consultado. Um CPF (11 dígitos) é apontado como erro com o motivo.
- CNPJs repetidos na entrada são consultados uma vez só.

## Limites e cuidados

- **Serviços de terceiros.** BrasilAPI e ReceitaWS são APIs públicas mantidas por terceiros, não pela Receita Federal. Podem ficar indisponíveis, mudar de formato ou limitar o número de requisições por minuto. Por isso o intervalo entre consultas é configurável com `--intervalo`. Se houver muitos erros, aumente o intervalo.
- **Não substitui documento oficial.** Para comprovação oficial (por exemplo, em processos licitatórios ou contratuais), use a consulta e a emissão de certidão no site oficial da Receita Federal.
- **Dados da sua lista.** As planilhas, CSVs e relatórios gerados podem conter CNPJs e nomes de empresas reais. Não os suba para repositórios públicos. O `.gitignore` deste repositório já ignora os arquivos de saída mais comuns.
- **Uso responsável.** Respeite os termos de uso das APIs e não reduza o intervalo entre consultas a ponto de sobrecarregar os serviços.

## Estrutura

```
consulta-cnpj-terminal/
├── consulta_cnpj_terminal.py   # o script
├── README.md
└── .gitignore
```
