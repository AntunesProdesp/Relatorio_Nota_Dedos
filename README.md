# Relatório de Notas dos Dedos (CDB_LOCAL)

O script `agrupa_notas_cdb_local.py` junta os consolidados diários `CDB_LOCAL_AAAAMMDD` em uma base única com as notas dos 10 dedos, uma linha por coleta, e gera **um arquivo para cada cliente**.

## O que o script faz

1. Procura na pasta de origem os arquivos `CDB_LOCAL_AAAAMMDD` (`.xlsx`, `.xlsm`, `.xlsb`, `.xls` ou `.csv`).
2. Lê apenas as colunas necessárias, o que deixa a leitura bem mais rápida.
3. Monta uma linha por coleta, com as colunas abaixo.
4. Se a mesma `idColeta` aparecer em mais de um arquivo, mantém só a do arquivo mais recente.
5. Filtra o período pela `dataInclusao` (opcional).
6. Salva um `.csv` e um `.xlsx` para cada cliente.

### Colunas do resultado

| Coluna | Origem no consolidado |
|---|---|
| `CLIENTE` | `cliente` |
| `NOME_POSTO` | `NOME FANTASIA` (nome do posto) |
| `FINALIDADE` | `finalidade` |
| `CPF_CIDADAO` | `cpf` (só dígitos, completado com zeros até 11) |
| `ID_COLETA` | `idColeta` |
| `ANO`, `MES_ANO`, `DATA` | `dataInclusao` (ver `ORIGEM_DATA`) |
| `1 - polegar da mao direita` … `10 - minimo da mao esquerda` | colunas dos dedos, com o mesmo nome |

Os nomes das colunas de origem são comparados sem considerar acentos, maiúsculas ou espaços extras.

## Instalação

Requer Python 3.9 ou mais recente.

```bash
pip install -r requirements.txt
```

O `python-calamine` é opcional. Sem ele, o script lê os `.xlsx` com o `openpyxl`, que é mais lento.

## Como usar

```bash
python agrupa_notas_cdb_local.py
```

Para ler os consolidados de outra pasta, passe o caminho como argumento:

```bash
python agrupa_notas_cdb_local.py "C:\outra\pasta"
```

As pastas padrão ficam no drive `T:`, que precisa estar mapeado na máquina.

### Arquivos gerados

Para cada cliente, na pasta `PASTA_SAIDA`:

```
AGRUPADO_CDB_LOCAL_<CLIENTE>.csv    (separador ";", UTF-8, datas DD/MM/AAAA)
AGRUPADO_CDB_LOCAL_<CLIENTE>.xlsx   (com filtro e cabeçalho congelado)
```

O nome do cliente vai para o nome do arquivo sem acentos e sem caracteres especiais. Por exemplo, `Secretaria da Saúde / SP` vira `SECRETARIA_DA_SAUDE_SP`. O `.xlsx` não é gerado quando passa de 1.048.575 linhas (limite do Excel); nesse caso, use o `.csv`.

## Configuração

As opções ficam no bloco `CONFIGURACAO`, no início do script:

| Opção | Padrão | Para que serve |
|---|---|---|
| `PASTA_ORIGEM` | `T:\...\CDB LOCAL\2026` | Pasta com os consolidados `CDB_LOCAL_AAAAMMDD`. |
| `PASTA_SAIDA` | `T:\...\Relatorios_Automatizados` | Onde salvar. Com `None`, salva na subpasta `_AGRUPADO` da pasta de origem. |
| `NOME_SAIDA` | `AGRUPADO_CDB_LOCAL` | Início do nome dos arquivos gerados. |
| `GERAR_XLSX` | `True` | Gera o `.xlsx` além do `.csv`. |
| `BUSCAR_SUBPASTAS` | `False` | Procura os consolidados também nas subpastas. |
| `SEPARAR_POR_CLIENTE` | `True` | Gera um arquivo por cliente. |
| `GERAR_ARQUIVO_GERAL` | `False` | Gera também um arquivo com todos os clientes juntos. |
| `ORIGEM_DATA` | `"inclusao"` | De onde vem a `DATA`: `"inclusao"` (`dataInclusao`), `"coleta"` (`dadosColeta.data`) ou `"arquivo"` (data no nome do arquivo). |
| `DATA_INICIO` / `DATA_FIM` | `None` | Filtro de período pela `DATA`, no formato `"DD/MM/AAAA"`. `None` = sem limite. |
| `REMOVER_COLETAS_DUPLICADAS` | `True` | Mantém só a versão mais recente de uma `idColeta` repetida entre arquivos. |
| `FUSO_HORARIO` | `America/Sao_Paulo` | Fuso usado para converter datas que vêm em UTC (ex.: `2026-10-23T17:33:00Z`). |

Quando a data escolhida está vazia ou inválida em uma linha, o script usa a data do nome do arquivo e mostra um aviso.

### Exemplo: só as inclusões de outubro de 2026

```python
DATA_INICIO = "01/10/2026"
DATA_FIM = "31/10/2026"
```

## Problemas comuns

- **"Pasta nao encontrada"**: o drive `T:` não está mapeado ou o caminho em `PASTA_ORIGEM` mudou.
- **"Nao consegui gravar ... (arquivo aberto?)"**: feche o arquivo no Excel e rode de novo.
- **"colunas nao encontradas"**: o consolidado não tem alguma das colunas esperadas. Elas ficam em branco no resultado e o restante é processado normalmente.
- **Um arquivo com erro**: ele é listado no final e não interrompe os outros.
