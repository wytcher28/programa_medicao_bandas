# Programa Medidas de Banda

Programa em Python para processar a planilha `dados_log_bruto.xlsx`, aba `Dados`, e gerar uma nova planilha `dados_log_bruto_resultado.xlsx` com as abas `Resultado`, `Dados` e `Contas`.

## O que o programa faz

1. Lê a planilha `dados_log_bruto.xlsx`.
2. Usa sempre a aba `Dados` como fonte dos cálculos.
3. Cria a aba `Contas` com os dados originais e duas novas colunas:
   - `Transação = Bytes / Sessions`, com arredondamento para cima.
   - `Perfil`, conforme a regra:
     - `4k` para `Transação <= 8192`
     - `16k` para `Transação <= 32768`
     - `64k` para `Transação > 32768`
4. Ordena a aba `Contas` pela coluna A em ordem crescente.
5. Cria a aba `Resultado`, sem a linha `Grand Total`.
6. Ordena os perfis da aba `Resultado` pela coluna A em ordem crescente.
7. Mantém a linha `Total Geral` no final, sem entrar na ordenação.
8. Gera todos os números sem casas decimais, aplicando arredondamento para cima quando necessário.

## Arquivos principais

- `medidas_banda.py`: código-fonte principal.
- `requirements.txt`: dependências do projeto.
- `compilar_windows.bat`: script para gerar o executável no Windows.
- `dados_log_bruto.xlsx`: planilha de entrada.

## Como executar pelo Python

Abra o Prompt de Comando ou PowerShell dentro da pasta do projeto e execute:

```bash
pip install -r requirements.txt
python medidas_banda.py
```

O arquivo de saída será criado na mesma pasta:

```text
dados_log_bruto_resultado.xlsx
```

## Como gerar o executável

No Windows, execute:

```bash
compilar_windows.bat
```

O executável será criado em:

```text
dist\MedidasBanda.exe
```

## Como usar o executável

Coloque na mesma pasta:

```text
MedidasBanda.exe
dados_log_bruto.xlsx
```

Depois execute `MedidasBanda.exe`.

O programa criará automaticamente:

```text
dados_log_bruto_resultado.xlsx
```

## Regras importantes

A planilha de entrada deve possuir:

- Nome do arquivo: `dados_log_bruto.xlsx`
- Aba obrigatória: `Dados`
- Colunas obrigatórias na aba `Dados`:
  - `Application`
  - `Bytes`
  - `Sessions`

Se a planilha mudar, o programa recalcula tudo novamente com base nos dados atuais da aba `Dados`.


Existe um caminho para compilar o programa e transformar ele em um exe, porém precisa ser executado em um micro que nao bloqueie




