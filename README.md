# Transferir dados entre duas planilhas Excel

Programa em Python para copiar valores de uma ou varias colunas da planilha 1 para a planilha 2, sem alterar formatacao.

Agora o projeto possui duas interfaces:

- Interface web moderna (`web_app.py`) com mapeamento visual por dropdown.
- Interface desktop Tkinter (`main.py`) ja existente.

## O que ele faz

- Voce escolhe:
  - arquivo de origem (planilha 1)
  - arquivo de destino (planilha 2)
  - arquivo de saida (gera uma copia da planilha 2)
  - aba de origem e aba de destino
  - mapeamento de colunas de origem -> destino
  - separador para combinar duas colunas em uma
  - linha inicial de origem e linha inicial de destino
- Copia os dados percorrendo as linhas com dados, linha a linha.
- Escreve somente os valores no destino (nao modifica estilos de celulas).

## Exemplo

Para enviar `A2` da planilha 1 para `A3` da planilha 2:

- Mapeamento: `A:A`
- Linha origem inicial: `2`
- Linha destino inicial: `3`

Exemplo com varias colunas:

- Mapeamento por letras: `B:A,D:B,F:C,H:D,I:E`
- Mapeamento por numeros: `2:1,4:2,6:3,8:4,9:5`

Exemplo combinando duas colunas de origem em uma de destino:

- `B+C:A` (combina colunas B e C e grava em A)
- `2+4:1` (combina colunas 2 e 4 e grava na coluna 1)
- Defina o separador no campo "Separador para combinacao" (ex: ` - `, `/`, `|`)

## Requisitos

- Python 3.10+

Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Como executar

### Versao web (recomendada)

```bash
python web_app.py
```

Depois abra no navegador:

```text
http://127.0.0.1:5000
```

Recursos da versao web:

- Detecta automaticamente abas e colunas ao enviar os arquivos.
- Mostra as colunas disponiveis para origem e destino.
- Permite montar mapeamentos visualmente.
- Permite mapeamento simples (`A -> B`) e combinado (`A + C -> D`) com separador.
- Faz download da planilha de destino atualizada.

### Versao desktop (Tkinter)

```bash
python main.py
```

## Observacoes

- O programa gera um novo arquivo de saida por seguranca.
- Se marcar "Copiar apenas linhas preenchidas", ele ignora linhas em que todas as colunas de origem do mapeamento estao vazias.
- Mapeamentos simples (ex: `A:B`) mantem o tipo original do valor (numero, data, texto).
- Mapeamentos combinados (ex: `A+C:D`) geram texto com o separador definido.
- Arquivos suportados: `.xlsx` e `.xlsm`.
