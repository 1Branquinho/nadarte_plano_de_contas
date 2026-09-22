# omie_compare

Compara Categorias e Departamentos entre 6 apps (empresas) da API OMIE, e
verifica se cada código tem lançamento (Contas a Pagar / Contas a Receber)
num período configurável.

## Como rodar

```bash
pip install -r requirements.txt
python comparar_omie.py
```

O período consultado é controlado pela constante `PERIODO_DIAS` no topo de
`comparar_omie.py` (padrão: 365 dias, ou seja, últimos 12 meses a partir de
hoje). **Recomenda-se testar primeiro com um período curto (ex.: 90 dias)**
antes de rodar com 365, porque:

- O volume de títulos de Contas a Pagar/Receber pode ser grande, e cada
  título é uma chamada paginada por empresa.
- A OMIE tem limite de requisições (rate limit); o cliente já faz retry com
  backoff, mas um período menor reduz o tempo total e o risco de esbarrar
  no limite antes de você validar que os campos foram detectados
  corretamente.

Para testar com 90 dias, edite temporariamente:

```python
PERIODO_DIAS = 90
```

## Credenciais

As credenciais dos 6 apps ficam em `config.yaml`, que **não é versionado**
(está no `.gitignore`) por conter segredos de produção. Se precisar recriar
o arquivo em outra máquina, use o mesmo formato:

```yaml
apps:
  NOME_DO_APP:
    app_key: "..."
    app_secret: "..."
```

## Saída

O script gera `omie_comparacao_<data_hora>.xlsx` com as seguintes abas:

- **Resumo**: para cada app, quantidade de categorias e departamentos
  cadastrados, quantos desses códigos têm pelo menos um lançamento no
  período, e se houve erro de autenticação/API para aquele app.
- **Categorias**: tabela código × app, mostrando a descrição cadastrada em
  cada empresa para aquele código. Linhas ficam destacadas em **amarelo**
  quando a descrição diverge entre as empresas que têm o código, ou quando
  o código não existe em todas as 6 empresas.
- **Departamentos**: mesma lógica da aba Categorias, para departamentos.
- **Uso_Categorias**: tabela código × app com o status de cada código em
  cada empresa: `Com lançamento`, `Sem lançamento` ou `Não existe`
  (quando o código nem está cadastrado naquela empresa).
- **Uso_Departamentos**: mesma lógica, para departamentos.

## Importante: valide a detecção dinâmica de campos

A API OMIE não documenta de forma 100% consistente os nomes de campo em
todas as respostas, e o layout pode variar por endpoint. Em vez de assumir
nomes fixos (tipo `categoria_cadastro` ou `cCodCateg`), o script:

1. Acha a lista de registros dentro de uma resposta pegando a primeira
   chave cujo valor é uma lista de dicionários.
2. Acha o campo de "código" e "descrição" de categorias/departamentos
   procurando chaves normalizadas (sem acento, minúsculas) que contenham
   `codigo`/`cod` e `descricao`/`nome`, respectivamente — evitando
   confundir com `codigo_pai`.
3. Detecta uso em lançamentos percorrendo **recursivamente** toda a
   estrutura de cada título financeiro, coletando valores de qualquer
   chave cujo nome normalizado contenha `categ` ou `departamento`, em vez
   de assumir um caminho fixo como `distribuicao[].cCodCateg`.

Essa heurística funciona bem com o formato observado na prática, mas
**pode errar** se a OMIE mudar nomes ou se algum endpoint tiver um formato
inesperado. Por isso, toda chamada bruta é salva em `raw/<app>_<tipo>.json`
(`categorias`, `departamentos`, `contas_pagar`, `contas_receber`). Antes de
confiar cegamente no resultado, abra alguns desses arquivos e confira:

- Se o "código" mostrado na planilha bate com o campo de código real do
  JSON (e não, por exemplo, um ID interno tipo `codigo_pai` ou
  `nCodTitulo`).
- Se a "descrição" é um texto legível (nome da categoria/departamento) e
  não algum outro campo textual.
- Se os códigos de categoria/departamento usados nos lançamentos (aba
  `Uso_*`) realmente aparecem nos títulos correspondentes dentro do JSON
  bruto de `contas_pagar`/`contas_receber`.

Se algo parecer errado, ajuste `find_code_field`, `find_description_field`
ou `collect_used_codes` em `comparar_omie.py` e rode de novo.
