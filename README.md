# Validador de relatorios XLSX

Aplicacao em Python para verificar relatorios `.xlsx` e listar os itens encontrados.

## Tipos de relatorio detectados automaticamente

- `Locais indicados`: usa `IdLocalProva`, `UF`, `Cidade`, `DescricaoLocalProva` e `NumeroLogradouro`.
- `Salas`: usa `IdLocalProva`, `UF`, `Cidade`, `LocalProva` e `Sala`.
- `Meta capacidade`: usa `CodigoIBGE`, `SiglaUF`, `DescricaoCidade`, `MetaCapacidade`, `CapacidadeIndicada` e `TaxaIndicacao`.
- `Junções`: usa `IdLocalProva`, `UF`, `Cidade`, `DescricaoLocalProva`, `DescricaoIndicacaoJuncao` e `Blocos`.

## Regras implementadas

- `Sala`: identifica qualquer nome de sala que contenha `Sala`.
- `NumeroLogradouro`: identifica numero do logradouro preenchido como `SN` ou `S/N`, ignorando `S/N°` e `S/N.º`.
- `DescricaoIndicacaoJuncao`: identifica bloco sem junção quando a descrição está vazia.
- `Blocos`: identifica blocos cujo nome não começa com `Bloco`, `Prédio` ou `Pavilhão`.
- `DescricaoLocalProva`: identifica nomes com sinais de texto fora do padrão UTF-8/ABNT2 Português (Brasil).
- `Complemento`: identifica complementos contendo `escola`, `prédio`, `bairro`, `casa`, `faculdade`, `centro`, `sem complemento` ou `sn`.
- `Meta capacidade`: classifica a porcentagem:
  - `0%` a `80%`: `Escolas Não Indicadas`
  - `90%` a `99%`: `Falta Cadeiras`
  - `100%`: `Meta Alcançada`
  - acima de `100%`: `Capacidade extra`

## Campos exibidos

- A tabela ajusta automaticamente os campos exibidos conforme o relatorio detectado.
- Para `Meta capacidade`: `ID`, `UF`, `Cidade`, `Nome`, `Meta`, `Capacidade atual`, `Porcentagem` e `Tag`.
- Para relatorios de erros: `ID`, `UF`, `Cidade`, `Nome`, `Campo`, `Exibido`, `Erro` e, quando existir, `Tag`.
- O campo `Tipo` nao e exibido.

Ao clicar no valor da coluna `ID`, o ID e copiado automaticamente para a area de transferencia.

O botao de tema alterna entre modo claro e modo escuro.

## Aplicacao web

Instale as dependencias:

```powershell
pip install -r requirements.txt
```

Execute a versao web:

```powershell
streamlit run streamlit_app.py
```

Na pagina web, envie um ou mais arquivos `.xlsx`, filtre os resultados e baixe a lista em CSV ou XLSX.
