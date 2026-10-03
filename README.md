# Leitor de etiquetas de PVB

Lê fotos de etiquetas de rolos de PVB com OCR local (RapidOCR) e extrai nº do rolo, cor, espessura, largura,
comprimento, área, pesos, data de fabricação, validade e inspetor. Reconhece hoje dois modelos de etiqueta:
"PVB Interlayer" e Shandong Qilu.

## Uso

```
.venv\Scripts\python leitor_etiquetas.py [pasta_das_fotos] [-o saida.xlsx|saida.csv]
```

- `.xlsx`: planilha para uma pessoa ler. As linhas com aviso ficam em amarelo.
- `.csv`: formato para outro programa importar. É o que o estoque de PVB da planilha SGQ usa.

**Código de saída**: `0` = arquivo gravado; `1` = pasta sem imagem ou arquivo não gravado; `2` = falha
inesperada.

**Cache de leitura**: o leitor grava `.leitura_etiquetas.json` (oculto) na pasta das fotos. Ler a mesma pasta de
novo devolve na hora as fotos já lidas e não alteradas (nome, tamanho e data iguais); só as novas ou alteradas
passam pelo OCR. Versão nova do leitor (`versao.txt`) invalida o cache inteiro. Leitura com erro não é guardada.
`--sem-cache` força ler tudo. Pasta sem permissão de gravação: funciona igual, só sem cache.

**Nº do rolo da Qilu**: quando a 1ª leitura não dá 2 letras válidas no começo, o leitor relê a faixa do número.
Letras que não dá para afirmar saem como `??`, sempre com aviso ("confira a foto"). Testes:
`.venv\Scripts\python test_rolo_qilu.py` e `.venv\Scripts\python test_cache.py`.

## Integração com o estoque de PVB

O formato do CSV é um contrato com o estoque: separador `;`, UTF-8, datas ISO, ponto decimal, colunas lidas
**pelo nome**. Coluna nova só pode entrar no fim. A definição completa está em
`pcp_brglass\specs\090-importar-etiquetas-pvb\contracts\csv-leitor.md`.

## Empacotar e publicar

Os computadores do setor não têm Python. O leitor vai para eles como pasta empacotada:

```
powershell -ExecutionPolicy Bypass -File empacotar.ps1
```

O script gera `dist\leitor_etiquetas\` (~250 MB) com um `versao.txt` novo. Para publicar, copie a pasta
inteira para `<unidade da rede>\leitor etiquetas\`. Cada máquina mantém uma cópia local e só copia de novo
quando o `versao.txt` muda. Por isso **toda publicação passa pelo script**, que troca a versão.

## Desenvolvimento

```
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt pyinstaller
```

Para um modelo novo de etiqueta, acrescente os rótulos em `CAMPOS` e, se a cor vier com outro nome, a
equivalência em `mdlPvbDominio.CorDaEtiquetaPvb` (estoque).
