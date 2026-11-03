# Editor de Planilha

Programa que arruma automaticamente a planilha recebida (CSV ou Excel): oculta as colunas desnecessárias, realça acessos repetidos e calcula o VALOR_BRUTO. Segue o mesmo modelo do programa de notas fiscais: um `.exe` pronto, na pasta compartilhada, que qualquer pessoa usa sem instalar nada. O Python só é necessário **uma vez, em um único computador**, para gerar o `.exe`.

## Primeira vez (só quem vai preparar o programa)
1. Em um computador Windows com Python instalado (python.org, marcando "Add python.exe to PATH"), rode `fonte\CRIAR_EXE.bat`.
2. Ele instala as bibliotecas, gera o `PlanilhaEngine.exe` e o copia para a pasta `dist`.
3. Pronto: coloque a pasta inteira na pasta compartilhada da empresa. Ninguém mais precisa de Python.

## Como usar (todos os usuários)
1. Abra a pasta do programa.
2. **Arraste a planilha (CSV ou Excel) para cima de `Editar_Planilha.bat`** (ou dê dois cliques e cole o caminho do arquivo).
3. A planilha pronta aparece na pasta `resultados`, com final `_EDITADA.xlsx`. A original não é alterada.

O nome do arquivo (inclusive o mês, ex.: 202610) não importa. O programa roda no próprio computador, sem internet, e não precisa do Excel para gerar a planilha.

## O que a planilha pronta tem
- **Colunas visíveis:** CODIGO_FATURA, ACESSO, CIRCUITO, UF_DO_CIRCUITO, PRECO_VIGENTE_DO_CIRCUITO, VALOR_LIQUIDO e VALOR_BRUTO (nova). As outras ficam ocultas, não apagadas (clique direito nas letras das colunas > Reexibir).
- **ACESSO:** valores repetidos ficam com fundo vermelho (formatação condicional). As células ficam em formato Geral, então dá para digitar fórmulas à mão; o realce considera o resultado da fórmula.
- **Formato monetário:** PRECO_VIGENTE_DO_CIRCUITO, VALOR_LIQUIDO e VALOR_BRUTO usam o formato Contábil em reais (R$), com separador de milhares e 2 casas.
- **PRECO_VIGENTE_DO_CIRCUITO:** arredondado para 2 casas decimais.
- **VALOR_BRUTO:** `VALOR_LIQUIDO / ((1 - ICMS da UF) x (1 - 3,65%))`, arredondado para 2 casas. É uma fórmula do Excel, então dá para clicar na célula e conferir. Fica em branco se o valor líquido estiver vazio ou a UF não estiver na tabela.
- **Aba `ALIQUOTAS`:** a tabela de ICMS por UF e o percentual de PIS/COFINS usados pela fórmula.
- Cabeçalho em negrito com filtro, primeira linha congelada.

## Como funciona por dentro

### Arquivos
```
Editar_Planilha.bat          <- o que o usuário usa
dist/
  PlanilhaEngine.exe         <- o programa (gerado uma vez pelo CRIAR_EXE.bat)
  perfis/planilha_profile.json   <- as regras (colunas, alíquotas, 3,65%)
fonte/
  planilha_engine.py         <- código-fonte do programa (Python)
  CRIAR_EXE.bat              <- gera o .exe (só com Python instalado)
resultados/                  <- criada automaticamente com os arquivos prontos
opcional_powershell/         <- versão alternativa sem Python (precisa do Excel; mais lenta)
```

### Passo a passo
1. **O `.bat`** confere se o `.exe` e o perfil existem, recebe o arquivo (arrastado ou digitado), copia o `.exe` para uma pasta local do computador (`%LOCALAPPDATA%\PlanilhaEngine_Cache`, só recopia se houver versão mais nova) e o executa com: perfil, arquivo de entrada, pasta de saída.
2. **O programa lê o perfil** (`planilha_profile.json`).
3. **Leitura da planilha:** se for CSV, detecta o separador (`;` ou `,`) e a codificação (UTF-8 ou Windows-1252). Se for Excel, lê a primeira aba. A primeira linha é o cabeçalho; linhas totalmente vazias são ignoradas.
4. **Localiza as colunas pelo nome do cabeçalho**, ignorando acentos e maiúsculas. Por isso não importa a posição das colunas. Se faltar alguma, o programa para e lista as que não encontrou.
5. **Monta a planilha nova** diretamente no arquivo `.xlsx` (sem abrir o Excel). Os campos entram como texto, para não perder zeros à esquerda nem virar notação científica. Só as colunas numéricas (PRECO_VIGENTE_DO_CIRCUITO e VALOR_LIQUIDO) viram número; aceita `1.234,56` e `1234.56`. O preço é arredondado para 2 casas.
6. **Cria a aba `ALIQUOTAS`** com a tabela do perfil.
7. **Cria o VALOR_BRUTO** como última coluna, com a fórmula do Excel em todas as linhas (o Excel calcula ao abrir).
8. **Aplica o realce vermelho** de valores duplicados na coluna ACESSO.
9. **Oculta** todas as colunas que não estão na lista de visíveis, aplica formatação e salva como `<nome original>_EDITADA.xlsx` em `resultados`.
10. Mostra `[OK]` ou `[ERRO]` com o motivo.

### O perfil (`planilha_profile.json`)
| Campo | O que faz |
|---|---|
| `colunas_visiveis` | Colunas que continuam visíveis. Todas as outras são ocultadas. |
| `coluna_duplicados` | Coluna em que repetidos ficam em vermelho. |
| `valor_bruto.nome` | Nome da coluna calculada. |
| `valor_bruto.coluna_base` | Coluna usada como base da conta (VALOR_LIQUIDO). |
| `valor_bruto.coluna_uf` | Coluna cuja UF define o ICMS (UF_DO_CIRCUITO). |
| `valor_bruto.pis_cofins` | Percentual do segundo fator (0.0365 = 3,65%). |
| `colunas_numericas` | Colunas convertidas de texto para número. |
| `aliquotas_icms` | ICMS por UF, em decimal (0.185 = 18,5%). |

Para mudar uma alíquota ou a lista de colunas, edite esse arquivo no Bloco de Notas, mantendo as aspas e vírgulas. Não é preciso mexer no script.

## Problemas comuns
| Mensagem | O que fazer |
|---|---|
| `Colunas não encontradas na planilha` | O arquivo não é o esperado ou o cabeçalho não está na primeira linha. |
| `Não consegui salvar` | Feche a planilha `_EDITADA` se estiver aberta no Excel e rode de novo. |
| `PlanilhaEngine.exe nao encontrado` | O `.exe` ainda não foi gerado. Rode `fonte\CRIAR_EXE.bat` em um computador com Python. |
| `Python não encontrado` (no CRIAR_EXE) | Instale o Python marcando "Add python.exe to PATH". |
| VALOR_BRUTO em branco numa linha | Valor líquido vazio, ou UF fora da tabela `ALIQUOTAS`. |

## Limitações
- Lê só a primeira aba, se a entrada for Excel.
- A UF usada no ICMS é a do circuito, e a fórmula do VALOR_BRUTO considera os 3,65% como segundo fator da divisão.
- Para mudar o código (não só o perfil), é preciso gerar o `.exe` de novo com `CRIAR_EXE.bat`.
- Alguns antivírus desconfiam de `.exe` gerado por PyInstaller; se acontecer, peça à TI para liberar.
