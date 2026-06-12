# Resposta 03 - Projeção de Indicadores EMBRAPII para 2027

Este repositório contém o script Python utilizado para estimar, com base nos dados históricos da EMBRAPII, os três indicadores solicitados na Questão 03 do processo seletivo:

1. número de projetos contratados em 2027;
2. valor total dos projetos contratados em 2027;
3. número de projetos concluídos em 2027.

## Técnica utilizada

A projeção foi feita por meio de uma regressão linear com tendência temporal e sazonalidade mensal. O modelo utiliza a série histórica mensal da planilha para aprender:

- a tendência geral ao longo do tempo;
- o comportamento típico de cada mês;
- a variação residual observada no histórico.

Para estimar a faixa de incerteza de 95%, foi aplicado bootstrap dos resíduos do modelo. Em termos práticos, o script reamostra os erros observados no histórico e calcula várias possíveis somas anuais para 2027.

## Arquivos principais

- `kennedy_anderson_resposta03_projecao_2027_refeito.py`: script principal da análise.
- `Embrapii_seleção_analista_2026_questao03_Estimativa.xlsx`: planilha de entrada da questão, quando permitido pelo processo seletivo.

## Dependências

Instale as bibliotecas abaixo antes de executar o script:

```bash
pip install pandas numpy matplotlib openpyxl
```

## Como executar

Coloque o script Python e a planilha Excel na mesma pasta e execute:

```bash
python resposta03_projecao_2027_refeito.py
```

Para informar manualmente o caminho da planilha:

```bash
python resposta03_projecao_2027_refeito.py --input "caminho/para/Embrapii_seleção_analista_2026_questao03_Estimativa.xlsx"
```

Para salvar os arquivos sem abrir os gráficos na tela:

```bash
python resposta03_projecao_2027_refeito.py --no-show
```

## Saídas geradas pelo script

O script cria a pasta `outputs_questao03` com duas subpastas.

### Tabelas

- `resultado_projecao_2027.csv`
- `previsao_mensal_2027.csv`
- `historico_anual_observado.csv`
- `resumo_base_metodo.csv`

### Visuais usados no relatório

- `tabela_resultados_2027.png`
- `tabela_resumo_base_metodo.png`
- `grafico_historico_anual_projecao_2027.png`
- `grafico_intervalos_confianca_2027.png`
- `grafico_distribuicao_mensal_2027.png`

## Resultados principais

| Indicador | Estimativa 2027 | Limite inferior 95% | Limite superior 95% |
|---|---:|---:|---:|
| Projetos contratados no ano | 823 | 745 | 921 |
| Valor contratado no ano | R$ 1.458.563.844 | R$ 1.249.184.914 | R$ 1.802.042.603 |
| Projetos concluídos no ano | 332 | 257 | 411 |

## Observações metodológicas

Os resultados devem ser interpretados como estimativas de apoio ao planejamento, não como metas fechadas. A técnica assume que os padrões históricos de tendência e sazonalidade continuam relevantes em 2027. Mudanças estratégicas, alterações orçamentárias, novas políticas públicas ou eventos extraordinários podem deslocar os resultados para fora da faixa estimada.

## Link do repositório

Substitua o texto abaixo pelo link público do GitHub antes do envio oficial:

`https://github.com/kennedyanst/embrapii-resposta-03`
