# Metodologia

## 1. Escopo

O Xbox Sales Pulse analisa 295 registros de assinaturas da base pública usada no desafio DIO “Criando um Dashboard de Vendas do Xbox com Excel”. O objetivo é transformar os dados em indicadores de receita, recorrência, mix de planos, periodicidade e adicionais.

## 2. Preservação da fonte

A planilha original permanece em `data/base.xlsx` e não é editada pelo gerador.

- SHA-256: `0b60de55f2e54ef2759b6363fe1ef1436ea7ce3de5493ddd5f0e987ac8052a13`
- Registros: 295
- Primeiro ID: 3231
- Último ID: 3525
- Período: 01/01/2024 a 16/12/2024

O script identifica a aba que contém o cabeçalho `Subscriber ID`, lê os campos da base e grava uma nova pasta de trabalho.

## 3. Normalização

Os registros são levados para a aba `Base` com os 13 campos originais e três campos técnicos:

- `Recalculated Total`: resultado independente da regra financeira;
- `Validation`: `OK` quando o resultado recalculado coincide com `Total Value`;
- `Month Start`: primeiro dia do mês da assinatura, usado na série temporal.

Também é criado `data/assinaturas_xbox.csv` em UTF-8 com BOM, facilitando inspeção e reutilização.

## 4. Regra financeira

Cada registro é validado por:

```text
Recalculated Total = Subscription Price
                   + EA Play Season Pass Price
                   + Minecraft Season Pass Price
                   − Coupon Value
```

Resultado: **295 de 295 registros válidos**.

## 5. Indicadores

Os indicadores principais são calculados com `SUMIFS` e `COUNTIFS`, combinando três filtros:

- plano;
- periodicidade;
- renovação automática.

Quando um filtro possui o valor `Todos`, seu critério é convertido em `*`. Os filtros são intervalos nomeados ligados às células verdes do `Dashboard`.

### Indicadores do painel

- Receita total: soma de `Total Value`;
- Assinantes: contagem de IDs;
- Ticket médio: receita dividida pela quantidade de assinantes;
- Auto renovação: clientes com `Auto Renewal = Yes` divididos pelo total filtrado;
- Base validada: registros com `Validation = OK` divididos pelo total filtrado.

### Componentes financeiros

- Receita-base: soma de `Subscription Price`;
- EA Play: soma de `EA Play Season Pass Price`;
- Minecraft: soma de `Minecraft Season Pass Price`;
- Cupons: soma de `Coupon Value`.

## 6. Perguntas do desafio

As respostas são calculadas diretamente na aba `Análise`:

| Tema | Critério | Resultado |
|---|---|---:|
| Planos anuais | `Subscription Type = Annual` | US$ 1.754,00 |
| Anual com renovação | `Annual` e `Yes` | US$ 1.537,00 |
| Anual sem renovação | `Annual` e `No` | US$ 217,00 |
| EA Play | Soma do preço do adicional | US$ 2.940,00 |
| Minecraft | Soma do preço do adicional | US$ 3.880,00 |

## 7. Séries dos gráficos

Os quatro gráficos do painel usam intervalos auxiliares da aba `Análise`:

1. receita mensal;
2. receita por plano;
3. receita por periodicidade;
4. EA Play e Minecraft por plano.

As séries respeitam os filtros ativos. Quando o gráfico segmenta uma dimensão, essa própria dimensão é substituída pela categoria da linha para permitir comparação entre as categorias, enquanto os demais filtros permanecem aplicados.

## 8. Reprodutibilidade

O gerador usa `openpyxl`. Como a biblioteca cria fórmulas, mas não possui um mecanismo de cálculo do Excel, o fluxo também grava caches OOXML com resultados calculados independentemente em Python. Assim:

- a primeira visualização já mostra os totais validados;
- o Excel recalcula todas as fórmulas ao abrir o arquivo;
- qualquer alteração nos filtros atualiza o painel no Excel.

O atributo de cálculo da pasta de trabalho é configurado como automático e com recálculo completo na abertura.

## 9. Cuidados de interpretação

- janeiro e fevereiro têm somente dois registros cada;
- dezembro termina em 16/12/2024;
- os valores mensais não devem ser tratados como meses completos equivalentes;
- a base é educacional e não representa resultados oficiais da Microsoft ou do Xbox.

## 10. Autoria

Design, estrutura, fórmulas, gerador, validação, documentação e identidade **Xbox Sales Pulse** foram produzidos por **0Barone (@gabrielnasciimentoleandro-arch)** para este projeto.