# Relatório de validação

## Resultado

**VALIDAÇÃO APROVADA**

O arquivo `xbox_sales_pulse.xlsx` passou pelas verificações automatizadas de conteúdo, cálculos, estrutura e integridade definidas em `scripts/verificar_dashboard.py`.

## Resumo da execução

```text
Arquivo: xbox_sales_pulse.xlsx
Autoria: 0Barone (@gabrielnasciimentoleandro-arch)
Abas: 6 (Dashboard, Análise, Base, Dicionário, Guia, Apoio)
Base: 295 registros, 295 IDs únicos e 295 valores conferidos
Receita: $ 7.633,00
Ticket médio: $ 25,87
Auto renovação: 50,2%
Planos anuais: $ 1.754,00
  ├─ com renovação: $ 1.537,00
  └─ sem renovação: $ 217,00
EA Play: $ 2.940,00
Minecraft: $ 3.880,00
Cupons: $ 2.122,00
Gráficos: 4
Filtros: 3
Fórmulas: 961
Links internos: 26
Macros: ausentes
Vínculos externos de pasta de trabalho: ausentes
```

## Verificações executadas

### Fonte e dados

- presença da base original;
- SHA-256 exato da base;
- 295 registros e IDs únicos;
- correspondência dos 13 campos originais em todas as linhas (com normalização explícita de `-` para zero em campos monetários);
- intervalo de IDs de 3231 a 3525;
- totais de receita, adicionais e cupons;
- exportação CSV com 295 linhas de dados.

### Regras de negócio

- recálculo do valor total em cada registro;
- 295 resultados com status `OK`;
- receita e quantidade por plano;
- receita e quantidade por periodicidade;
- receita e quantidade por renovação;
- valores das perguntas centrais do desafio.

### Pasta de trabalho

- seis abas na ordem prevista;
- aba `Apoio` oculta;
- tabela estruturada `tblAssinaturas`;
- três validações de dados;
- quatro gráficos;
- 16 intervalos nomeados;
- 961 fórmulas;
- pelo menos 25 links internos de navegação;
- metadados de autoria.

### Integridade e segurança

- arquivo XLSX válido como pacote ZIP/OOXML;
- todos os itens do pacote íntegros;
- quatro definições de gráfico no OOXML;
- ausência de VBA/macros;
- ausência de external links da pasta de trabalho;
- ausência de erros armazenados como `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!` ou `#NULL!`;
- fórmulas sem padrões inválidos de construção;
- caches numéricos conferidos.

### Artefatos visuais

- existência de `dashboard-geral.png`;
- existência de `data-view.png`;
- resolução mínima de 2.500 × 1.400 pixels para as prévias.

## Como repetir

Na raiz do projeto:

```bash
python scripts/gerar_dashboard.py
python scripts/gerar_previas.py
python scripts/gerar_relatorio.py
python scripts/verificar_dashboard.py
```

Saída esperada da última etapa:

```text
VALIDAÇÃO APROVADA
```

## Observação sobre recálculo

A geração ocorre sem automatizar o Microsoft Excel. Os resultados iniciais foram calculados independentemente em Python e gravados como caches OOXML; o arquivo também solicita recálculo automático completo quando aberto no Excel. Essa abordagem permite validar o artefato em ambiente local, sem servidor ou API paga.