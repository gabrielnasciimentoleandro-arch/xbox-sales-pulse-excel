# Contribuindo

Contribuições são bem-vindas para código, documentação, acessibilidade e testes.

## Fluxo sugerido

1. Crie um fork do repositório.
2. Abra uma branch curta e descritiva.
3. Preserve `data/base.xlsx` sem alterações.
4. Execute o fluxo completo antes de enviar a mudança:

```bash
python scripts/gerar_dashboard.py
python scripts/gerar_previas.py
python scripts/gerar_relatorio.py
python scripts/verificar_dashboard.py
```

5. Confirme que a saída termina em `VALIDAÇÃO APROVADA`.
6. Abra um pull request explicando a motivação e o impacto visual ou analítico.

## Diretrizes

- não inclua dashboards ou ativos de terceiros sem licença compatível;
- não adicione macros ou vínculos externos sem justificativa e documentação;
- mantenha a análise reproduzível sem API paga;
- evite expor dados pessoais reais;
- atualize a documentação quando houver mudança de regra de negócio.
