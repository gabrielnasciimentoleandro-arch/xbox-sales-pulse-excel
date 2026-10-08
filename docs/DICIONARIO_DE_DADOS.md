# Dicionário de dados

| Campo | Tipo | Descrição |
|---|---|---|
| `Subscriber ID` | inteiro | Identificador único da assinatura |
| `Name` | texto | Nome associado ao registro |
| `Plan` | categoria | Core, Standard ou Ultimate |
| `Start Date` | data | Início da assinatura |
| `Auto Renewal` | categoria | Yes ou No |
| `Subscription Price` | moeda | Valor-base do plano |
| `Subscription Type` | categoria | Monthly, Quarterly ou Annual |
| `EA Play Season Pass` | categoria | Adesão ao EA Play |
| `EA Play Season Pass Price` | moeda | Receita do adicional EA Play |
| `Minecraft Season Pass` | categoria | Adesão ao Minecraft |
| `Minecraft Season Pass Price` | moeda | Receita do adicional Minecraft |
| `Coupon Value` | moeda | Desconto aplicado |
| `Total Value` | moeda | Valor final do registro |
| `Recalculated Total` | moeda | Total refeito pela fórmula de auditoria |
| `Validation` | texto | `OK` ou `REVIEW` |
| `Month Start` | data | Primeiro dia do mês para agregação temporal |

## Regra de validação

```text
Total Value = Subscription Price
            + EA Play Season Pass Price
            + Minecraft Season Pass Price
            − Coupon Value
```

## Categorias observadas

- Planos: Core, Standard, Ultimate;
- Periodicidade: Monthly, Quarterly, Annual;
- Renovação automática: Yes, No;
- Adesão a adicionais: Yes, No;
- Cupons: Sem Cupom, Xbox10, Xbox20, Xbox30, Xbox50.
