# Simulador de fundos imobiliários

Planilha com cara de aplicativo. A pessoa preenche o amarelo e lê o cinza. Não precisa saber fórmula para usar.

Abra `simulador-fiis.xlsx`. A tela é a aba **Simulador**. A aba **Apoio** guarda a lista de perfis e os percentuais.

Os números publicados são de exemplo: salário de R$ 6.000, aporte de R$ 1.500, 10 anos e 0,80% ao mês. Não são dados pessoais e não são recomendação de investimento.

## As cinco perguntas

| Pergunta | Onde responde | Nesta simulação |
| --- | --- | --- |
| Quanto investir por mês? | Célula amarela `B8` e resposta 1 em `F6` | R$ 1.500,00 |
| Por quantos anos? | Célula amarela `B9` e resposta 2 em `F7` | 10 anos |
| Qual a taxa de rendimento mensal? | Célula amarela `B10` e resposta 3 em `F8` | 0,80% ao mês |
| Quanto de patrimônio vai acumular? | Resposta 4, célula cinza `F9` | R$ 300.326,20 |
| Quanto vai receber de dividendos por mês? | Resposta 5, célula cinza `F10` | R$ 2.402,61 |

O salário fica em `B6`. A sugestão de aporte (30% do salário) fica em `B7`: R$ 1.800,00. Ela não substitui o aporte, a menos que a pessoa queira usar a segunda simulação.

Também aparecem o total aportado (`F11` = R$ 180.000,00) e o ganho acima desse total (`F12` = R$ 120.326,20).

## Como o VF entra

O patrimônio usa `VF` (no arquivo a fórmula está em inglês, `FV`; o Excel em português mostra `VF`):

```excel
=VF(taxa_mensal; prazo_anos*12; -aporte)
```

| Argumento | O que é aqui | Valor do exemplo |
| --- | --- | --- |
| taxa | Rendimento de um mês | 0,80% |
| nper | Quantidade de meses (`anos × 12`) | 120 |
| pgto | Aporte com sinal negativo, porque é dinheiro que sai do bolso | -1.500 |
| vp | Omitido. A conta começa do zero | 0 |
| tipo | Omitido. Vale 0: o aporte entra no fim de cada mês | 0 |

Com o pagamento negativo, o `VF` devolve o saldo positivo no fim do prazo.

O dividendo do mês é esse patrimônio vezes a taxa mensal:

```excel
=patrimonio*taxa_mensal
```

É o rendimento de um mês sobre o saldo acumulado, sem aporte novo.

## Cenários

A seção 3 repete a mesma conta para 2, 5, 10, 20 e 30 anos. `taxa_mensal` e `aporte` são intervalos nomeados, então não mudam de coluna quando a fórmula é arrastada. Só o prazo, na célula de cima, muda.

Patrimônio com aporte de R$ 1.500:

| 2 anos | 5 anos | 10 anos | 20 anos | 30 anos |
| --- | --- | --- | --- | --- |
| R$ 39.514,73 | R$ 114.935,80 | R$ 300.326,20 | R$ 1.081.696,83 | R$ 3.114.619,85 |

## Segunda simulação: 30% do salário

As linhas 18 e 19 trocam o aporte informado por `sugestao_aporte` (30% de R$ 6.000 = R$ 1.800). Prazo e taxa continuam os mesmos.

Em 10 anos esse caminho chega a **R$ 360.391,44** de patrimônio e **R$ 2.883,13** de dividendos por mês.

## Como o PROCV entra

A aba **Apoio** tem quatro colunas: chave, perfil, tipo de fundo e percentual. A chave é uma fórmula:

```excel
=B4&"|"&C4
```

Exemplo: `Moderado|Logística`.

Na aba Simulador, cada linha monta a mesma chave (`perfil` + `|` + tipo) e busca o percentual:

```excel
=PROCV(D23; tabela_alocacao; 4; FALSO)
```

| Argumento | Função |
| --- | --- |
| D23 | Chave composta, o valor procurado |
| tabela_alocacao | Intervalo `Apoio!A4:D27` |
| 4 | Devolve a quarta coluna, o percentual |
| FALSO | Correspondência exata |

Sem a chave, um `PROCV` só pelo tipo "Logística" acharia sempre a primeira logística da tabela, a do perfil Conservador. A chave `Moderado|Logística` aponta para uma linha só.

O valor em reais é `percentual × aporte`. Trocar o perfil na lista muda essa divisão. O patrimônio não muda, porque ele depende do aporte, do prazo e da taxa, não do perfil.

No exemplo Moderado, os R$ 1.500 ficam assim:

| Tipo | Percentual | Valor |
| --- | --- | --- |
| Papel (recebíveis) | 25% | R$ 375,00 |
| Logística | 20% | R$ 300,00 |
| Shoppings | 15% | R$ 225,00 |
| Lajes corporativas | 15% | R$ 225,00 |
| Híbrido | 15% | R$ 225,00 |
| Fundo de fundos | 10% | R$ 150,00 |
| Soma | 100% | R$ 1.500,00 |

No perfil Arrojado, com o mesmo aporte, o papel cai para 10% (R$ 150) e shoppings, lajes e híbrido sobem para 20% (R$ 300 cada). A soma continua 100%. O gráfico de pizza usa a coluna de valores, então muda junto com o perfil.

## De onde vêm os percentuais

São um exemplo criado para este exercício, não uma carteira recomendada e não uma cópia de planilha de terceiros. A regra didática foi esta:

- Conservador concentra em papel (recebíveis): 45%.
- Moderado equilibra papel e tijolo.
- Arrojado reduz papel para 10% e abre espaço para shoppings, lajes e híbrido.
- Foco em renda fica perto do conservador, com um pouco mais de híbrido.

| Tipo | Conservador | Moderado | Arrojado | Foco em renda |
| --- | --- | --- | --- | --- |
| Papel (recebíveis) | 45% | 25% | 10% | 40% |
| Logística | 20% | 20% | 20% | 20% |
| Shoppings | 10% | 15% | 20% | 10% |
| Lajes corporativas | 10% | 15% | 20% | 10% |
| Híbrido | 10% | 15% | 20% | 15% |
| Fundo de fundos | 5% | 10% | 10% | 5% |
| Soma | 100% | 100% | 100% | 100% |

A célula `Apoio!I4` (`percentual_sugestao`) guarda os 30%. A sugestão de aporte é `salario * percentual_sugestao`.

## Intervalos nomeados

| Nome | Célula | Uso |
| --- | --- | --- |
| salario | Simulador!B6 | Salário mensal de exemplo |
| sugestao_aporte | Simulador!B7 | 30% do salário |
| aporte | Simulador!B8 | Quanto investir por mês |
| prazo_anos | Simulador!B9 | Horizonte em anos |
| taxa_mensal | Simulador!B10 | Rendimento mensal da carteira |
| rendimento_carteira | Simulador!B10 | O mesmo B10, com o nome da configuração |
| perfil | Simulador!B11 | Perfil escolhido na lista |
| patrimonio | Simulador!F9 | Resultado do VF |
| dividendos_mensais | Simulador!F10 | Patrimônio × taxa |
| percentual_sugestao | Apoio!I4 | 30% |
| lista_perfis | Apoio!F4:F7 | Validação de dados da lista |
| tabela_alocacao | Apoio!A4:D27 | Matriz do PROCV |

## O que mudou em relação à ferramenta do Expert

Não havia a planilha resolvida do Expert neste projeto, então a ferramenta foi montada pelo enunciado, com escolhas próprias:

- Banner, grade escondida, amarelo para entrada e cinza para conta.
- As cinco respostas reunidas num cartão, além das células de entrada.
- Cenários de 2, 5, 10, 20 e 30 anos, com a segunda simulação de 30% do salário na mesma grade.
- Quatro perfis e seis tipos de fundo, com percentuais deste exercício.
- Gráfico de pizza ligado ao valor do aporte.
- Conferência de que o perfil soma 100%.
- Total aportado e ganho acima do total aportado.

## Prints

Vale fotografar a aba Simulador duas vezes, com os mesmos R$ 1.500, 10 anos e 0,80% ao mês:

1. Perfil **Moderado** — papel em 25% e patrimônio de R$ 300.326,20.
2. Perfil **Arrojado** — papel em 10% e o mesmo patrimônio, porque o aporte não mudou.

A soma do perfil tem de aparecer como 100% nos dois.
