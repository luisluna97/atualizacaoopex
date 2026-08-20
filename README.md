# Ferramentas RP

Aplicativos internos de planejamento operacional.

| Ferramenta | O que faz |
|---|---|
| **Conversor de Malha** | Converte as malhas da GOL, AZUL, LATAM e as manuais no CSV padrão *Malha RP* |
| **Atualizador de OPEX** | Atualiza o staff (grupo RAMPA) pela folha e/ou a quantidade de voos pela malha |

## Rodar na sua máquina

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre em `http://localhost:8501`.

## Colocar online (Streamlit Community Cloud — grátis)

1. Suba este repositório no GitHub — **privado**, já que a ferramenta lida com dados de pessoal
2. Entre em <https://share.streamlit.io> com a conta do GitHub
3. **New app** → escolha o repositório, o branch e o arquivo `app.py`
4. **Deploy**

A cada `git push`, o app atualiza sozinho.

## Atualizador de OPEX

As duas partes são independentes — envie só o que quiser atualizar:

- **OPEX + folha (FPRE109)** → atualiza só o staff
- **OPEX + malha (CSV)** → atualiza só os voos
- **OPEX + as duas** → atualiza tudo

Mexe apenas nas colunas de quantidade: **X** nas abas de base (staff) e
**# Flights** na aba Voos&Tarifas. Tarifas, fórmulas, gráficos e Power Pivot
ficam intactos; a coluna de receita recalcula sozinha ao abrir.

### Staff — regras de situação

| Situação na folha | No OPEX |
|---|---|
| Trabalhando, Férias, Atestado | conta normalmente |
| Auxílio Doença, Acidente, Maternidade, Licenças | a função fica com **0** |
| Aposentadoria por Invalidez | não entra |

Aceita o `.xlsx` e também o `.xls` bruto que o sistema exporta (formato antigo
BIFF2, que o Excel abre mas o pandas sozinho não lê).

A CH semanal em decimal vira a mensal do OPEX (7,5 → 180h · 8,75 → 210h ·
9,1667 → 220h). O formato antigo (`210:00`) também é aceito.

### Voos — tipo de atendimento

- **LATAM**: `TST.N` conta como **PNT**; o resto é TST
- **Demais**: `PNT` quando o tempo de solo passa de 4h

O rótulo do equipamento segue o que já existe naquela base e cliente
(a LATAM usa `319/320`, a AZUL usa `320` para o mesmo A320).

### Proteções

- Se uma base **não aparecer na folha**, ela é mantida como está (não zera) e
  o app avisa. Base sumir da folha quase sempre é problema de arquivo, não
  gente que saiu
- Rodar duas vezes seguidas não muda nada na segunda
- Atualizar só o staff não toca nos voos, e vice-versa
- Malha vazia ou arquivo errado: o app recusa em vez de gravar zeros

### O que ele avisa

- **Linhas que faltam criar**: função na folha sem linha correspondente na aba
- **Funções só com afastados**: devem entrar com 0
- **Voos sem linha de tarifa**: combinação de base/cliente/equipamento que não
  existe na aba Voos&Tarifas e por isso não foi lançada

## Segurança

O `.gitignore` bloqueia `.xlsx`, `.xls`, `.csv` e `.mhtml` — nenhum arquivo de
folha, OPEX ou malha vai para o repositório, nem por engano.

## Estrutura

```
app.py                          página inicial
pages/
  1_Conversor_de_Malha.py
  2_Atualizador_de_OPEX.py
requirements.txt
.gitignore
.streamlit/config.toml
```
