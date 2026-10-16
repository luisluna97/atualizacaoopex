# Ferramentas RP

Aplicativos internos de planejamento operacional.

| Ferramenta | O que faz |
|---|---|
| **Conversor de Malha** | Converte as malhas da GOL, AZUL, LATAM e as manuais no CSV padrão *Malha RP* |
| **Atualizador de OPEX** | Atualiza o staff (todos os grupos) pela folha e/ou a quantidade de voos pela malha |

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
| Auxílio Doença, Acidente Trabalho, Maternidade, Licenças, Suspensão Ajuizamento Rec Trab | a função fica com **0** |
| Aposentadoria por Invalidez | não entra |

Vale para **todos os grupos** da aba de base (Rampa, Gerência, PAX,
Administrativo, Safety, Security etc.) — exceto em **MGF**, que é a sede: só a
parte operacional (rampa + Jovem Aprendiz) entra na conferência, o resto
(cargos corporativos de sede) fica de fora, sem tocar e sem reportar.

Aceita o `.xlsx` e também o `.xls` bruto que o sistema exporta (formato antigo
BIFF2, que o Excel abre mas o pandas sozinho não lê).

Quando o nome da função no OPEX não bate letra por letra com o nome na folha
(ex.: "GERENTE OPERACIONAL FILIAL BYO" no OPEX vs "GERENTE OPERACIONAL" na
folha, ou "COORDENADOR GERAL SPAX" vs "COORDENADOR DE SPAX JR"), o app tenta
reconhecer que é a mesma pessoa antes de desistir — por variação de nível
(JR/PL/SR), pelo sufixo de base/filial, ou por palavras de ligação (DE, DA,
GERAL etc.). Quando reconhece, ele realinha a linha existente (ajusta CH e
quantidade) em vez de reportar como gente nova ou gente que saiu.

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
- Se uma **função do OPEX não tiver correspondência nenhuma na folha**
  (nome, nível ou base-filial não reconhecidos), a linha é mantida como está
  e reportada — nunca zerada sozinha
- Rodar duas vezes seguidas não muda nada na segunda
- Atualizar só o staff não toca nos voos, e vice-versa
- Malha vazia ou arquivo errado: o app recusa em vez de gravar zeros

### O que ele avisa

- **Linhas que faltam criar**: função na folha sem nenhuma linha correspondente
  na aba — não cria a linha sozinho (bugaria o Excel), só lista pra criação
  manual no site/planilha; normalmente é pouca coisa
- **Função realinhada**: linha que já existia, só mudou de CH ou de redação do
  nome — foi atualizada, não reportada como nova
- **Função sem correspondência**: tem gente lançada na linha do OPEX, mas a
  folha não reconhece esse nome de função em nenhuma CH — o app **não zera
  essa linha sozinho**, só avisa pra conferir se é gente que saiu ou só um
  nome escrito diferente
- **Funções só com afastados**: devem entrar com 0
- **Voos sem linha de tarifa**: combinação de base/cliente/equipamento que não
  existe na aba Voos&Tarifas e por isso não foi lançada

## Conversor de Malha — o que ele confere

- **Arquivo duplicado da mesma cia**: se dois arquivos (ex.: dois da AZUL) cobrem
  a mesma base e o mesmo dia, o app mostra um erro vermelho — os voos sairiam
  contados em dobro. Suba só o arquivo mais novo
- **Manual que não chegou**: envie as manuais do mês anterior em *Manuais do mês
  anterior* e o app repete o padrão **por dia da semana** (um voo que operou em
  pelo menos metade das ocorrências daquele dia da semana, com o horário mais
  comum). É uma projeção — o app avisa e mostra a tabela do que foi repetido.
  O nome do arquivo precisa começar com a sigla da base (`FLN_...`)
- Cabeçalhos da GOL com sublinhado (`Inbound_Flt_Num`) ou com espaço são aceitos
- Voo sem número na origem sai com o campo vazio (não mais a palavra `nan`)

No Atualizador, clientes que o OPEX junta numa linha só (ex.: `SKY + JETSMART` em
FLN) recebem os voos da malha de cada cia (`SKY`, `JETSMART`) normalmente.

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
