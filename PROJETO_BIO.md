# Projeto BIO — Freelancer em Power BI e Excel

**Local no computador:** `D:\PROJETOS_CODE\BIO`
**Repositório:** [jaajanjos/Docker-and-Kubernetes](https://github.com/jaajanjos/Docker-and-Kubernetes/tree/claude/funny-brown-vetnlo), branch `claude/funny-brown-vetnlo`
**Atualizado em:** 20/10/2026

> Este documento descreve o conteúdo do projeto BIO (portfólio freelancer e base de dados de estudo).
> As pastas numeradas (`07_Dockerfile` … `23B_...`) e o `README.md` pertencem ao curso de Docker e Kubernetes e não fazem parte deste projeto.

---

## 1. Objetivo

Estruturar a atuação como **Especialista em Dashboards Power BI e Automação em Excel** nas plataformas Fiverr, Workana, 99Freelas, GetNinjas/20aBuscar e Freelancer.com, com:

1. Portfólio de serviços em **4 pacotes fechados**
2. Base de dados fictícia de e-commerce para **estudo e amostras de portfólio**
3. Roteiro dos **conhecimentos técnicos** necessários

---

## 2. Estrutura de pastas

```
D:\PROJETOS_CODE\BIO
│
├── PROJETO_BIO.md                    ← este documento
├── PROJETO_BIO.pdf                   ← este documento em PDF
│
├── Portfolio_Freelancer\
│   ├── PORTFOLIO_PACOTES.md          ← texto editável do portfólio
│   └── PORTFOLIO_PACOTES.pdf         ← portfólio formatado (8 páginas, A4)
│
├── E-COMMERCE EM EXCEL\
│   ├── Base_Ficticia_ECommerce.xlsx  ← base fictícia (10 abas)
│   └── gerar_base.py                 ← script Python que recria a base
│
└── (pastas e README.md do curso Docker/Kubernetes)
```

---

## 3. Portfólio de Pacotes (`Portfolio_Freelancer`)

Material pronto para copiar e colar nas plataformas: título do perfil, bios longa e curta, palavras-chave, pacotes, extras, modelo de proposta, perguntas frequentes e política de trabalho.

| Pacote | Serviço | Preço (R$) | Prazo |
|---|---|---|---|
| 1 | Limpeza e Padronização de Bases de Dados | 80 – 350 | 1 – 4 dias |
| 2 | Relatório de Vendas e Estoque no Excel | 150 – 550 | 2 – 6 dias |
| 3 | Dashboard Executivo no Power BI | 300 – 1.200 | 3 – 8 dias |
| 4 | Manutenção e Atualização Recorrente | 200 – 750/mês | mensal / quinzenal / semanal |

Cada pacote tem três níveis (Básico, Padrão e Premium), com preço em R$ e US$.

**Palavras-chave principais:** Dashboard Power BI · Organizar planilha · Relatório de vendas · Tratamento de dados

**Para alterar:** edite `PORTFOLIO_PACOTES.md` e gere o PDF novamente (ou peça ao Claude).

---

## 4. Base Fictícia de E-commerce (`E-COMMERCE EM EXCEL`)

**Período:** 01/01/2025 a 30/09/2026 · 9.798 linhas de venda · 5.991 pedidos · 46 produtos · 450 clientes · 8 vendedores
**Totais:** faturamento líquido ≈ R$ 1,96 milhão · margem ≈ 48% · ticket médio ≈ R$ 348

| Aba | Tipo | Conteúdo |
|---|---|---|
| LEIA-ME | — | Guia das abas e convenções |
| Resumo | Painel | KPIs, faturamento por mês, categoria e canal, 2 gráficos |
| dProdutos | Dimensão | Categoria, fornecedor, custo, preço, margem, estoque mínimo |
| dClientes | Dimensão | Cidade, UF, região, data de cadastro |
| dVendedores | Dimensão | Canal principal, região, meta mensal |
| dCalendario | Dimensão | Data, ano, mês, trimestre, dia da semana |
| fVendas | Fato | 1 linha por item de pedido (Item = 1 marca a 1ª linha do pedido) |
| fCompras | Fato | Entradas de estoque (NF de fornecedores) |
| fEstoque | Cálculo | Inicial + entradas − saídas, alerta OK / REPOR / SEM ESTOQUE |
| Base_Suja | Treino | Cadastro bagunçado de propósito (Pacote 1) |

**Convenções:**
- **Texto azul:** valores que podem ser alterados (custo, preço, estoque mínimo, estoque inicial)
- **Texto preto:** fórmulas, que recalculam sozinhas
- Pedidos **Cancelado** ficam fora do faturamento, do lucro e das saídas de estoque
- O **frete** fica só na 1ª linha de cada pedido e não entra no Valor_Liquido
- **Sazonalidade:** Black Friday (24 a 30/11, volume ×3) e dezembro forte; crescimento de cerca de 25% ao ano

**Tabelas nomeadas** (para Power Query e Power BI): `tbVendas`, `tbProdutos`, `tbClientes`, `tbVendedores`, `tbCalendario`, `tbCompras`, `tbEstoque`

**Modelo estrela no Power BI:**
```
tbProdutos[ID_Produto]   1 ──< tbVendas[ID_Produto]
tbClientes[ID_Cliente]   1 ──< tbVendas[ID_Cliente]
tbVendedores[ID_Vendedor]1 ──< tbVendas[ID_Vendedor]
tbCalendario[Data]       1 ──< tbVendas[Data]
```

### Recriar a base

```powershell
cd "D:\PROJETOS_CODE\BIO\E-COMMERCE EM EXCEL"
pip install openpyxl
python gerar_base.py
```

O script usa semente fixa (`random.seed(42)`), então gera sempre os mesmos dados. Para gerar dados diferentes, troque o número 42. Ao abrir no Excel as fórmulas são calculadas automaticamente; se aparecerem zeradas, pressione **Ctrl+Alt+F9**.

---

## 5. Amostras de portfólio a produzir

| # | Amostra | Pacote | Aba de origem | Entregável |
|---|---|---|---|---|
| 1 | "Antes x Depois" da limpeza de cadastro | 1 | Base_Suja | Prints lado a lado + GIF |
| 2 | Planilha de Vendas e Estoque | 2 | Resumo, fEstoque | Print do painel + vídeo 30s |
| 3 | Dashboard Executivo de e-commerce | 3 | fVendas + dimensões | Prints de 3 páginas + vídeo |

---

## 6. Conhecimentos técnicos necessários

| Bloco | Principais tópicos | Pacotes |
|---|---|---|
| Excel avançado | PROCX/ÍNDICE+CORRESP, SOMASES, funções de texto e data, Tabelas, Tabela Dinâmica, Segmentação, Formatação Condicional, Validação, Proteção | 1, 2 |
| Power Query | Importar fontes, limpar, mesclar/acrescentar consultas, transformar colunas em linhas, colunas condicionais, linguagem M | 1, 2, 3 |
| Power BI | Modelo estrela, tabela calendário, DAX (CALCULATE, inteligência de tempo, SUMX, RANKX), visuais, drill-through, tema, publicação, atualização agendada, Gateway, RLS | 3, 4 |
| Dados e automação | SQL (MySQL, PostgreSQL, SQL Server), Google Sheets, exportações de e-commerce, Python (pandas, openpyxl, requests) | 3, 4, extras |
| Negócio | Levantamento de requisitos, storytelling com dados, vídeo de entrega, LGPD, precificação e escopo | todos |

### Roteiro de 8 semanas (1 a 2 h por dia)

| Semana | Foco | Libera |
|---|---|---|
| 1–2 | Excel: tabelas, fórmulas, tabela dinâmica, formatação condicional | Pacotes 1 e 2 (Básico) |
| 3 | Power Query | Pacote 1 completo |
| 4 | Painel no Excel + conceitos de vendas/estoque | Pacote 2 completo |
| 5 | Power BI: modelagem + visuais | Pacote 3 (Básico) |
| 6–7 | DAX | Pacote 3 (Padrão) |
| 8 | Power BI Service, Gateway, RLS | Pacote 3 Premium + Pacote 4 |

**Fontes de estudo:** Microsoft Learn (trilha PL-300, gratuita) · YouTube: Hashtag Treinamentos, Karine Lago, Leonardo Karpinski · Kaggle (datasets de vendas)

---

## 7. Sincronizar com o GitHub

**Primeira vez** (pasta vazia):
```powershell
New-Item -ItemType Directory -Force -Path "D:\PROJETOS_CODE\BIO"
cd "D:\PROJETOS_CODE\BIO"
git clone -b claude/funny-brown-vetnlo https://github.com/jaajanjos/Docker-and-Kubernetes.git .
```

**Baixar atualizações:**
```powershell
cd "D:\PROJETOS_CODE\BIO"
git pull
```

**Enviar alterações feitas no seu computador:**
```powershell
cd "D:\PROJETOS_CODE\BIO"
git add .
git commit -m "Descreva a alteração"
git push
```

**Sem Git:** baixe o ZIP em
https://github.com/jaajanjos/Docker-and-Kubernetes/archive/refs/heads/claude/funny-brown-vetnlo.zip

> Os arquivos ainda não estão na branch `main`. Para juntá-los, abra um pull request da branch `claude/funny-brown-vetnlo` para `main`.

---

## 8. Histórico

| Data | Entrega |
|---|---|
| 20/10/2026 | Portfólio de pacotes (Markdown) |
| 20/10/2026 | Portfólio em PDF |
| 20/10/2026 | Base fictícia de e-commerce em Excel |
| 20/10/2026 | Script `gerar_base.py` e esta documentação |

---

## 9. Próximos passos

- [ ] Fazer o `git clone`/`git pull` em `D:\PROJETOS_CODE\BIO`
- [ ] Criar os perfis nas plataformas usando `PORTFOLIO_PACOTES.pdf`
- [ ] Limpar a aba **Base_Suja** com Power Query (amostra 1)
- [ ] Montar o painel de vendas e estoque no Excel (amostra 2)
- [ ] Construir o dashboard no Power BI com as tabelas `tb*` (amostra 3)
- [ ] Publicar os prints e vídeos das amostras nos perfis
