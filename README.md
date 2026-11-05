# Gestão de Manutenção Predial (IEMA)

Sistema web para acompanhar as manutenções prediais do Instituto Estadual de Meio Ambiente (IEMA), da solicitação ao pagamento. O fluxo segue a Instrução de Serviço de manutenção predial (minuta) e os dois fluxogramas que a acompanham: **modalidade Programado** e **modalidade Contínuo** (demandas rotineiras).

O site é estático: HTML, CSS e JavaScript puros, sem instalação, sem banco de dados e sem bibliotecas externas. Publica direto no **GitHub Pages**.

![Painel](docs/img/painel.png)

## O que o sistema faz

| Tela | Para que serve |
|---|---|
| **Painel** | Quantas demandas estão em cada uma das 7 etapas, atrasos, próximos términos de obra, saldo da OS Contínuo e movimentações recentes. |
| **Demandas** | Cadastro e consulta, com filtros por modalidade, etapa, prazo, local, setor, sistema, criticidade, situação e ano. Exporta CSV. |
| **Cronograma de obras** | Gráfico de Gantt com o tempo real em cada etapa (cores) e a execução prevista (tracejado). Escala em semanas, meses ou trimestres, agrupado por local, unidade, etapa ou modalidade. |
| **Etapas do fluxo** | Quadro com as 7 etapas. Os cartões podem ser arrastados para a etapa seguinte ou anterior. Cada cartão mostra quanto do prazo da etapa já foi consumido. |
| **Setores atendidos** | Demandas, atrasos e valores por setor, por local (sede e Unidades de Conservação) e por sistema afetado, com mapa de calor. |
| **Área do Gestor** | Dashboard, 7 relatórios gerados automaticamente, controle de saldo da OS Contínuo, alertas e registro de atividades. Acesso por PIN. |
| **Configurações** | Contrato, valor da OS Contínuo, prazos por etapa, listas de setores, locais, fiscais e gestores, PIN e importação/exportação de dados. |

### As 7 etapas (mesmas cores dos fluxogramas)

| Etapa | Responsável | Base na IS | Para avançar, o sistema exige |
|---|---|---|---|
| 1. Solicitação | Unidade Demandante | Arts. 7º e 8º | Campos obrigatórios e confirmação da solicitação completa |
| 2. Deliberação | Diretor Setorial e Diretor-Geral | Art. 9º | Registro formal da deliberação |
| 3. Instrução técnica e orçamentária | SUBAD, Contratada e Fiscal | Arts. 10 a 12 | Levantamento técnico, planilha, memorial, Checklist (Anexo I) e valor estimado |
| 4. Autorização (OS) | Ordenador de Despesas | Arts. 13 a 19 | Nº da OS; no Programado, OS assinada; no Contínuo, **saldo suficiente** |
| 5. Execução | Contratada | Arts. 14 e 21 | Relatório de Finalização de Demanda |
| 6. Fiscalização e ateste | Fiscal do Contrato | Arts. 22 e 24 | Ateste (Anexo III) e valor medido |
| 7. Pagamento | Gestor do Contrato | Art. 23 | Nota Fiscal conferida e número informado |

Também há os desvios do fluxo: **devolver etapa** (com motivo), **suspender** (solicitação incompleta, Art. 8º, § 3º), **não aprovar** (na deliberação) e **reativar**. Tudo fica no histórico da demanda e no registro de atividades.

### Saldo da OS Contínuo (Arts. 18 e 19)

O valor global anual (R$ 300.000,00 na minuta, editável) é comprometido quando uma demanda Contínuo entra em execução. Se o saldo não cobre a demanda, o sistema **bloqueia** o avanço e orienta a emissão de OS complementar, que pode ser registrada na Área do Gestor. Há alerta quando o comprometimento passa do percentual configurado (80% por padrão).

### Relatórios automáticos (Área do Gestor)

1. Relatório gerencial consolidado
2. Prazos e atrasos
3. Atendimento por setor
4. Execução da OS Contínuo
5. Cronograma de obras (previsto x realizado, com desvio em dias)
6. Financeiro e medições
7. Ficha da demanda (dados, documentos e histórico)

Todos podem ser filtrados por exercício, modalidade e local, e saem em **PDF** (Imprimir > Salvar como PDF), **HTML** autossuficiente ou **CSV**.

![Relatórios](docs/img/gestor-relatorios.png)

## Publicar no GitHub

1. Crie um repositório e envie todos os arquivos desta pasta (mantendo a estrutura).
2. Em **Settings > Pages**, em *Build and deployment*, escolha **Source: GitHub Actions**.
3. Faça um *push* na branch `main`. O workflow `.github/workflows/pages.yml` publica o site e mostra o endereço ao final.

Para testar no computador sem publicar:

```bash
python3 -m http.server 8000
# abra http://localhost:8000
```

Abrir o `index.html` com dois cliques também funciona; nesse modo o navegador não lê `data/demandas.json` e o sistema começa com os dados de exemplo.

## Onde ficam os dados

- **No navegador de cada usuário** (localStorage). Cada pessoa vê o que cadastrou naquele navegador.
- **No repositório**, de forma opcional, em `data/demandas.json`. Quem abre o site sem dados locais recebe essa versão. Para atualizar: em Configurações, use **Exportar dados (JSON)**, substitua `data/demandas.json` e faça *commit*. Usuários que já têm dados locais usam **Recarregar data/demandas.json**.
- Sem `data/demandas.json`, o sistema gera **dados fictícios de demonstração** com datas relativas ao dia de hoje. Antes do uso real, use **Apagar todas as demandas** em Configurações.

> **Atenção à privacidade.** Em repositório público, o conteúdo de `data/demandas.json` fica público. Para dados reais, use repositório privado com Pages restrito à organização, ou mantenha os dados apenas nos navegadores e compartilhe os arquivos exportados pelos canais oficiais.

> **Sobre o PIN da Área do Gestor.** Ele evita acessos acidentais, mas não é controle de segurança: num site estático, quem tem o endereço consegue ler o código. O controle real de acesso deve ser feito no GitHub.

O canal oficial de solicitação continua sendo o definido pela Instrução de Serviço (Art. 7º). Este sistema é uma ferramenta de acompanhamento e não substitui o processo administrativo.

## Estrutura

```
├── index.html                  página única da aplicação
├── assets/
│   ├── css/styles.css          identidade visual, temas claro/escuro, impressão
│   ├── img/favicon.svg
│   └── js/
│       ├── core.js             etapas, regras da IS, prazos, saldo, armazenamento
│       ├── seed.js             dados fictícios de demonstração
│       ├── charts.js           gráficos SVG (sem bibliotecas)
│       ├── ui.js               detalhe da demanda, formulário, modais
│       ├── views.js            painel, demandas, cronograma, etapas, setores, configurações
│       ├── gestor.js           dashboard, relatórios, OS Contínuo, alertas, registro
│       └── app.js              navegação, tema e atalhos
├── data/LEIA-ME.md             como versionar os dados no repositório
├── docs/
│   ├── MANUAL.md               guia de uso por perfil
│   ├── MODELO_DE_DADOS.md      formato do JSON
│   ├── fluxogramas/            fluxogramas Programado e Contínuo (PDF)
│   └── img/                    capturas de tela
└── .github/workflows/pages.yml publicação automática
```

## Atalhos

- `/` busca demandas
- `N` abre uma nova demanda
- `Esc` fecha o detalhe ou a janela aberta

## Personalização rápida

- **Prazos por etapa, valor da OS Contínuo, setores, locais, fiscais e gestores:** tela Configurações.
- **Cores das etapas:** `C.ETAPAS` em `assets/js/core.js` e as variáveis `--e1` a `--e7` em `assets/css/styles.css`.
- **Regras para avançar de etapa:** função `R.requisitos` em `assets/js/core.js`.
- **Documentos da instrução:** lista `C.DOCS` em `assets/js/core.js`.
