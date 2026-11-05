# Gestão de Manutenção Predial · IEMA - SUBAD

Sistema para registrar e acompanhar as demandas de manutenção predial do IEMA, com **banco de dados em planilha Google**, site publicado no **GitHub Pages** e **backup mensal automático por e-mail**. Não há servidor próprio: a ponte entre o site e a planilha é um script do Google (Apps Script) que fica dentro da própria planilha.

## Arquivos do pacote

| Arquivo | Onde vai | Para que serve |
|---|---|---|
| `index.html` | GitHub | O sistema completo (telas, gráficos e relatórios) |
| `config.js` | GitHub | Endereço de conexão com a planilha Google |
| `README.md` | GitHub | Este guia |
| `apps-script/Codigo.gs` | Planilha Google | Código que transforma a planilha em banco de dados e envia os backups |

Sem o endereço em `config.js`, o sistema abre em **modo demonstração**, com dados fictícios salvos só no navegador. Usuários de teste, todos com a senha `iema123`: `master`, `gestor`, `solicitante`, `contratada` e `leitura`.

---

## Instalação (cerca de 15 minutos)

### 1. Criar a planilha
1. Com a conta Google institucional, abra o Google Planilhas e crie uma **planilha em branco**.
2. Dê um nome, por exemplo **Manutenção Predial IEMA**.

### 2. Colar o código
1. Na planilha, abra **Extensões > Apps Script**.
2. Apague todo o conteúdo do arquivo `Código.gs` que aparece.
3. Abra o arquivo `apps-script/Codigo.gs` do pacote, copie **todo** o conteúdo e cole no editor.
4. Antes de salvar, troque a senha provisória na primeira linha de configuração:
   `var SENHA_INICIAL_MASTER = 'Iema@2026';`
5. Clique em **Salvar** (ícone de disquete).
6. Recomendado: em **Configurações do projeto** (ícone de engrenagem), defina o fuso horário **(GMT-03:00) São Paulo**.

### 3. Instalar a estrutura
1. No editor, na lista de funções ao lado do botão **Executar**, escolha **instalar**.
2. Clique em **Executar**.
3. O Google pede autorização: escolha a sua conta, clique em **Avançado > Acessar o projeto** e em **Permitir**.
   - O script pede acesso à planilha, ao Google Drive (pasta de backups), ao envio de e-mails e a "serviço externo", usado para gerar o arquivo Excel do backup.
4. Volte à planilha. Devem aparecer as abas Demandas, Notas Fiscais, Empenhos, Lotes, Usuários, Perfis, Histórico e Configurações, além do menu **Manutenção Predial**.
5. A instalação também cria o gatilho que envia o backup no **1º dia útil de cada mês**.

### 4. Publicar o script como App da Web
1. No editor do Apps Script, clique em **Implantar > Nova implantação**.
2. Na engrenagem de "Selecione o tipo", escolha **App da Web**.
3. Preencha assim:
   - **Executar como:** Eu
   - **Quem pode acessar:** Qualquer pessoa
4. Clique em **Implantar** e copie o **URL do app da Web**. Ele termina em `/exec`.

O URL é público, mas o sistema só entrega dados para quem entra com login e senha. Se a opção **Qualquer pessoa** não aparecer, peça ao administrador do Google Workspace para liberá-la; sem ela, o site não consegue falar com a planilha.

### 5. Informar o endereço no `config.js`
Abra `config.js` e cole o URL entre as aspas:

```js
window.IEMA_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/XXXXXXXX/exec'
};
```

### 6. Publicar no GitHub
1. Crie um repositório no GitHub (botão **+ > New repository**).
2. Clique em **uploading an existing file** e envie `index.html`, `config.js` e `README.md`. O `Codigo.gs` não precisa ir para o GitHub.
3. Clique em **Commit changes**.
4. Em **Settings > Pages**, escolha **Deploy from a branch**, depois **main** e **/ (root)**, e clique em **Save**.
5. Em 1 a 3 minutos, o endereço do site aparece no topo da página de Pages.

### Já instalou a versão anterior? Atualize assim
1. No editor do Apps Script, apague o código antigo, cole o novo `Codigo.gs` e salve.
2. Execute a função **instalar** de novo. Ela não apaga dados: cria as abas do Looker Studio, acrescenta os novos campos em Configurações, troca a lista antiga de setores pelas siglas e fixa os alertas em 70% quando ainda estiverem vazios.
3. Vá em **Implantar > Gerenciar implantações**, clique no lápis, escolha **Versão: Nova versão** e **Implantar**. O endereço `/exec` continua o mesmo.
4. Envie o novo `index.html` ao GitHub, substituindo o anterior. O `config.js` não muda.

### 7. Primeiro acesso
1. Abra o site e entre com o login `master` e a senha provisória definida no passo 2.
2. O sistema pede uma senha pessoal. Defina-a.
3. Em **Área do Gestor > Configurações**, cadastre:
   - processo de gestão, processo de execução, número do contrato, número e ano da ata, contratada e vigência;
   - os **lotes**, com número, descrição e valor (o valor global da vigência passa a ser a soma deles);
   - os **empenhos**, novos e reforços;
   - o valor global da vigência (se não usar valores nos lotes) e os percentuais de alerta de execução;
   - os setores e unidades (já vêm as siglas do IEMA) e os e-mails que recebem o backup.
4. Em **Área do Gestor > Usuários e perfis**, crie os usuários de cada pessoa.

---

## Como o sistema funciona

### Registro de demandas
- Toda demanda é obrigatoriamente vinculada a um **lote**. Quem tem permissão financeira pode criar um lote novo direto do formulário.
- O **prazo estimado** é informado em **dias, semanas ou meses**, e o término previsto é calculado automaticamente.
- Em demandas **programadas**, o **valor estimado é obrigatório**; o formulário avisa enquanto ele não for preenchido.
- Unidade demandante, local e setores atendidos usam a lista de **siglas dos setores e unidades** (com filtro para localizar a sigla).
- Há campos para o **nº da Ordem de Serviço** e o **nº do Atendimento (Help Desk)**.
- Valores em reais são digitados só com números: os pontos, a vírgula e o "R$" aparecem automaticamente.

### Andamento por marcação
- O gestor marca cada etapa na própria demanda: **Demanda inicial**, **Aprovação**, **Execução**, **Ateste** e **Pagamento**.
- A Aprovação aparece apenas para demandas **programadas** e é opcional.
- Ao marcar, o sistema grava a data e o usuário.
- Para alterar ou desmarcar uma etapa já marcada, o sistema pergunta se o gestor deseja realizar a alteração e pede a **senha de quem está fazendo a ação**, junto com o motivo.
- Depois do pagamento, qualquer dado continua editável, sempre com senha e motivo (financeiro, reparo, cadastro, prazo ou outro).

### Notas fiscais, empenhos e saldos
- Depois do ateste, registra-se a **Nota Fiscal**: número, valor do serviço, data e empenho.
- O valor da nota é deduzido automaticamente do **empenho** (somado aos seus reforços), do **lote** e do **valor global da vigência**.
- O sistema não aceita uma nota maior que o saldo do empenho.
- Há dois alertas de execução, exibidos no painel, no dashboard, na aba Financeiro e nos relatórios financeiros:
  - **valor global** (e cada lote) executado em **70% ou mais**;
  - **empenho** (somado aos reforços) executado em **70% ou mais**.
- Os dois alertas já começam **fixos em 70%**. Para mudar, use **Alterar valores** em Configurações: o sistema pergunta se o gestor realmente deseja alterar e, se sim, pede a senha e o motivo.
- O valor global da vigência tem o botão **Salvar**; depois de gravado, só muda pelo mesmo botão **Alterar valores**.

### Exclusão de demandas
- Gestores (perfis com permissão de excluir) apagam uma demanda pelo botão **Excluir** no detalhe ou pelo ícone **×** na lista de demandas.
- O sistema pergunta se deseja realmente excluir, mostra se há Notas Fiscais vinculadas (que também serão excluídas e voltarão ao saldo dos empenhos) e pede o motivo e a senha do gestor.
- A exclusão fica registrada no histórico. Para manter o registro da demanda, prefira **Cancelar demanda**.

### Setores e unidades
Em **Área do Gestor > Configurações > Listas de apoio**:
- **Inserir setor ou unidade** acrescenta uma ou várias siglas (separadas por vírgula) e já salva.
- **Importar Excel** lê um arquivo .xlsx, .xls ou .csv com a coluna "Sigla" (ou as siglas na primeira coluna), mostra quantas são novas e permite **substituir** a lista ou **acrescentar** as novas. A lista é atualizada logo após o envio.
- A leitura do Excel usa uma biblioteca carregada da internet na primeira vez.

### Histórico
Toda inserção e alteração fica registrada na aba **Histórico** da planilha e na Área do Gestor, com:
- data e hora;
- usuário e perfil;
- campo alterado, valor anterior e valor novo;
- motivo, quando houver.

### Perfis
O gestor master cria perfis e define, para cada um, o que pode fazer: ver todas as demandas ou só as próprias, registrar, editar cada grupo de campos, marcar etapas, lançar notas e empenhos, acessar a Área do Gestor, alterar configurações e gerenciar usuários.

Os perfis já criados são:

| Perfil | O que pode fazer |
|---|---|
| **Gestor Master** | Tudo, inclusive gerenciar usuários e perfis |
| **Gestor** | Tudo, exceto gerenciar usuários e perfis |
| **Solicitante** | Registra e acompanha as próprias demandas; pode editá-las até o gestor marcar a demanda inicial |
| **Contratada (consulta)** | Somente leitura das demandas e do cronograma |
| **Somente leitura** | Consulta, incluindo dashboards e relatórios |

### Segurança da sessão
Após **10 minutos sem uso** do mouse ou do teclado, o sistema avisa e, um minuto depois, encerra a sessão, exigindo login e senha novamente.

### Dashboards e relatórios
- Filtros por período: hoje, últimos 7, 30 ou 90 dias, este mês, mês anterior, um mês específico, trimestres, ano, vigência ou datas livres.
- Comparação com o período anterior, com o mesmo período do ano anterior ou com qualquer mês.
- São 11 relatórios, entre gerenciais, financeiros e de controle. Saem em **PDF** (download direto, ajustado à página A4; tabelas largas saem em paisagem), **HTML**, **Excel** e **CSV**, além da impressão.
- No PDF, a largura de cada coluna é calculada pelo conteúdo: números, datas, valores e códigos ficam numa linha só, e os textos longos ocupam o espaço restante.
- No detalhe de cada demanda, o botão **Ficha (PDF)** baixa a ficha completa da demanda.
- O PDF direto usa uma biblioteca carregada da internet na primeira vez. Sem internet, o sistema abre a impressão para salvar como PDF.

### Google Looker Studio
- A planilha mantém três abas só para o Looker Studio: **Looker - Demandas**, **Looker - Notas Fiscais** e **Looker - Saldos**. Elas são refeitas a cada gravação do sistema e têm as mesmas informações usadas nos filtros do gestor (período, mês, ano, lote, modalidade, situação, prazo, unidade, local, sistema, empenho e alertas).
- Na Área do Gestor, aba **Dashboard**, o botão **Abrir no Google Looker Studio** mostra um link para cada aba: **Demandas**, **Notas Fiscais** e **Saldos**. Cada link cria um relatório já conectado àquela aba.
- Na primeira vez: use o link de **Demandas**. No Looker, acrescente as outras abas em **Adicionar dados > Google Planilhas**, inclua os controles de filtro e os gráficos e clique em **Editar e compartilhar** para salvar. Cole o endereço do relatório salvo em **Configurações > Google Looker Studio**; depois disso, o botão abre direto esse relatório.
- Os filtros usados no Looker Studio servem apenas para consulta e apresentação: não alteram o sistema nem a tela do gestor.
- O Looker Studio atualiza os dados conforme a frequência configurada na fonte de dados; para ver na hora, use **Atualizar dados** no próprio Looker.

### Backup: duas formas
1. **Automático:** no 1º dia útil de cada mês (dias úteis desconsideram fins de semana e feriados), a planilha envia para os e-mails configurados a planilha completa em Excel (.xlsx), a base em JSON e as tabelas em CSV.
   - Cópias ficam na pasta **Backups - Manutenção Predial IEMA** do Google Drive.
   - Destinatários iniciais: `cleilson.costa@iema.es.gov.br` e `subad@iema.es.gov.br`.
   - Também é possível enviar na hora, pelo botão **Enviar backup por e-mail agora** ou pelo menu **Manutenção Predial** da planilha.
2. **Manual:** na Área do Gestor, aba **Backup**, há botões para baixar em Excel, JSON e CSV e para restaurar um JSON (somente o gestor master).

### Atualização automática
Cada alteração é gravada na planilha na mesma hora. O site confere a cada 45 segundos se houve mudança feita por outro usuário e atualiza a tela sozinho.

---

## Manutenção

- **Atualizar o código do Apps Script:** cole a nova versão, salve e vá em **Implantar > Gerenciar implantações > editar (lápis) > Versão: Nova versão > Implantar**. O URL continua o mesmo.
- **Esqueceu a senha do master:** na planilha, use o menu **Manutenção Predial > Redefinir senha do usuário master**. A senha volta a ser a provisória do código.
- **Não altere os títulos da primeira linha** das abas da planilha. O sistema localiza as colunas por esses títulos. Novas colunas podem ser acrescentadas no fim.
- **Acesso à planilha:** compartilhe a planilha só com quem administra o sistema. Quem edita a planilha diretamente altera os dados sem passar pelas regras e sem registro no histórico.
- **Senhas:** são guardadas apenas como código criptográfico (SHA-256 com salt), em colunas ocultas da aba Usuários.

## Problemas comuns

| Sintoma | Causa provável e solução |
|---|---|
| "Sem conexão com a planilha Google" | O URL em `config.js` está errado, ou o App da Web não está com acesso "Qualquer pessoa". |
| "A aba ... não existe" | Execute a função `instalar` de novo. Ela também atualiza a estrutura sem apagar dados. |
| Alteração no código não teve efeito | Crie uma **nova versão** da implantação (veja "Manutenção"). |
| Backup não chegou | Confira os e-mails em Configurações e se o backup automático está marcado. Veja também **Execuções** no editor do Apps Script. |
| O site abre em modo demonstração | O `config.js` não foi enviado ao GitHub ou está sem o URL. |
