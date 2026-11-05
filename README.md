# Gestão de Manutenção Predial (IEMA)

Sistema para acompanhar as manutenções prediais do IEMA, da solicitação ao pagamento, nas modalidades **Programado** e **Contínuo**, conforme a Instrução de Serviço de manutenção predial.

Tudo está em um único arquivo: **`index.html`**.

## Publicar no GitHub

1. Crie um repositório público (botão **+ > New repository**).
2. Clique em **uploading an existing file**, envie o `index.html` e este `README.md` e clique em **Commit changes**.
3. Abra **Settings > Pages**. Em *Source*, escolha **Deploy from a branch**; em *Branch*, escolha **main** e **/ (root)**. Clique em **Save**.
4. Em 1 a 3 minutos, o endereço do site aparece no topo da página de Pages.

Para usar sem publicar, basta abrir o `index.html` no navegador.

## Primeiro uso

Em **Configurações**:
- preencha o contrato e o número da OS Contínuo;
- ajuste os prazos por etapa e as listas de setores, locais, fiscais e gestores;
- troque o PIN da Área do Gestor (o inicial é **1234**);
- clique em **Apagar todas as demandas** para remover os dados fictícios de exemplo.

## O que tem no sistema

- **Painel:** demandas por etapa, atrasos, próximos términos e saldo da OS Contínuo.
- **Demandas:** cadastro, filtros e exportação em CSV.
- **Cronograma de obras:** gráfico de Gantt com o tempo em cada etapa e a execução prevista.
- **Etapas do fluxo:** quadro de arrastar, com prazo consumido em cada cartão.
- **Setores atendidos:** demandas e valores por setor, local e sistema.
- **Área do Gestor:** dashboard, 7 relatórios automáticos (PDF, HTML e CSV), saldo da OS Contínuo, alertas e registro de atividades.

## Dados

As demandas ficam salvas **no navegador** de cada usuário, e não no GitHub. Para levar os dados para outro computador ou fazer cópia de segurança, use **Configurações > Exportar dados (JSON)** e depois **Importar dados (JSON)**.

O PIN da Área do Gestor evita acessos acidentais, mas não é um controle de segurança.
