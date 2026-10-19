# Handoff Vagas

As vagas do site ficam numa aba própria, **Vagas do site**, dentro da Handoff — Base. Um script separado (`vagas.gs`), no mesmo Apps Script das ligas, confere essa aba todo dia. Nada disso usa o Claude.

O script não toca nas abas nem nas automações das ligas:
- Todas as funções têm o prefixo `vagas`.
- Ele não cria outro `onOpen`. O menu entra por um gatilho próprio.
- Não muda o fuso da planilha nem a ordem das abas.
- Só apaga os próprios gatilhos.

| Arquivo | O que é |
|---|---|
| `vagas.gs` | Script das vagas: cria a aba, confere os links todo dia e manda o e-mail |
| `site/index.html` | O site, para publicar no GitHub Pages |

## Instalação (uma vez, cerca de 10 minutos)

1. **Script.**
   - Na Handoff — Base, abra **Extensões > Apps Script**.
   - Clique no **+** ao lado de *Arquivos*, escolha **Script** e dê o nome **vagas**.
   - Cole o conteúdo de `vagas.gs` e salve. Não mexa nos outros arquivos.
2. **Instalar.**
   - No topo, escolha a função **vagasInstalar** e clique em **Executar**.
   - Autorize. O Google pode pedir permissões novas, porque agora o script também abre links e manda e-mail.
   - Isso cria a aba **Vagas do site** no fim da planilha, com as 50 vagas. Também cria o menu **Handoff Vagas** e o gatilho das 7h, e roda a primeira conferência.
3. **Publicar só a aba das vagas.**
   - Em **Arquivo > Compartilhar > Publicar na Web**, escolha **Vagas do site**. **Não escolha "Documento inteiro"**, ou as abas das ligas ficam públicas.
   - Escolha o formato **CSV** e clique em **Publicar**. Copie o link.
   - Não mude o compartilhamento da planilha para "qualquer pessoa com o link". Publicar só a aba já basta.
4. **Site.**
   - No `site/index.html`, cole o link entre as aspas de `window.HANDOFF_CSV_URL = "";`.
   - Suba o arquivo num repositório público do GitHub, por exemplo `handoff-vagas`.
   - Ative o Pages em **Settings > Pages > Deploy from a branch > main / (root)**.
   - O endereço fica `https://SEU-USUARIO.github.io/handoff-vagas/`, e é esse link que vai na bio.

## Rotina diária

- **E-mail.** Às 7h chega o e-mail **"Handoff Vagas · N para conferir, M fechando"**. Ele só chega quando há algo para olhar.
- **O que conferir.** Na aba **Vagas do site**, filtre **conferir = SIM**:
  - **motivo** diz o que olhar: página mudou, prazo passou ou link fora do ar.
  - **datas_na_pagina** mostra as datas que o script achou na página, como pista.
- **Corrigir.** Corrija a linha, selecione-a e use **Handoff Vagas > Marcar linhas selecionadas como conferidas**.
- **Esconder uma vaga.** Para tirar uma vaga do site sem apagar, coloque **publicar = não**.
- **Vaga nova.** Adicione uma linha no fim. As colunas com lista suspensa só aceitam os valores certos.
- **Quando o site atualiza.** O site mostra as mudanças em alguns minutos. É o tempo que o Google leva para atualizar o CSV publicado.
- **Conferir na hora.** Use **Handoff Vagas > Conferir agora**.

## O que o script faz e o que não faz

- Abre cada link e marca os que saíram do ar e os sites que bloqueiam checagem automática. Estes últimos você abre na mão.
- Detecta mudança só nos trechos que falam de prazo, inscrição, edital e datas. Notícias e banners não disparam alerta.
- Em PDF, detecta apenas se o arquivo foi trocado.
- **Não lê o prazo sozinho.** Ele aponta o que mudou e você confirma, porque cada site escreve datas de um jeito.
- Avisa quando o prazo de uma vaga recorrente passou, para você procurar o próximo ciclo.
- Leva poucos minutos por dia e divide a cota diária do Google com as automações das ligas, sem chegar perto do limite.

## Colunas da aba Vagas do site

**Você edita:** publicar, titulo, instituicao, categoria, pais, cidade, modalidade, publico, custo, inscricao_abre, inscricao_fecha (dd/mm/aaaa), periodo_programa, recorrente, ciclo, resumo, link, conferido_em.

**O script preenche (cinza):** conferir, motivo, situacao_link, pagina_mudou_em, datas_na_pagina, ultima_checagem, assinatura.
