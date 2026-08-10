# Design — Automação de Download de NFC-e em Python

**Data:** 2026-08-10
**Base:** `especificacao_download_nfce.md` (extraída de `Autom Download NF V4.xlsm`)
**Status:** aprovado para implementação

## 1. Objetivo

Substituir a planilha VBA `Autom Download NF V4.xlsm` por um script Python que baixa
PDFs de NFC-e do portal Tec Fiscal / Linx (`https://tecfiscal.linx.com.br`) a partir de
uma lista de chaves de acesso, mantendo controle de progresso retomável e conciliação
com os arquivos já presentes em disco.

Volume alvo: ~2.300 chaves por rodada, ~4.200 arquivos já existentes no diretório de
destino.

## 2. Decisões de arquitetura

| Decisão | Escolha | Motivo |
|---|---|---|
| Estrutura | **Arquivo único** `download_nfce.py` | Escolha do usuário. Tradução direta, menos cerimônia. |
| Navegador | **Playwright + Chromium** | Espera automática por elemento clicável (elimina os sleeps fixos do VBA), driver embutido (sem conflito de versão de chromedriver), captura de download nativa via `expect_download`. |
| Login | **Híbrido** | Tenta automático com credenciais do `.env`; cai para manual se falhar ou se aparecer captcha/2FA. |
| Dados | **CSV** | `chaves.csv` (entrada) + `controle.csv` (estado). Sem escrita no `.xlsm`. |
| Python | **3.13** (já instalado) | — |

Dependências: `playwright`, `openpyxl`, `python-dotenv`, `pytest` (dev).

### 2.1 Arquivos entregues

```
download_nfce.py                  # todo o programa
requirements.txt
.env.example                      # modelo de credenciais (o .env real não vai pro git)
.gitignore                        # ignora .env, controle.csv, pdfs/, logs/
README.md                         # instalação e uso
tests/test_download_nfce.py       # testes das funções puras
```

### 2.2 Seções internas do `download_nfce.py`, nesta ordem

1. Imports e constantes (URLs, timeouts padrão)
2. `SELETORES` — dicionário com todos os localizadores da página
3. Configuração (leitura do `.env`)
4. Funções de CSV (chaves e controle)
5. Varredura de diretório
6. Importação do `.xlsm`
7. Funções de portal (login, filtro, busca, seleção, download)
8. Loop principal
9. `main()` com `argparse`

`SELETORES` é a primeira constante do arquivo justamente porque é a parte com maior
probabilidade de quebrar: mudança de layout no portal se conserta num só lugar.

## 3. Interface de linha de comando

```
python download_nfce.py importar --xlsm ARQUIVO [--aba Base] [--saida chaves.csv]
python download_nfce.py baixar [opções]
python download_nfce.py conferir [--controle controle.csv] [--destino ./pdfs]
python download_nfce.py status [--controle controle.csv]
python download_nfce.py inspecionar
```

Opções de `baixar`:

| Opção | Padrão | Efeito |
|---|---|---|
| `--chaves` | `chaves.csv` | Arquivo de entrada com as chaves |
| `--controle` | `controle.csv` | Arquivo de estado |
| `--destino` | `./pdfs` | Pasta onde os PDFs são salvos |
| `--limite N` | sem limite | Processa no máximo N chaves nesta rodada |
| `--lote N` | `150` | Reinicia o navegador a cada N chaves processadas |
| `--tentativas N` | `6` | Tentativas por chave antes de marcar `erro` |
| `--timeout S` | `15` | Timeout padrão (segundos) para localizar elementos |
| `--retentar` | desligado | Inclui na rodada também as chaves `nao_encontrado` (as `erro` já entram por padrão) |
| `--headless` | desligado | Roda sem janela. Exige login automático (ver §5.2) |

Equivalência com o VBA original:

| VBA | Python |
|---|---|
| Botão "Iniciar Download" | `baixar` |
| Botão "Conferir no Diretório" | `conferir` |
| Ler coluna B no olho | `status` |
| (não existia) | `importar`, `inspecionar` |

## 4. Modelo de dados

### 4.1 `chaves.csv`

Coluna única `chave`, uma chave de 44 dígitos por linha, com cabeçalho.
Gerado por `importar` a partir da aba `Base` do `.xlsm` (coluna A, da linha 2 em diante).

Regras da importação:
- Remove qualquer caractere não numérico antes de validar (a planilha tem chaves
  formatadas com espaços em algumas linhas).
- Aceita apenas chaves com exatamente 44 dígitos; as demais são reportadas com o
  número da linha de origem e descartadas.
- Remove duplicatas preservando a ordem de primeira aparição.
- Ao final imprime: total lido, aceitas, inválidas, duplicadas.

### 4.2 `controle.csv`

Colunas: `chave`, `status`, `tentativas`, `arquivo`, `erro`, `atualizado_em`.

| Status | Significado | Entra na próxima rodada? |
|---|---|---|
| `pendente` | Nunca tentada (equivale a coluna B vazia) | Sim |
| `baixado` | Arquivo salvo e confirmado em disco (equivale a `"Realizado"`) | Não |
| `nao_encontrado` | O portal executou a busca e não retornou nenhuma nota | Só com `--retentar` |
| `erro` | Falha técnica: timeout, elemento ausente, sessão perdida | Sim |

`arquivo` guarda o nome do arquivo salvo (apenas para `baixado`).
`erro` guarda a mensagem truncada em 200 caracteres.
`atualizado_em` é ISO-8601 local.

Chaves presentes em `chaves.csv` e ausentes do `controle.csv` são inseridas como
`pendente` no início da rodada. Chaves presentes no controle e ausentes da entrada são
preservadas (não se apaga histórico).

### 4.3 Divergências deliberadas em relação ao VBA

**Não existe status `"Extraído"`.** No VBA ele era gravado antes do download e nunca
atualizado depois, o que tornava a coluna B incapaz de dizer o que realmente baixou —
origem da necessidade de conferir a pasta manualmente. Aqui o status só é gravado
depois que o resultado da chave é conhecido.

**`nao_encontrado` é distinto de `erro`.** O VBA marcava `gv_pula = "X"` nos dois casos.
Separar permite responder "quantas notas de fato não existem no portal" versus "quantas
valem nova tentativa".

**O controle é regravado após cada chave.** Reescrever 2.300 linhas custa milissegundos;
em troca, uma queda de energia ou `Ctrl+C` perde no máximo uma chave.

## 5. Fluxo do comando `baixar`

1. Carrega `.env` e valida a pasta de destino (cria se não existir).
2. Carrega `chaves.csv` e `controle.csv`, sincroniza as chaves novas como `pendente`.
3. Monta a fila: `pendente` + `erro`, mais `nao_encontrado` se `--retentar`.
4. Inicia o navegador e faz login (§5.2).
5. Para cada chave da fila:
   1. Se já existe arquivo `<chave>.*` no destino → marca `baixado` e segue (curto-circuito
      de idempotência).
   2. Abre o painel de filtro.
   3. Limpa o campo de texto e digita a chave.
   4. Aplica o filtro.
   5. Aguarda o resultado estabilizar (§5.3).
   6. Se a tabela está vazia → `nao_encontrado`, sem retry e sem screenshot; segue.
   7. Marca o checkbox da primeira linha.
   8. Clica em download capturando o arquivo (§5.4).
   9. Fecha o modal residual, se houver.
   10. Grava `baixado` com o nome do arquivo.
   11. A cada `--lote` chaves processadas, recicla a sessão (§5.5).
6. Ao final imprime o resumo (mesmo formato do comando `status`).

### 5.1 Retry por chave

Cada chave tem até `--tentativas` execuções dos passos 2 a 9 do laço, com 3 segundos
entre elas. Uma falha técnica em qualquer etapa aborta a tentativa e a reinicia do
passo 2 do laço (reabrir o filtro), porque a página pode ter ficado num estado
intermediário.

A cada 3 erros seguidos na mesma chave, a página é recarregada (`goto` na URL de notas)
antes da próxima tentativa. Um erro de layout pode deixar a página num estado do qual o
retry normal (reabrir o filtro) nunca se recupera; recarregar reinicia o processo do
zero e evita ficar preso repetindo o mesmo erro até esgotar as tentativas.

Esgotadas as tentativas: grava `erro` com a mensagem, salva screenshot em
`logs/erros/<chave>.png` e segue para a próxima chave. **Uma chave nunca derruba a
rodada.**

`nao_encontrado` não consome tentativas — é um resultado válido da busca, não uma falha.

`Ctrl+C` é tratado: grava o controle, fecha o navegador e sai com o resumo parcial.

### 5.2 Login híbrido

1. Navega para `https://tecfiscal.linx.com.br/#/login`.
2. Se `TECFISCAL_USUARIO` e `TECFISCAL_SENHA` estão definidos, preenche e submete.
3. Aguarda até 20s por um marcador de sessão ativa (presença do botão de filtro na tela
   de notas, ou URL contendo `/app/emissao/notas`).
4. Se o marcador não aparecer:
   - **Modo com janela:** imprime `Faça login no navegador e pressione ENTER aqui.` e
     bloqueia no `input()`. Depois revalida o marcador. Este é o equivalente ao `MsgBox`
     do VBA.
   - **Modo `--headless`:** não há como intervir. Aborta com mensagem explícita
     orientando rodar sem `--headless`.
5. Se após a intervenção manual o marcador ainda não aparecer, aborta. Não há laço
   infinito de tentativas de login.
6. Se `TECFISCAL_CNPJ` estiver definido, seleciona a filial correspondente após o login.
   Opcional: o fluxo ativo do VBA não fazia essa seleção (só o módulo morto `base_busca`),
   então o padrão é não mexer na filial que o portal já traz.

### 5.3 Esperas

Todos os sleeps fixos do VBA (0.5s a 5s) são substituídos pela espera automática do
Playwright por elemento visível e habilitado, com timeout de `--timeout` segundos.

A única espera que precisa de tratamento próprio é a de resultado da busca (os 4s fixos
do VBA), porque "tabela vazia" é um resultado legítimo e não pode ser confundido com
"tabela ainda carregando". Sequência: aguarda o spinner (`nz-spin`) desaparecer, e então
aguarda ou uma linha na `tbody` ou o indicador de resultado vazio. O primeiro dos dois a
aparecer decide o caminho.

### 5.4 Captura do download

```python
with page.expect_download(timeout=60_000) as info:
    clicar(SELETORES["botao_download"])
download = info.value
extensao = Path(download.suggested_filename).suffix or ".pdf"
destino_arquivo = destino / f"{chave}{extensao}"
download.save_as(destino_arquivo)
```

Depois valida que o arquivo existe e tem tamanho maior que zero; caso contrário, trata
como falha técnica (entra no retry).

Duas consequências: o nome do arquivo passa a ser determinístico (`<chave>.pdf`), o que
é o que torna o curto-circuito de idempotência do passo 1 do laço confiável; e se o portal
entregar outro formato (`.zip`, `.xml`), a extensão original é preservada em vez de
assumir `.pdf` como o VBA fazia.

### 5.5 Reciclagem de sessão

A cada `--lote` chaves (padrão 150), fecha o contexto e o navegador, reabre e refaz o
login. Endereça o item 3 da §6 da especificação e o mesmo problema de acúmulo de memória
que causava "Memória insuficiente" no VBA.

É este ponto que justifica o login híbrido: numa rodada de 2.300 chaves são ~15
reciclagens, e login exclusivamente manual exigiria 15 intervenções. Com credenciais no
`.env`, o relogin é automático; se elas não estiverem configuradas, o modo manual é
solicitado novamente a cada reciclagem.

## 6. Comando `conferir`

Equivalente ao botão "Conferir no Diretório", como reconciliação e não como etapa
obrigatória do fluxo:

1. Varre `--destino` recursivamente (incluindo subpastas, como o VBA fazia).
2. De cada nome de arquivo extrai a primeira sequência de 44 dígitos encontrada.
3. Toda chave do controle com arquivo correspondente vira `baixado`.
4. Imprime três números: chaves com arquivo, chaves sem arquivo, arquivos sem chave
   correspondente no controle ("órfãos").

Uso principal: primeira execução sobre a pasta que já contém os ~4.200 arquivos
históricos, para não rebaixar o que já existe.

## 7. Comando `inspecionar`

Abre o navegador, executa o login (§5.2), navega até a tela de notas e chama
`page.pause()` — abre o Playwright Inspector, que permite apontar elementos na página e
obter o localizador correspondente.

Existe porque os XPaths da §5 da especificação são posicionais e não foi possível
validá-los sem acesso autenticado ao portal. É a ferramenta para corrigir `SELETORES`
quando o layout mudar.

## 8. Seletores

`SELETORES` mapeia cada ação a uma **lista ordenada** de localizadores. O código tenta
cada um em ordem e usa o primeiro que resolver para um elemento visível.

Ordem de preferência: `get_by_role` / `aria-label` / texto → seletor CSS estável
(`nz-`, `lx-`, `data-testid`) → XPath posicional da especificação como último recurso.

Ações necessárias:

| Chave | Ação |
|---|---|
| `botao_abrir_filtro` | Abre o painel de filtro |
| `campo_chave` | Campo de texto do filtro |
| `botao_aplicar_filtro` | Aplica o filtro / dispara a busca |
| `spinner` | Indicador de carregamento da tabela |
| `linha_resultado` | Primeira linha da tabela de resultado |
| `resultado_vazio` | Indicador de "nenhum registro" |
| `checkbox_nota` | Checkbox de seleção da primeira linha |
| `botao_download` | Botão de download |
| `botao_fechar_modal` | Fecha o modal pós-download |
| `marcador_logado` | Elemento que confirma sessão ativa |
| `campo_usuario`, `campo_senha`, `botao_entrar` | Tela de login |
| `seletor_filial` | Seleção de filial por CNPJ (opcional) |

Quando nenhum localizador da lista resolve, a mensagem de erro nomeia a chave que
falhou (ex.: `Não encontrei 'botao_download'`), para que a correção seja direta.

## 9. Configuração e segurança

`.env` (fora do controle de versão, listado no `.gitignore`):

```
TECFISCAL_USUARIO=
TECFISCAL_SENHA=
TECFISCAL_CNPJ=
```

`.env.example` versionado com as chaves vazias. Isso substitui a prática do VBA de
gravar a senha em texto puro na célula (2, 200) da planilha (item 2 da §6). O script
nunca imprime a senha em log nem em mensagem de erro.

Se `TECFISCAL_SENHA` estiver vazio, o script funciona — apenas sempre pelo caminho de
login manual.

## 10. Testes

`tests/test_download_nfce.py`, com `pytest`. Cobre as funções que não dependem do
navegador:

**CSV de controle**
- Round-trip: gravar e reler preserva todos os campos.
- Sincronização: chave nova em `chaves.csv` entra como `pendente`.
- Sincronização: chave que saiu de `chaves.csv` permanece no controle.
- Montagem da fila: inclui `pendente` e `erro`, exclui `baixado`.
- Montagem da fila com `--retentar`: inclui também `nao_encontrado`.
- Retomada: controle parcialmente preenchido gera fila apenas com o que resta.

**Importação do `.xlsm`**
- Extrai chaves válidas da coluna A a partir da linha 2.
- Descarta chave com menos ou mais de 44 dígitos, reportando a linha.
- Remove máscara (espaços, pontos) antes de validar.
- Remove duplicatas preservando a ordem.

**Varredura de diretório**
- Encontra arquivos em subpastas.
- Extrai a chave de 44 dígitos do nome do arquivo.
- Ignora arquivos cujo nome não contém chave (contados como órfãos).
- Casa a chave independentemente da extensão (`.pdf`, `.zip`).

Os testes usam `tmp_path` do pytest; um `.xlsx` mínimo é gerado com `openpyxl` no
próprio teste, sem depender do `.xlsm` real.

**Fora de cobertura automatizada:** as funções de portal. É o custo aceito da escolha de
arquivo único sem camada de abstração de página; a mitigação é o comando `inspecionar`
mais uma primeira rodada com `--limite 3` para validação manual ponta a ponta.

## 11. Critérios de aceite

1. `importar` gera `chaves.csv` com as chaves válidas da aba `Base` do `.xlsm` e reporta
   o que descartou.
2. `conferir` sobre a pasta histórica marca como `baixado` toda chave que já tem arquivo.
3. `baixar --limite 3` completa três chaves ponta a ponta e os PDFs aparecem no destino
   nomeados como `<chave>.pdf`.
4. Interromper com `Ctrl+C` e rodar `baixar` de novo retoma sem reprocessar o que já
   está `baixado`.
5. Uma chave que falha não interrompe a rodada, e aparece como `erro` com screenshot em
   `logs/erros/`.
6. `status` mostra a contagem por status.
7. `pytest` passa.

## 12. Riscos

| Risco | Mitigação |
|---|---|
| Os XPaths da especificação estão desatualizados (não pudemos validar sem acesso logado) | Listas de localizadores com fallback; comando `inspecionar`; erro nomeia o seletor que falhou |
| O portal tem captcha ou 2FA no login | Modo manual do login híbrido cobre; documentado que `--headless` não é usável nesse caso |
| Rate limiting ou bloqueio do portal por volume | `--limite` permite fatiar a rodada; o estado retomável torna rodadas parciais normais |
| O portal entrega ZIP em vez de PDF | Extensão vem de `suggested_filename`; a checagem de idempotência usa `<chave>.*` |
