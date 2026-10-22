# Automação de Download de NFC-e

Baixa PDFs de NFC-e do portal Tec Fiscal / Linx a partir de uma lista de chaves de
acesso, com controle de progresso retomável. Substitui a planilha VBA
`Autom Download NF V4.xlsm`.

Design: [docs/superpowers/specs/2026-08-10-download-nfce-design.md](docs/superpowers/specs/2026-08-10-download-nfce-design.md).

## Instalação

O jeito mais simples é dar duplo-clique em **`RUN.bat`**. Ele funciona em qualquer
máquina Windows: acha o Python (ou instala o 3.13 pelo `winget`), cria um ambiente
virtual próprio em `.venv`, instala o `requirements.txt` e o Chromium do Playwright, e
roda o robô. Na segunda execução em diante a instalação é pulada; ela só se repete se o
`requirements.txt` mudar. Se o proxy bloquear o download do Chromium, o robô usa o
Microsoft Edge da máquina.

Manualmente (Python 3.9+, recomendado 3.13):

```powershell
py -3.13 -m pip install -r requirements.txt
py -3.13 -m playwright install chromium
```

### Login: o portal tem reCAPTCHA

Verificado em 2026-08-10 contra o portal real. Isso tem três consequências práticas:

- **`--headless` não é usável.** Sem janela não há como resolver o captcha.
- **O `.env` é quase inútil sozinho.** Mesmo com usuário e senha preenchidos, o captcha
  interrompe o login automático. O `.env` continua valendo caso o portal deixe de exigir
  captcha para o seu usuário.
- **Use sempre `--perfil`.** Com um diretório de perfil, os cookies de sessão sobrevivem
  às reciclagens de navegador _e_ entre execuções: você resolve o captcha uma vez e as
  rodadas seguintes entram direto. Sem isso, uma rodada de 2.300 chaves pediria captcha
  a cada 150 chaves — cerca de 15 vezes.

O robô não pede ENTER: ele detecta sozinho quando você terminou de logar e segue.

Credenciais (opcionais):

```powershell
Copy-Item .env.example .env
notepad .env
```

O `.env` não vai para o controle de versão. A senha nunca é impressa em log nem em
mensagem de erro. O `perfil/` também é ignorado pelo git — ele guarda cookies de sessão.

### Onde colocar os arquivos

Aponte `--destino` e `--controle` para uma pasta **fora do OneDrive**. O controle é
regravado a cada chave — milhares de vezes por rodada —, e cada regravação numa pasta
sincronizada é uma chance de o arquivo estar bloqueado. O programa tolera isso (insiste
e, no limite, escreve sem atomicidade), mas o certo é não provocar.

## Uso

### 1. Importar as chaves da planilha

```powershell
py -3.13 download_nfce.py importar --xlsm "Autom Download NF V4.xlsm"
```

Lê a coluna A da aba `Base`, da linha 2 em diante; remove máscara, descarta o que não
tiver 44 dígitos (reportando a linha) e elimina duplicatas. Gera `chaves.csv`.

### 2. Reconciliar com o que já está em disco

```powershell
py -3.13 download_nfce.py conferir --destino "D:\NFCe"
```

Faça isso **antes do primeiro `baixar`**, apontando para a pasta histórica: toda chave
que já tem arquivo em disco entra no controle como `baixado` e não é rebaixada. A
varredura é recursiva e casa qualquer nome de arquivo que contenha a chave, com
qualquer extensão.

### 3. Baixar

```powershell
# Primeira rodada: valide o caminho completo com poucas chaves
py -3.13 download_nfce.py baixar --limite 3 --perfil perfil --destino "D:\NFCe"

# Rodada cheia
py -3.13 download_nfce.py baixar --perfil perfil --destino "D:\NFCe"
```

Para o dia a dia, dê duplo-clique em [baixar.bat](baixar.bat): ele roda
`py -3.13 download_nfce.py baixar --perfil perfil` (destino padrão `./pdfs`) sem precisar
abrir terminal nem digitar comando.

Na primeira vez uma janela abre na tela de login: preencha e resolva o captcha. O robô
detecta a sessão e segue sozinho. Nas execuções seguintes, o `--perfil` reaproveita a
sessão e o login é pulado.

Pode interromper com `Ctrl+C` a qualquer momento: o controle é gravado a cada chave, e
a próxima execução retoma de onde parou.

O padrão é baixar o **DANFE** (PDF). Para os XMLs fiscais, use `--formato xml` com um
`--destino` separado — a idempotência é por chave, não por extensão, então misturar os
dois formatos na mesma pasta faz um esconder o outro.

### 4. Acompanhar

```powershell
py -3.13 download_nfce.py status
```

## Comandos

| Comando       | O que faz                                          | Equivalente no VBA            |
| ------------- | -------------------------------------------------- | ----------------------------- |
| `importar`    | Gera `chaves.csv` a partir do `.xlsm`              | (não existia)                 |
| `baixar`      | Percorre a fila e baixa os PDFs                    | Botão "Iniciar Download"      |
| `conferir`    | Reconcilia o controle com os arquivos em disco     | Botão "Conferir no Diretório" |
| `status`      | Contagem por status                                | Ler a coluna B no olho        |
| `inspecionar` | Abre o navegador logado com o Playwright Inspector | (não existia)                 |

Opções de `baixar`:

| Opção              | Padrão         | Efeito                                                                                    |
| ------------------ | -------------- | ----------------------------------------------------------------------------------------- |
| `--chaves`         | `chaves.csv`   | Arquivo de entrada                                                                        |
| `--controle`       | `controle.csv` | Arquivo de estado                                                                         |
| `--destino`        | `./pdfs`       | Pasta onde os PDFs são salvos                                                             |
| `--limite N`       | sem limite     | Processa no máximo N chaves nesta rodada                                                  |
| `--lote N`         | `150`          | Reinicia o navegador a cada N chaves                                                      |
| `--tentativas N`   | `6`            | Tentativas por chave antes de marcar `erro` (a cada 3 erros seguidos, recarrega a página) |
| `--timeout S`      | `15`           | Timeout (segundos) para localizar elementos                                               |
| `--retentar`       | desligado      | Inclui também as chaves `nao_encontrado`                                                  |
| `--formato`        | `danfe`        | `danfe` = PDF; `xml` = arquivo fiscal                                                     |
| `--perfil DIR`     | nenhum         | Perfil do Chromium: preserva a sessão. Use sempre                                         |
| `--espera-login S` | `240`          | Segundos concedidos ao login manual                                                       |
| `--headless`       | desligado      | Sem janela. Inviável neste portal (captcha)                                               |

## Estado (`controle.csv`)

Colunas: `chave`, `status`, `tentativas`, `arquivo`, `erro`, `atualizado_em`.

| Status           | Significado                                              | Entra na próxima rodada? |
| ---------------- | -------------------------------------------------------- | ------------------------ |
| `pendente`       | Nunca tentada                                            | Sim                      |
| `baixado`        | Arquivo salvo e confirmado em disco                      | Não                      |
| `nao_encontrado` | O portal buscou e não retornou nenhuma nota              | Só com `--retentar`      |
| `erro`           | Falha técnica (timeout, elemento ausente, sessão perdida) ou o portal recusou gerar o documento | Sim |

Quando o portal mostra o toast **"Não foi possível gerar a danfe..."** (erro nos dados da
nota, como `O valor da coluna 'xCpl' na tabela 'entrega' é DBNull`), o robô não insiste:
marca `erro` na primeira tentativa, mostra um AVISO, grava a mensagem completa em
`logs/erros.log` e o print em `logs/erros/<chave>.png`, e segue para a próxima chave.
Na próxima execução essas chaves voltam para a fila e são tentadas de novo.

Não existe o `"Extraído"` do VBA: o status só é gravado depois que o resultado da chave
é conhecido, então a coluna sempre diz a verdade sobre o que baixou.

Os arquivos são salvos como `<chave>.<extensão original>` — em geral `<chave>.pdf`. É
esse nome determinístico que permite pular, sem consultar o portal, as chaves que já
estão em disco.

## Quando o portal mudar de layout

Os erros nomeiam o seletor que falhou:

```
Não encontrei 'botao_download' na página. Rode `inspecionar` e ajuste SELETORES.
```

Então:

```powershell
py -3.13 download_nfce.py inspecionar --perfil perfil
```

Isso abre o navegador já logado e o Playwright Inspector. Use **Pick locator** para
apontar o elemento na tela e copiar o localizador, e acrescente-o à lista da ação
correspondente no dicionário `SELETORES`, no topo de
[download_nfce.py](download_nfce.py). Cada ação tem uma lista ordenada de candidatos e o
código usa o primeiro que resolver — dá para adicionar o novo sem remover o antigo.

Prefixos aceitos nos descritores: `papel:<role>:<nome>`, `rotulo:`, `placeholder:`,
`texto:`, `titulo:`. Sem prefixo, vale CSS ou `xpath=...`.

⚠️ **`papel:` casa por substring.** A barra de seleção tem quatro botões cujo nome
acessível começa com `download` (XML, DANFE, Resumo Sem Valor Fiscal, Relatório de
Notas), porque a ligadura do ícone entra no nome. Um `papel:button:Download` casaria
com os quatro e pegaria o primeiro do DOM. Ao adicionar um localizador por papel,
escolha um trecho que só exista no elemento certo.

Falhas técnicas geram screenshot em `logs/erros/<chave>.png`.

## Testes

```powershell
py -3.13 -m pytest tests -q
```

Cobrem as funções puras: CSV de controle, sincronização, montagem da fila, importação
do `.xlsm` e varredura de diretório. As funções de portal não têm cobertura
automatizada — para validá-las use `inspecionar` mais uma rodada com `--limite 3`.
