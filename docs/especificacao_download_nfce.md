# Especificação funcional — Automação de Download de NFC-e (para reescrita em Python)

Documento extraído do VBA original de `Autom Download NF V4.xlsm`, descrevendo todas as regras de negócio, fluxo e detalhes técnicos, para servir de base à reimplementação em Python.

## 1. Objetivo

Planilha Excel com VBA que automatiza, via navegador Chrome controlado por robô (SeleniumBasic), o download de arquivos de Notas Fiscais de Consumidor Eletrônica (NFC-e) a partir do portal **Tec Fiscal / Linx** (`https://tecfiscal.linx.com.br`), para uma lista de chaves de acesso previamente carregada em planilha, e depois cruza os arquivos baixados com o que já existe em uma pasta local para saber o que falta baixar.

## 2. Estrutura de dados (planilhas)

### 2.1 Aba "Base" — lista de trabalho

| Coluna | Conteúdo |
|---|---|
| A (`CHAVE_NFCE`) | Chave de acesso da NFC-e (44 dígitos), uma por linha, a partir da linha 2 |
| B (`Situação Download`) | Status do processamento daquela linha. Valores usados pelo robô: vazio (ainda não processado), `"Extraído"` (chave já digitada no filtro do site, processo em andamento), `"Realizado"` (download concluído com sucesso) |
| C (`Diretório`) | Preenchida pela função "Conferir no Diretório" com fórmula que verifica se já existe um PDF com o nome da chave na pasta local (`"OK"` se existir, vazio se não) |
| I1 | Rótulo fixo `"CNPF:"` |
| J1 | CNPJ da empresa/filial em uso (texto livre, editável pelo usuário, ex: `89.237.911/0031-66`) |
| Célula (linha 1, coluna 200) | Usuário salvo do login (se a opção "Salvar Usuário" estiver marcada) |
| Célula (linha 2, coluna 200) | Senha salva do login (se a opção "Salvar Senha" estiver marcada) — **gravada em texto puro**, com um apóstrofo `'` na frente para forçar formatação de texto |

A última linha usada é sempre calculada por `Range("A100000").End(xlUp).Row` (equivalente a "Ctrl+Seta para cima" a partir da linha 100000 da coluna A).

### 2.2 Aba "Arq_Diretório" — inventário de arquivos já baixados

Gerada do zero toda vez que o usuário aciona "Conferir no Diretório":

- **A1**: texto `"Arquivos do Diretório: <caminho escolhido>"`, negrito, tamanho 12.
- **A3:E3**: cabeçalho — `Caminho:`, `Nome:`, `Data Criação:`, `Data último Acesso:`, `Data última Modificação:` (negrito, centralizado, quebra de linha).
- A partir da linha 4: uma linha por arquivo encontrado (recursivamente, incluindo subpastas), com caminho da pasta, nome do arquivo, data de criação, data de último acesso e data de modificação.

## 3. Fluxo funcional

### 3.1 Botão "Iniciar Download" (planilha Base)

1. Limpa a coluna C (`Diretório`) e escreve o cabeçalho `"Diretório"` em C1.
2. Abre o formulário de login (`login.Show`).

### 3.2 Formulário de login

Campos: `Usuário` (texto), `Senha` (texto, mascarado), checkbox `Salvar Usuário`, checkbox `Salvar Senha`, botão `Conectar`.

**Ao abrir o formulário:**
- Se já existir usuário salvo na célula (1,200), preenche o campo e marca "Salvar Usuário".
- Se já existir senha salva na célula (2,200), preenche o campo e marca "Salvar Senha".

**Ao clicar em "Conectar":**
1. Desabilita o botão (evita duplo clique).
2. Grava usuário/senha nas células (1,200)/(2,200) da aba Base **somente se** a respectiva caixa "Salvar" estiver marcada; senão, limpa a célula.
3. Abre uma instância do Chrome e navega para `https://tecfiscal.linx.com.br/#/login`.
4. Exibe um `MsgBox` pedindo **"Faça login"** — ou seja, **o preenchimento de usuário/senha e o clique em entrar no site são feitos manualmente pelo usuário** (o código de login automático via XPath está todo comentado/desativado no arquivo original). O robô só continua depois que o usuário clica OK nessa caixa de mensagem.
5. Aguarda 1 segundo.
6. Recalcula a última linha usada na aba Base e entra em loop **para cada linha de 2 até a última**, processando apenas as linhas em que a coluna B (Situação Download) ainda está vazia. Para cada uma:
   1. `abre_filtro` — clica no botão que abre o painel de filtros da tela de emissão de notas.
   2. Aguarda 1s.
   3. `insere_chave` — digita a chave de acesso (coluna A da linha) no campo de busca por texto do filtro, e marca a linha como `"Extraído"` na coluna B.
   4. Aguarda 1s.
   5. `Fecha_filtro` — clica no botão que fecha/aplica o painel de filtro (dispara a busca).
   6. Aguarda 4s (tempo para a busca carregar o resultado).
   7. `Seleciona_nota` — marca o checkbox da primeira linha do resultado da tabela.
   8. Aguarda 2s.
   9. Se nada falhou até aqui (`gv_pula` continua vazio): `baixa_nota` — clica no botão de download da nota selecionada.
   10. Aguarda 5s (tempo para o navegador processar/baixar o arquivo).
   11. Se nada falhou: `Fecha_tela` — fecha o modal/tela que ficou aberta após o download.
   12. Aguarda 1s.
   13. Passa para a próxima linha.

**Rotinas auxiliares chamadas no loop acima**, todas seguindo o mesmo padrão "tenta clicar; se der erro, espera 10s e tenta de novo uma única vez; se continuar falhando, marca `gv_pula = "X"` e pula o restante das etapas daquela linha (sem marcar como "Realizado")":

- `abre_filtro`: clica no botão de abrir filtro.
- `insere_chave`: limpa o campo de texto do filtro e digita a chave; marca a linha como "Extraído".
- `Fecha_filtro`: clica no botão de aplicar/fechar filtro.
- `Seleciona_nota`: clica no checkbox de seleção da nota encontrada.
- `baixa_nota`: clica no botão de download.
- `Fecha_tela`: fecha a tela/modal residual.

Importante: **neste fluxo (o que é realmente usado pela interface), a linha só é marcada como "Extraído" — nunca como "Realizado"**. (A marcação como "Realizado" só acontece no módulo alternativo `base_busca`, que está descrito abaixo mas não é chamado por nenhum botão nem evento — é código morto/versão anterior mantida no arquivo.)

Ao fechar o formulário (evento `Terminate`), a aplicação executa `End` (encerra a execução de todas as macros em andamento).

### 3.3 Botão "Conferir no Diretório" (planilha Base)

1. Vai para a aba "Arq_Diretório", limpa tudo.
2. Abre um seletor de pasta do Windows (`Shell.Application.BrowseForFolder`) para o usuário escolher a pasta onde os PDFs são salvos.
3. Lista recursivamente todos os arquivos da pasta escolhida (e subpastas) na aba "Arq_Diretório", com caminho, nome e datas.
4. Volta para a aba "Base" e, para cada linha usada, escreve na coluna C a fórmula:
   `=SE(ÉERRO(PROCV(A2&".pdf"; Arq_Diretório!B:B; 1; FALSO)); ""; "OK")`
   (ou seja, `"OK"` se existir um arquivo `<chave>.pdf` na pasta escolhida, vazio caso contrário).
5. Converte essa coluna de fórmula em valores fixos (copiar → colar especial → valores).

Isso serve como uma **checagem cruzada manual**: depois de rodar o download, o usuário confere quais chaves já têm PDF salvo em disco, independentemente do que a coluna B diz.

### 3.4 Módulo `base_busca` (não usado pela interface atual — mantido só como referência de uma versão anterior)

Contém uma rotina `Processo` que percorre a mesma lista de chaves e, para cada uma sem status na coluna B, chama `process_donload`, que:

1. Navega direto para `https://tecfiscal.linx.com.br/app/emissao/notas` (sem passar pela tela de login manual).
2. Verifica se já existe algum item selecionado na tabela e desmarca.
3. Abre o filtro, digita a chave, seleciona a filial pelo CNPJ (célula J1) se necessário, aplica o filtro (`clica_pesquisa`), marca o item retornado, e baixa (`clica_download`).
4. Se não der erro, marca a linha como `"Realizado"` e atualiza um contador de progresso (`login.status`).
5. **Continha o bug corrigido nesta conversa**: a cada suposto lote de "100" registros processados sem sucesso prévio, deveria reiniciar a sessão do navegador (`bot.Quit` + `Refresh`) para liberar memória — mas a condição no código verificava `= 10000` em vez de `100`, então a sessão nunca era reciclada em tempo hábil, causando o erro "Memória insuficiente" do VBA após processar muitos registros.

Recomendação: **não é necessário portar `base_busca` para o Python**, já que não é o caminho usado hoje. Mas vale considerar a ideia de "renovar a sessão do navegador periodicamente" como boa prática na nova aplicação, mesmo sem esse bug specific do VBA existir em Python.

## 4. Timers e tolerância a falha

- Todas as esperas no VBA são **tempos fixos** (`Application.Wait`-like, via `Timer`+`DoEvents`), não esperas inteligentes por elemento pronto. Valores usados: 0.5s, 1s, 2s, 3s, 4s, 5s e 10s (no retry), dependendo da etapa.
- Padrão de tratamento de erro em quase todas as rotinas de clique: tenta uma vez; se falhar, espera mais um tempo fixo (geralmente 10s) e tenta de novo; se falhar de novo, desiste daquela linha (`gv_pula = "X"`) sem marcar sucesso, mas segue para a próxima linha da planilha.
- **Recomendação para a versão Python**: trocar os tempos fixos por esperas explícitas (`WebDriverWait` + `expected_conditions`) até o elemento existir/estar clicável, o que tende a ser mais rápido e mais robusto que sleeps fixos.

## 5. Identificação dos elementos na página (XPaths usados no site Tec Fiscal / Linx)

Estes XPaths são absolutos (posicionais) e portanto **frágeis** — qualquer mudança no layout do site os quebra. Ainda assim, seguem como referência do que cada clique faz, no fluxo do formulário de login (fluxo ativo):

| Ação | XPath usado |
|---|---|
| Abrir painel de filtro | `.../lx-page-filter/div/div[1]/button` |
| Campo de texto do filtro (chave de acesso) | `.../lx-page-filter/div/div[2]/form/app-lx-filter-text[2]/div/div[2]/input` |
| Fechar/aplicar filtro | `.../lx-page-filter/div/div[1]/div/div/div[2]/button[2]/span[2]/span` |
| Checkbox de seleção da nota no resultado | `.../linx-table/div/nz-table/nz-spin/div/div/nz-table-inner-default/div/table/tbody/tr/td[1]/label/span[1]/input` |
| Botão de download | `.../linx-table-toolbar/div/div[2]/div/div[2]/nz-button-group/button[2]/span[2]/span` |
| Botão de fechar modal pós-download | `/html/body/div/div[2]/div/div/div[2]/div/div/div[1]/div/button/span` |

URLs envolvidas:
- Login: `https://tecfiscal.linx.com.br/#/login`
- Tela de emissão/consulta de notas: `https://tecfiscal.linx.com.br/app/emissao/notas` (e variante com `#/`)

**Recomendação**: na reescrita em Python, priorizar seletores mais estáveis (texto do botão, `data-testid`, `aria-label`, classe CSS específica) sempre que possível, em vez de XPath posicional completo — mas como o site é de terceiros (Linx), pode ser necessário inspecionar a página atual para confirmar se esses caminhos ainda são válidos.

## 6. Pontos de atenção para a nova aplicação Python

1. **Login manual vs. automático**: hoje o login é manual (o robô só abre o Chrome e espera o usuário clicar "OK" depois de logar). Vale decidir com o usuário se a nova versão deve automatizar esse passo também (usuário/senha já existem salvos, então é viável) ou manter manual (mais seguro caso haja captcha/2FA).
2. **Credenciais em texto puro**: o arquivo original grava a senha em texto puro numa célula da planilha. Isso é uma prática de segurança ruim. Na versão Python, recomenda-se usar variável de ambiente, arquivo `.env` (fora do controle de versão) ou um cofre de segredos, nunca gravar a senha em uma planilha ou arquivo de dados em texto puro.
3. **Renovação de sessão do navegador**: implementar reinício periódico do WebDriver a cada N notas processadas (ex.: a cada 100–300) para evitar acúmulo de memória em execuções longas (milhares de chaves).
4. **Idempotência**: o controle de "já processado" é feito pela coluna B da planilha (`vazio` / `"Extraído"` / `"Realizado"`). Reaproveitar essa lógica de estado (pode ser uma coluna de planilha, um banco SQLite, ou um CSV de controle) para permitir retomar de onde parou em caso de interrupção.
5. **Conferência com o diretório local**: manter a função de comparar a lista de chaves com os arquivos já existentes na pasta de destino (por nome do arquivo = chave + extensão), para saber o que falta baixar de fato, independente do status da planilha.
6. **Volume**: a base atual tem cerca de 2.300 chaves na aba "Base" e mais de 4.200 arquivos já inventariados na aba "Arq_Diretório" — a solução deve suportar esse volume confortavelmente (o problema original de "memória insuficiente" era decorrência direta da falta de reciclagem de sessão nesse volume).
7. **Formato de entrada/saída**: manter compatibilidade com os arquivos `.xlsx`/`.xlsm` existentes (ex.: usando `openpyxl` ou `pandas`) para não obrigar o usuário a mudar de ferramenta para consultar a lista de chaves e o status.

## 7. Resumo do bug já corrigido no VBA (contexto, não é necessário no Python)

No módulo legado `base_busca`, a condição de reciclagem de sessão comparava o contador com `10000` em vez de `100` (conforme o próprio comentário do código indicava a intenção original), fazendo a sessão do Chrome nunca ser reciclada a tempo, o que estourava a memória do processo do Excel/VBA (erro "Memória insuficiente"). Isso já foi corrigido diretamente no arquivo `.xlsm` original. Esse bug é mencionado aqui apenas como contexto histórico — não existe no fluxo ativo (`login.CommandButton1_Click`) e não precisa ser replicado nem corrigido no Python, já que a arquitetura nova deve nascer com reciclagem de sessão bem definida (ver item 3 acima).
