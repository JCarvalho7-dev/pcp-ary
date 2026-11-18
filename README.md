# PCP & Telemetria de Produção

Aplicativo Python/Streamlit completo, sem logotipo ou nome de empresa na interface. O design usa off-white mineral, grafite e terracota. As quatro tabelas e todos os dados originais estão incorporados em `app.py`; nenhum arquivo de dados é necessário para o sandbox.

## Executar

Requer Python 3.11 ou superior. Dentro desta pasta:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

No Linux/macOS, use `.venv/bin/python` nos dois últimos comandos. Abra o endereço indicado pelo Streamlit, normalmente `http://localhost:8501`.

O acesso pessoal usa o login Aryane e a senha definida no pedido, verificada por bcrypt. A opção **Manter conectado neste navegador** vem marcada e emite um cookie assinado válido por até 365 dias. Fechar a aba ou o navegador preserva o acesso; **Sair deste navegador**, limpar os cookies ou alterar a senha encerra o acesso lembrado. Cada navegador e perfil mantém seu próprio acesso. Não é necessário criar conta de usuário no aplicativo, nem usar ChatGPT para acessá-lo.

O sandbox local inicia quando a conexão Google não está configurada e `production_mode` está desativado. Todos os cálculos e alterações do sandbox funcionam em memória, por sessão de navegador. Atualizar os dados não apaga essa sessão. Encerrá-la ou reiniciar o serviço pode apagar as alterações locais. Na hospedagem, use `production_mode = true`: o app exige Google Sheets conectado e uma chave privada de cookie, e não permite trabalhar em memória por engano.

## Módulos

| Módulo | O que entrega |
| --- | --- |
| Torre de Controle | Quatro KPIs, Kanban, cadastro de OP, avanço de status, inventário, exportação e central de automações |
| Motor MRP | Explosão por receita, perdas, demanda agregada, déficits, compra com reserva mínima e requisições XLSX/CSV |
| Sequenciador | Cenário cronológico, agrupamento por família, contagem de trocas, horas teóricas e exportação da sequência |
| Simulador ATP | Disponibilidade imediata ou prazo pela fila compatível, turno, dias úteis e alerta de insuficiência de materiais |
| Qualidade | Lotes aguardando laudo, aprovação, bloqueio, transferência de saldo, conciliação e registro da sessão |

A aba **Automações & exceções** da Torre reúne atrasos, entregas do dia, insuficiência de materiais, laudos pendentes, divergências de quarentena e risco de prazo por linha. A reposição considera OPs já em fila para reduzir a criação duplicada de produção. O painel se reavalia a cada 60 segundos enquanto a sessão está ativa. Isso não é um agendador de tarefas em segundo plano.

## Google Sheets

1. Baixe **Modelo de integração · 4 abas** pela barra lateral ou use `modelo-quatro-abas.xlsx` incluído no pacote.
2. Importe o arquivo no Google Sheets. Preserve nomes, ordem e grafia dos cabeçalhos. Configure `Data_Entrega` como texto ISO `AAAA-MM-DD`; mantenha números como valores numéricos.
3. Crie uma conta de serviço no Google Cloud, habilite a **Google Sheets API**, gere sua chave e compartilhe somente a planilha desejada com o e-mail da conta como **Editor**.
4. Preencha o exemplo completo em `.streamlit/secrets.toml`. A URL da base criada já está preenchida. Todos os parâmetros de login, cookie e hospedagem devem aparecer antes de `[gcp_service_account]`.
5. Reinicie o app. A barra superior deve mostrar **GOOGLE SHEETS SINCRONIZADO**.

As únicas abas operacionais são `Estoque_Atual`, `Fichas_Tecnicas_BOM`, `Ordens_Producao` e `Capacidade_Linhas`. O app não cria nem apaga abas da planilha e não sobrescreve sua base para inicializar a conexão. Corrija o modelo na origem se houver falha de validação.

As leituras usam `@st.cache_data(ttl=60)`. Toda escrita invalida o cache imediatamente, inclusive quando há falha de confirmação. Aprovar um lote envia status e dois saldos em uma única chamada atômica `spreadsheets.batchUpdate`. Antes de escrever, o aplicativo relê a base diretamente. Linhas em branco na planilha não deslocam os endereços de gravação.

O bloqueio por processo evita concorrência entre sessões de uma mesma instância. Google Sheets não oferece aqui uma comparação e troca transacional entre a leitura e a escrita. Use **uma instância, um processo de aplicação**, evite editar as mesmas células simultaneamente por fora do app e migre para banco transacional antes de múltiplas réplicas ou grande volume de apontamentos. O teste da integração utiliza um cliente simulado: nenhuma conta Google real foi conectada nesta entrega.

## Regras e limites operacionais

- Novas OPs começam como `Programada`. Fluxo: `Programada → Em_Processo → Aguardando_Laudo → Liberada → Expedida`. É possível bloquear uma ordem antes da expedição. Reabertura e correções excepcionais são feitas na origem após conciliação; o app impede retrocessos que poderiam creditar um lote duas vezes.
- A aprovação é permitida somente quando a OP está `Aguardando_Laudo`. Transfere `Qtd_Ton` para o disponível, limita a quarentena a zero e altera o status para `Liberada`. Repetir a aprovação da mesma OP não credita estoque novamente.
- Conforme a regra do pedido, se a quantidade de um lote exceder a quarentena cadastrada, a aprovação credita o lote integral. O sistema alerta antes e registra a divergência. Essa regra deve ser revisada ao conectar dados reais.
- O estoque é agregado por SKU. Não há saldo separado por lote, baixa automática de insumo, apontamento de produção, entrada em quarentena ou baixa de expedição. Essas movimentações precisam ser registradas na origem. Avançar o status da OP não simula movimentações físicas adicionais.
- MRP: `Qtd_Ton × Qtd_Por_Ton × (1 + Perda_Tecnica_Pct/100)`. Considera OPs programadas e, opcionalmente, em processo. BOM de nível único, sem reservas de vendas e sem pedidos de compra em trânsito.
- O cálculo mantém a fração técnica de embalagens. A compra sugerida arredonda unidades para cima. Toneladas, kg e unidades não são somadas entre si e aparecem em escalas separadas nos gráficos.
- ATP: usa primeiro o saldo disponível e produz apenas a diferença. Prazo: `(fila da família + quantidade líquida nova)/capacidade + setup`, dividido por 16 h/dia e arredondado para cima. Segunda a sexta, data conservadora D+N dias úteis, sem feriados, manutenção, frete ou laboratório. Aos fins de semana, o planejamento começa na próxima segunda e aplica D+N. Não reserva saldo nem garante atendimento comercial. A fila de produção considera a quantidade integral da OP, pois não existe apontamento parcial.
- Sequenciamento: a primeira OP não conta como troca; cada transição recebe o tempo de setup da família de destino. Com uma linha dedicada por família nos dados iniciais, a economia global é **teórica**, não uma economia confirmada de horas da planta. A programação por equipamento é calculada separadamente na Torre.
- Os registros exibidos na Qualidade são apenas da sessão. Não há trilha permanente por usuário, anexo de laudo, assinatura eletrônica ou validação laboratorial automatizada.
- O login foi simplificado para uma única pessoa. O acesso lembrado usa cookie assinado; a senha não é gravada em texto no navegador. Em produção, `auth_cookie_key` deve ser privada e estável entre deploys.
- A telemetria atual é derivada das OPs e dos saldos. Não há conexão com sensores, CLPs, OPC UA ou MQTT.

## Publicar

**Opção selecionada: Streamlit Community Cloud + Google Sheets.** É uma hospedagem independente do ChatGPT; o aplicativo não faz chamadas à OpenAI. O endereço do app será um subdomínio `streamlit.app`, confirmado após a publicação. Não houve criação ou publicação de Site na hospedagem integrada.

Coloque o projeto em um repositório privado do GitHub. Acesse Streamlit Community Cloud, selecione **`streamlit_app.py`** como arquivo principal e Python 3.11 ou superior. Essa entrada força o modo de produção e impede gravações em memória. O acesso de visitantes na plataforma deve permitir chegar à tela de login do aplicativo sem exigir conta Google/Streamlit; os dados e as funções do sistema continuam protegidos pelo login pessoal. Cadastre os valores em **Settings → Secrets**, mantendo `production_mode = true`, `authentication_required = true`, uma `auth_cookie_key` aleatória e estável, a URL da planilha e o bloco da conta de serviço. O `.gitignore` ignora o arquivo local de secrets. A conta GitHub e a configuração inicial do Google pertencem à pessoa que administra a hospedagem; o uso diário exige somente o login do aplicativo.

Base já criada no Google Drive: https://docs.google.com/spreadsheets/d/12GKooqAJfh4OjPjFITPJ_fmJedkQhwdsVXJzSQt6bP0/edit

Ainda é necessário autorizar a publicação na conta GitHub/Streamlit e configurar uma conta de serviço Google para o aplicativo hospedado. A conexão usada aqui para preparar a planilha não transfere automaticamente credenciais ao servidor. Consulte `PUBLICACAO-EXTERNA.txt` para a configuração inicial.

**Servidor próprio:** instale `requirements.txt` em um ambiente virtual e execute `python -m streamlit run app.py --server.address=0.0.0.0 --server.port=8501` atrás de proxy HTTPS, com o acesso definido pela equipe. Também há um `Dockerfile` opcional. Não houve publicação em nuvem nesta entrega; o app é entregue pronto para configuração e deploy.

## Verificar

```powershell
python -m unittest discover -s tests -v
```

Os testes verificam dados iniciais, fórmulas MRP, ausência de BOM, fila vazia, ATP, fins de semana, setup, aprovação idempotente, divergência de quarentena, reposição, exportações, gravação atômica simulada, invalidação do cache, linhas vazias e navegação/login/formulários reais por `streamlit.testing.v1.AppTest`.

Valores de referência para **18/11/2026**: 10 OPs ativas, 159 ton em fila, 2 lotes/32 ton aguardando laudo e 5 itens abaixo do mínimo. MRP somente Programada: hidrato 75,5812 ton; silício 32,136 ton; agente branqueador 243,6 kg; sacos 1.827 un. Comparação de sequências: 23,5 h contra 14,5 h, ganho teórico de 9 h.

## Pesquisa de automação

Veja `PESQUISA-AUTOMACAO.txt` para as fontes oficiais consultadas e um roteiro priorizado para reduzir tarefas manuais sem prometer integrações ainda não conectadas.
