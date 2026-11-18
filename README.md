# PCP pessoal

Aplicativo independente em Python e Streamlit para organizar ordens, materiais, produção e qualidade. Não utiliza login, API ou hospedagem do ChatGPT.

## Publicar

No Streamlit Community Cloud, selecione a branch `main` e o arquivo `streamlit_app.py`. Use Python 3.11 ou superior. Essa entrada exige login configurado, chave privada de cookie e conexão com Google Sheets antes de permitir o uso.

O arquivo `.streamlit/secrets.example.toml` mostra a estrutura da configuração. Preencha os valores reais somente nos Secrets privados da hospedagem. A versão atual do código obtém o login, o hash da senha e a URL da base operacional somente dos Secrets privados. Configure credenciais novas e exclusivas para a hospedagem.

O acesso lembrado usa cookie assinado por até 365 dias. Sair, limpar cookies ou trocar a senha encerra esse acesso. A chave `auth_cookie_key` precisa permanecer estável entre deploys. As OPs e os saldos são persistidos no Google Sheets; simulações são recalculadas. O registro de eventos exibido na Qualidade pertence à sessão.

As quatro abas são `Estoque_Atual`, `Fichas_Tecnicas_BOM`, `Ordens_Producao` e `Capacidade_Linhas`. Os dados incorporados ao código são o conjunto de demonstração; não substituem a conexão persistente na versão online. Compartilhe a planilha real como Editor somente com a conta de serviço usada pelo aplicativo.

## Executar e verificar

```text
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
python -m unittest discover -s tests -v
```

Há cinco módulos, com uma visão inicial Meu dia, MRP, comparação de sequências, consulta ATP e liberação de lotes. O aplicativo foi projetado para um usuário e uma instância. O acesso de visitantes na hospedagem precisa permitir chegar à tela própria de login; os dados e comandos permanecem protegidos no servidor.
