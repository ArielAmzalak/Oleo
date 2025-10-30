# Oleo

Aplicativo Streamlit para registrar coletas de amostras de óleo da Oliveira Energia.

## Funcionalidades principais

- O campo **n.º da Amostra** é o primeiro da ficha e dispara a busca automática na planilha Google assim que um número é digitado.
- Se a amostra já existir, todos os campos do formulário são preenchidos com os dados cadastrados anteriormente, permitindo revisar e corrigir facilmente.
- Ao enviar, o registro é salvo na mesma linha da planilha (preservando os campos de status), gerando também o PDF atualizado.
- Caso o número ainda não exista, uma nova linha é criada com as informações informadas.

## Execução local

Instale as dependências e execute o aplicativo com Streamlit:

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Certifique-se de configurar as credenciais do Google Sheets por meio de `token.json` ou da variável/segredo `GOOGLE_CLIENT_SECRET`.
