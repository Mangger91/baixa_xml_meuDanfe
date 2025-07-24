# 🧾 Automação de Download de XML (MeuDanfe)

Automação Python para baixar XMLs de NF-e a partir de uma planilha de chaves de acesso via o site [meudanfe.com.br](https://meudanfe.com.br).

---

## 🚀 Funcionalidades

✅ Lê chaves de uma planilha Excel  
✅ Acessa o site meudanfe.com.br e preenche automaticamente  
✅ Baixa o XML (diretamente ou via botão "Baixar XML")  
✅ Timeout de 3 minutos por chave (reinicia o navegador se travar)  
✅ Atualiza o status na planilha: `SUCESSO`, `FALHA`, `ERRO`  
✅ Gera log `.txt` na pasta `Logs/`

---

## 📁 Estrutura

baixa_xml_meuDanfe/
├── baixar_xmls.py # Script principal
├── Chaves_de_Acesso.xlsx # Planilha com as chaves de acesso
├── XML/ # Pasta onde os XMLs são salvos
├── Logs/ # Logs das execuções
├── README.md # Instruções e explicações
└── .gitignore # Arquivos ignorados no Git


---

## ⚙️ Requisitos

- Python 3.10 ou superior
- Google Chrome instalado
- Instalar dependências:

```bash
pip install undetected-chromedriver selenium openpyxl


▶️ Como executar

1. Preencha a planilha Chaves_de_Acesso.xlsx com as chaves de acesso na coluna A.

2. Execute o script:
python baixar_xmls.py

3. Os XMLs serão salvos na pasta XML/.

4. O status de cada chave será salvo na planilha.

5. O log da execução será gerado em Logs/.


📝 **Observação**: Caso deseje apenas validar os XMLs baixados com os que constam na planilha, execute o script `validador.py`.


✍️ Autor
Roberto Mangger Jr
Supervisor de TI | Automação em Python | Integrações API
🔗 GitHub - @Mangger91

📄 Licença
Este projeto é de uso livre para fins educacionais ou internos em empresas.
Não é afiliado oficialmente ao site MeuDanfe.


---

### ✅ 2. Sugestões extras (se quiser melhorar ainda mais)

- ✅ Adicionar um print ou GIF mostrando o funcionamento (ex: usando [Licecap](https://www.cockos.com/licecap/))
- ✅ Adicionar badge de Python version:

```markdown
![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
