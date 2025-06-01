# Projeto: Automação Web

Este projeto realiza a automação de buscas de ofertas de produtos em sites de comparação de preços, como o Buscapé, utilizando Python, Selenium e Pandas. O objetivo é facilitar a pesquisa de produtos dentro de uma faixa de preço desejada, filtrando termos indesejados e exportando os resultados para uma planilha Excel. Caso ofertas sejam encontradas, o sistema envia automaticamente um e-mail com a tabela de ofertas.

## Como funciona

1. **Leitura dos Produtos:**  
   O sistema lê uma lista de produtos e parâmetros de busca a partir de um arquivo Excel (`buscas.xlsx`).

2. **Automação Web:**  
   Utiliza o Selenium para acessar o site Buscapé, pesquisar os produtos, filtrar resultados por termos banidos e faixa de preço, e coletar nome, preço e link das ofertas.

3. **Exportação dos Resultados:**  
   As ofertas encontradas são salvas em um arquivo Excel (`Ofertas.xlsx`).

4. **Envio de E-mail:**  
   Caso haja ofertas, um e-mail é enviado automaticamente com a tabela de ofertas no corpo do e-mail.

## Observações Importantes

- **Limitações de Automação:**  
  Este projeto foi desenvolvido e testado para funcionar com o site Buscapé.  
  Para outros sites como Google Shopping, Amazon, Magazine Luiza, Americanas, entre outros, o processo de automação pode não funcionar corretamente.  
  Esses sites costumam ter proteções contra automação (como CAPTCHAs, bloqueios de bot, carregamento dinâmico, etc.), o que pode impedir ou dificultar a coleta automática de dados via Selenium.

- **Uso de APIs:**  
  Para realizar buscas automatizadas nesses sites de forma confiável e ética, é recomendado utilizar APIs oficiais (quando disponíveis).  
  O uso de automação web em sites que não permitem pode violar os termos de uso dessas plataformas.

## Requisitos

- Python 3.x
- Bibliotecas: `pandas`, `selenium`, `webdriver_manager`, `pywin32`
- Microsoft Outlook instalado (para envio de e-mails)

## Como usar

1. Instale as dependências:
2. Preencha o arquivo `buscas.xlsx` com os produtos e parâmetros desejados.
3. Execute o script principal (`Automacao Web.ipynb`).
4. Verifique o arquivo `Ofertas.xlsx` e o e-mail enviado.

---

**Atenção:**  
Automatizar buscas em sites sem permissão pode violar os termos de uso. Sempre verifique as políticas dos sites antes de realizar automações.