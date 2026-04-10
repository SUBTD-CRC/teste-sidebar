# 1746 Gerenciador de Sidebar

Esta é uma aplicação construída para facilitar a visualização dos Temas, Subtemas e Serviços da interface da Sidebar do Portal 1746. 

A ferramenta permite arrastar e soltar (drag and drop) categorizações, fazer edições em tempo real para os nomes e descrições ao cidadão e, ao final, exportar as definições diretas para uma planilha automatizada em formato **Excel**.

## ✨ Funcionalidades

- **Visualização em 3 Níveis (Tema > Subtema > Serviço):** Utiliza um layout de colunas independentes inspirado na interface atual, eliminando listas infinitas.
- **Rearranjo Hierárquico Inteligente:** Reordene a sequência dos itens visualmente utilizando "arrastar e soltar". Mova Temas, troque as ordens dos Serviços e altere Subtemas à vontade.
- **Painel de Propriedades (Inspector):** Adicione ou edite nomes e textos de "Descrição para o Cidadão" facilmente com um design focado na utilidade e legibilidade.
- **Persistência Local (LocalStorage):** As alterações nunca são perdidas se você não fechar intencionalmente, tudo sendo armazenado na sessão do navegador.
- **Importação/Exportação para Excel (.xlsx):** Recurso nativo que traduz instantaneamente toda a estrutura de 3 dimensões em uma estrutura de planilha e vice-versa, com total compatibilidade no seu pacote Office.

## 🚀 Tecnologias Integradas

- **Front-End:** HTML5 Avançado, CSS3 Puro, JavaScript ES6 (Sem Frameworks Reativos pesados)
- **Bibliotecas Usadas:** 
  - `sortablejs` (Para comportamento fluido de drag and drop moderno).
  - `xlsx` (SheetJS) para a integração 100% nativa no navegador sem uso backend de Planilhas Excel.

---

## 💻 Como Rodar o Projeto Localmente

Se você deseja fazer alterações no CSS ou no comportamento JavaScript, primeiro você precisa instalar as dependências de desenvolvedor:

1. Certifique-se de ter o [Node.js](https://nodejs.org/en/) e NPM instalados no seu computador.
2. Navegue até o diretório do projeto:
    ```bash
    cd sidebar-app
    ```
3. Instale as bibliotecas necessárias:
    ```bash
    npm install
    ```
4. Inicie o servidor para carregar a ferramenta (isso exibirá no `localhost:5173` no seu computador e qualquer modificação ativará um auto-recarregamento):
    ```bash
    npm run dev
    ```

---

## 📂 Visão Rápida da Estrutura

- `index.html`: Shell e base em HTML para o app, marcação limpa.
- `style.css`: Totalidade das variáveis cores, modo noturno/misto, componentes estilizados de layout.
- `main.js`: Lógica funcional principal de eventos, CRUD, e mapeamento das dependências.
- `initial_data.json`: Payload com definições em JSON puro iniciais exportados do `.csv` bruto da Prefeitura - para iniciar um novo quadro limpo.
- `vite.config.js`: Regras de Base-URL exclusivas do Github para permitir portabilidade dos arquivos de scripts e estilos.
