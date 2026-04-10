import { defineConfig } from 'vite';

export default defineConfig({
  // Garante que os caminhos para o CSS e JS funcionem no GitHub Pages,
  // não importando o nome do repositório
  base: './',
});
