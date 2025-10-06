# SisNUCI - Google Apps Script

Este projeto é um sistema de manipulação de dados em planilhas Google, com interface web customizada, controle de acesso por e-mail e funcionalidades de dashboard, edição, paginação e filtragem.

## Estrutura do Projeto


## Ambientes

- `src/` - Ambiente de produção
   - Contém o código que está em uso no sistema oficial do Google Apps Script (GAS).
- `src-homologacao/` - Ambiente de homologação
   - Utilizado para testes, validações e ajustes antes de publicar no ambiente de produção.

Cada ambiente possui seus próprios arquivos:
   - `Código.js` - Script principal
   - `index.html` - Interface web
   - `appsscript.json` - Configuração do projeto Apps Script
   - `.clasp.json` - Configuração do Clasp (apenas se necessário)

## Estrutura do Projeto

- `src/` - Código-fonte do Apps Script (produção)
   - `Código.js` - Script principal
   - `index.html` - Interface web
- `src-homologacao/` - Código-fonte do Apps Script (homologação)
   - `Código.js` - Script principal
   - `index.html` - Interface web
- `appsscript.json` - Configuração do projeto Apps Script
- `.clasp.json` - Configuração do Clasp

## Como usar

1. Instale o [Node.js](https://nodejs.org/) e o [clasp](https://github.com/google/clasp):
   ```
   npm install -g @google/clasp
   ```
2. Faça login com sua conta Google:
   ```
   clasp login
   ```
3. Clone ou crie o projeto:
   ```
   clasp clone <ID_DO_PROJETO>
   ```
4. Edite os arquivos em `src/` normalmente.
5. Envie as alterações para o Google Apps Script:
   ```
   clasp push
   ```
6. Sincronize alterações feitas na web:
   ```
   clasp pull
   ```

## Licença

MIT
