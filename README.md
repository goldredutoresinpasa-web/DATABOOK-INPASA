# Data Book Digital - Gold Redutores

Estrutura profissional para publicar páginas HTML no GitHub Pages.

## Como usar

1. Crie um repositório no GitHub.
2. Envie todos estes arquivos para o repositório.
3. Vá em Settings > Pages.
4. Em Branch, selecione `main` e pasta `/root`.
5. Salve.
6. O GitHub gerará um link parecido com:

https://seuusuario.github.io/nome-do-repositorio/

## Estrutura

- `index.html`: página inicial com a lista de equipamentos.
- `assets/css/style.css`: aparência visual do site.
- `assets/img`: logos e fotos dos equipamentos.
- `equipamentos/ME631010/index.html`: página individual do equipamento.
- `modelo/modelo-equipamento.html`: modelo para copiar e criar novas TAGs.

## Como criar nova TAG

1. Copie a pasta `equipamentos/ME631010`.
2. Renomeie para a nova TAG, exemplo: `ME631011`.
3. Edite o arquivo `index.html` dentro da pasta.
4. Coloque a foto na pasta `assets/img`.
5. Adicione o novo link no `index.html` principal.
