---
tipo: projeto
status: em andamento
tags: [redator, documentação]
atualizado: 2026-09-28
---

# Plano de lançamento

Nota de um cofre do Obsidian aberta no Redator. Os links [[Roteiro]] e
[[Equipe|quem participa]] abrem a nota de destino com Ctrl+clique, e as caixas
coloridas abaixo são callouts.

> [!tip] Dica
> Salvar sem editar não muda nenhum byte do arquivo. Editar um parágrafo
> reescreve só as linhas dele.

> [!warning] Atenção
> A versão 1.1 depende da leitura de arquivos `.doc` antigos.

## Tarefas

- [x] Instalador com Inno Setup
- [x] Build automático a cada push na `main`
- [ ] Leitura de `.doc` e `.rtf`

## Cronograma

| Etapa | Data | Situação |
|---|---|---|
| Testes | 26/09 | concluído |
| Lançamento | 30/09 | em andamento |

```js
const nota = redator.abrir('Nota de exemplo.md');
nota.salvar(); // mesmos bytes do original
```
