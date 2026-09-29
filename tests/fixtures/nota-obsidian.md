---
tipo: doc
area: testes
tags: [redator, markdown, obsidian]
aliases:
  - Nota de teste
---

# Nota de teste do Redator

Parágrafo com **negrito**, *itálico*, __negrito com sublinhado__, _itálico com sublinhado_,
~~tachado~~ e `código em linha`. A linha foi quebrada no meio de propósito,
como muita gente faz em arquivos de documentação.

Links: [site](https://example.com "título"), <https://example.org>, https://example.net solto,
[âncora](#segunda-seção) e referência [ref][r1].

[r1]: https://example.com/ref

Obsidian: [[Outra nota]], [[Pasta/Nota com ponto.js|apelido]] e ![[anexos/figura.png]].

> [!info] Título do callout
> Corpo do callout com `código` e [[link]].
>
> Segundo parágrafo do callout.

> [!warning]-
> Callout recolhido sem título.

> [!NOTA] Tipo em português
> Texto.

> Citação comum
> em duas linhas.
>
> > Citação dentro de citação.

## Segunda seção

* item com asterisco
* outro item
  * subitem com dois espaços
  * outro subitem

1. primeiro
2. segundo
   - misturado
3. terceiro

1) parêntese
2) parêntese dois

- [ ] tarefa aberta
- [x] tarefa feita

Setext também
-------------

```js
function soma(a, b) {
	return a + b; // tabulação no começo
}
```

~~~
cerca com til
~~~

    código indentado
    com quatro espaços

| Coluna | Alinhada |   Direita |
|:-------|:--------:|----------:|
| a      |    b     |         c |
| pipe \| escapado | `x\|y` | [[nota\|alias]] |

<!-- comentário que precisa sobreviver -->

<details>
<summary>Detalhes</summary>

Conteúdo escondido.

</details>

Genéricos como Result<T> e Array<Objeto> ficam como estão, e <br> quebra linha.
Tecla <kbd>Ctrl</kbd> e nota de rodapé[^1].

[^1]: Texto da nota.

***

Escapes: \*não é itálico\*, 1\. não é lista, \# não é título, snake_case_nome.
Linha com quebra forçada no fim  
e continuação.
