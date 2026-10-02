**Plano de refinamento do Redator**

O objetivo é tornar o Redator mais confiável para documentos reais, mais fácil de usar e mais simples de manter. O trabalho cobre salvamento e recuperação, fidelidade de formatos, paginação, desempenho, interface, acessibilidade, segurança, testes e distribuição. A execução começa pelos riscos de perda de trabalho e segue para compatibilidade e acabamento.

A auditoria inicial examinou a versão 1.0.8, no commit `817ae9b`, em 2 de outubro de 2026. A suíte existente passou em uma cópia isolada: **60 verificações aprovadas, nenhuma falha**, incluindo OCR do Windows. Esse resultado estabelece uma referência para as próximas mudanças; a cobertura atual ainda não verifica vários cenários descritos neste plano.

O projeto já possui preservação de partes DOCX, mapas de origem para Markdown, recuperação automática, paginação com cache e integração nativa com Windows. Essas soluções são a base do refinamento. Cada etapa deve produzir mudanças pequenas o suficiente para revisar, com verificações do comportamento que motivou a alteração.

**Entrega 1.1.0 — melhorias implementadas**

A execução deste plano produziu uma entrega que cobre as falhas reproduzidas de salvamento, recuperação, colagem e interação, além de melhorias de distribuição. As etapas abaixo continuam como roteiro completo; os itens de arquitetura e validação externa ainda exigem trabalho próprio.

| Frente | Entregue nesta versão | Trabalho restante |
|---|---|---|
| Salvamento | Fila por destino e anexos, temporários exclusivos, flush, publicação por rename, hash antes de substituir, escolhas explícitas para conflito, manutenção de edições posteriores ao snapshot | Concorrência com programas externos continua sujeita ao intervalo entre a última conferência e o rename |
| Recuperação | Snapshot único com checksum, leitura legada, contenção de corrida entre gravar/limpar/fechar, adoção após parser concluir, preservação de cópia corrompida e fallback DOCX para recursos novos incompatíveis com texto/CSV/Markdown | Interface para diagnóstico e recuperação manual de cópias danificadas |
| Colagem e mídia | Bytes de imagens incorporados, identidade de origem, remapeamento de estilos/listas, descarte de URLs do modelo e nomes de anexos Markdown sem colisão | Medir memória em importadores externos e documentos com muitas imagens |
| Histórico | Margens, papel, orientação, tabulação padrão, numeração de página e opções de cabeçalho podem ser desfeitas/refeitas; edições de cabeçalho/rodapé durante serialização continuam pendentes | Histórico compartilhado entre corpo e editores de cabeçalho; modelo de várias seções |
| Paginação | Notas de tabelas acompanham a página da linha da referência, sem duplicação | Cabeçalhos de tabela repetidos, linhas maiores que uma página e geometria de várias seções |
| Formatos | ODT/OTT intactos preservam bytes; perdas identificadas são avisadas antes de substituir, inclusive no formato original; metadados de autor e diagnósticos acompanham o fluxo | Preservar recursos avançados ODT após edição; corpus externo de Word/LibreOffice e matriz mais ampla |
| Interface | Teclado em menus/abas/botões divididos, foco modal, ARIA, leitura idempotente, unidades cm/in/pt, progresso persistente, adaptação a janelas estreitas e alto contraste | Verificação com leitor de tela, escala alta, tema escuro e refinamento dos controles de layout |
| Serviços e segurança | IPC valida janela/frame/URL; anexos recusam traversal e junction; auxiliares têm timeout/controle de falha; preferências inválidas são normalizadas | Limites de expansão e tempo dos importadores; política completa para recursos remotos |
| Build e distribuição | Recursos PDF versionados, testes Node/Electron com diagnóstico, teste no executável empacotado, checks em PR, permissões menores, releases históricas com checksum | Validação em Windows limpo e assinatura de código quando houver certificado |

A interface foi capturada e inspecionada no Electron em 1366×768, 768×600 e 560×420, incluindo diálogos. Na janela estreita, a faixa de opções e a página permitem rolagem, e os botões de confirmação permanecem acessíveis. Isso não substitui a verificação com leitor de tela nem toda a matriz de escalas proposta adiante.

A validação final do código passou em **52 testes Node e 154 verificações no Electron**, com nenhuma falha. Inclui OCR nativo, preservação de bytes, reabertura de formatos, falhas de escrita, conflitos, corrida de recuperação, colagem, histórico e teclado. O instalador desta entrega é gerado localmente com Inno Setup 6.7.3, usando Electron 44.4.5 e Node 24.19.0.

As mesmas **154 verificações passaram no executável empacotado**, incluindo recursos PDF, OCR e serviços carregados do pacote. A atualização local para 1.1.0 terminou com código de saída zero; o `app.asar` instalado tem o mesmo SHA-256 do pacote testado. Os arquivos JSON existentes de preferências permaneceram iguais. O atalho da Área de Trabalho aponta para a instalação atualizada, e `Redator-Setup.exe` foi colocado na Área de Trabalho.

As prioridades seguem o impacto: **P0** protege documentos e trabalho pendente; **P1** corrige funções essenciais, compatibilidade e acesso por teclado; **P2** melhora desempenho, acabamento e manutenção. Uma hipótese identificada no código precisa de reprodução antes de ser tratada como defeito confirmado.

1. **Estabelecer a referência de qualidade e ampliar a cobertura**

   Registrar build, testes, corpus de documentos e limitações conhecidas. Separar verificações rápidas de lógica e persistência dos testes integrados no Electron. O executor deve registrar falhas de inicialização, stderr, código de saída e timeout; testes dependentes de recursos opcionais devem informar quando foram pulados.

   Ampliar o corpus com DOCX, DOC, ODT e RTF produzidos por Word e LibreOffice. Incluir documentos com revisões, objetos preservados, cabeçalhos, tabelas mescladas, seções, imagens, notas, listas multinível e fontes ausentes. Manter arquivos de texto com diferentes codificações, BOM, quebras de linha e nomes com acentos.

   **Aceite:** executar a suíte de referência sem regressões; novas verificações devem falhar antes da correção correspondente e passar depois. Cada fixture precisa descrever a origem, o recurso exercitado e o resultado esperado. Exportações precisam reabrir e preservar conteúdo, estrutura e recursos relevantes; tamanho do arquivo sozinho não comprova sucesso.

   **Prioridade e esforço:** P1, médio. Começa imediatamente e acompanha todas as etapas.

2. **Garantir salvamento coerente e proteger o original**

   A reprodução isolada do método de salvamento confirmou um defeito: uma edição feita enquanto a escrita aguarda pode ser marcada como salva, embora o disco contenha a versão anterior. A recuperação também pode ser apagada nesse cenário. O salvamento deve capturar a revisão correspondente ao conteúdo escrito e limpar apenas essa revisão.

   Serializar pedidos de salvamento e usar temporários exclusivos por operação. Tratar fechamento durante gravação, vários Ctrl+S, Salvar como e troca de documento. O fechamento só deve prosseguir quando a última revisão estiver salva ou o usuário tiver escolhido descartar as alterações.

   Substituir a tentativa de gravação direta sobre o original após falha de rename por um fluxo que preserve a versão anterior. Para Markdown com imagens auxiliares, preparar e validar os anexos antes de declarar sucesso; falhas devem deixar uma mensagem clara e manter as alterações pendentes. Usar o mesmo serviço de escrita para documentos, PDF e dados internos quando aplicável.

   Detectar alterações externas antes de gravar, além da conferência ao ganhar foco. Comparar a versão esperada e oferecer recarregar, salvar uma cópia ou resolver o conflito. A comparação deve considerar conteúdo quando tamanho e data não forem suficientes. A implementação precisa reconhecer limites de concorrência com programas que não cooperam com bloqueios.

   **Aceite:** digitar durante um salvamento lento conserva o estado pendente; vários pedidos não colidem; fechar não descarta uma revisão nova. Simular disco cheio, permissão negada, arquivo bloqueado, falha de substituição e erro de anexo. O arquivo anterior deve permanecer íntegro quando a operação falhar. Uma alteração externa deve gerar conflito antes de ser sobrescrita.

   **Prioridade e esforço:** P0, médio. É o primeiro marco de execução.

3. **Fortalecer recuperação, preferências e dados locais**

   Versionar snapshots de recuperação e manter a última cópia íntegra durante a preparação da próxima. Hoje dados e metadados são substituídos em etapas separadas; a inicialização também remove arquivos órfãos. Esse fluxo precisa ser validado com interrupções em cada etapa.

   Definir um limite máximo de atraso para a recuperação, inclusive durante digitação contínua. Mostrar falha de recuperação e permitir tentar novamente. Guardar os recursos necessários para restaurar imagens, cabeçalhos e conteúdo convertido, mantendo codificação e mapa de origem de texto quando aplicável.

   Validar o formato de preferências, recentes e estado da janela. Usar escrita segura e migrações simples de versão. Um JSON sintaticamente válido com tipo inesperado não deve impedir a inicialização. Ajustar posições de janela ao remover ou trocar monitores.

   **Aceite:** interromper a gravação em cada etapa e recuperar a última revisão íntegra; salvar com sucesso remove somente a recuperação correspondente. Preferências corrompidas ou incompatíveis devem produzir valores padrão e diagnóstico útil, sem impedir o uso. Testar nomes longos, caminhos de rede e pastas sincronizadas.

   **Prioridade e esforço:** P1, médio. Depende do serviço de persistência da etapa 2.

4. **Preservar imagens, estilos e listas ao copiar entre documentos**

   A inspeção encontrou um risco concreto no fluxo de colagem: imagens usam caminhos internos do ZIP como identidade. Dois DOCX podem ter imagens diferentes chamadas `word/media/image1.png`; o destino pode escolher sua própria imagem ao receber atributos da origem. Reproduzir esse cenário e transportar recursos com identidade explícita.

   Importar e remapear mídias, relacionamentos, estilos e numeração. IDs iguais em documentos diferentes não significam definições iguais. Manter as opções de preservar formatação da origem, adotar a formatação do destino e colar como texto. O HTML da área de transferência deve continuar utilizável em outros aplicativos.

   Gerenciar URLs de mídia por documento. Liberar recursos ao trocar ou fechar documentos, respeitando tarefas em andamento e a área de transferência. Copiar uma imagem e fechar a origem deve continuar permitindo colá-la corretamente.

   **Aceite:** copiar entre arquivos com caminhos e IDs iguais, mas recursos distintos; salvar e reabrir preserva pixels, dimensões, texto alternativo, estilos e listas. Repetir com listas multinível, reinício de numeração, cabeçalhos e origem fechada. Abrir e fechar documentos ricos em imagens não deve acumular URLs indefinidamente.

   **Prioridade e esforço:** P0 para integridade da imagem, P1 para estilos e listas; médio a grande.

5. **Completar o modelo de documento e o histórico de edição**

   Unificar ações sobre texto e metadados em transações que possam ser desfeitas. Atualmente margens, orientação e outras configurações modificam o modelo fora do histórico ProseMirror. O usuário precisa desfazer as ações na ordem em que as executou.

   Representar seções por intervalo do documento, com tamanho de página, orientação, margens, herança de cabeçalho e rodapé, primeira página e numeração. Hoje a geometria usa uma seção global; seções intermediárias podem ter seu XML preservado sem aparecer corretamente na tela.

   Inventariar recursos DOCX como editáveis, preservados ou convertidos. Ampliar a preservação de revisões e wrappers XML de bloco, mantendo autores, IDs e relacionamentos quando houver suporte. Definir o comportamento de campos, sumário, objetos opacos e macros convertidas para DOCX conforme o conteúdo real.

   **Aceite:** texto → margens → orientação → cabeçalho deve ser desfeito e refeito na sequência correta. Documento retrato → paisagem → retrato deve respeitar as três seções na tela, no PDF e ao reabrir o DOCX. Abrir e salvar fixtures sem editar precisa conservar os recursos que a matriz declara preservados.

   **Prioridade e esforço:** P1, grande. Dividir histórico, seções e preservação em entregas independentes.

6. **Corrigir paginação de tabelas, notas e objetos**

   Paginar tabelas com cabeçalhos repetidos, regras de divisão de linha, células longas e mesclagens verticais. O fluxo atual quebra entre linhas e pode aceitar uma linha maior que a área da página. As notas de uma tabela também são agregadas à página final; cada nota deve acompanhar sua referência.

   Validar parágrafos com viúvas e órfãs, manter com o próximo, quebras explícitas, imagens flutuantes, tabulações, recuos e cabeçalhos extensos. Tratar substituição de fontes e mudanças de zoom sem deslocamentos indevidos. A preparação para impressão deve esperar fontes, imagens e paginação estabilizarem.

   **Aceite:** tabela de várias páginas com cabeçalho repetido, células mescladas, linha longa e notas distribuídas; conteúdo não invade rodapé nem desaparece. Tela e PDF devem concordar em quebras, dimensões e posição dos elementos do corpus. Acrescentar verificações de corte e sobreposição, além da contagem de páginas.

   **Prioridade e esforço:** P1, grande. Depende do modelo de seções para os casos que atravessam seções.

7. **Tornar compatibilidade e conversões previsíveis**

   Criar uma matriz por recurso e formato, incluindo abertura, edição, preservação e exportação. O aviso de perda deve se basear no conteúdo do documento e no destino. A gravação de ODT reconstrói um pacote e usa um conjunto limitado de propriedades de página; recursos incompatíveis precisam ser identificados antes de substituir um arquivo.

   Ampliar a fidelidade ODT por etapas, incluindo masters de página, cabeçalhos, objetos e partes não editadas. Para PDF, medir ordem de leitura, tabelas, OCR, fontes e imagens em documentos reais; mostrar progresso e permitir cancelamento. Conferir RTF, DOC antigo, EPUB, MHT e EML com fixtures externas e diagnósticos de conversão.

   Preservar as garantias existentes de Markdown, CSV e texto: bytes sem edição, codificação, BOM, finais de linha e alteração restrita ao trecho editado. Testar múltiplas edições e salvamentos sucessivos, renomeação, links relativos e imagens auxiliares.

   **Aceite:** perdas previstas aparecem antes de substituir o original e oferecem salvar uma cópia. Cada formato exportado reabre com conteúdo e estrutura conferidos. A matriz deve corresponder aos testes, e o README deve declarar os limites observados.

   **Prioridade e esforço:** P1, grande. Começar pelos formatos usados com mais frequência e pelos riscos de perda silenciosa.

8. **Refinar interação, acessibilidade e aparência**

   Corrigir a transição de leitura e edição. A sequência Leitura → Leitura → Impressão foi reproduzida em isolamento e deixa um documento editável protegido. Separar a permissão do arquivo do modo de visualização e tornar as transições idempotentes.

   Dar aos menus foco inicial, setas, Enter, Espaço, Escape e navegação de submenus. Fazer botões divididos funcionarem pelo teclado. Implementar foco contido em diálogos, fundo inerte, fechamento da camada superior e retorno ao acionador. Expor nomes e estados de abas, controles, pesquisa, zoom e avisos. Usar os padrões de [menus](https://www.w3.org/WAI/ARIA/apg/patterns/menubar/) e [diálogos modais](https://www.w3.org/WAI/ARIA/apg/patterns/dialog-modal/) do W3C como referência.

   Implementar cm, polegadas e pontos na apresentação e entrada, mantendo unidades internas estáveis. Atualizar os atalhos exibidos quando a preferência mudar. Separar progresso persistente de mensagens transitórias; abertura, salvamento e OCR devem continuar indicando atividade enquanto estiverem em andamento.

   Adaptar a faixa de opções, os painéis e os diálogos a janelas pequenas e escala alta. Consolidar cores, espaçamento, tipografia, bordas, foco e estados. Ampliar áreas de interação de controles pequenos. Evoluir temas claro, escuro e alto contraste após estabilizar os componentes e testar a legibilidade do documento em cada um.

   **Aceite:** abrir um recente, navegar por títulos, inserir tabela, escolher cor e formatar um parágrafo usando apenas teclado. Validar com Narrador ou NVDA. Conferir 560×420, 768×600, 1024×768 e 1366×768, com escala de 100%, 150% e 200%; todas as ações devem permanecer alcançáveis. Trocar unidades não altera as dimensões físicas. Uma operação acima de quatro segundos mantém o progresso visível.

   **Prioridade e esforço:** P1 para funções e acessibilidade, P2 para acabamento; médio a grande. Pode avançar em paralelo com persistência e formatos.

9. **Medir e melhorar desempenho e arquitetura**

   Medir tempo de abrir, salvar, pesquisar e paginar; latência de digitação; tarefas longas; uso de memória; e tamanho do pacote. Usar documentos de 10, 100 e 500 páginas, documentos com muitas imagens e texto de 10 MB. Registrar máquina, versão, corpus e variação entre execuções.

   Priorizar o que os perfis mostrarem: paginação incremental, redução de consultas ao DOM, busca cancelável fora do processo principal, processamento pesado em worker e descarte de recursos. Virtualização exige uma prova de que seleção, impressão, notas e paginação continuam corretas. Definir metas depois da referência medida, usando percentis para evitar que uma média esconda travamentos.

   Separar gradualmente persistência, recuperação, janelas, preferências, integrações Windows, ciclo do documento e comandos. Reduzir dependência do contexto global de renderização em operações assíncronas. Centralizar o cadastro de formatos e extensões hoje repetido no processo principal e no renderer. Acrescentar tipos ou contratos nas fronteiras mais sujeitas a erro.

   **Aceite:** conservar os benchmarks existentes e registrar a melhora ou ausência de regressão no corpus. Abrir e fechar documentos repetidamente deve estabilizar memória. Cada extração de módulo precisa manter comportamento e cobertura; mudanças de arquitetura devem apoiar uma necessidade observada.

   **Prioridade e esforço:** P2, grande e incremental. Medição começa na primeira rodada; otimização depende dos resultados.

10. **Reforçar segurança e integridade de entrada**

    Manter isolamento de contexto, sandbox e Node desativado no renderer. Validar remetente e frame de IPC, tipos e limites de argumentos. Confinar anexos ao diretório autorizado e tratar caminhos normalizados, traversal, links simbólicos e destinos de rede de acordo com o fluxo de arquivo. A inspeção confirmou ausência de validação em partes dessa fronteira; não foi demonstrada exploração.

    Restringir navegação, protocolos e abertura externa; aplicar uma política explícita a recursos remotos de documentos. Revisar HTML, SVG, XML, ZIP e imagens de entrada, com limites de tamanho, expansão e tempo que permitam documentos legítimos grandes. Tratar falha, timeout e reinicialização dos auxiliares de OCR, ortografia e conversão.

    Revisar dependências e atualizações com testes de compatibilidade. Usar as [recomendações oficiais do Electron](https://www.electronjs.org/docs/latest/tutorial/security) como referência para IPC, protocolos, CSP e permissões. Diagnósticos e relatórios de falha devem evitar conteúdo de documentos e dados pessoais desnecessários.

    **Aceite:** testes negativos rejeitam frame desconhecido, argumentos inválidos e anexo fora do destino. Arquivos malformados produzem erro útil sem travar o aplicativo. Acesso à rede deve respeitar a configuração do produto. As medidas precisam ser verificadas no aplicativo empacotado.

    **Prioridade e esforço:** P1, médio a grande. Confinamento de anexos acompanha persistência; demais limites avançam com os importadores.

11. **Tornar build e distribuição reproduzíveis**

    Executar checks em pull requests e manter permissões mínimas no CI. Atualizar recursos PDF quando a dependência mudar; build limpo e incremental devem usar os mesmos recursos. Incluir novos serviços no empacotamento e verificar que o aplicativo não depende de arquivos do repositório ou de Node instalado.

    Publicar versões com tags próprias, SHA correto, notas e checksum. O fluxo atual apaga a release `latest` antes de recriá-la; preservar o download anterior até a nova publicação estar pronta. Concluir build e instalador antes de gravar a atualização de versão na branch.

    Testar instalação, atualização, associações, abertura por duplo clique, impressão e desinstalação em Windows limpo. Adotar assinatura de código quando houver certificado e configuração de CI disponíveis. Manter documentação de desenvolvimento, arquitetura, formatos, diagnóstico e contribuição alinhada ao comportamento.

    **Aceite:** PR executa a suíte sem publicar; falha de build ou publicação conserva a versão anterior. Release identifica o commit efetivamente empacotado, mantém histórico e oferece checksum. Instalação e atualização preservam preferências e documentos, e a desinstalação respeita a escolha sobre dados locais.

    **Prioridade e esforço:** P1 para CI e publicação, P2 para documentação e assinatura; médio. A assinatura depende de recurso externo.

A execução deve produzir marcos verificáveis. O primeiro entrega salvamento, conflitos e regressões de estado. O segundo reúne colagem, controles de teclado, unidades e adaptação da interface. O terceiro trata histórico, seções, tabelas e fidelidade com corpus externo. O quarto fecha desempenho, segurança, instalador e documentação. Checks de qualidade acompanham cada marco.

| Frente | Depende de | Pode avançar em paralelo | Tamanho relativo |
|---|---|---|---|
| Salvamento e conflitos | Referência de testes | Interface e corpus | Médio |
| Recuperação e dados locais | Serviço de persistência | Colagem e IPC | Médio |
| Colagem e recursos | Fixtures cruzadas | Interface e CI | Médio a grande |
| Histórico e seções | Contratos do modelo | Compatibilidade de formatos | Grande |
| Tabelas e notas | Geometria de seções | Corpus e validação visual | Grande |
| Interface e acessibilidade | Componentes de interação | Persistência e formatos | Médio a grande |
| Desempenho | Medições e casos reais | Arquitetura incremental | Grande |
| Segurança e distribuição | Checks e pacote verificável | Demais frentes | Médio a grande |

Esses tamanhos são comparações de complexidade, sem promessa de calendário. Ao iniciar cada entrega, registrar escopo, casos de teste e dependências; atualizar o plano conforme as medições e os documentos reais revelarem novas necessidades.

O refinamento está concluído para uma entrega quando os testes existentes e as regressões correspondentes passam, os fluxos afetados são conferidos no Electron, o pacote inclui todos os recursos necessários e as limitações ficam documentadas. Para mudanças de layout e formatos, incluir comparação visual e reabertura em aplicativo independente. Para mudanças de persistência, incluir falhas e concorrência. O commit deve descrever o problema e o comportamento resultante; o push deve conter apenas alterações revisadas.

As evidências iniciais mais relevantes estão nos pontos abaixo. As linhas se referem ao commit auditado e podem mudar durante a execução.

| Evidência | Fonte inicial | Estado da conclusão |
|---|---|---|
| Edição durante gravação marcada como salva | `src/app.js`, método `writeTo`, linha 737 | Confirmado com método real e dependências simuladas |
| Leitura repetida conserva bloqueio ao sair | `src/app.js`, método `setViewMode`, linha 1382 | Confirmado com método real e dependências simuladas |
| Temporário fixo, fallback direto e falha de anexo omitida | `main.js`, handler `file:write`, linha 414 | Confirmado por leitura; falhas reais entram na regressão |
| Conferência externa apenas no ganho de foco | `src/app.js`, linha 273 e método `checkExternalChange` | Confirmado no fluxo |
| Identidade de mídia por caminho interno | `src/editor/nodeviews.js`, linha 9; `src/docx/writer.js`, linha 90 | Cenário de colagem precisa de teste integrado |
| Geometria global de seção | `src/docx/reader.js`, linha 759; `src/editor/layout.js`, linha 444 | Limitação confirmada no modelo e renderização |
| Tabela dividida só entre linhas e notas agregadas | `src/editor/layout.js`, método `placeTable`, linha 846 | Limitação confirmada no algoritmo |
| Preferência de unidades não aplicada | `src/ui/dialogs.js`, linha 738; `src/ui/ruler.js`, linha 63 | Confirmado por busca e leitura |
| Menus e modais sem navegação completa de foco | `src/ui/popup.js`, linha 23; `src/ui/dialogs.js`, linha 17 | Confirmado por leitura; conferir com teclado e leitor de tela |
| Publicação apaga release anterior | `.github/workflows/release.yml`, linha 61 | Confirmado no workflow |

O corpus externo, a comparação visual com Word e LibreOffice, a instalação em Windows limpo e as medições completas de desempenho continuam sendo requisitos de validação. A suíte inicial aprovada e as reproduções isoladas não substituem essas verificações.
