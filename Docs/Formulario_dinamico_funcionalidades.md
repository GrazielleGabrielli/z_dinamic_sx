# Form

## Edição nativa

Ao editar a página do SharePoint, o painel lateral exibe as configurações principais da página e da WebPart.

Nessa área ficam disponíveis os campos e botões usados para ajustar o formulário dinâmico.

### 1. Input Título

Campo nativo do SharePoint usado para definir o título da página.

O valor informado nesse campo será exibido como o título da página onde o formulário está publicado.

Exemplo:

`Formulário de Férias`

### 2. Edição Dinâmica

Área responsável por abrir a configuração do formulário dinâmico.

#### Editar formulário

Botão usado para acessar a edição do formulário.

Ao clicar em `Editar formulário`, o usuário abre a área onde os campos, regras e comportamento do formulário podem ser configurados.

Use essa opção quando for necessário ajustar a estrutura do formulário, como campos exibidos, obrigatoriedade, organização e regras de preenchimento.

### 3. Configuração geral

Área responsável por abrir as configurações gerais da WebPart.

#### FlexView — Editar configuração

Botão usado para acessar as configurações gerais da solução.

Ao clicar em `FlexView — Editar configuração`, o usuário abre o painel de configuração da WebPart, onde são definidos os ajustes principais de funcionamento.

Use essa opção quando for necessário revisar ou alterar configurações gerais da visualização e do comportamento da WebPart na página.

## Configurar formulário e regras

O modal `Configurar formulário e regras` concentra as principais opções de configuração do formulário dinâmico.

Ele é dividido em abas para organizar cada tipo de ajuste e facilitar a manutenção da configuração.

### JSON

A opção `JSON (ver / colar)` permite visualizar ou colar a configuração completa do formulário.

Esse recurso é utilizado principalmente para copiar configurações entre ambientes, como homologação e produção.

Com isso, uma configuração validada em homologação pode ser copiada e colada em produção, refletindo os mesmos campos, regras e comportamentos sem a necessidade de refazer os ajustes manualmente.

### Abas disponíveis

#### Estrutura

Aba usada para configurar a estrutura principal do formulário.

Nela são organizados os campos, seções e definições principais que compõem a tela do formulário.

#### Regras dos campos

Aba usada para configurar comportamentos dos campos.

Ela permite definir regras relacionadas à exibição, obrigatoriedade, bloqueio e demais comportamentos condicionais dos campos.

#### Componentes

Aba usada para configurar componentes adicionais do formulário.

Esses componentes complementam a experiência da tela e podem ser usados para organizar ou enriquecer a interface.

#### Anexos

Aba usada para configurar o comportamento de anexos no formulário.

Ela centraliza os ajustes relacionados ao uso de arquivos vinculados ao registro.

#### Botões

Aba usada para configurar os botões disponíveis no formulário.

Ela define quais ações estarão acessíveis para o usuário durante o uso do formulário.

#### Auditoria e versões

Aba usada para consultar ou configurar recursos relacionados ao histórico do formulário.

Ela apoia o acompanhamento de alterações, versões e registros de auditoria quando disponíveis.

#### Listas vinculadas

Aba usada para configurar relações com listas vinculadas.

Ela permite conectar o formulário a outras listas usadas como apoio ou complemento da informação principal.

#### Quebra de permissões

Aba usada para configurar regras relacionadas a permissões do item.

Ela permite tratar cenários em que o registro precisa ter permissões específicas, diferentes da lista principal.

## Funcionalidades

### Estrutura

A aba `Estrutura` reúne as configurações que definem como o formulário será apresentado para o usuário.

Nessa área são ajustados pontos visuais e comportamentos de navegação, como largura do formulário, alinhamento na página, espaçamento interno e regras para avançar ou voltar entre etapas.

Essa aba deve ser usada quando for necessário controlar a experiência geral de preenchimento do formulário.

Exemplos de uso:

- Ajustar o formulário para ocupar toda a largura disponível da página.
- Centralizar o formulário para melhorar a leitura.
- Criar espaçamento interno para deixar os campos mais confortáveis visualmente.
- Impedir que o usuário avance para a próxima etapa sem preencher campos obrigatórios.
- Permitir que o usuário volte etapas anteriores sem bloquear a navegação.

#### Layout do formulário

A seção `Layout do formulário` controla a forma como o formulário aparece visualmente na página.

Ela ajuda a adaptar o formulário ao tipo de conteúdo, quantidade de campos e experiência esperada para o usuário final.

##### Largura

Define como a largura do formulário será calculada dentro da área disponível da página.

Quando a opção estiver configurada como `Porcentagem da área disponível`, o formulário ocupará uma porcentagem da largura disponível no espaço onde ele foi inserido.

Essa opção é útil quando o formulário precisa se adaptar melhor a diferentes tamanhos de tela ou layouts de página.

Exemplo:

Um formulário simples, com poucos campos, pode ocupar menos espaço para ficar mais confortável visualmente.

Um formulário maior, com muitos campos ou seções, pode ocupar toda a largura disponível para facilitar a leitura.

##### Porcentagem da largura

Define quanto da área disponível será ocupada pelo formulário.

O valor deve ser informado de `1` a `100`.

Exemplos:

- `100`: o formulário ocupa toda a largura disponível.
- `80`: o formulário ocupa 80% da largura disponível.
- `60`: o formulário ocupa 60% da largura disponível.

Uso recomendado:

- Use `100` quando o formulário tiver muitos campos ou precisar aproveitar toda a tela.
- Use valores menores quando quiser deixar o formulário mais compacto.
- Use valores entre `70` e `90` quando quiser equilibrar leitura e aproveitamento de espaço.

##### Alinhamento horizontal

Define onde o formulário ficará posicionado dentro da área disponível.

As opções de alinhamento ajudam a controlar a posição do formulário quando ele não ocupa 100% da largura.

Exemplos de uso:

- `Início (esquerda)`: mantém o formulário alinhado à esquerda da página.
- `Centro`: posiciona o formulário no centro da área disponível.
- `Fim (direita)`: posiciona o formulário à direita da área disponível.

Uso recomendado:

Para formulários utilizados por clientes ou usuários finais, o alinhamento central costuma deixar a tela mais equilibrada quando a largura for menor que 100%.

Quando o formulário ocupar 100% da largura, o alinhamento terá pouco ou nenhum efeito visual.

##### Padding

Define o espaçamento interno do formulário, em pixels.

Esse espaçamento cria uma margem interna entre as bordas do formulário e o conteúdo exibido dentro dele.

Na prática, o `Padding` ajuda a deixar os campos menos colados nas bordas e melhora a leitura.

Exemplos:

- `0`: sem espaçamento interno adicional.
- `20`: espaçamento leve.
- `50`: espaçamento maior, deixando o conteúdo mais afastado das bordas.

Uso recomendado:

- Use valores menores em formulários com muitos campos, para aproveitar melhor o espaço.
- Use valores maiores em formulários mais simples, quando quiser uma apresentação mais limpa.
- Evite valores muito altos quando a tela tiver pouco espaço disponível.

#### Navegação entre etapas

A seção `Navegação entre etapas` controla como o usuário pode avançar ou voltar dentro de formulários divididos em etapas.

Essas opções são importantes quando o formulário funciona como um fluxo, por exemplo:

- Dados do solicitante.
- Informações da solicitação.
- Anexos.
- Revisão final.
- Confirmação.

Com essas configurações, é possível definir se o usuário poderá avançar livremente ou se precisará preencher e validar as informações antes de seguir.

##### Exigir obrigatórios preenchidos para avançar

Define se o usuário precisa preencher os campos obrigatórios da etapa atual antes de avançar para a próxima etapa.

Quando ativado, o formulário bloqueia o avanço caso exista algum campo obrigatório sem preenchimento.

Quando desativado, o usuário pode avançar mesmo que ainda existam campos obrigatórios pendentes.

Exemplo de uso:

Em um formulário de solicitação de férias, a etapa com `Data de início` e `Data de término` pode exigir preenchimento antes de permitir que o usuário avance.

Uso recomendado:

Ative essa opção quando a ordem do preenchimento for importante ou quando a próxima etapa depender das informações da etapa atual.

##### Ao avançar, aplicar todas as regras de validação nos campos da etapa

Define se, ao avançar, o formulário deve conferir todas as validações configuradas nos campos da etapa atual.

Essa opção vai além da obrigatoriedade. Ela também considera outras regras que possam existir para os campos.

Exemplos:

- Verificar se uma data está dentro de um período permitido.
- Verificar se um valor informado atende a uma regra definida.
- Verificar se determinado campo foi preenchido conforme uma condição.

Uso recomendado:

Ative essa opção quando a etapa possuir regras importantes que precisam ser respeitadas antes de seguir.

Em formulários mais simples, essa opção pode ficar desativada caso a validação completa só precise acontecer no momento final do envio.

##### Permitir voltar etapa sem validar a atual

Define se o usuário pode voltar para uma etapa anterior sem que a etapa atual seja validada.

Quando ativado, o usuário pode retornar para revisar informações anteriores mesmo que a etapa atual ainda esteja incompleta.

Quando desativado, o formulário pode exigir que a etapa atual esteja válida antes de permitir a navegação de volta.

Exemplo de uso:

Um usuário está na etapa de anexos, mas percebe que precisa corrigir uma informação nos dados iniciais. Com essa opção ativada, ele consegue voltar sem precisar concluir os anexos primeiro.

Uso recomendado:

Mantenha essa opção ativada quando quiser facilitar a revisão das etapas anteriores.

Desative somente quando o processo exigir controle rígido da ordem de preenchimento.

#### Botão Nova etapa

O botão `Nova etapa` é usado para adicionar uma nova etapa ao formulário.

Uma etapa funciona como uma divisão do formulário em partes menores, facilitando o preenchimento quando existem muitos campos ou quando o processo precisa seguir uma sequência lógica.

Ao criar uma nova etapa, o formulário passa a ter uma nova área de preenchimento, que pode receber campos específicos daquela parte do processo.

Exemplos de etapas:

- `Dados do solicitante`
- `Informações da solicitação`
- `Detalhes do pedido`
- `Documentos e anexos`
- `Revisão e envio`

Essa funcionalidade é útil para transformar formulários longos em um fluxo mais organizado, evitando que o usuário veja todos os campos de uma vez.

Exemplo de uso:

Em um formulário de férias, é possível criar uma etapa para os dados do colaborador, outra para o período solicitado e outra para revisão final antes do envio.

Uso recomendado:

- Use etapas quando o formulário tiver muitos campos.
- Use etapas quando o preenchimento precisar seguir uma ordem.
- Use etapas para separar assuntos diferentes dentro do mesmo formulário.
- Evite criar etapas demais quando o formulário for simples, para não tornar o preenchimento mais longo do que o necessário.

#### Aba Ocultos

A aba `Ocultos` é usada para organizar campos que fazem parte do registro, mas que não devem aparecer visualmente no formulário para o usuário.

Esses campos continuam existindo na estrutura do formulário e podem ser enviados junto com os metadados do item, mesmo sem serem exibidos na tela.

Na prática, isso permite manter informações importantes no registro sem exigir que o usuário visualize ou preencha esses campos diretamente.

Exemplos de uso:

- Campo interno usado para controle do processo.
- Status inicial definido automaticamente.
- Identificador auxiliar usado em integrações.
- Informação técnica que precisa ser gravada no item.
- Campo usado como apoio para regras ou organização dos dados.

Exemplo prático:

Em um formulário de solicitação, o campo `Origem da solicitação` pode ficar em `Ocultos` com um valor padrão, como `Portal`, sem aparecer para quem está preenchendo.

Uso recomendado:

- Use `Ocultos` para campos que precisam ser enviados no registro, mas não precisam ser exibidos ao usuário.
- Use essa área para metadados de apoio, controle interno ou informações preenchidas automaticamente.
- Evite colocar em `Ocultos` campos que o usuário precisa revisar antes de enviar o formulário.

#### Etapas criadas

As etapas criadas pelo botão `Nova etapa` representam as partes visíveis do formulário.

Cada etapa pode agrupar campos relacionados a um mesmo assunto, deixando o preenchimento mais organizado e fácil para o usuário.

Exemplos:

- Uma etapa para dados pessoais.
- Uma etapa para informações da solicitação.
- Uma etapa para dados financeiros.
- Uma etapa para anexos.
- Uma etapa para revisão final.

As etapas podem ser arrastadas para alterar a ordem em que aparecem no formulário.

Isso permite reorganizar o fluxo de preenchimento sem precisar recriar a estrutura do zero.

Exemplo prático:

Se a etapa `Anexos` foi criada antes da etapa `Dados da solicitação`, ela pode ser arrastada para depois, deixando o formulário em uma ordem mais natural para o usuário.

Uso recomendado:

- Coloque primeiro as etapas com informações básicas.
- Deixe etapas de complemento, anexos ou revisão para o final.
- Agrupe campos do mesmo assunto na mesma etapa.
- Evite misturar campos de temas diferentes em uma única etapa.

##### Campos dentro das etapas

Dentro de cada etapa ficam os campos que serão exibidos ao usuário naquela parte do formulário.

Esses campos podem ser organizados conforme a ordem desejada de preenchimento.

A ordenação dos campos ajuda a conduzir o usuário de forma mais clara, começando pelas informações mais simples e avançando para dados mais específicos.

Exemplo de organização:

- `Nome do colaborador`
- `Matrícula`
- `Área`
- `Gestor responsável`

Nesse exemplo, os campos seguem uma ordem de identificação antes de avançar para dados mais específicos da solicitação.

Os campos também podem ser movidos entre etapas quando fizer sentido reorganizar o formulário.

Exemplo prático:

Se o campo `Centro de custo` estiver na etapa `Dados do solicitante`, mas fizer mais sentido ficar em `Informações da solicitação`, ele pode ser reposicionado para deixar a etapa mais coerente.

Uso recomendado:

- Ordene os campos na mesma sequência em que o usuário deve preencher.
- Mantenha juntos os campos que tratam do mesmo assunto.
- Evite etapas com campos demais quando o formulário puder ser dividido em partes menores.
- Revise a ordem das etapas e dos campos antes de publicar o formulário para o cliente final.

##### Botão Configurar da etapa

O botão `Configurar` da etapa abre o painel `Visibilidade da etapa`.

Esse painel define em quais situações a etapa deve aparecer no formulário.

Ele é usado quando uma etapa não deve estar disponível para todos os momentos ou para todos os usuários.

Exemplos de uso:

- Mostrar uma etapa apenas durante a criação do item.
- Mostrar uma etapa apenas quando o formulário estiver em edição.
- Mostrar uma etapa apenas para consulta, no modo de visualização.
- Mostrar uma etapa somente quando um campo tiver determinado valor.
- Mostrar uma etapa somente para usuários de um grupo específico.

Uso recomendado:

- Use essa configuração quando a etapa precisa aparecer em situações específicas.
- Mantenha a etapa visível em todos os modos quando ela fizer parte do fluxo principal.
- Use condições para evitar que o usuário veja etapas que não se aplicam ao caso dele.

##### Modos de formulário

A área `Modos de formulário` define em quais modos a etapa será exibida.

Os modos representam o momento em que o formulário está sendo usado.

Opções disponíveis:

- `Criar`: quando o usuário está preenchendo um novo registro.
- `Editar`: quando o usuário está alterando um registro já existente.
- `Ver`: quando o usuário está apenas consultando as informações do registro.

Quando todos os modos estão marcados, a etapa aparece em todos os cenários.

Quando apenas alguns modos estão marcados, a etapa aparece somente nos modos escolhidos.

Exemplos:

- Uma etapa de `Dados iniciais` pode aparecer em `Criar`, `Editar` e `Ver`.
- Uma etapa de `Aprovação` pode aparecer apenas em `Editar` e `Ver`.
- Uma etapa de `Instruções de preenchimento` pode aparecer apenas em `Criar`.

Uso recomendado:

- Use `Criar` para etapas importantes no cadastro inicial.
- Use `Editar` para etapas que precisam ser ajustadas depois que o item já existe.
- Use `Ver` para etapas que devem aparecer na consulta do registro.
- Desmarque modos quando a etapa não fizer sentido naquele momento.

##### Condições de exibição da etapa

A opção `Só mostrar esta etapa quando as condições abaixo forem verdadeiras` permite controlar a exibição da etapa com base em regras.

Quando essa opção está desativada, a etapa aparece conforme os modos de formulário configurados.

Quando essa opção está ativada, a etapa só aparece se a condição definida for atendida.

Exemplo:

Uma etapa chamada `Dados do veículo` pode aparecer somente quando o campo `Tipo de solicitação` for igual a `Transporte`.

Se o usuário escolher outro tipo de solicitação, essa etapa não será exibida.

Uso recomendado:

- Use condições quando a etapa depende de uma resposta anterior.
- Use condições para simplificar o formulário e mostrar apenas o que é necessário.
- Evite criar condições muito complexas quando uma divisão mais simples do formulário resolver o caso.

##### Lógica entre condições

A opção `Lógica entre condições` define como o formulário deve interpretar mais de uma condição.

Opções disponíveis:

- `Todas (E)`: a etapa só aparece se todas as condições forem verdadeiras.
- `Pelo menos uma (OU)`: a etapa aparece se qualquer uma das condições for verdadeira.

Exemplo com `Todas (E)`:

A etapa `Aprovação financeira` pode aparecer somente quando:

- `Tipo de solicitação` é igual a `Compra`.
- `Valor` é maior que `10000`.

Nesse caso, as duas condições precisam ser atendidas.

Exemplo com `Pelo menos uma (OU)`:

A etapa `Informações adicionais` pode aparecer quando:

- `Tipo de solicitação` é igual a `Urgente`.
- `Prioridade` é igual a `Alta`.

Nesse caso, basta uma das condições ser atendida para a etapa aparecer.

##### Campo

O campo `Campo` define qual informação do formulário será analisada na condição.

É a partir desse campo que o sistema decide se a etapa deve ou não ser exibida.

Exemplo:

Selecionar `Tipo de solicitação` para controlar se uma etapa específica deve aparecer.

##### Operador

O campo `Operador` define como o valor do campo será avaliado.

Exemplos de operadores:

- `é igual a`
- `é diferente de`
- `contém`
- `não contém`
- `maior que`
- `menor que`
- `está vazio`
- `não está vazio`
- `é verdadeiro`
- `é falso`

Exemplo:

Para mostrar uma etapa quando o valor da solicitação for maior que `10000`, selecione o campo de valor, escolha o operador `maior que` e informe `10000` no valor.

##### Comparar com

O campo `Comparar com` define o tipo de comparação usada na condição.

Opções disponíveis:

- `Texto fixo`: compara o campo com um valor digitado manualmente.
- `Outro campo`: compara o campo selecionado com outro campo do formulário.
- `Token`: compara com uma informação dinâmica disponível no contexto.
- `Membro do grupo`: exibe a etapa quando o usuário pertence ao grupo informado.
- `Fora do grupo`: exibe a etapa quando o usuário não pertence ao grupo informado.

Exemplos:

- Use `Texto fixo` para comparar com valores como `Sim`, `Não`, `Urgente` ou `Aprovado`.
- Use `Outro campo` quando uma etapa depender da comparação entre duas informações do formulário.
- Use `Membro do grupo` para exibir uma etapa apenas para uma equipe específica.
- Use `Fora do grupo` para esconder ou mostrar etapas conforme o usuário não pertença a determinado grupo.

##### Valor ou grupo

O campo `Valor` recebe o conteúdo usado na comparação.

Quando a condição usa `Membro do grupo` ou `Fora do grupo`, esse campo passa a representar o nome do grupo.

Exemplos:

- `Urgente`
- `Aprovado`
- `10000`
- `Gestores`
- `Equipe Financeira`

Para operadores como `está vazio`, `não está vazio`, `é verdadeiro` e `é falso`, o valor não precisa ser preenchido, pois a própria condição já define o comportamento.

##### Adicionar condição

O botão `Adicionar condição` permite incluir mais uma regra para controlar a exibição da etapa.

Cada condição adicionada pode avaliar um campo, operador e valor diferente.

Exemplo:

A etapa `Aprovação do gestor` pode aparecer quando:

- `Tipo de solicitação` é igual a `Férias`.
- `Dias solicitados` é maior que `15`.

Com isso, a etapa só aparece quando a solicitação for de férias e tiver mais de 15 dias, se a lógica estiver como `Todas (E)`.

##### Remover condição

O botão de remover condição permite excluir uma regra que não deve mais ser considerada.

Essa opção é útil quando a etapa foi simplificada ou quando uma regra deixou de fazer sentido para o processo.

##### Exemplo completo

Cenário: exibir a etapa `Aprovação financeira` somente para compras acima de `10000`.

Configuração:

- Modos de formulário: `Criar`, `Editar` e `Ver`.
- Ativar `Só mostrar esta etapa quando as condições abaixo forem verdadeiras`.
- Lógica entre condições: `Todas (E)`.
- Condição 1: campo `Tipo de solicitação`, operador `é igual a`, comparar com `Texto fixo`, valor `Compra`.
- Condição 2: campo `Valor`, operador `maior que`, comparar com `Texto fixo`, valor `10000`.

Resultado:

A etapa só será exibida quando o usuário estiver tratando uma solicitação de compra com valor maior que `10000`.

##### Botão Configurar Colunas

O botão `Configurar Colunas` abre um modal para definir como os campos daquela etapa serão distribuídos na tela.

Essa configuração controla o comportamento visual dos campos em linha, permitindo definir se um campo ocupará a linha inteira, metade da linha, um terço da linha ou outro tamanho disponível.

Na prática, ela ajuda a organizar a etapa em colunas, deixando o formulário mais compacto, legível e adequado ao tipo de informação exibida.

Exemplos de uso:

- Colocar `Nome completo` ocupando a linha inteira.
- Colocar `Data de início` e `Data de término` lado a lado.
- Colocar `DDD`, `Telefone` e `Ramal` na mesma linha.
- Fazer campos mais importantes ocuparem mais espaço.
- Ajustar o layout de forma diferente para criação, edição e visualização.

##### Modal Configurar Colunas

O modal `Configurar Colunas` mostra a etapa selecionada e todos os campos que pertencem a ela.

Cada campo aparece com opções de largura baseadas em uma grade de `12` colunas.

Essa grade funciona como uma divisão da linha disponível.

Exemplos:

- `12`: o campo ocupa a linha inteira.
- `6`: o campo ocupa metade da linha.
- `4`: o campo ocupa um terço da linha.
- `3`: o campo ocupa um quarto da linha.
- `2`: o campo ocupa uma parte menor da linha.
- `8`: o campo ocupa uma área maior que metade da linha.

Exemplo prático:

Se dois campos forem configurados com `6`, eles podem aparecer lado a lado na mesma linha.

Se um campo for configurado com `12`, ele ocupa a linha inteira e o próximo campo começa em outra linha.

##### Modos do modal

O modal possui os modos `Novo`, `Editar` e `Ver`.

Esses modos permitem configurar a disposição dos campos de forma diferente dependendo do momento de uso do formulário.

Opções disponíveis:

- `Novo`: layout usado quando o usuário está criando um novo registro.
- `Editar`: layout usado quando o usuário está alterando um registro existente.
- `Ver`: layout usado quando o usuário está apenas consultando o registro.

Exemplos:

- No modo `Novo`, os campos podem ocupar mais espaço para facilitar o preenchimento.
- No modo `Editar`, os campos podem ser organizados de forma parecida com o cadastro original.
- No modo `Ver`, os campos podem ficar mais compactos, facilitando a leitura das informações.

Uso recomendado:

- Configure primeiro o modo `Novo`, pensando no preenchimento.
- Depois revise o modo `Editar`, pensando em manutenção dos dados.
- Por fim, ajuste o modo `Ver`, pensando em leitura e consulta.

##### Campos da etapa

Dentro do modal, cada campo da etapa pode receber uma configuração própria de colunas.

Isso permite que campos diferentes tenham larguras diferentes na mesma etapa.

Exemplo:

- `Descrição da solicitação`: `12`, ocupando a linha inteira.
- `Data de início`: `6`, ocupando metade da linha.
- `Data de término`: `6`, ocupando metade da linha.
- `Quantidade de dias`: `3`, ocupando uma área menor.

Resultado:

O formulário fica mais organizado, com campos longos recebendo mais espaço e campos curtos ocupando menos largura.

##### Faixas de largura da tela

O modal permite configurar o comportamento dos campos por faixa de largura da tela.

Isso ajuda o formulário a se adaptar melhor em telas menores ou maiores.

Na prática, o mesmo campo pode se comportar de uma forma em telas grandes e de outra forma em telas menores.

Exemplo:

Em uma tela grande, `Data de início` e `Data de término` podem aparecer lado a lado.

Em uma tela menor, esses campos podem ficar um abaixo do outro para facilitar a leitura.

Uso recomendado:

- Em telas menores, prefira campos mais largos para evitar que o conteúdo fique apertado.
- Em telas maiores, use colunas para aproveitar melhor o espaço.
- Revise campos com textos longos, observações e descrições para garantir boa leitura.

##### Herança entre faixas

Quando uma faixa de largura não recebe uma configuração específica, ela herda o comportamento da faixa menor.

Isso significa que não é necessário configurar todas as faixas manualmente quando o mesmo comportamento visual já atende bem.

Uso recomendado:

- Configure apenas as diferenças necessárias.
- Use a herança para manter o layout mais simples.
- Ajuste faixas específicas apenas quando algum campo ficar desconfortável em determinado tamanho de tela.

##### Exemplo completo

Cenário: configurar a etapa `Período de férias`.

Configuração no modo `Novo`:

- `Tipo de solicitação`: `12`.
- `Data de início`: `6`.
- `Data de término`: `6`.
- `Quantidade de dias`: `4`.
- `Observações`: `12`.

Resultado:

O usuário vê o tipo da solicitação em uma linha completa, as datas lado a lado, a quantidade de dias em um espaço menor e as observações ocupando a linha inteira.

Essa organização deixa a etapa mais clara e evita que campos curtos ocupem espaço desnecessário.

##### Botão Colunas do campo

O botão `Colunas` aparece em cada campo dentro da etapa.

Ele abre o modal `Colunas na Linha`, usado para configurar o espaço que aquele campo específico ocupará dentro da linha do formulário.

Essa opção é útil quando apenas um campo precisa de ajuste individual, sem alterar todos os campos da etapa.

Exemplos de uso:

- Fazer o campo `Descrição` ocupar a linha inteira.
- Deixar `Data de início` ocupando metade da linha.
- Fazer `Quantidade` ocupar menos espaço que campos de texto.
- Ajustar um campo que ficou apertado em determinada largura de tela.

No modal `Colunas na Linha`, o campo pode ser configurado por faixa de largura da tela e por modo de formulário.

Modos disponíveis:

- `Novo`: quando o registro está sendo criado.
- `Ver`: quando o registro está sendo consultado.
- `Editar`: quando o registro está sendo alterado.

Valores de coluna disponíveis:

- `12`: ocupa a linha inteira.
- `8`: ocupa uma área maior que metade da linha.
- `6`: ocupa metade da linha.
- `4`: ocupa um terço da linha.
- `3`: ocupa um quarto da linha.
- `2`: ocupa uma área menor da linha.

Exemplo prático:

O campo `Observações` pode ser configurado com `12` no modo `Novo`, para facilitar a digitação, e também com `12` no modo `Ver`, para facilitar a leitura do texto completo.

Já campos curtos, como `Quantidade de dias`, podem usar `3` ou `4`, ocupando menos espaço na linha.

Uso recomendado:

- Use `Colunas` quando precisar ajustar um campo específico.
- Use `12` para campos longos, descrições e observações.
- Use `6` para pares de campos relacionados, como datas inicial e final.
- Use valores menores para campos curtos, como quantidade, código ou ramal.
- Revise o comportamento em telas menores para garantir boa leitura.

##### Botão Remover do campo

O botão `Remover` retira o campo daquela estrutura do formulário.

Ao remover um campo, ele deixa de aparecer naquela etapa e deixa de ser exibido para o usuário dentro do formulário.

Essa ação não significa excluir a coluna da lista do SharePoint. Ela apenas remove o campo da organização visual do formulário dinâmico.

Exemplos de uso:

- Remover um campo que foi adicionado por engano.
- Retirar da etapa um campo que não deve ser exibido ao usuário.
- Limpar campos que não fazem mais parte do processo.
- Reorganizar o formulário antes de adicionar o campo em outra etapa.

Campos obrigatórios da lista podem ter a remoção bloqueada.

Isso acontece porque campos obrigatórios precisam continuar presentes em alguma etapa do formulário para evitar problemas no preenchimento ou no salvamento do registro.

Uso recomendado:

- Antes de remover, confirme se o campo realmente não deve aparecer no formulário.
- Se o campo ainda for necessário, mova-o para outra etapa em vez de remover.
- Não remova campos obrigatórios sem revisar a regra de preenchimento do processo.
- Após remover campos, revise a etapa para garantir que a sequência de preenchimento continue clara.

Exemplo prático:

Se o campo `Centro de custo` foi colocado na etapa `Dados do solicitante`, mas não deve aparecer nessa parte do formulário, ele pode ser removido dali e depois incluído na etapa correta, como `Informações da solicitação`.

##### Botão Remover etapa

O botão `Remover etapa` exclui a etapa criada da estrutura do formulário.

Ele deve ser usado quando uma etapa não é mais necessária no fluxo de preenchimento.

Ao remover uma etapa, ela deixa de aparecer para o usuário e deixa de fazer parte da navegação do formulário.

Exemplos de uso:

- Remover uma etapa criada por engano.
- Excluir uma etapa que deixou de fazer parte do processo.
- Simplificar um formulário que ficou dividido em partes demais.
- Reorganizar o formulário após mover os campos para outras etapas.

Antes de remover:

- Verifique se a etapa ainda possui campos importantes.
- Mova os campos que devem continuar no formulário para outra etapa.
- Confirme se alertas ou banners daquela etapa ainda serão necessários.
- Revise se a remoção não quebra a sequência de preenchimento esperada pelo usuário.

Exemplo prático:

Se a etapa `Dados complementares` deixou de ser necessária, os campos que ainda forem úteis podem ser movidos para `Informações da solicitação`. Depois disso, a etapa `Dados complementares` pode ser removida.

Uso recomendado:

- Remova etapas vazias ou sem função clara.
- Evite remover etapas sem revisar os campos que estão dentro dela.
- Use a remoção para manter o formulário simples e objetivo.
- Após remover, revise a ordem das etapas restantes.

#### Adicionar alerta

A funcionalidade `Adicionar alerta` permite incluir uma mensagem de destaque dentro da estrutura do formulário.

O alerta serve para chamar a atenção do usuário sobre alguma informação importante durante o preenchimento.

Ele pode ser usado dentro de uma etapa específica ou em uma área fixa, dependendo de onde a mensagem precisa aparecer.

Exemplos de uso:

- Avisar que determinados campos são obrigatórios.
- Orientar o usuário antes de preencher uma etapa.
- Informar uma regra importante do processo.
- Destacar cuidados antes do envio do formulário.
- Comunicar uma restrição, prazo ou condição especial.

Exemplo prático:

Em um formulário de solicitação de férias, pode ser incluído um alerta informando que o pedido deve ser enviado com antecedência mínima definida pela empresa.

Uso recomendado:

- Use alertas para mensagens curtas e objetivas.
- Use alertas quando a informação impactar diretamente o preenchimento.
- Posicione o alerta próximo dos campos relacionados à orientação.
- Evite excesso de alertas para não deixar o formulário poluído visualmente.

##### Título

Campo usado para definir o título principal do alerta.

O título deve resumir rapidamente o motivo do aviso, para que o usuário entenda a importância da mensagem antes de ler o conteúdo completo.

Exemplos:

- `Atenção`
- `Prazo de solicitação`
- `Informação importante`
- `Dados obrigatórios`

Uso recomendado:

- Use títulos curtos.
- Use títulos claros e objetivos.
- Evite títulos muito longos, pois o detalhe deve ficar na mensagem.

##### Mensagem

Campo usado para escrever o conteúdo do alerta.

Essa mensagem explica ao usuário o que ele precisa saber, conferir ou fazer durante o preenchimento do formulário.

Exemplo:

`Solicitações de férias devem ser enviadas com no mínimo 30 dias de antecedência.`

Uso recomendado:

- Escreva a mensagem em linguagem simples.
- Informe exatamente o que o usuário precisa fazer.
- Evite textos longos demais.
- Use o alerta para orientar, não para substituir instruções completas do processo.

##### Campos no alerta

Campo usado para selecionar um ou mais campos que serão destacados dentro do alerta.

Essa opção ajuda quando o aviso está relacionado a campos específicos do formulário.

Exemplo:

Em um alerta sobre período de férias, podem ser selecionados os campos `Data de início` e `Data de término`.

Na prática, o usuário entende que aquele aviso está conectado diretamente aos campos indicados.

Uso recomendado:

- Selecione apenas os campos relacionados ao aviso.
- Use essa opção quando o alerta precisar reforçar atenção sobre campos importantes.
- Evite selecionar muitos campos para não perder o foco da mensagem.

##### Tipo

Campo usado para definir o estilo visual do alerta.

O tipo ajuda o usuário a entender a natureza da mensagem.

Opções disponíveis:

- `Informação`: usado para orientações gerais.
- `Sucesso`: usado para mensagens positivas ou confirmação.
- `Aviso`: usado para chamar atenção sobre uma regra, prazo ou cuidado.
- `Erro`: usado para indicar problema, bloqueio ou situação que exige correção.

Exemplos:

- Use `Informação` para orientar o preenchimento.
- Use `Aviso` para destacar uma regra importante.
- Use `Erro` quando a mensagem indicar que algo precisa ser corrigido.
- Use `Sucesso` para indicar uma condição positiva ou confirmação.

Uso recomendado:

- Use o tipo de acordo com a gravidade da mensagem.
- Evite usar `Erro` para mensagens apenas informativas.
- Use `Aviso` quando o usuário precisa prestar atenção antes de continuar.

##### Mostrar só quando a condição abaixo for verdadeira

Opção usada para exibir o alerta somente em determinadas situações.

Quando ativada, o alerta passa a depender de uma condição para aparecer no formulário.

Quando desativada, o alerta fica visível normalmente, conforme a posição definida.

Exemplo:

Um alerta pode aparecer somente quando o campo `Tipo de solicitação` for igual a `Urgente`.

Nesse caso, usuários que escolherem outro tipo de solicitação não verão o alerta.

Uso recomendado:

- Use essa opção quando o aviso não for necessário para todos os usuários.
- Use condições para evitar excesso de mensagens na tela.
- Mantenha alertas sempre visíveis quando a informação for geral e importante para todos.

##### Campo

Campo usado para escolher qual informação do formulário será avaliada na condição.

É o campo que o sistema irá observar para decidir se o alerta deve aparecer ou não.

Exemplo:

Selecionar o campo `Tipo de solicitação` para mostrar um alerta apenas quando o tipo escolhido exigir uma orientação específica.

##### Operador

Campo usado para definir como o valor será comparado.

O operador determina a regra da condição.

Exemplos de operadores:

- `é igual a`
- `é diferente de`
- `contém`
- `não contém`
- `maior que`
- `menor que`
- `está vazio`
- `não está vazio`
- `é verdadeiro`
- `é falso`

Exemplo prático:

Para exibir um alerta quando o valor de uma compra for maior que `10000`, selecione o campo de valor, use o operador `maior que` e informe `10000` no valor.

##### Comparar com

Campo usado para definir com o que a condição será comparada.

Opções disponíveis:

- `Texto fixo`: compara o campo com um valor digitado manualmente.
- `Outro campo`: compara o campo escolhido com outro campo do formulário.
- `Token`: compara o campo com uma informação dinâmica disponível no contexto.

Exemplos:

- Use `Texto fixo` para comparar com uma palavra, número ou status definido.
- Use `Outro campo` quando a regra depender da comparação entre dois campos.
- Use `Token` quando a regra depender de uma informação dinâmica do ambiente.

##### Valor

Campo usado para informar o valor da comparação.

Esse campo muda de importância conforme o operador escolhido.

Exemplo:

Se a condição for `Status é igual a Aprovado`, o valor informado será `Aprovado`.

Para operadores como `está vazio`, `não está vazio`, `é verdadeiro` e `é falso`, o campo de valor não precisa ser preenchido, pois a própria condição já define o comportamento esperado.

##### Ícone

Campo opcional usado para informar o nome de um ícone visual para o alerta.

O ícone ajuda a reforçar o tipo da mensagem e chamar atenção do usuário.

Exemplos de uso:

- Ícone de informação para orientações.
- Ícone de aviso para regras importantes.
- Ícone de erro para mensagens críticas.

Uso recomendado:

- Use ícones simples e coerentes com a mensagem.
- Não use ícone quando ele não agregar clareza ao alerta.
- Mantenha consistência visual entre alertas semelhantes.

##### Destacar visualmente

Opção usada para deixar o alerta com mais destaque na tela.

Quando ativada, o alerta ganha mais presença visual e chama mais atenção do usuário.

Quando desativada, o alerta aparece de forma mais discreta.

Uso recomendado:

- Ative para avisos importantes.
- Use com moderação para não fazer todos os alertas parecerem urgentes.
- Deixe desativado para mensagens simples de orientação.

##### Fechável

Opção usada para permitir que o usuário feche o alerta durante o uso do formulário.

Quando ativada, o usuário pode dispensar a mensagem após ler.

Quando desativada, o alerta permanece visível enquanto a condição de exibição for atendida ou enquanto estiver posicionado no formulário.

Uso recomendado:

- Ative quando o alerta for apenas informativo.
- Desative quando a mensagem precisa permanecer visível durante o preenchimento.
- Evite permitir fechamento em avisos críticos ou obrigatórios.

##### Posição no formulário

Campo usado para definir onde o alerta será exibido.

Opções disponíveis:

- `Na etapa (ordem com os campos)`: o alerta aparece junto com os campos, respeitando a ordem em que foi posicionado.
- `Fixo no topo (sticky)`: o alerta fica no topo e pode acompanhar a rolagem da tela.
- `Fixo em baixo (sticky)`: o alerta fica na parte inferior e pode acompanhar a rolagem da tela.

Uso recomendado:

- Use `Na etapa` quando o alerta estiver relacionado a campos daquela etapa.
- Use `Fixo no topo` quando o aviso precisar ficar visível durante o preenchimento.
- Use `Fixo em baixo` quando o alerta funcionar como lembrete próximo da navegação ou envio.

##### Zona fixa

Campo exibido quando o alerta está configurado para posição fixa.

Ele define se o alerta ficará na parte superior ou inferior da área do formulário.

Opções disponíveis:

- `Fixo no topo`
- `Fixo em baixo`

Uso recomendado:

- Use topo para instruções que precisam ser vistas antes do preenchimento.
- Use embaixo para lembretes, reforços ou mensagens próximas das ações finais.

##### Posicionamento

Campo exibido quando o alerta está configurado como fixo.

Ele define como o alerta se comporta visualmente em relação ao conteúdo.

Opções disponíveis:

- `Fixo (acompanha ao scroll)`: o alerta permanece visível enquanto o usuário rola a tela.
- `Absoluto (sobre o conteúdo)`: o alerta fica sobreposto ao conteúdo.
- `No espaço (fluxo normal)`: o alerta ocupa seu espaço normal na página, sem sobrepor os campos.

Uso recomendado:

- Use `Fixo (acompanha ao scroll)` quando o aviso precisa continuar visível.
- Use `No espaço (fluxo normal)` quando quiser evitar sobreposição.
- Use `Absoluto (sobre o conteúdo)` apenas quando houver necessidade visual específica, pois pode cobrir parte do formulário.

#### Adicionar banner

A funcionalidade `Adicionar banner` permite incluir uma área visual de destaque no formulário.

O banner pode ser usado para apresentar uma mensagem institucional, orientação geral, identificação da etapa ou comunicação visual mais destacada.

Diferente do alerta, que normalmente é usado para avisos pontuais, o banner pode ter um papel mais visual e informativo na experiência do formulário.

Exemplos de uso:

- Exibir uma mensagem de boas-vindas no início do formulário.
- Identificar o objetivo do formulário.
- Apresentar uma orientação geral antes do preenchimento.
- Destacar uma campanha interna ou comunicado da empresa.
- Separar visualmente uma parte importante do formulário.

Exemplo prático:

Em um formulário de férias, o banner pode apresentar a mensagem `Solicitação de Férias`, junto com uma orientação breve sobre como preencher corretamente as informações.

Uso recomendado:

- Use banners para informações gerais ou de apresentação.
- Use banners no início do formulário ou de uma etapa importante.
- Mantenha o texto claro e direto.
- Evite usar muitos banners em sequência para não dificultar a leitura dos campos.

##### URL da imagem

Campo usado para informar o endereço da imagem que será exibida no banner.

Essa imagem pode representar uma identificação visual do formulário, uma comunicação interna ou uma orientação para o usuário.

Exemplo de uso:

Em um formulário de benefícios, o banner pode usar uma imagem institucional relacionada à campanha ou ao tipo de solicitação.

Uso recomendado:

- Use imagens claras e alinhadas ao objetivo do formulário.
- Prefira imagens com boa qualidade e tamanho adequado para web.
- Evite imagens com textos muito pequenos, pois podem ficar difíceis de ler em telas menores.

##### Largura

O campo `Largura (%)` define quanto da largura disponível o banner deve ocupar.

O valor é informado em porcentagem, de `1` a `100`.

Exemplos:

- `100`: o banner ocupa toda a largura disponível da etapa ou área onde foi inserido.
- `80`: o banner ocupa 80% da largura disponível.
- `50`: o banner ocupa metade da largura disponível.

Uso recomendado:

- Use `100` quando o banner for usado como cabeçalho visual da etapa ou do formulário.
- Use valores menores quando o banner for apenas um apoio visual.
- Evite larguras muito pequenas quando a imagem tiver texto ou detalhes importantes.

##### Altura

O campo `Altura (px)` define a altura do banner em pixels.

Essa configuração ajuda a controlar o espaço vertical ocupado pela imagem no formulário.

Exemplos:

- `120`: banner mais baixo, útil para faixas simples.
- `240`: banner médio, adequado para destaque visual.
- `400`: banner maior, indicado apenas quando a imagem precisa ter mais presença.

Uso recomendado:

- Use alturas menores para banners informativos simples.
- Use alturas médias quando o banner for parte importante da apresentação.
- Evite alturas muito grandes para não empurrar os campos principais para baixo.

##### Posição no formulário

O campo `Posição no formulário` define onde o banner será exibido em relação à etapa e ao formulário.

Opções disponíveis:

- `Na etapa (ordem com os campos)`: o banner aparece junto com os campos, respeitando a ordem em que foi posicionado.
- `Fixo no topo (sticky)`: o banner fica no topo e pode acompanhar a rolagem da tela.
- `Fixo em baixo (sticky)`: o banner fica na parte inferior e pode acompanhar a rolagem da tela.

Exemplo de uso:

Um banner com instruções gerais pode ficar no topo da etapa. Já um banner com lembrete de envio pode ficar fixo embaixo, próximo da área final de ação.

Uso recomendado:

- Use `Na etapa` quando o banner fizer parte do conteúdo normal do formulário.
- Use `Fixo no topo` quando a informação precisar continuar visível durante o preenchimento.
- Use `Fixo em baixo` quando o banner servir como lembrete ou reforço durante a navegação.

##### Posicionamento

O campo `Posicionamento` aparece quando o banner é configurado como fixo.

Ele define como o banner se comporta visualmente dentro da página.

Opções disponíveis:

- `Fixo (acompanha ao scroll)`: o banner permanece visível enquanto o usuário rola a tela.
- `Absoluto (sobre o conteúdo)`: o banner fica sobreposto em relação ao conteúdo.
- `No espaço (fluxo normal)`: o banner ocupa seu espaço normal na página, sem sobrepor o conteúdo.

Uso recomendado:

- Use `Fixo (acompanha ao scroll)` quando o usuário precisa ver a mensagem durante todo o preenchimento.
- Use `No espaço (fluxo normal)` quando quiser evitar sobreposição com campos.
- Use `Absoluto (sobre o conteúdo)` apenas quando houver necessidade visual específica, pois pode cobrir parte do formulário se não for bem ajustado.

#### Etapa Fixos

A etapa `Fixos` é uma etapa especial usada para manter campos, alertas ou banners em uma posição fixa no formulário.

Ela aparece sempre logo após a etapa `Ocultos` e antes das etapas criadas pelo usuário.

Essa etapa não entra como uma etapa comum do fluxo de navegação. Ela serve para exibir informações de apoio no topo ou no rodapé do formulário, mantendo esses elementos disponíveis durante o preenchimento.

Exemplos de uso:

- Exibir um alerta com instruções importantes antes dos campos principais.
- Mostrar um banner institucional ou informativo no início do formulário.
- Manter uma mensagem fixa no rodapé com orientações sobre o envio.
- Destacar avisos que precisam continuar visíveis durante o preenchimento.

##### Título da etapa

O campo `Título da etapa (fixos)` permite alterar o nome exibido para a etapa.

Esse título ajuda o administrador a identificar a função da etapa durante a configuração do formulário.

##### Configurar

O botão `Configurar` abre as opções de visibilidade da etapa.

Nessa configuração é possível definir em quais modos a etapa será exibida, como criação, edição ou visualização.

Também é possível aplicar uma condição para controlar quando os elementos da etapa devem aparecer.

Exemplo de uso:

Um aviso pode ser exibido somente no modo de criação, enquanto uma orientação de consulta pode aparecer apenas no modo de visualização.

##### Configurar Colunas

O botão `Configurar Colunas` permite ajustar a organização dos campos dentro da etapa.

Essa opção é útil quando os campos fixos precisam ser exibidos em mais de uma coluna ou quando a disposição visual precisa ser alinhada com o restante do formulário.

##### Incluir em Fixos

A área `Incluir em Fixos` permite adicionar campos, alertas e banners à etapa.

Os itens adicionados nessa área podem ser posicionados no topo ou no rodapé do formulário, dependendo da configuração de cada componente.

Para alertas e banners, a configuração pode incluir a zona fixa e o tipo de posicionamento.

Exemplos de posicionamento:

- `Topo`: exibe o item acima do conteúdo principal do formulário.
- `Rodapé`: exibe o item abaixo do conteúdo principal do formulário.
- `Sticky`: mantém o item visível enquanto o usuário rola a tela.
- `Absoluto`: posiciona o item em relação ao contêiner do formulário.
- `Fluxo normal`: exibe o item seguindo a ordem natural do conteúdo.

Uso recomendado:

- Use `Fixos` para informações de apoio que não pertencem a uma etapa específica do fluxo.
- Use alertas fixos quando o usuário precisar lembrar de uma regra durante todo o preenchimento.
- Use banners fixos para comunicação visual, instruções gerais ou identificação do formulário.
- Evite colocar muitos campos nessa etapa para não ocupar espaço excessivo da tela.

### Regras dos campos

A aba `Regras dos campos` é usada para acessar e organizar as configurações de comportamento de cada campo do formulário.

Nessa área, o usuário encontra a lista de campos disponíveis para configurar regras de uso, como exibição, edição, obrigatoriedade e outros comportamentos específicos.

Antes de abrir a configuração de um campo, é possível organizar a lista para facilitar a localização do campo desejado.

#### Filtrar campo

O campo `Filtrar campo` permite digitar parte do nome de um campo para localizar rapidamente o item desejado na lista.

Esse filtro é útil quando o formulário possui muitos campos e o usuário não quer procurar manualmente pela lista completa.

A busca não diferencia letras maiúsculas e minúsculas.

Exemplo:

Digitar `data`, `Data` ou `DATA` pode localizar campos como `Data de início`, `Data de término` ou `Data de retorno`.

A busca também considera o nome interno do campo, quando aplicável.

Exemplo:

Se um campo aparece para o usuário como `Centro de custo`, mas possui um nome interno relacionado, a busca pode ajudar a localizá-lo por qualquer uma dessas referências.

Além disso, a busca ignora acentos.

Exemplo:

Digitar `area` pode localizar o campo `Área`.

Quando nenhum campo corresponde ao texto digitado, a lista informa que nenhum campo foi encontrado.

Uso recomendado:

- Use o filtro para localizar rapidamente um campo específico.
- Digite apenas uma parte do nome quando não souber o nome completo.
- Use sem se preocupar com maiúsculas, minúsculas ou acentos.
- Limpe o filtro para voltar a visualizar todos os campos.

#### Botão Crescente

O botão `Crescente (A–Z)` ordena os campos em ordem alfabética crescente.

Essa opção organiza a lista começando pelos campos com nomes mais próximos do início do alfabeto.

Exemplo:

- `Área`
- `Centro de custo`
- `Data de início`
- `Nome do colaborador`

Uso recomendado:

- Use quando quiser localizar um campo pelo nome.
- Use quando a lista tiver muitos campos e precisar de uma ordem mais simples.
- Use como visualização padrão para facilitar a procura.

#### Botão Decrescente

O botão `Decrescente (Z–A)` ordena os campos em ordem alfabética decrescente.

Essa opção organiza a lista começando pelos campos com nomes mais próximos do final do alfabeto.

Exemplo:

- `Status`
- `Solicitante`
- `Período`
- `Centro de custo`

Uso recomendado:

- Use quando o campo procurado estiver mais próximo do final da lista alfabética.
- Use para alternar rapidamente a forma de visualização.
- Use quando a equipe preferir revisar os campos na ordem inversa.

#### Botão Tipo de dado

O botão `Tipo de dado (agrupa por tipo)` organiza os campos conforme o tipo de informação que cada campo armazena.

Em vez de ordenar apenas pelo nome, essa opção agrupa campos semelhantes.

Exemplos de tipos de dados:

- Texto.
- Número.
- Data.
- Sim ou não.
- Pessoa.
- Escolha.
- Lookup.

Exemplo prático:

Campos de data, como `Data de início`, `Data de término` e `Data de retorno`, podem aparecer próximos entre si.

Campos de texto, como `Nome`, `Descrição` e `Observações`, também podem ficar agrupados.

Uso recomendado:

- Use quando quiser configurar regras em campos do mesmo tipo.
- Use para revisar todos os campos de data de uma vez.
- Use para identificar rapidamente campos numéricos, campos de texto ou campos de escolha.
- Use quando a regra que será aplicada depende do tipo de informação do campo.

#### Botão Regras

O botão `Regras` abre o painel lateral `Configurar regras` do campo selecionado.

Esse painel é usado para configurar o comportamento individual de cada campo no formulário.

As opções exibidas podem mudar conforme o tipo de dado do campo.

Exemplo:

Um campo de texto possui opções de validação, transformação e máscara.

Um campo de data possui opções de limite de data, comparação com outros campos e mensagem de erro.

Um campo de lookup possui opções para controlar o texto exibido, detalhes adicionais e filtros das opções.

Uso recomendado:

- Abra `Regras` quando precisar ajustar o comportamento de um campo específico.
- Configure primeiro os campos mais importantes do processo.
- Revise o comportamento em `Novo`, `Editar` e `Ver`.
- Use regras apenas quando houver uma necessidade clara para o usuário final.

#### Ações comuns do painel de regras

Algumas ações aparecem no painel de regras independentemente do tipo de campo.

##### Aplicar

O botão `Aplicar` salva as regras configuradas para o campo atual.

Use esse botão depois de concluir os ajustes do campo.

##### Cancelar

O botão `Cancelar` fecha o painel sem aplicar as alterações feitas naquele momento.

Use essa opção quando abrir o painel apenas para consulta ou quando não quiser manter as mudanças realizadas.

##### Pré-visualização

A `Pré-visualização` informa quantas regras serão geradas para aquele campo com base nas configurações feitas.

Ela ajuda a confirmar se as opções selecionadas realmente estão criando regras para o formulário.

Exemplo:

Se um campo foi marcado como obrigatório e também recebeu uma validação de tamanho mínimo, a pré-visualização pode indicar que mais de uma regra será gerada.

### Regras por tipo de campo

Cada tipo de campo pode abrir seções diferentes dentro do painel de regras.

As seções abaixo explicam o que aparece para cada tipo de dado e como usar cada grupo de configuração.

#### Text

Campos do tipo `text` são usados para textos curtos, como nome, matrícula, código, e-mail ou telefone.

Ao abrir `Regras` em um campo de texto, o painel pode exibir as seções abaixo.

##### Exibição

Controla como o campo aparece no formulário.

Principais opções:

- `Mostrar em`: define se o campo aparece em `Novo`, `Editar` e `Ver`.
- `Placeholder`: texto exibido dentro do campo antes do preenchimento.
- `Texto de ajuda`: orientação exibida para auxiliar o usuário.
- `Valor padrão`: valor inicial preenchido automaticamente quando o campo estiver vazio.
- `Expressão`: permite calcular ou montar um valor para o campo.
- `Somente leitura`: mostra o campo sem permitir alteração.
- `Ocultar no formulário`: esconde o campo da tela.

Exemplo:

Um campo `E-mail` pode ter placeholder `nome@empresa.com` e texto de ajuda `Informe o e-mail corporativo`.

##### Validação

Define regras para verificar se o valor digitado está correto.

Principais opções:

- `Modelo: data não no passado`: aplica um modelo de regra para impedir datas anteriores, quando fizer sentido para o campo.
- `Modelo: validar e-mail`: aplica um modelo de validação para e-mail.
- `Obrigatório`: exige o preenchimento do campo.
- `Mín. caracteres`: define o tamanho mínimo do texto.
- `Máx. caracteres`: define o tamanho máximo do texto.
- `Regex`: define um padrão específico de preenchimento.
- `Mensagem se falhar o padrão`: texto exibido quando o valor não atende ao padrão.

Exemplo:

Um campo `CPF` pode exigir uma quantidade específica de caracteres e uma mensagem orientando o formato correto.

##### Transformação

Define se o texto digitado será transformado automaticamente.

Opções disponíveis:

- `Maiúsculas`: transforma o texto em letras maiúsculas.
- `Minúsculas`: transforma o texto em letras minúsculas.
- `Capitalizar`: deixa as palavras com início em maiúscula.

Exemplo:

Um campo `Nome completo` pode usar `Capitalizar` para padronizar a apresentação do nome.

##### Máscaras

Define uma máscara de preenchimento para orientar a digitação.

Opções disponíveis:

- `Nenhuma`: não aplica máscara.
- `CPF`: aplica formato de CPF.
- `Telefone (BR)`: aplica formato de telefone brasileiro.
- `CEP`: aplica formato de CEP.
- `CNPJ`: aplica formato de CNPJ.
- `Personalizada`: permite informar um padrão específico.

Exemplo:

Um campo `Telefone` pode usar a máscara `Telefone (BR)` para facilitar o preenchimento.

##### Desativar / ativar o campo

Controla quando o campo deve ficar bloqueado ou editável.

Principais opções:

- `Desativar este campo quando a condição for verdadeira`: bloqueia o campo conforme uma regra.
- `Tornar editável quando a condição for verdadeira`: libera edição quando uma regra for atendida.
- `Campo`: define qual campo será usado na condição.
- `Operador`: define como a comparação será feita.
- `Comparar`: define se a comparação será com texto fixo, outro campo, token ou grupo do SharePoint.
- `Valor`: define o valor usado na comparação.
- `Grupos do SharePoint`: permite escolher grupos para regras baseadas em permissão ou perfil.

Exemplo:

O campo `Justificativa` pode ficar desativado até que o usuário selecione `Sim` no campo `Precisa justificar`.

##### Condicionais

Permite criar grupos de regras para controlar o comportamento do campo conforme condições.

Principais opções:

- `Adicionar grupo de regra`: cria um novo grupo de condição.
- `Aplicar esta regra apenas nos modos`: define se a regra vale para `Criar`, `Editar` ou `Ver`.
- `Incluir grupos SharePoint`: aplica a regra somente para usuários de grupos selecionados.
- `Excluir grupos SharePoint`: impede que a regra seja aplicada para usuários de grupos selecionados.
- `Operador lógico entre condições`: define se todas as condições precisam ser atendidas ou se basta uma.
- `Condições`: define campo, operador, comparação e valor.
- `Ação quando as condições se verificam`: define o que acontece quando a regra é atendida.

Exemplo:

O campo `Motivo da urgência` pode aparecer somente quando `Prioridade` for igual a `Alta`.

#### Multiline

Campos do tipo `multiline` são usados para textos maiores, como descrição, observações, justificativas ou comentários.

Eles possuem as mesmas seções principais do tipo `text`, com um ajuste adicional na exibição.

##### Exibição

Além das opções comuns de exibição, o campo `multiline` possui a opção `Linhas do textarea`.

Essa opção define a altura inicial do campo de texto longo.

Exemplo:

Um campo `Observações` pode ter `5` linhas para dar mais espaço de digitação ao usuário.

##### Validação

Funciona como no campo `text`.

Pode exigir preenchimento, tamanho mínimo, tamanho máximo, padrão e mensagem de erro.

Exemplo:

Um campo `Justificativa` pode ser obrigatório e exigir pelo menos `20` caracteres.

##### Transformação

Permite aplicar maiúsculas, minúsculas ou capitalização ao texto.

Use com cuidado em textos longos, pois pode alterar a forma como o usuário escreveu a descrição.

##### Máscaras

Permite aplicar máscara, embora normalmente seja mais usada em campos de texto curto.

Para textos longos, geralmente a opção `Nenhuma` é a mais indicada.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

Exemplo:

O campo `Comentário do gestor` pode ficar editável apenas para usuários de um grupo específico.

##### Condicionais

Permite mostrar, ocultar ou ajustar o comportamento do campo conforme respostas de outros campos, modos do formulário ou grupos do SharePoint.

#### Choice

Campos do tipo `choice` são usados quando o usuário escolhe uma opção em uma lista de valores.

Exemplos:

- Status.
- Prioridade.
- Tipo de solicitação.
- Categoria.

##### Exibição

Permite definir em quais modos o campo aparece, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Prioridade` pode aparecer em `Novo` e `Editar`, mas ficar apenas visível em `Ver`.

##### Desativar / ativar o campo

Permite bloquear ou liberar a escolha conforme uma condição.

Exemplo:

O campo `Status` pode ficar bloqueado para usuários comuns e editável apenas para gestores.

##### Observação sobre condições entre colunas

Condições mais específicas entre colunas podem depender da configuração JSON do gestor.

No painel de regras, o foco fica nas opções principais de exibição, valor padrão, bloqueio e validações aplicáveis.

#### Multichoice

Campos do tipo `multichoice` permitem selecionar mais de uma opção.

Exemplos:

- Benefícios desejados.
- Áreas envolvidas.
- Tipos de documento.

##### Exibição

Controla quando o campo aparece, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Áreas envolvidas` pode ter texto de ajuda informando que mais de uma opção pode ser selecionada.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme outro valor do formulário ou grupo do SharePoint.

Exemplo:

O campo `Benefícios adicionais` pode ficar editável somente quando `Tipo de contratação` for igual a `CLT`.

##### Observação sobre condições entre colunas

Assim como em `choice`, regras mais específicas entre colunas podem depender da configuração JSON do gestor.

#### Number

Campos do tipo `number` armazenam valores numéricos.

Exemplos:

- Quantidade.
- Dias.
- Percentual.
- Pontuação.

##### Exibição

Controla modos de exibição, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Quantidade de dias` pode ter valor padrão `1` e texto de ajuda orientando o preenchimento.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

Exemplo:

O campo `Quantidade` pode ficar bloqueado quando o status for `Finalizado`.

##### Validação numérica

Define limites mínimos e máximos para o valor.

Principais opções:

- `Mínimo`: menor valor permitido.
- `Máximo`: maior valor permitido.

Exemplo:

O campo `Quantidade de dias` pode aceitar no mínimo `1` e no máximo `30`.

#### Currency

Campos do tipo `currency` armazenam valores monetários.

Exemplos:

- Valor da compra.
- Orçamento.
- Reembolso.
- Custo estimado.

##### Exibição

Controla modos de exibição, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Valor solicitado` pode ter texto de ajuda informando que o valor deve ser preenchido em reais.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

Exemplo:

O campo `Valor aprovado` pode ficar editável apenas para o grupo financeiro.

##### Validação numérica

Define valor mínimo e máximo aceito.

Exemplo:

Um campo `Reembolso` pode aceitar valores entre `0` e `5000`.

#### Boolean

Campos do tipo `boolean` representam respostas de sim ou não, verdadeiro ou falso.

Exemplos:

- Requer aprovação.
- Possui anexo.
- Solicitação urgente.

##### Exibição

Controla quando o campo aparece, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Para valor padrão, use valores como `true` ou `false`.

Exemplo:

O campo `Solicitação urgente` pode iniciar como `false`.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

Exemplo:

O campo `Aprovado` pode ficar editável somente para usuários do grupo `Gestores`.

##### Observação sobre visibilidade condicional

Para campos booleanos, a visibilidade condicional pode ser configurada pelo painel ou por JSON do gestor, conforme a necessidade da regra.

#### Datetime

Campos do tipo `datetime` armazenam datas ou datas com horário.

Exemplos:

- Data de início.
- Data de término.
- Prazo.
- Data de retorno.

##### Exibição

Controla modos de formulário, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Data de início` pode ter valor padrão baseado na data atual.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

Exemplo:

O campo `Data de aprovação` pode ficar editável apenas quando o status for `Em análise`.

##### Limites em relação a hoje

Define limites de data com base na data atual.

Principais opções:

- `Mín. dias a partir de hoje`: menor data permitida em relação ao dia atual.
- `Máx. dias a partir de hoje`: maior data permitida em relação ao dia atual.
- `Bloquear fins de semana`: impede seleção de sábados e domingos.
- Dias específicos da semana: permite bloquear segunda, terça, quarta, quinta, sexta, sábado ou domingo.

Exemplo:

Um formulário de férias pode exigir que a data de início seja pelo menos `30` dias depois da data atual.

##### Comparação com outros campos

Permite comparar a data com outro campo do formulário.

Principais opções:

- `Data >= campo`: exige que a data seja maior ou igual à data de outro campo.
- `Data <= campo`: exige que a data seja menor ou igual à data de outro campo.

Exemplo:

`Data de término` pode precisar ser maior ou igual a `Data de início`.

##### Mensagem de erro

Define o texto exibido quando a data informada não atende às regras.

Exemplo:

`A data de término deve ser maior ou igual à data de início.`

#### Url

Campos do tipo `url` armazenam endereços de link.

Exemplos:

- Link do documento.
- Site de referência.
- Endereço de evidência.

##### Exibição

Controla modos de exibição, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Link do documento` pode ter placeholder `https://`.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

Exemplo:

O link pode ficar obrigatório ou editável apenas quando o tipo da solicitação exigir evidência externa.

##### Validação de texto

Permite validar o conteúdo informado como texto.

Principais opções:

- `Mín. caracteres`.
- `Máx. caracteres`.
- `Regex`.
- `Mensagem se falhar o padrão`.

Exemplo:

É possível usar um padrão para orientar que o link comece com `https://`.

#### Lookup

Campos do tipo `lookup` exibem opções vindas de outra lista.

Exemplos:

- Centro de custo.
- Projeto.
- Departamento.
- Cliente.

##### Exibição

Controla modos de formulário, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Projeto` pode ter texto de ajuda orientando o usuário a selecionar o projeto correto.

##### Desativar / ativar o campo

Permite bloquear ou liberar o lookup conforme condições.

Exemplo:

O campo `Centro de custo` pode ficar bloqueado até que a área seja selecionada.

##### Lista ligada (texto das opções)

Define qual campo da lista ligada será usado como texto principal das opções.

Principais opções:

- `Campo para o texto das opções`: escolhe qual informação será exibida na lista.
- `Propriedade a exibir`: define uma propriedade específica quando o campo escolhido também for usuário ou lookup.

Exemplo:

Em vez de exibir apenas o título padrão do item, o lookup pode exibir o nome do projeto ou outro campo mais claro para o usuário.

##### Detalhe abaixo da seleção

Permite mostrar informações complementares abaixo da opção selecionada.

Exemplo:

Ao selecionar um `Projeto`, o formulário pode exibir abaixo o `Código`, a `Área` ou o `Responsável`.

##### Filtrar opções

Permite filtrar as opções do lookup com base em outro campo do formulário.

Principais opções:

- `Campo pai`: campo do formulário usado como referência.
- `Comparador`: regra usada para comparar os valores.
- `Campo na lista filho`: campo da lista ligada que será comparado com o campo pai.

Exemplo:

Depois que o usuário seleciona uma `Área`, o campo `Projeto` pode mostrar apenas projetos daquela área.

#### Lookupmulti

Campos do tipo `lookupmulti` permitem selecionar vários itens de outra lista.

Exemplos:

- Projetos relacionados.
- Documentos vinculados.
- Áreas envolvidas.

##### Exibição

Controla modos de formulário, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

##### Desativar / ativar o campo

Permite bloquear ou liberar a seleção múltipla conforme condições.

##### Lista ligada (texto das opções)

Funciona como no campo `lookup`, definindo qual texto será exibido para cada opção.

Quando o campo de origem possui múltiplos valores, os valores podem aparecer concatenados para formar o texto da opção.

##### Detalhe abaixo da seleção

Permite mostrar informações extras relacionadas aos itens selecionados.

##### Filtrar opções

Permite limitar as opções disponíveis com base em outro campo do formulário.

Exemplo:

O usuário seleciona uma unidade e o campo passa a listar apenas documentos vinculados àquela unidade.

#### User

Campos do tipo `user` são usados para selecionar uma pessoa ou grupo.

Exemplos:

- Solicitante.
- Gestor responsável.
- Aprovador.

##### Exibição

Controla modos de formulário, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

Exemplo:

O campo `Gestor responsável` pode ter texto de ajuda orientando a selecionar o gestor imediato.

##### Desativar / ativar o campo

Permite bloquear ou liberar a seleção da pessoa conforme condições.

Exemplo:

O campo `Aprovador` pode ficar editável apenas para usuários do grupo `Administradores`.

#### Usermulti

Campos do tipo `usermulti` permitem selecionar mais de uma pessoa ou grupo.

Exemplos:

- Participantes.
- Responsáveis.
- Equipe envolvida.

##### Exibição

Controla modos de formulário, placeholder, texto de ajuda, valor padrão, expressão, somente leitura e ocultação.

##### Desativar / ativar o campo

Permite bloquear ou liberar a seleção múltipla conforme condições.

Exemplo:

O campo `Equipe envolvida` pode ficar editável somente quando o tipo de solicitação for `Projeto`.

#### Calculated

Campos do tipo `calculated` são campos calculados pelo SharePoint.

Normalmente, esses campos são usados para exibir valores derivados de outras informações.

Exemplos:

- Total calculado.
- Prazo calculado.
- Situação calculada.

##### Exibição

Controla em quais modos o campo aparece e permite incluir texto de ajuda.

Como o valor é calculado, a edição direta costuma ser limitada.

##### Desativar / ativar o campo

Quando disponível, pode ser usada para controlar se o campo aparece bloqueado conforme condições.

Uso recomendado:

- Use campos calculados principalmente para consulta.
- Evite tratá-los como campos de preenchimento manual.
- Explique o significado do cálculo no texto de ajuda quando necessário.

#### Taxonomy

Campos do tipo `taxonomy` representam termos de metadados gerenciados.

Exemplos:

- Categoria corporativa.
- Classificação documental.
- Área de conhecimento.

##### Exibição

Controla modos de formulário, texto de ajuda, valor padrão quando aplicável, somente leitura e ocultação.

Exemplo:

O campo `Classificação documental` pode aparecer em `Novo` e `Editar`, mas ficar somente leitura em `Ver`.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

Exemplo:

Uma classificação pode ficar bloqueada depois que o documento for aprovado.

#### Taxonomymulti

Campos do tipo `taxonomymulti` permitem selecionar mais de um termo de metadados gerenciados.

Exemplos:

- Palavras-chave.
- Categorias múltiplas.
- Temas relacionados.

##### Exibição

Controla modos de formulário, texto de ajuda, valor padrão quando aplicável, somente leitura e ocultação.

##### Desativar / ativar o campo

Permite bloquear ou liberar a seleção dos termos conforme condições.

Exemplo:

O campo `Temas relacionados` pode ficar editável apenas durante a criação do registro.

#### Unknown

Campos do tipo `unknown` são campos cujo tipo não foi identificado de forma específica pela solução.

Nesses casos, o painel apresenta opções mais genéricas.

##### Exibição

Controla modos de formulário, placeholder, texto de ajuda, valor padrão, somente leitura e ocultação.

##### Desativar / ativar o campo

Permite bloquear ou liberar o campo conforme condições.

##### Validação de texto

Permite aplicar validações básicas como tamanho mínimo, tamanho máximo, padrão e mensagem de erro.

Uso recomendado:

- Use regras simples em campos `unknown`.
- Valide o comportamento no formulário antes de publicar.
- Quando possível, confirme o tipo correto do campo na lista SharePoint.

### Expressões nas regras dos campos

As expressões são usadas para preencher, calcular ou montar valores automaticamente dentro do formulário.

Elas aparecem principalmente nas configurações de `Valor padrão` e `Expressão` dentro da seção `Exibição` das regras de um campo.

#### Onde as expressões podem ser usadas

As expressões podem ser usadas em diferentes situações:

- Definir um valor inicial para um campo.
- Preencher um campo com informações do usuário atual.
- Preencher uma data automaticamente.
- Calcular valores numéricos.
- Montar textos usando valores de outros campos.
- Buscar informações de campos lookup ou pessoa.
- Definir uma pasta de anexos quando a configuração de anexos usa biblioteca.

#### Diferença entre Valor padrão e Expressão

##### Valor padrão

O `Valor padrão` é aplicado quando o campo ainda está vazio.

Ele serve para sugerir ou preencher automaticamente um valor inicial, mas o usuário ainda pode alterar o campo se ele estiver editável.

Exemplo:

Um campo `Solicitante` pode iniciar com o usuário atual.

Um campo `Data da solicitação` pode iniciar com a data de hoje.

Uso recomendado:

- Use para facilitar o preenchimento.
- Use para reduzir digitação manual.
- Use quando o valor inicial pode ser alterado pelo usuário.

##### Expressão

A `Expressão` é usada para calcular ou montar o valor do campo.

Ela é indicada quando o valor depende de outro campo, de uma regra ou de informações do contexto.

Exemplo:

Um campo `Total` pode ser calculado com base em `Quantidade` e `Valor unitário`.

Um campo `Resumo` pode montar um texto juntando o nome do solicitante, o tipo da solicitação e a data.

Uso recomendado:

- Use quando o campo deve ser preenchido por regra.
- Use quando o valor depende de outros campos.
- Use quando a informação precisa seguir um padrão.

#### Sempre expressão ao vivo

A opção `Sempre expressão ao vivo` faz a expressão ser recalculada mesmo em edição e visualização.

Quando essa opção está ativa, o campo pode ignorar o valor gravado anteriormente e exibir o resultado atualizado da expressão.

Exemplo:

Um campo `Dias restantes` pode ser recalculado sempre que o formulário for aberto, considerando a data atual.

Uso recomendado:

- Use quando o valor precisa refletir o momento atual.
- Use para cálculos que dependem de datas, contexto ou outros campos atualizados.
- Evite usar quando o valor precisa preservar exatamente o que foi gravado no momento do envio.

#### Referência a outros campos

Para usar o valor de outro campo, informe o nome interno entre chaves duplas.

Formato:

`{{NomeInternoDoCampo}}`

Exemplo:

`{{Quantidade}}`

`{{ValorUnitario}}`

Uso em uma expressão numérica:

`{{Quantidade}} * {{ValorUnitario}}`

Resultado esperado:

Se `Quantidade` for `3` e `ValorUnitario` for `100`, o resultado será `300`.

Uso recomendado:

- Use nomes internos dos campos.
- Use campos compatíveis com o tipo de cálculo.
- Confirme o nome interno quando o campo tiver acentos ou espaços no nome exibido.

#### Expressões numéricas

Expressões numéricas são usadas em campos de número ou moeda.

Elas permitem usar operações matemáticas simples.

Operadores comuns:

- `+`: soma.
- `-`: subtração.
- `*`: multiplicação.
- `/`: divisão.
- `(` e `)`: agrupamento.

Exemplos:

`{{Quantidade}} * {{ValorUnitario}}`

`{{ValorTotal}} / 2`

`({{ValorA}} + {{ValorB}}) / 2`

Uso recomendado:

- Use em campos de número ou moeda.
- Use para totais, médias, diferenças e cálculos simples.
- Evite fórmulas muito longas quando a regra puder ser dividida em campos auxiliares.

#### Expressões de texto

Para montar texto usando campos, use o prefixo `str:`.

Formato:

`str: texto {{Campo}}`

Exemplo:

`str: Solicitação de {{Solicitante}} para {{Departamento}}`

Resultado esperado:

O formulário monta um texto combinando partes fixas com valores preenchidos em outros campos.

Uso recomendado:

- Use para gerar descrições automáticas.
- Use para montar títulos padronizados.
- Use para criar resumos de solicitação.

Exemplo prático:

`str: Férias de {{Colaborador}} - {{DataInicio}} até {{DataFim}}`

#### Tokens de usuário

Tokens de usuário usam informações do usuário atual.

Tokens disponíveis:

- `[me]`: ID numérico do usuário atual.
- `[myId]`: igual a `[me]`.
- `[myName]`: nome do usuário atual.
- `[myEmail]`: e-mail do usuário atual.
- `[myLogin]`: login do usuário atual.
- `[myDepartment]`: departamento do usuário, quando disponível.
- `[myJobTitle]`: cargo do usuário, quando disponível.

Exemplos:

`[myName]`

`[myEmail]`

`str: Solicitação aberta por [myName]`

Uso recomendado:

- Use `[myName]` para preencher nome do solicitante.
- Use `[myEmail]` para registrar o e-mail do usuário.
- Use `[me]` ou `[myId]` em campos de pessoa ou lookup que esperam identificador numérico.

#### Tokens de data

Tokens de data usam datas relativas ao momento atual.

Tokens disponíveis:

- `[today]`: data de hoje.
- `[now]`: data e hora atuais.
- `[tomorrow]`: dia seguinte.
- `[yesterday]`: dia anterior.
- `[startOfMonth]`: primeiro dia do mês atual.
- `[endOfMonth]`: último dia do mês atual.
- `[startOfYear]`: primeiro dia do ano atual.
- `[endOfYear]`: último dia do ano atual.

Exemplos:

`[today]`

`[today] + 7`

`{{DataInicio}} + 30`

Uso recomendado:

- Use `[today]` para campos de data da solicitação.
- Use `[now]` quando precisar de data e hora.
- Use sufixos como `+ 7`, `+ 14` ou `+ 30` para prazos futuros.

Exemplo prático:

Um campo `Prazo final` pode usar:

`[today] + 7`

Assim, o prazo será preenchido com sete dias após a data atual.

#### Diferença em dias entre datas

Para calcular diferença entre duas datas, use:

`{{DAYS:CampoDataA:CampoDataB}}`

Exemplo:

`{{DAYS:DataInicio:DataFim}}`

Resultado esperado:

O formulário calcula a quantidade de dias entre as duas datas.

Uso recomendado:

- Use para calcular duração de férias.
- Use para calcular prazo de atendimento.
- Use para calcular tempo entre abertura e conclusão.

#### Tokens literais

Tokens literais representam valores simples.

Tokens disponíveis:

- `[empty]`: texto vazio.
- `[null]`: valor nulo.
- `[true]`: verdadeiro.
- `[false]`: falso.

Exemplos:

`[true]`

`[false]`

Uso recomendado:

- Use `[true]` ou `[false]` em campos sim/não.
- Use `[empty]` quando o campo precisa iniciar vazio por regra.
- Use `[null]` quando o valor precisa ser tratado como nulo.

#### Token de parâmetro da URL

O token `[query:nome]` permite buscar um valor informado na URL da página.

Formato:

`[query:nome]`

Exemplo:

Se a página for aberta com:

`?origem=portal`

O campo pode usar:

`[query:origem]`

Resultado esperado:

O formulário preenche o campo com o valor `portal`.

Uso recomendado:

- Use quando a página recebe parâmetros externos.
- Use para identificar origem da solicitação.
- Use para pré-preencher campos conforme links enviados ao usuário.

#### Referências a campos lookup e pessoa

Campos lookup e pessoa podem expor propriedades específicas.

Formato:

`{{Campo/Propriedade}}`

Exemplos:

`{{Projeto/Title}}`

`{{Projeto/Id}}`

`{{Responsavel/EMail}}`

`{{Responsavel/LoginName}}`

Uso recomendado:

- Use `/Title` para exibir o nome ou título relacionado.
- Use `/Id` quando precisar do identificador.
- Use `/EMail` para campos de pessoa.
- Use `/LoginName` quando o login for necessário para integração ou controle.

Exemplo prático:

`str: Projeto selecionado: {{Projeto/Title}}`

#### Expressões para campos lookup, lookupmulti, user e usermulti

Campos de lookup e pessoa geralmente esperam identificadores.

Por isso, em muitos cenários, as expressões mais comuns são:

- `[me]`
- `[myId]`
- Referências a outros campos lookup ou pessoa.

Exemplo:

Um campo `Solicitante` do tipo pessoa pode usar:

`[me]`

Assim, o formulário identifica o usuário atual.

Uso recomendado:

- Use `[me]` para preencher pessoa atual.
- Use referências de lookup quando o valor deve vir de outro campo relacionado.
- Evite usar texto livre em campos que esperam identificador.

#### Expressões para campos de data

Campos de data aceitam tokens de data, referências a outros campos de data e acréscimos de dias.

Exemplos:

`[today]`

`[tomorrow]`

`{{DataInicio}} + 7`

`[today] + {{QuantidadeDias}}`

Uso recomendado:

- Use `[today]` para data inicial automática.
- Use `{{OutraData}} + N` para calcular prazos.
- Use campos numéricos quando o prazo varia conforme uma informação preenchida pelo usuário.

#### Expressões para campos texto, escolha, URL e taxonomia

Campos de texto, escolha, URL e taxonomia podem usar textos montados com `str:`, tokens e referências a outros campos.

Exemplos:

`str: Aberto por [myName]`

`str: {{TipoSolicitacao}} - {{Departamento}}`

`str: https://empresa.com/processo?id={{ID}}`

Uso recomendado:

- Use para padronizar títulos, descrições e links.
- Use para preencher classificações conforme regras.
- Valide o resultado antes de publicar, principalmente em campos de escolha ou taxonomia.

#### Expressões para número e moeda

Campos numéricos e de moeda devem usar expressões numéricas.

Exemplos:

`{{Quantidade}} * {{ValorUnitario}}`

`{{ValorTotal}} - {{Desconto}}`

`({{ValorA}} + {{ValorB}}) / 2`

Uso recomendado:

- Use apenas campos numéricos ou de moeda nas contas.
- Evite misturar texto com cálculo numérico.
- Use parênteses quando precisar controlar a ordem do cálculo.

#### Expressões para booleano

Campos booleanos representam verdadeiro ou falso.

Exemplos:

`[true]`

`[false]`

Uso recomendado:

- Use `[true]` para marcar uma opção automaticamente.
- Use `[false]` para deixar uma opção desmarcada por padrão.
- Use condições de ativação, desativação ou visibilidade quando o valor depender de outras respostas.

#### Expressões para pastas de anexos

Quando os anexos usam biblioteca de documentos com árvore de pastas configurada, a expressão pode apontar para uma pasta específica.

Formato:

`attfolder:idDaPasta`

Uso esperado:

Essa expressão gera o caminho da pasta ligada ao item, conforme a estrutura configurada na aba `Anexos`.

Uso recomendado:

- Use quando o formulário precisa salvar ou referenciar arquivos em uma pasta específica.
- Use em conjunto com a configuração de biblioteca de documentos.
- Valide a árvore de pastas antes de usar em produção.

#### Sugestões com @

Nos campos que aceitam expressão, digitar `@` abre sugestões disponíveis.

As sugestões ajudam o usuário a inserir tokens, campos e referências sem precisar decorar todos os formatos.

Exemplos de sugestões:

- Tokens de usuário.
- Tokens de data.
- Campos do formulário.
- Campos numéricos.
- Campos lookup.
- Pastas de anexos, quando disponíveis.

Uso recomendado:

- Digite `@` para procurar o token ou campo desejado.
- Prefira selecionar a sugestão em vez de digitar manualmente.
- Use as sugestões para evitar erro no nome interno do campo.

#### Boas práticas para expressões

- Comece com expressões simples.
- Teste o resultado em homologação antes de copiar para produção.
- Use nomes internos corretos dos campos.
- Use `str:` quando quiser montar texto.
- Use operadores matemáticos apenas em campos numéricos.
- Use tokens de data apenas em campos compatíveis com data ou texto.
- Evite expressões muito longas quando a regra puder ser quebrada em partes menores.
- Documente a finalidade da expressão no texto de ajuda do campo quando o comportamento não for óbvio para o usuário.

### Componentes

A aba `Componentes` reúne configurações visuais e comportamentais do formulário.

Ela não altera diretamente os campos da lista, mas define como algumas partes da experiência serão exibidas para o usuário.

Nessa aba é possível configurar a forma de visualização da listagem, os indicadores de carregamento, o visual das etapas e o histórico de auditoria.

#### Visualização e listagem (gestor)

O collapse `Visualização e listagem (gestor)` controla como a listagem de registros aparece acima do formulário.

Essa área é usada quando a página permite consultar registros existentes e alternar entre diferentes formas de visualização.

##### Modo de visualização padrão

Define qual visualização será aberta primeiro quando o usuário acessar a página.

Opções disponíveis:

- `Tabela`: exibe os registros em linhas e colunas.
- `Cartões`: exibe os registros em formato visual de cards.

Uso recomendado:

- Use `Tabela` quando o usuário precisa comparar muitos dados ao mesmo tempo.
- Use `Cartões` quando a leitura visual por item for mais importante que a comparação em colunas.
- Use `Cartões` para experiências mais simples e com menos campos por registro.

Exemplo:

Uma lista de solicitações administrativas pode abrir em `Tabela`, pois o usuário normalmente precisa comparar status, datas e responsáveis.

Um catálogo de itens pode abrir em `Cartões`, pois a apresentação visual de cada item pode ser mais importante.

##### Controlo no ecrã para alternar tabela / cartões

Define como o usuário poderá alternar entre os modos de visualização disponíveis.

Opções disponíveis:

- `Botões segmentados`: mostra botões com ícones para alternar entre tabela e cartões.
- `Lista suspensa compacta`: mostra uma lista menor para escolher o modo de visualização.

Uso recomendado:

- Use `Botões segmentados` quando quiser deixar a troca de visualização mais visível.
- Use `Lista suspensa compacta` quando quiser economizar espaço na tela.

Exemplo:

Em uma tela com bastante espaço, os botões segmentados facilitam a troca rápida entre tabela e cartões.

Em uma tela mais compacta, a lista suspensa ocupa menos espaço e mantém a interface mais limpa.

#### Carregar formulário / dados

O collapse `Carregar formulário / dados` define o visual exibido enquanto o formulário ou os dados estão sendo carregados.

Essa configuração melhora a percepção do usuário durante momentos de espera.

Em vez de parecer que a tela travou, o formulário mostra um indicador visual informando que os dados ainda estão sendo carregados.

##### Estilo de loading (dados)

Define o tipo de indicador exibido durante o carregamento.

Opções disponíveis:

- `Spinner Fluent`: indicador padrão de carregamento.
- `Spinner grande`: indicador maior, com mais destaque.
- `Blocos shimmer`: mostra blocos simulando a estrutura da tela enquanto carrega.
- `Barra de progresso indeterminada`: mostra uma barra animada de progresso.
- `Cartão com avatar + linhas`: mostra uma prévia visual em formato de cartão.

Uso recomendado:

- Use `Spinner Fluent` para telas simples.
- Use `Spinner grande` quando o carregamento precisa ficar mais evidente.
- Use `Blocos shimmer` quando quiser indicar que a estrutura da tela está sendo montada.
- Use `Barra de progresso indeterminada` quando o tempo de carregamento pode variar.
- Use `Cartão com avatar + linhas` quando a experiência tiver aparência de cards ou registros individuais.

Exemplo:

Se o formulário carrega dados de uma lista com muitos campos, o `shimmer` pode deixar a espera mais agradável, pois indica que o conteúdo está sendo preparado.

##### Pré-visualização do loading

Dentro do collapse existe uma área de pré-visualização.

Ela mostra como o estilo escolhido será exibido para o usuário.

Use essa prévia para escolher o indicador mais adequado antes de salvar a configuração.

#### Gravar — loading ao gravar (padrão)

O collapse `Gravar — loading ao gravar (padrão)` define o comportamento visual exibido quando o usuário salva ou envia o formulário.

Essa configuração é importante para indicar que a ação está em andamento e evitar que o usuário clique várias vezes no botão de gravação.

##### Estilo de loading ao gravar

Define como o formulário mostra que está salvando as informações.

Opções disponíveis:

- `Sobreposição + spinner`: mostra uma camada sobre o formulário com indicador de carregamento.
- `Barra de progresso no topo`: mostra uma barra no topo da área.
- `Shimmer sobre o formulário`: mostra uma animação de carregamento sobre a estrutura do formulário.
- `Spinner por baixo dos botões`: mostra o indicador próximo aos botões de ação.
- `Faixa informativa`: mostra uma mensagem em formato de faixa.

Uso recomendado:

- Use `Sobreposição + spinner` quando quiser deixar claro que o usuário deve aguardar.
- Use `Barra de progresso no topo` quando quiser uma indicação mais discreta.
- Use `Spinner por baixo dos botões` quando a ação estiver diretamente ligada aos botões finais.
- Use `Faixa informativa` quando quiser informar o estado da gravação com uma mensagem mais visível.

Exemplo:

Em um formulário de solicitação, ao clicar em `Enviar`, a sobreposição com spinner evita que o usuário altere campos enquanto o envio está sendo processado.

#### Etapas — layout e navegação

O collapse `Etapas — layout e navegação` controla a aparência das etapas do formulário e dos botões usados para avançar ou voltar.

Essa área é usada quando o formulário possui mais de uma etapa e precisa de uma apresentação visual clara para o usuário.

##### Cor de destaque

Define a cor usada no passador de etapas e nos botões de navegação.

Essa cor ajuda a alinhar o formulário com a identidade visual da página ou da empresa.

Uso recomendado:

- Use a cor principal do tema quando quiser manter o padrão visual do site.
- Use uma cor de destaque quando o formulário precisa se diferenciar.
- Evite cores que dificultem a leitura ou reduzam o contraste.

##### Layout das etapas no formulário

Define o estilo visual usado para mostrar as etapas do formulário.

Exemplos de estilos disponíveis:

- `Trilho lateral`: mostra as etapas em coluna, indicado para formulários longos.
- `Segmentos`: mostra etapas como pílulas horizontais.
- `Linha do tempo`: mostra o progresso em formato de linha.
- `Cartões`: exibe cada etapa com visual mais destacado.
- `Migalhas`: mostra o caminho das etapas em formato de navegação.
- `Separadores`: apresenta as etapas como abas.
- `Contorno`: mostra etiquetas com borda.
- `Compacto`: reduz o espaço ocupado pelas etapas.
- `Passo numerado`: mostra etapas numeradas em sequência.
- `Minimal`: usa uma apresentação mais leve e discreta.

Uso recomendado:

- Use `Trilho lateral` para formulários longos.
- Use `Segmentos` para formulários empresariais com poucas ou médias etapas.
- Use `Linha do tempo` quando o progresso do usuário for importante.
- Use `Compacto` quando houver pouco espaço vertical.
- Use `Minimal` para formulários simples.

Exemplo:

Um formulário com cinco etapas pode usar `Linha do tempo` para mostrar claramente em qual parte do processo o usuário está.

##### Botões Etapa anterior / Próxima etapa

Define o estilo dos botões usados para navegar entre as etapas.

Exemplos de estilos disponíveis:

- `Fluent padrão`: botões comuns do Fluent UI.
- `Pílulas`: botões largos com formato arredondado.
- `Bolinhas e setas`: navegação com indicadores e setas.
- `Só ícones`: botões compactos com ícones.
- `Ligações de texto`: navegação com aparência de link.
- `Extremos`: botão anterior à esquerda e próximo à direita.
- `Empilhado`: botões em coluna, útil em telas estreitas.
- `Contorno`: botões com borda.
- `Barra cinza`: botões agrupados em uma faixa.
- `Compacto`: botões menores, com menos espaçamento.

Uso recomendado:

- Use `Fluent padrão` para manter o visual familiar.
- Use `Extremos` quando quiser separar bem voltar e avançar.
- Use `Empilhado` para formulários em telas estreitas.
- Use `Só ícones` apenas quando o contexto estiver claro para o usuário.

Exemplo:

Em um formulário usado em celular, o estilo `Empilhado` pode facilitar o toque nos botões de navegação.

#### Histórico de auditoria

O collapse `Histórico de auditoria` configura o botão e a forma de apresentação do histórico do item.

Essa funcionalidade permite que o usuário consulte registros de alterações, ações ou versões relacionadas ao formulário, conforme a configuração de auditoria existente.

##### Ativar botão de histórico de auditoria

Liga ou desliga o botão de histórico no formulário.

Quando está inativo, o usuário não vê o botão de histórico.

Quando está ativo, novas opções aparecem para configurar como o histórico será aberto e exibido.

Uso recomendado:

- Ative quando o cliente precisa acompanhar alterações ou registros de ação.
- Deixe inativo em formulários simples que não exigem rastreabilidade.

##### Abrir histórico como

Define onde o histórico será exibido quando o usuário clicar no botão.

Opções disponíveis:

- `Painel lateral`: abre o histórico em um painel ao lado.
- `Modal`: abre o histórico em uma janela central.
- `Secção no formulário`: mostra o histórico dentro do próprio formulário, abaixo dos botões.

Uso recomendado:

- Use `Painel lateral` quando quiser manter o formulário visível enquanto consulta o histórico.
- Use `Modal` quando o histórico precisa de mais foco.
- Use `Secção no formulário` quando quiser manter tudo na mesma tela.

##### Aspeto do botão no formulário

Define como o botão de histórico será apresentado.

Opções disponíveis:

- `Só texto`: exibe apenas o texto do botão.
- `Só ícone`: exibe apenas o ícone.
- `Ícone e texto`: exibe ícone e texto juntos.

Uso recomendado:

- Use `Só texto` quando quiser máxima clareza.
- Use `Só ícone` quando houver pouco espaço e o ícone for conhecido.
- Use `Ícone e texto` quando quiser equilíbrio entre clareza e visual.

##### Texto do botão

Define o texto exibido no botão de histórico.

Exemplo:

`Histórico`

Também pode ser alterado para algo mais específico, como:

- `Ver alterações`
- `Auditoria`
- `Histórico do item`

##### Nome acessível

Quando o botão usa apenas ícone, o `Nome acessível` serve como identificação para tooltip e leitores de tela.

Essa configuração ajuda na acessibilidade, pois permite que o usuário entenda a função do botão mesmo sem texto visível.

##### Ícone Fluent

Define o nome do ícone usado no botão de histórico.

Exemplos:

- `History`
- `Clock`
- `TimelineProgress`

Uso recomendado:

- Use ícones relacionados a tempo, histórico ou registros.
- Mantenha o ícone simples e fácil de reconhecer.

##### Subtítulo / ajuda

Define uma mensagem de apoio exibida no painel de histórico ou como tooltip.

Use esse campo para explicar o que o usuário encontrará ao abrir o histórico.

Exemplo:

`Consulte alterações e registros relacionados a este item.`

##### Grupos do SharePoint

Permite limitar o botão de histórico a usuários de grupos específicos do SharePoint.

Também existe um filtro para localizar grupos pelo nome.

Uso recomendado:

- Use quando o histórico deve ser visível apenas para gestores, administradores ou equipes de auditoria.
- Deixe sem grupos quando todos os usuários do formulário puderem consultar o histórico.

Exemplo:

O botão de histórico pode ser exibido apenas para os grupos `Gestores` e `Administradores`.

##### Estilo da lista de registos

Define como os registros do histórico serão apresentados.

Opções disponíveis:

- `Lista`: registros empilhados em blocos.
- `Linha do tempo`: registros apresentados em sequência cronológica visual.
- `Cartões`: registros em cards com destaque visual.
- `Compacto`: apresentação mais densa, ocupando menos espaço.

Uso recomendado:

- Use `Lista` para uma apresentação simples e clara.
- Use `Linha do tempo` quando a ordem dos eventos for importante.
- Use `Cartões` quando quiser destacar cada registro.
- Use `Compacto` quando houver muitos registros e pouco espaço.

##### Pré-visualização do estilo

O collapse mostra uma pré-visualização do estilo escolhido para os registros.

Use essa prévia para validar se o formato está adequado antes de salvar.

### Anexos

A aba `Anexos` reúne as configurações relacionadas ao envio de arquivos pelo formulário.

Nessa aba é possível definir onde os arquivos serão armazenados e, quando aplicável, como eles serão vinculados ao item principal do formulário.

#### Destino do upload

O collapse `Destino do upload` define onde os arquivos enviados pelo usuário serão gravados.

Essa configuração é importante porque determina se os arquivos ficarão como anexos do próprio item da lista ou se serão enviados para uma biblioteca de documentos.

##### Anexos ao item

A opção `Anexos ao item (lista principal)` grava os arquivos diretamente no item da lista principal.

Esse é o comportamento mais simples e direto.

Quando o usuário envia arquivos pelo formulário, eles ficam anexados ao próprio registro criado ou editado.

Exemplo:

Em um formulário de solicitação de férias, o usuário pode anexar um comprovante ou documento de apoio diretamente ao item da solicitação.

Uso recomendado:

- Use quando os arquivos pertencem diretamente ao registro.
- Use quando não há necessidade de organizar os arquivos em pastas.
- Use quando o cliente quer uma configuração mais simples.
- Use para formulários com poucos arquivos por item.

Ponto de atenção:

Como os arquivos ficam anexados ao item da lista, a organização segue o próprio registro do SharePoint.

##### Biblioteca de documentos

A opção `Biblioteca de documentos` envia os arquivos para uma biblioteca do SharePoint.

Esse modo é indicado quando os arquivos precisam ficar organizados em uma biblioteca, com possibilidade de estrutura de pastas e metadados próprios.

Ao selecionar essa opção, o formulário exibe configurações adicionais para escolher a biblioteca e informar como os arquivos serão ligados ao item principal.

Exemplo:

Em um formulário de gestão documental, os arquivos enviados podem ser armazenados em uma biblioteca chamada `Documentos`, em vez de ficarem apenas como anexos do item.

Uso recomendado:

- Use quando os arquivos precisam ser organizados em biblioteca.
- Use quando os documentos precisam ter controle, visualização ou gestão própria.
- Use quando o cliente já trabalha com bibliotecas documentais.
- Use quando há necessidade de estruturar pastas por item ou processo.

##### Biblioteca de documentos

Campo usado para escolher em qual biblioteca os arquivos serão salvos.

A lista mostra as bibliotecas disponíveis no site.

Exemplo:

Selecionar a biblioteca `Documentos de Solicitações` para armazenar todos os arquivos enviados pelo formulário.

Uso recomendado:

- Escolha uma biblioteca criada para esse tipo de processo.
- Evite misturar documentos de processos diferentes na mesma biblioteca sem necessidade.
- Confirme se os usuários possuem permissão para gravar arquivos na biblioteca escolhida.

##### Lookup para a lista principal

O campo `Lookup para a lista principal (vínculo ao item)` define qual coluna da biblioteca liga o arquivo ao item principal do formulário.

Esse vínculo permite identificar a qual registro cada arquivo pertence.

Na prática, a biblioteca precisa ter uma coluna de lookup apontando para a lista principal usada pelo formulário.

Exemplo:

A biblioteca `Documentos de Solicitações` pode ter uma coluna `Solicitação` que aponta para a lista principal `Solicitações`.

Quando um arquivo é enviado, ele fica salvo na biblioteca e vinculado ao item correto da lista.

Uso recomendado:

- Configure esse campo quando usar `Biblioteca de documentos`.
- Garanta que a biblioteca tenha uma coluna de lookup para a lista principal.
- Use nomes claros para o lookup, como `Solicitação`, `Item relacionado` ou `Registro principal`.

Ponto de atenção:

Se não existir uma coluna de lookup na biblioteca apontando para a lista principal, será necessário criá-la antes de concluir essa configuração.

##### Mensagens de carregamento e validação

Durante a configuração, o painel pode exibir mensagens de carregamento ou aviso.

Exemplos:

- Carregando bibliotecas disponíveis.
- Carregando campos da biblioteca selecionada.
- Resolvendo a lista principal do formulário.
- Avisando que não existe lookup válido para vincular os arquivos ao item.

Essas mensagens ajudam a identificar se a configuração está pronta ou se ainda falta algum ajuste no SharePoint.

#### Aspeto e pré-visualização do controlo

O collapse `Aspeto e pré-visualização do controlo` define como o campo de anexos será exibido para o usuário no formulário.

Essa configuração não altera o destino dos arquivos. Ela controla apenas a aparência do controle de upload e a forma como os arquivos selecionados aparecem na tela.

Use essa área para escolher uma experiência visual mais simples, mais destacada ou mais compacta, conforme o tipo de formulário.

##### Tipo de layout do input de anexos

Define o formato visual usado para o usuário selecionar ou arrastar arquivos.

Opções disponíveis:

- `Clássico`: exibe o controle de anexo no formato mais tradicional.
- `Zona destacada`: cria uma área de destaque para clicar ou arrastar arquivos.
- `Cartão com ícone e sombra`: apresenta o controle em formato de cartão visual.
- `Faixa azul + área de largar`: exibe uma faixa visual com área para soltar arquivos.
- `Compacto`: mostra um botão mais simples e arquivos em formato de chips.

Uso recomendado:

- Use `Clássico` quando quiser uma experiência simples e familiar.
- Use `Zona destacada` quando o envio de arquivos for uma parte importante do formulário.
- Use `Cartão com ícone e sombra` quando quiser dar mais destaque visual ao anexo.
- Use `Faixa azul + área de largar` para orientar claramente que o usuário pode arrastar arquivos.
- Use `Compacto` quando o formulário tiver pouco espaço ou quando anexos forem opcionais.

Exemplo:

Em um formulário de envio de documentos, a `Zona destacada` pode facilitar o entendimento de que o usuário deve anexar arquivos.

Em um formulário simples, onde o anexo é opcional, o layout `Compacto` pode deixar a tela mais limpa.

##### Pré-visualização dos ficheiros selecionados

Define como os arquivos selecionados serão exibidos depois que o usuário adiciona os anexos.

Opções disponíveis:

- `Só nome do ficheiro`: mostra apenas o nome do arquivo.
- `Nome e tamanho`: mostra o nome e o tamanho do arquivo.
- `Ícone por tipo + nome`: mostra um ícone conforme o tipo do arquivo, junto com o nome.
- `Miniatura ou ícone + nome`: mostra miniatura para imagens ou ícone para outros arquivos.
- `Pré-visualização grande`: mostra os arquivos em cartões maiores.

Uso recomendado:

- Use `Só nome do ficheiro` quando quiser a visualização mais simples.
- Use `Nome e tamanho` quando o usuário precisa conferir também o peso do arquivo.
- Use `Ícone por tipo + nome` para facilitar a identificação visual do tipo de arquivo.
- Use `Miniatura ou ícone + nome` quando imagens forem comuns no formulário.
- Use `Pré-visualização grande` quando os anexos forem parte central da análise.

Exemplo:

Em um formulário de evidências com imagens, a opção `Miniatura ou ícone + nome` ajuda o usuário a conferir rapidamente se anexou a imagem correta.

Em um formulário administrativo com PDFs e documentos, `Nome e tamanho` costuma ser suficiente.

##### Pré-visualização com arquivos de teste

A área de pré-visualização permite testar como o controle ficará antes de salvar a configuração.

Nessa área, é possível adicionar arquivos de teste para visualizar o comportamento do layout escolhido e da forma de apresentação dos arquivos.

Essa pré-visualização serve apenas para apoiar a configuração visual.

Uso recomendado:

- Teste o layout antes de liberar o formulário para o cliente.
- Adicione arquivos de tipos diferentes para validar a apresentação.
- Confira se o controle fica claro para o usuário final.
- Use a pré-visualização para escolher entre uma experiência mais compacta ou mais destacada.

Exemplo:

Ao selecionar `Cartão com ícone e sombra` e `Ícone por tipo + nome`, a pré-visualização mostra como o usuário verá os arquivos após selecioná-los.

#### Extensões permitidas

O collapse `Extensões permitidas` define quais tipos de arquivos o usuário poderá anexar no formulário.

Essa configuração serve para restringir o upload a formatos específicos, de acordo com a necessidade do processo.

Quando nenhuma extensão é marcada, o formulário aceita qualquer tipo de anexo.

##### Nenhuma extensão selecionada

Se nenhuma opção for marcada, não haverá restrição por extensão.

Nesse caso, o usuário poderá anexar arquivos de qualquer tipo, desde que o SharePoint e as permissões do ambiente permitam.

Uso recomendado:

- Use quando o processo aceita vários tipos de arquivo.
- Use quando não há regra específica sobre formato.
- Use em formulários mais genéricos, onde o usuário pode precisar enviar documentos variados.

Exemplo:

Um formulário de atendimento pode aceitar imagens, PDFs, planilhas, documentos e arquivos compactados. Nesse caso, pode ser melhor deixar nenhuma extensão selecionada.

##### Quando marcar extensões

Ao marcar uma ou mais extensões, o formulário passa a aceitar somente os tipos selecionados.

Isso ajuda a controlar melhor o conteúdo recebido e evita arquivos fora do padrão esperado.

Exemplo:

Se forem marcadas apenas as opções `PDF`, `Word .doc` e `Word .docx`, o usuário só poderá enviar arquivos desses formatos.

Uso recomendado:

- Marque extensões quando o processo exige um formato específico.
- Marque apenas os tipos realmente aceitos pelo cliente.
- Use restrição para reduzir envio de arquivos incorretos.
- Revise as extensões antes de publicar o formulário.

##### PDF e documentos Word

Grupo usado para permitir documentos de texto e arquivos formais.

Extensões disponíveis:

- `PDF`
- `Word .doc`
- `Word .docx`

Uso recomendado:

- Use para contratos, declarações, comprovantes e documentos oficiais.
- Use `PDF` quando o documento não deve ser facilmente alterado.
- Use Word quando o usuário pode enviar documentos editáveis.

##### Excel

Grupo usado para permitir planilhas.

Extensões disponíveis:

- `.xls`
- `.xlsx`

Uso recomendado:

- Use para formulários que recebem planilhas de controle.
- Use para importação de dados, relatórios ou levantamentos.
- Evite liberar Excel quando o processo não precisa de planilhas.

##### PowerPoint

Grupo usado para permitir apresentações.

Extensões disponíveis:

- `.ppt`
- `.pptx`

Uso recomendado:

- Use quando o processo recebe apresentações institucionais, comerciais ou materiais de apoio.
- Evite liberar se o formulário for voltado apenas para documentos administrativos simples.

##### Imagens

Grupo usado para permitir arquivos de imagem.

Extensões disponíveis:

- `PNG`
- `JPEG .jpg`
- `JPEG .jpeg`
- `GIF`
- `WebP`
- `SVG`

Uso recomendado:

- Use para evidências visuais.
- Use para fotos, prints, imagens de comprovantes ou anexos gráficos.
- Use junto com pré-visualização por miniatura quando imagens forem comuns.

Exemplo:

Um formulário de vistoria pode permitir `PNG`, `JPG` e `JPEG` para envio de fotos.

##### Texto e tabelas

Grupo usado para permitir arquivos simples de texto e dados tabulares.

Extensões disponíveis:

- `Texto .txt`
- `CSV`

Uso recomendado:

- Use `TXT` para arquivos simples de texto.
- Use `CSV` quando o processo aceita dados exportados de sistemas ou planilhas.
- Use com cuidado quando os dados precisarem seguir um padrão específico.

##### Arquivos e correio

Grupo usado para permitir arquivos compactados ou mensagens de e-mail.

Extensões disponíveis:

- `ZIP`
- `Outlook .msg`

Uso recomendado:

- Use `ZIP` quando o usuário precisa enviar vários arquivos juntos.
- Use `.msg` quando o processo exige anexar mensagens de e-mail.
- Evite liberar `ZIP` quando o processo precisa analisar cada arquivo separadamente.

##### Vídeo

Grupo usado para permitir arquivos de vídeo.

Extensão disponível:

- `MP4`

Uso recomendado:

- Use quando o processo precisa receber evidências em vídeo.
- Use para registros visuais, demonstrações ou comprovações.
- Considere o tamanho dos arquivos antes de liberar vídeo para todos os usuários.

##### Boas práticas

- Libere apenas as extensões necessárias para o processo.
- Se o formulário for genérico, deixe sem seleção para aceitar qualquer tipo.
- Se o formulário for específico, marque somente os formatos esperados.
- Combine a restrição de extensão com instruções claras no texto de ajuda do campo.
- Teste o envio com arquivos reais antes de publicar para o cliente.

### Botões

A aba `Botões` reúne as configurações relacionadas aos botões de ação do formulário.

Nessa aba é possível definir onde a barra de botões será exibida e configurar ações específicas que o usuário poderá executar.

#### Onde mostrar a barra de botões

A seção `Onde mostrar a barra de botões` define a posição da barra de botões dentro do formulário.

Essa configuração controla onde os botões aparecem visualmente para o usuário, considerando a posição vertical e horizontal.

Ela é útil para adaptar a experiência ao tamanho do formulário, ao tipo de processo e ao comportamento esperado do usuário.

##### Vertical

O campo `Vertical` define se a barra de botões ficará na parte superior ou inferior do formulário.

Opções disponíveis:

- `Inferior`: exibe a barra de botões na parte de baixo do formulário.
- `Superior`: exibe a barra de botões na parte de cima do formulário.

Uso recomendado:

- Use `Inferior` quando o usuário deve preencher o formulário antes de executar uma ação.
- Use `Inferior` para formulários de cadastro, solicitação ou envio.
- Use `Superior` quando as ações precisam ficar disponíveis logo ao abrir o formulário.
- Use `Superior` em telas de consulta, aprovação ou gerenciamento, onde o usuário pode precisar agir rapidamente.

Exemplo:

Em um formulário de solicitação de férias, a posição `Inferior` faz sentido porque o usuário normalmente preenche os dados antes de enviar.

Em uma tela de aprovação, a posição `Superior` pode facilitar o acesso rápido aos botões `Aprovar` ou `Reprovar`.

##### Horizontal

O campo `Horizontal` define o alinhamento da barra de botões na tela.

Opções disponíveis:

- `Esquerda`: alinha os botões à esquerda.
- `Direita`: alinha os botões à direita.

Uso recomendado:

- Use `Esquerda` quando quiser manter os botões próximos do início do conteúdo.
- Use `Esquerda` em formulários com leitura da esquerda para a direita e ações mais simples.
- Use `Direita` quando quiser aproximar os botões do padrão comum de confirmação no fim da área.
- Use `Direita` para ações finais, como salvar, enviar ou concluir.

Exemplo:

Em um formulário de cadastro, a barra `Inferior` e `Direita` pode deixar os botões próximos do ponto natural de conclusão.

Em uma tela administrativa, a barra `Superior` e `Esquerda` pode facilitar o acesso às ações assim que a tela abrir.

##### Combinações comuns

Exemplos de combinações:

- `Inferior` + `Direita`: indicado para formulários de preenchimento e envio.
- `Inferior` + `Esquerda`: indicado para formulários simples ou fluxos internos.
- `Superior` + `Direita`: indicado para telas de consulta com ações rápidas.
- `Superior` + `Esquerda`: indicado para telas administrativas ou de gerenciamento.

Uso recomendado:

- Escolha a posição pensando no momento em que o usuário executa a ação.
- Para ações finais, prefira posicionar a barra no fim do fluxo.
- Para ações de consulta ou gestão, prefira deixar a barra mais acessível no topo.
- Mantenha um padrão entre formulários parecidos para facilitar o uso pelo cliente.

#### Adicionar botão

O botão `Adicionar botão` cria uma nova ação personalizada na barra de botões do formulário.

Ao clicar nessa opção, um novo botão é adicionado à lista de botões configuráveis.

Por padrão, o novo botão é criado com o nome `Novo botão` e pode ser ajustado depois conforme a finalidade desejada.

Essa funcionalidade é usada quando o formulário precisa oferecer ações além dos botões padrão.

Exemplos de uso:

- Criar um botão `Aprovar`.
- Criar um botão `Reprovar`.
- Criar um botão `Enviar para análise`.
- Criar um botão `Cancelar solicitação`.
- Criar um botão `Gerar protocolo`.
- Criar um botão para executar ações internas configuradas no formulário.

##### O que acontece ao adicionar

Quando um botão é criado, ele aparece como um novo bloco dentro da aba `Botões`.

Esse bloco pode ser expandido para configurar o texto, aparência, tipo de operação, condições de exibição e ações que serão executadas.

O botão criado passa a fazer parte da barra de botões do formulário, respeitando a posição vertical e horizontal definida anteriormente.

Exemplo:

Se a barra estiver configurada como `Inferior` e `Direita`, o novo botão será exibido nessa área junto com os demais botões configurados.

##### Reordenação dos botões

Os botões criados podem ser reordenados.

Isso permite controlar a sequência em que as ações aparecem para o usuário.

Exemplo:

Em um fluxo de aprovação, a ordem pode ser:

- `Aprovar`
- `Reprovar`
- `Solicitar ajuste`

Uso recomendado:

- Coloque primeiro as ações mais usadas.
- Mantenha ações críticas em posições fáceis de identificar.
- Evite criar muitos botões quando poucas ações resolvem o processo.
- Use nomes claros para que o usuário entenda exatamente o que cada botão faz.

##### Quando usar

Use `Adicionar botão` quando o processo exigir uma ação específica dentro do formulário.

Exemplos:

- Aprovação de solicitação.
- Mudança de status.
- Execução de automações internas.
- Redirecionamento após uma ação.
- Registro de uma decisão no histórico.

##### Boas práticas

- Dê nomes objetivos aos botões.
- Evite textos genéricos como `Executar` quando a ação não for óbvia.
- Use poucos botões para não confundir o usuário.
- Revise a ordem dos botões antes de publicar o formulário.
- Teste cada botão em homologação antes de liberar em produção.

#### Clonar botão

O botão `Clonar` cria uma cópia de um botão já configurado.

Essa funcionalidade é útil quando é necessário criar outro botão parecido, aproveitando a configuração existente como base.

Ao clonar, o sistema cria um novo botão logo abaixo do botão original.

O novo botão recebe um novo identificador interno e o nome passa a incluir `(cópia)`.

Exemplo:

Se o botão original se chama `Aprovar`, o botão clonado pode aparecer como `Aprovar (cópia)`.

##### O que é copiado

Ao clonar um botão, as configurações do botão original são copiadas para o novo botão.

Isso pode incluir:

- Texto do botão.
- Aparência.
- Tipo de operação.
- Condições de exibição.
- Ações configuradas.
- Comportamento após executar as ações.
- Configurações relacionadas ao carregamento.

Depois de clonar, o novo botão pode ser editado normalmente.

Exemplo:

Um botão `Aprovar` pode ser clonado para criar um botão `Reprovar`, aproveitando a mesma estrutura de configuração e ajustando apenas o texto, as ações e as condições.

##### Quando usar

Use `Clonar` quando dois botões têm comportamento parecido.

Exemplos:

- `Aprovar` e `Reprovar`.
- `Enviar para análise` e `Enviar para revisão`.
- `Salvar rascunho` e `Salvar e concluir`.
- `Notificar gestor` e `Notificar solicitante`.

Uso recomendado:

- Clone botões para ganhar tempo em configurações parecidas.
- Após clonar, revise o texto do botão.
- Ajuste as ações para evitar executar o mesmo comportamento do botão original por engano.
- Revise as condições de exibição do botão clonado.
- Teste o botão clonado antes de publicar o formulário.

##### Ponto de atenção

Como o botão clonado copia a configuração do original, ele pode trazer ações ou condições que não fazem sentido para o novo caso.

Por isso, sempre revise a cópia antes de salvar a configuração final.

#### Remover botão

O botão `Remover botão` exclui um botão personalizado da configuração do formulário.

Ao remover, o botão deixa de aparecer na barra de botões e suas configurações deixam de fazer parte da experiência do usuário.

Essa ação deve ser usada quando uma ação não é mais necessária no processo.

Exemplos de uso:

- Remover um botão criado por engano.
- Excluir uma ação que deixou de fazer parte do fluxo.
- Limpar botões antigos após uma mudança no processo.
- Remover um botão clonado que não será mais utilizado.

##### O que acontece ao remover

Quando o botão é removido, ele sai da lista de botões configurados.

Com isso, o usuário final não verá mais essa ação no formulário.

As configurações associadas ao botão removido também deixam de ser usadas, como:

- Texto do botão.
- Aparência.
- Tipo de operação.
- Condições de exibição.
- Ações configuradas.
- Comportamento após execução.

##### Cuidados antes de remover

Antes de remover um botão, revise se ele ainda é usado no processo.

Uso recomendado:

- Confirme se a ação realmente não será mais necessária.
- Verifique se outro botão já substitui essa função.
- Revise se o processo ainda terá uma forma de concluir, aprovar, enviar ou cancelar quando necessário.
- Evite remover botões críticos sem validar o fluxo completo.

Exemplo:

Se o botão `Enviar para análise` for removido, o usuário pode ficar sem uma forma de encaminhar o registro para a próxima etapa do processo, caso não exista outro botão equivalente.

##### Ponto de atenção

A remoção é indicada para simplificar a interface, mas deve ser feita com cuidado.

Se houver dúvida, clone ou ajuste o botão antes de remover definitivamente da configuração.

#### Configurações de um botão adicionado

Ao expandir um botão criado na aba `Botões`, o painel exibe as configurações daquele botão.

Essas configurações definem o nome exibido, o tipo de ação, o comportamento visual, em quais modos o botão aparece e se ele está ativo no formulário.

Esta seção descreve as opções desde `Texto do botão` até `Botão ativo`.

##### Texto do botão

O campo `Texto do botão` define o nome que será exibido para o usuário na barra de botões.

Esse texto deve deixar clara a ação que será executada.

Exemplos:

- `Salvar`
- `Enviar`
- `Aprovar`
- `Reprovar`
- `Solicitar ajuste`
- `Cancelar solicitação`

Uso recomendado:

- Use textos curtos e objetivos.
- Evite nomes genéricos quando a ação for importante.
- Prefira verbos de ação, como `Enviar`, `Aprovar`, `Cancelar` ou `Concluir`.
- Use o mesmo padrão de nome em formulários semelhantes.

Exemplo:

Em vez de usar `OK`, prefira `Enviar solicitação`, pois o usuário entende melhor o que acontecerá ao clicar.

##### Descrição curta

O campo `Descrição curta` aparece quando o botão está configurado como botão de histórico legado.

Ele serve como texto de apoio ou tooltip para explicar a função do botão.

Uso recomendado:

- Use para orientar o usuário quando o botão não tiver texto suficiente.
- Use descrições curtas.
- Evite repetir exatamente o mesmo texto do botão.

Ponto de atenção:

Para histórico, a configuração recomendada é usar o botão integrado na aba `Componentes`, na seção `Histórico de auditoria`.

##### Tipo de operação

O campo `Tipo de operação` define qual será o comportamento principal do botão.

Opções disponíveis:

- `Ações em cadeia`: executa uma sequência de ações configuradas no próprio botão (mostrar/ocultar campos, definir valores, juntar texto, **pedidos HTTP** a APIs, etc.).
- `Redirecionar`: envia o usuário para uma URL configurada.
- `Adicionar`: cria um novo item na lista.
- `Atualizar`: grava alterações no item atual.
- `Eliminar`: apaga o item atual.
- `Histórico`: opção legada para histórico, quando disponível.

Uso recomendado:

- Use `Ações em cadeia` quando o botão precisa executar várias ações internas.
- Use `Redirecionar` quando o botão deve levar o usuário para outra página.
- Use `Adicionar` para criar novos registros.
- Use `Atualizar` para salvar alterações no item atual.
- Use `Eliminar` apenas quando o processo permite exclusão de registros.

Exemplo:

Um botão `Enviar para análise` pode usar `Ações em cadeia` para alterar o status, preencher campos e depois enviar o formulário.

##### Loading ao gravar

O campo `Loading ao gravar` define qual indicador visual será exibido enquanto a ação do botão está sendo processada.

Opções disponíveis:

- `Padrão`: usa o comportamento definido na aba `Componentes`.
- `Sobreposição + spinner`: mostra uma camada de carregamento sobre o formulário.
- `Barra de progresso no topo`: mostra uma barra no topo.
- `Shimmer sobre o formulário`: mostra uma animação sobre o formulário.
- `Spinner por baixo dos botões`: mostra o carregamento próximo à barra de botões.
- `Faixa informativa`: mostra uma mensagem em faixa.

Uso recomendado:

- Use `Padrão` para manter consistência entre os botões.
- Use uma opção específica quando aquele botão tiver uma ação mais demorada ou crítica.
- Use indicadores mais visíveis em ações que gravam, enviam ou alteram status.

Exemplo:

Um botão `Enviar` pode usar `Sobreposição + spinner` para deixar claro que o usuário deve aguardar o envio terminar.

##### URL de destino

O campo `URL de destino` aparece quando o tipo de operação é `Redirecionar`.

Ele define para qual endereço o usuário será enviado ao clicar no botão.

A URL pode conter valores dinâmicos do formulário.

Exemplos:

- `https://empresa.sharepoint.com/sites/portal`
- `/sites/portal/SitePages/Resumo.aspx?item={{FormID}}`
- `/sites/portal/SitePages/Detalhe.aspx?status={{Status}}`

Uso recomendado:

- Use URLs completas ou caminhos claros.
- Use valores dinâmicos quando o destino depende do item atual.
- Teste o redirecionamento antes de liberar para produção.

##### Inserir valor dinâmico

A opção `Inserir valor dinâmico` ajuda a incluir campos ou tokens na URL.

Ela adiciona o valor escolhido ao final da URL ou substitui um placeholder vazio.

Exemplo:

Se a URL tiver:

`/SitePages/Detalhe.aspx?id={{}}`

Ao escolher `FormID`, o resultado pode ficar:

`/SitePages/Detalhe.aspx?id={{FormID}}`

Uso recomendado:

- Use quando a página de destino precisa receber o ID do item.
- Use quando o destino depende de um campo preenchido no formulário.
- Prefira selecionar o valor pela lista para evitar erro de digitação.

##### Mostrar o botão eliminar em

Essa opção aparece quando o tipo de operação é `Eliminar`.

Ela define em quais modos o botão de exclusão será exibido.

Opções disponíveis:

- `Modo ver`: mostra o botão quando o usuário está apenas visualizando o item.
- `Modo editar`: mostra o botão quando o usuário está editando o item.

Uso recomendado:

- Use com cuidado, pois a exclusão é uma ação sensível.
- Mostre em `Modo ver` quando a exclusão pode ser feita a partir da consulta.
- Mostre em `Modo editar` quando a exclusão deve acontecer durante a manutenção do registro.
- Combine com condições de exibição quando apenas alguns usuários podem excluir.

##### Cor do botão

O campo `Cor do botão` define o visual do botão com base no tema do site.

Essa configuração ajuda a diferenciar botões principais, secundários ou de atenção.

Exemplos de estilos:

- `Contorno`: visual neutro, com menos destaque.
- `Primária do tema`: botão com destaque principal.
- Cores secundárias ou variações do tema, conforme disponível no site.

Uso recomendado:

- Use a cor principal para a ação mais importante.
- Use contorno para ações secundárias.
- Evite dar o mesmo destaque para todos os botões.
- Use cores com cuidado em ações críticas, como reprovar ou excluir.

Exemplo:

Em um formulário de aprovação, `Aprovar` pode usar a cor principal, enquanto `Solicitar ajuste` pode usar um visual mais discreto.

##### Depois das ações

O campo `Depois das ações` aparece quando o botão usa `Ações em cadeia`.

Ele define o que deve acontecer após executar as ações configuradas.

Opções disponíveis:

- `Só executar ações`: executa as ações e não faz outro comportamento final automático.
- `Ações e depois rascunho`: executa as ações e salva como rascunho.
- `Ações e depois enviar`: executa as ações e envia o formulário.
- `Ações e depois fechar formulário`: executa as ações e fecha o formulário.

Uso recomendado:

- Use `Só executar ações` quando o botão apenas ajusta campos ou prepara o formulário.
- Use `Ações e depois rascunho` quando o usuário ainda poderá continuar depois.
- Use `Ações e depois enviar` para botões finais.
- Use `Ações e depois fechar formulário` quando a ação encerra o uso da tela.

Exemplo:

Um botão `Enviar para análise` pode alterar o status para `Em análise` e depois enviar o formulário.

##### Modos

A seção `Modos` define em quais modos do formulário o botão será exibido.

Opções disponíveis:

- `Criar`: mostra o botão quando o usuário está criando um novo item.
- `Editar`: mostra o botão quando o usuário está alterando um item existente.
- `Ver`: mostra o botão quando o usuário está apenas consultando o item.

Quando nenhum modo específico restringe o botão, ele pode ser considerado disponível para todos os modos aplicáveis.

Uso recomendado:

- Use `Criar` para ações de cadastro inicial.
- Use `Editar` para ações de manutenção ou andamento do processo.
- Use `Ver` para ações que podem ser executadas a partir da consulta do item.

Exemplo:

Um botão `Enviar solicitação` pode aparecer apenas em `Criar`.

Um botão `Aprovar` pode aparecer em `Editar` ou `Ver`, conforme o fluxo definido.

##### Botão ativo

A opção `Botão ativo` define se o botão está habilitado na configuração.

Quando marcada, o botão fica ativo e pode aparecer no formulário conforme os modos e condições configuradas.

Quando desmarcada, o botão fica desativado na configuração e não deve ser usado pelo usuário.

Uso recomendado:

- Mantenha ativo somente o que estiver pronto para uso.
- Desative botões em configuração ou em teste.
- Use para guardar uma configuração sem exibir o botão temporariamente.
- Antes de publicar, revise se todos os botões necessários estão ativos.

Exemplo:

Um botão `Reabrir solicitação` pode ficar configurado, mas desativado até que o processo esteja validado pelo cliente.

##### Modal de confirmação

A seção `Modal de confirmação` permite pedir uma confirmação do usuário antes de executar o botão.

Essa confirmação acontece como primeiro passo do clique.

Se o usuário cancelar, nenhuma ação do botão será executada.

Isso inclui ações em cadeia, gravação, redirecionamento, exclusão ou qualquer outro comportamento configurado depois.

Uso recomendado:

- Use em ações importantes ou irreversíveis.
- Use quando o clique pode alterar status, enviar informações ou excluir dados.
- Use quando o usuário precisa revisar uma mensagem antes de continuar.
- Use para pedir uma justificativa ou informação complementar no momento da ação.

Exemplo:

Antes de executar o botão `Reprovar`, o formulário pode abrir uma confirmação pedindo que o usuário revise a ação e informe uma justificativa.

##### Pedir confirmação antes de executar

A opção `Pedir confirmação antes de executar` ativa ou desativa o modal de confirmação do botão.

Quando ativada, o usuário precisa confirmar antes que o botão execute qualquer ação.

Quando desativada, o botão executa diretamente o fluxo configurado.

Comportamento importante:

- Confirmar: continua a execução do botão.
- Cancelar: interrompe tudo e não executa ações.

Exemplo:

Em um botão `Excluir item`, essa opção evita que o usuário apague um registro por engano.

##### Ícone / tipo

O campo `Ícone / tipo` define o estilo visual da confirmação.

Ele ajuda o usuário a entender a importância da mensagem exibida.

Opções disponíveis:

- `Informação`: usado para confirmações simples ou neutras.
- `Sucesso`: usado para ações positivas ou conclusivas.
- `Aviso`: usado para ações que exigem atenção.
- `Erro / crítico`: usado para ações sensíveis, perigosas ou irreversíveis.
- `Bloqueado`: usado quando a mensagem precisa transmitir impedimento ou restrição.

Uso recomendado:

- Use `Informação` para ações comuns.
- Use `Aviso` para ações que merecem revisão.
- Use `Erro / crítico` para exclusão, reprovação ou ações de alto impacto.
- Use `Sucesso` para confirmações positivas, como aprovar ou concluir.
- Use `Bloqueado` quando a confirmação está ligada a uma restrição de processo.

Exemplo:

Um botão `Aprovar` pode usar `Sucesso`.

Um botão `Excluir` pode usar `Erro / crítico`.

##### Mensagem

O campo `Mensagem` define o texto exibido dentro do modal de confirmação.

Essa mensagem deve explicar o que acontecerá se o usuário confirmar.

Exemplos:

- `Tem certeza que deseja enviar esta solicitação para análise?`
- `Ao confirmar, o item será aprovado e seguirá para a próxima etapa.`
- `Esta ação irá excluir o item atual. Deseja continuar?`

Uso recomendado:

- Escreva mensagens claras e diretas.
- Informe a consequência da confirmação.
- Evite textos longos demais.
- Use linguagem compatível com o processo do cliente.

Ponto de atenção:

A mensagem é obrigatória, exceto quando for escolhido um campo para preencher no modal.

Se a mensagem e o campo estiverem vazios, a confirmação não será gravada na configuração.

##### Campo da lista principal a preencher no modal

O campo `Campo da lista principal a preencher no modal` permite escolher um campo que o usuário deverá preencher dentro da confirmação.

Essa opção transforma o modal em uma etapa rápida de coleta de informação antes de executar o botão.

Exemplos de uso:

- Pedir `Motivo da reprovação`.
- Pedir `Comentário do aprovador`.
- Pedir `Data prevista`.
- Pedir `Valor aprovado`.
- Pedir uma marcação simples de confirmação.

Tipos de campo que podem aparecer nessa lista:

- Texto.
- Texto multilinha.
- URL.
- Número.
- Moeda.
- Sim ou não.
- Data.
- Escolha.

Campos ocultos, somente leitura ou incompatíveis não devem aparecer como opção para preenchimento no modal.

Uso recomendado:

- Use quando a ação precisa registrar uma justificativa.
- Use quando o usuário precisa complementar uma informação antes de executar.
- Use para reprovação, devolução, aprovação com observação ou encerramento com comentário.
- Escolha campos claros e relacionados à ação do botão.

Exemplo:

No botão `Reprovar`, o modal pode pedir o campo `Motivo da reprovação`.

O usuário só confirma a ação depois de preencher esse campo.

##### Como a confirmação se comporta no fluxo

O modal de confirmação é executado antes de todo o restante do botão.

Fluxo esperado:

- Usuário clica no botão.
- O modal de confirmação é exibido.
- Se o usuário cancelar, o fluxo para.
- Se o usuário confirmar, o botão continua.
- Se houver campo no modal, o valor informado é usado no processo.
- Depois disso, o botão executa as ações configuradas.

Exemplo:

Um botão `Solicitar ajuste` pode:

- Abrir confirmação.
- Pedir o campo `Comentário para ajuste`.
- Confirmar a ação.
- Atualizar o status para `Ajuste solicitado`.
- Salvar o formulário.

##### Boas práticas

- Use confirmação em ações críticas.
- Não use confirmação em botões simples demais, para não cansar o usuário.
- Deixe claro o que acontecerá ao confirmar.
- Use o tipo visual correto para a gravidade da ação.
- Se pedir um campo no modal, escolha um campo diretamente relacionado ao botão.
- Teste o fluxo confirmando e cancelando antes de publicar.

##### Último passo

A seção `Último passo` define o que deve acontecer depois que o fluxo do botão terminar com sucesso.

Ela é executada somente após o botão concluir suas ações sem erro.

Essa configuração é útil para definir o encerramento da experiência do usuário depois de uma ação.

Exemplos:

- Deixar o usuário na mesma tela.
- Redirecionar para outra página.
- Limpar o formulário para novo preenchimento.

##### Quando o fluxo do botão terminar sem erro

O campo `Quando o fluxo do botão terminar sem erro` define a ação final executada após o botão concluir o fluxo.

Opções disponíveis:

- `Nada`: não executa nenhuma ação final extra.
- `Redirecionar`: envia o usuário para uma URL configurada.
- `Limpar o formulário`: limpa os dados da tela após o sucesso.

Uso recomendado:

- Use `Nada` quando o usuário deve continuar na mesma tela.
- Use `Redirecionar` quando o processo deve levar o usuário para uma página de confirmação, listagem ou acompanhamento.
- Use `Limpar o formulário` quando o usuário pode cadastrar outro item em seguida.

Exemplo:

Após clicar em `Enviar solicitação`, o formulário pode redirecionar o usuário para uma página de acompanhamento.

##### URL de redirecionamento

O campo `URL de redirecionamento` aparece quando a opção final escolhida é `Redirecionar`.

Ele define para onde o usuário será levado depois que o botão terminar com sucesso.

Exemplos:

- `/sites/portal/SitePages/Obrigado.aspx`
- `/sites/portal/SitePages/MinhasSolicitacoes.aspx`
- `/sites/portal/SitePages/Detalhe.aspx?id={{FormID}}`

Uso recomendado:

- Use uma página de confirmação quando o usuário precisa saber que a ação foi concluída.
- Use uma página de listagem quando o usuário deve acompanhar os registros.
- Use uma página de detalhe quando o usuário deve consultar o item recém-atualizado.

Ponto de atenção:

O redirecionamento só deve acontecer se o fluxo do botão terminar sem erro.

Se houver falha na ação do botão, o usuário não deve ser enviado para a página final como se tudo tivesse sido concluído.

##### Só mostrar se todos os campos obrigatórios estiverem preenchidos

A opção `Só mostrar se todos os campos obrigatórios estiverem preenchidos` controla a exibição do botão com base no preenchimento dos campos obrigatórios.

Quando marcada, o botão só aparece quando todos os campos obrigatórios estiverem preenchidos.

Quando desmarcada, o botão pode aparecer mesmo que ainda existam campos obrigatórios pendentes.

Uso recomendado:

- Use em botões finais, como `Enviar`, `Concluir` ou `Aprovar`.
- Use quando não faz sentido executar a ação antes de completar os dados obrigatórios.
- Evite usar em botões auxiliares, como `Salvar rascunho`, se o rascunho puder ser salvo incompleto.

Exemplo:

O botão `Enviar solicitação` pode ficar oculto até que todos os campos obrigatórios sejam preenchidos.

Já o botão `Salvar rascunho` pode continuar visível mesmo com campos pendentes.

##### Só autor do item

A opção `Só autor do item` mostra o botão apenas quando o usuário atual é o criador do item.

Ela compara o usuário atual com o campo `AuthorId`, que representa o autor do registro no SharePoint.

Quando marcada, o botão aparece somente para quem criou o item.

Quando desmarcada, o botão não fica limitado ao autor por essa regra.

Uso recomendado:

- Use para ações que só o solicitante original deve executar.
- Use para botões como `Cancelar minha solicitação`, `Editar minha solicitação` ou `Reenviar`.
- Evite usar quando gestores, administradores ou aprovadores também precisam executar a ação.

Exemplo:

O botão `Cancelar solicitação` pode aparecer apenas para o autor do item, impedindo que outros usuários cancelem solicitações que não criaram.

##### Grupos do SharePoint

A seção `Grupos do SharePoint` permite limitar a exibição do botão a usuários que pertencem a grupos específicos do site.

Quando nenhum grupo é selecionado, o botão não fica restrito por grupo.

Quando um ou mais grupos são selecionados, o botão passa a ser exibido apenas para usuários desses grupos.

Exemplos de grupos:

- `Gestores`
- `Aprovadores`
- `Financeiro`
- `Administradores`
- `RH`

Uso recomendado:

- Use para botões de aprovação.
- Use para ações administrativas.
- Use para botões que alteram status sensíveis.
- Use quando apenas uma área específica pode executar determinada ação.

Exemplo:

O botão `Aprovar` pode aparecer apenas para usuários do grupo `Gestores`.

O botão `Validar pagamento` pode aparecer apenas para o grupo `Financeiro`.

##### Filtrar grupos por nome

O campo `Filtrar grupos por nome` ajuda a localizar grupos do SharePoint na lista.

Ele é útil quando o site possui muitos grupos.

Uso recomendado:

- Digite parte do nome do grupo para encontrar rapidamente.
- Use o filtro antes de marcar o grupo desejado.
- Limpe o filtro para voltar a visualizar todos os grupos.

##### Grupos guardados que não estão na lista do site

Quando um grupo salvo na configuração não é encontrado na lista atual do site, ele pode aparecer como grupo guardado.

Isso indica que a configuração possui uma referência antiga ou vinda de outro ambiente.

Exemplo:

Um grupo configurado em homologação pode não existir ainda em produção.

Uso recomendado:

- Revise grupos guardados ao copiar configurações entre ambientes.
- Remova grupos que não existem mais.
- Crie o grupo no ambiente correto quando ele ainda for necessário.

##### Mensagens da lista de grupos

Durante o uso, o painel pode exibir mensagens como:

- Carregando grupos do site.
- Erro ao carregar grupos.
- Nenhum grupo corresponde ao filtro.
- Nenhum grupo no site.

Essas mensagens ajudam a entender se a lista de grupos foi carregada corretamente ou se é necessário revisar permissões e configuração do ambiente.

##### Mostrar só quando as condições abaixo forem verdadeiras

A opção `Mostrar só quando as condições abaixo forem verdadeiras` permite exibir o botão apenas quando uma ou mais regras forem atendidas.

Quando essa opção está desmarcada, o botão pode aparecer conforme as demais configurações, como modo, ativo, autor e grupos.

Quando essa opção está marcada, o botão só aparece se as condições configuradas forem verdadeiras.

Uso recomendado:

- Use quando o botão depende do valor de um campo.
- Use quando o botão só deve aparecer em um status específico.
- Use quando a ação só faz sentido em determinada etapa do processo.
- Use para reduzir botões desnecessários na tela do usuário.

Exemplo:

O botão `Aprovar` pode aparecer somente quando o campo `Status` for igual a `Em análise`.

##### Lógica entre condições

O campo `Lógica entre condições` define como o sistema deve avaliar múltiplas condições.

Opções disponíveis:

- `Todas (E)`: todas as condições precisam ser verdadeiras para o botão aparecer.
- `Pelo menos uma (OU)`: basta uma das condições ser verdadeira para o botão aparecer.

Exemplo com `Todas (E)`:

O botão `Aprovar` aparece somente quando:

- `Status` é igual a `Em análise`.
- `Valor` é menor ou igual a `5000`.

Nesse caso, as duas regras precisam ser atendidas.

Exemplo com `Pelo menos uma (OU)`:

O botão `Solicitar ajuste` aparece quando:

- `Status` é igual a `Pendente`.
- `Status` é igual a `Em revisão`.

Nesse caso, basta uma das regras ser atendida.

##### Condições nos dados do formulário

A área `Condições nos dados do formulário` exibe as regras configuradas para controlar a visibilidade do botão.

Cada regra aparece como uma condição numerada.

Exemplo:

`Condição 1`

`Condição 2`

Cada condição possui os campos `Campo`, `Operador`, `Comparar com` e `Valor`.

##### Campo

O campo `Campo` define qual campo do formulário será usado na condição.

É esse campo que o sistema irá observar para decidir se o botão deve aparecer.

Exemplo:

Selecionar o campo `Status` para mostrar o botão apenas quando o status tiver determinado valor.

##### Operador

O campo `Operador` define como o valor será avaliado.

Exemplos de operadores:

- `é igual a`
- `é diferente de`
- `contém`
- `não contém`
- `começa com`
- `termina com`
- `maior que`
- `maior ou igual a`
- `menor que`
- `menor ou igual a`
- `está vazio`
- `não está vazio`
- `é verdadeiro`
- `é falso`

Uso recomendado:

- Use `é igual a` para status, tipo ou categoria.
- Use `contém` quando o campo pode ter texto maior.
- Use `maior que` ou `menor que` para valores numéricos.
- Use `está vazio` ou `não está vazio` para verificar preenchimento.

##### Comparar com

O campo `Comparar com` define qual será a referência usada na condição.

Opções disponíveis:

- `Texto fixo`: compara o campo com um valor digitado manualmente.
- `Outro campo`: compara o campo selecionado com outro campo do formulário.
- `Token`: compara o campo com uma informação dinâmica.

Exemplo com `Texto fixo`:

Mostrar o botão quando `Status` for igual a `Em análise`.

Exemplo com `Outro campo`:

Mostrar o botão quando `Data final` for maior que `Data inicial`.

Exemplo com `Token`:

Mostrar o botão quando um campo estiver relacionado ao usuário atual ou a uma informação dinâmica disponível.

##### Valor

O campo `Valor` recebe o valor usado na comparação.

Exemplos:

- `Em análise`
- `Aprovado`
- `5000`
- `[me]`
- `{{OutroCampo}}`

O campo `Valor` fica desabilitado quando o operador escolhido não precisa de valor complementar.

Exemplos:

- `está vazio`
- `não está vazio`
- `é verdadeiro`
- `é falso`

Nesses casos, o próprio operador já define o que será verificado.

##### Adicionar condição

O botão `Adicionar condição` cria uma nova regra para controlar a exibição do botão.

Use quando o botão precisa depender de mais de um critério.

Exemplo:

O botão `Validar pagamento` pode aparecer quando:

- `Status` é igual a `Aguardando financeiro`.
- `Comprovante anexado` é verdadeiro.

##### Remover condição

O ícone de remover exclui uma condição da lista.

Ele é útil quando uma regra deixou de fazer sentido ou quando o botão deve depender de menos critérios.

Uso recomendado:

- Remova condições antigas após mudanças no processo.
- Revise a lógica `E/OU` depois de remover uma condição.
- Teste o botão para confirmar se ele aparece nos cenários corretos.

##### Exemplo completo

Cenário: mostrar o botão `Aprovar` somente para solicitações em análise com valor até `5000`.

Configuração:

- Ativar `Mostrar só quando as condições abaixo forem verdadeiras`.
- Lógica entre condições: `Todas (E)`.
- Condição 1: campo `Status`, operador `é igual a`, comparar com `Texto fixo`, valor `Em análise`.
- Condição 2: campo `Valor`, operador `menor ou igual a`, comparar com `Texto fixo`, valor `5000`.

Resultado:

O botão só aparece quando a solicitação está em análise e o valor está dentro do limite definido.

##### Boas práticas

- Use condições para simplificar a tela do usuário.
- Evite deixar botões visíveis quando a ação não pode ser executada.
- Prefira regras simples e fáceis de validar.
- Teste cada cenário esperado antes de publicar.
- Combine condições com grupos do SharePoint quando a regra depender de dados e perfil de usuário.

##### Ações por ordem

A seção `Ações por ordem` permite configurar uma sequência de ações que o botão executará quando for clicado.

As ações são executadas na ordem em que aparecem.

Isso significa que a `Ação 1` é executada antes da `Ação 2`, a `Ação 2` antes da `Ação 3`, e assim por diante.

Essa ordem é importante porque uma ação pode preparar informações que serão usadas por ações seguintes.

Exemplo:

Um botão `Enviar para análise` pode executar:

- Ação 1: mostrar o campo `Comentário`.
- Ação 2: definir o campo `Status` como `Em análise`.
- Ação 3: juntar campos em um resumo.
- Ação 4 (opcional): pedido HTTP a uma API para registrar o envio e guardar um token devolvido para usar num pedido seguinte.

Uso recomendado:

- Organize as ações na sequência real do processo.
- Coloque primeiro ações que preparam campos ou valores.
- Coloque depois ações que dependem dos valores já definidos.
- Teste a ordem antes de publicar.

##### Adicionar ação

O botão `Adicionar ação` inclui uma nova ação no final da lista.

Ao adicionar, a nova ação passa a fazer parte da sequência executada pelo botão.

Uso recomendado:

- Use quando o botão precisa executar mais de uma tarefa.
- Adicione ações aos poucos e teste o resultado.
- Evite criar uma sequência muito longa sem necessidade.

##### Remover ação

O botão `Remover ação` exclui uma ação da sequência.

Ao remover, aquela etapa deixa de ser executada quando o botão for clicado.

Uso recomendado:

- Remova ações que não fazem mais parte do processo.
- Revise a ordem das ações restantes depois da remoção.
- Verifique se outra ação não dependia da ação removida.

Exemplo:

Se uma ação definia o campo `Status` e outra ação dependia desse status, remover a primeira pode afetar o comportamento da segunda.

##### Tipo da ação

O campo `Tipo` define o que aquela ação fará.

Tipos disponíveis:

- `Mostrar campos`
- `Ocultar campos`
- `Definir valor de um campo`
- `Juntar vários campos num campo`
- `Pedido HTTP`

Cada tipo exibe configurações próprias.

#### Tipo Mostrar campos

O tipo `Mostrar campos` faz com que campos selecionados passem a aparecer no formulário quando a ação for executada.

Essa ação é útil quando alguns campos devem ficar ocultos inicialmente e só aparecer depois que o usuário clicar em um botão.

Exemplos de uso:

- Mostrar campos de aprovação.
- Mostrar campos de justificativa.
- Mostrar campos complementares.
- Mostrar campos que estavam na aba `Ocultos`.

##### Campos

A área `Campos` permite selecionar quais campos serão exibidos pela ação.

Os campos aparecem em uma lista com caixas de seleção.

Ao marcar um campo, ele passa a fazer parte da ação.

Ao desmarcar, ele deixa de ser exibido por essa ação.

Exemplo:

Um botão `Solicitar ajuste` pode mostrar os campos:

- `Motivo do ajuste`
- `Comentário para o solicitante`
- `Prazo para correção`

##### Etapa onde mostrar

O campo `Etapa onde mostrar` aparece quando o formulário possui mais de uma etapa e a ação precisa mostrar campos que estavam apenas em `Ocultos`.

Essa configuração define em qual etapa os campos devem aparecer depois da execução do botão.

Uso recomendado:

- Use quando o campo está em `Ocultos` e precisa ser exibido em uma etapa específica.
- Escolha uma etapa coerente com o assunto do campo.
- Evite mostrar campos em uma etapa que não tenha relação com a ação.

Exemplo:

Campos de aprovação podem aparecer na etapa `Análise do gestor` após o botão `Iniciar aprovação`.

#### Tipo Ocultar campos

O tipo `Ocultar campos` esconde campos selecionados quando a ação for executada.

Essa ação é útil quando determinados campos deixam de ser necessários após uma decisão do usuário.

Exemplos de uso:

- Ocultar campos de justificativa quando a solicitação for aprovada.
- Ocultar campos de edição após finalizar o processo.
- Ocultar campos auxiliares que não precisam continuar visíveis.

##### Campos

A área `Campos` permite selecionar quais campos serão ocultados pela ação.

Ao marcar um campo, ele será escondido quando o botão executar essa ação.

Exemplo:

Um botão `Aprovar` pode ocultar o campo `Motivo da reprovação`, pois ele só faz sentido quando a solicitação for reprovada.

Uso recomendado:

- Use para simplificar a tela após uma ação.
- Oculte apenas campos que realmente não precisam aparecer.
- Evite ocultar campos importantes para conferência final.

#### Tipo Definir valor de um campo

O tipo `Definir valor de um campo` preenche ou altera o valor de um campo quando o botão é executado.

Essa é uma das ações mais usadas para automatizar o processo.

Exemplos de uso:

- Alterar `Status` para `Aprovado`.
- Preencher `Data de aprovação`.
- Registrar o usuário responsável.
- Preencher um comentário padrão.

##### Campo

O campo `Campo` define qual campo será preenchido ou alterado.

Exemplo:

Selecionar o campo `Status` para que o botão altere o status da solicitação.

Uso recomendado:

- Escolha o campo que representa o resultado da ação.
- Use campos claros e diretamente relacionados ao botão.
- Evite alterar campos que o usuário não espera que sejam modificados.

##### Valor

Quando o campo selecionado é do tipo escolha, pode aparecer um campo `Valor` em formato de lista.

Essa lista permite escolher uma das opções disponíveis do campo.

Exemplo:

Para o campo `Status`, o valor pode ser:

- `Em análise`
- `Aprovado`
- `Reprovado`
- `Cancelado`

Uso recomendado:

- Use a lista quando o campo possui opções fixas.
- Escolha exatamente o valor que representa a ação do botão.
- Revise se o valor existe no campo de escolha da lista.

##### Valor fixo ou str:{{Campo}}

Quando o campo não usa lista de opções, aparece o campo `Valor fixo ou str:{{Campo}}`.

Esse campo permite informar um valor manual ou montar um valor dinâmico.

Exemplos de valor fixo:

- `Aprovado`
- `Reprovado`
- `Em análise`
- `Solicitação enviada`

Exemplos com expressão:

- `str:Solicitação enviada por [myName]`
- `str:Pedido {{Title}} aprovado`
- `str:{{Departamento}} - {{TipoSolicitacao}}`

Uso recomendado:

- Use valor fixo quando o preenchimento é sempre o mesmo.
- Use `str:` quando o valor precisa combinar texto com dados do formulário.
- Use `{{NomeInterno}}` para inserir valores de outros campos.
- Use tokens como `[myName]`, `[myEmail]` ou `[today]` quando fizer sentido.

Exemplo prático:

Um botão `Aprovar` pode definir:

- Campo: `Status`
- Valor: `Aprovado`

Outro botão pode definir:

- Campo: `Resumo`
- Valor: `str:Solicitação {{Title}} aprovada por [myName]`

#### Tipo Juntar vários campos num campo

O tipo `Juntar vários campos num campo` monta um valor combinando informações de vários campos e grava o resultado em um campo destino.

Essa ação é útil para criar resumos, títulos, descrições padronizadas ou textos compostos.

Exemplos de uso:

- Criar um resumo da solicitação.
- Montar um título padronizado.
- Juntar número, área e tipo de solicitação.
- Gerar uma descrição com informações principais.

##### Campo destino

O campo `Campo destino` define onde o resultado será gravado.

Exemplo:

Selecionar `Resumo` para receber o texto montado a partir de outros campos.

Uso recomendado:

- Use um campo destinado a receber o texto final.
- Evite sobrescrever campos preenchidos manualmente sem necessidade.
- Informe no texto de ajuda do campo quando ele for preenchido automaticamente.

##### Modelo de texto

O campo `Modelo de texto` permite escrever um texto usando placeholders.

Formato dos placeholders:

`{{NomeInterno}}`

Exemplo:

`Número: {{Numero}} — Obra: {{Title}}`

O formulário substitui os placeholders pelos valores atuais dos campos.

Exemplo prático:

Modelo:

`str:Solicitação {{Title}} da área {{Area}} para {{TipoSolicitacao}}`

Resultado esperado:

Um texto preenchido com os valores atuais desses campos.

Uso recomendado:

- Use quando precisa de um texto com estrutura específica.
- Use placeholders para inserir campos na posição correta.
- Revise os nomes internos dos campos.
- Teste com registros reais para validar o resultado.

##### Campos na ordem

A área `Campos na ordem` permite escolher quais campos serão usados na junção.

Ela também permite ordenar os campos.

Quando o `Modelo de texto` está vazio, a junção pode usar a ordem dos campos selecionados e o separador configurado.

Exemplo:

Campos na ordem:

- `Numero`
- `Title`
- `Area`

Separador:

` - `

Resultado:

`123 - Solicitação de férias - RH`

##### Adicionar campo à ordem

O campo `Adicionar campo à ordem` permite incluir mais um campo na lista de campos usados na junção.

Uso recomendado:

- Adicione apenas os campos que devem compor o resultado final.
- Ordene os campos na mesma sequência em que devem aparecer no texto.

##### Subir e descer

Os botões de subir e descer alteram a ordem dos campos selecionados.

Essa ordem influencia o resultado quando o modelo de texto está vazio.

Exemplo:

Se a ordem for `Área`, `Tipo`, `Número`, o resultado seguirá essa sequência.

##### Acrescentar placeholder ao modelo

O botão de adicionar ao lado de um campo insere o placeholder desse campo no `Modelo de texto`.

Exemplo:

Ao clicar para acrescentar o campo `Title`, o modelo recebe:

`{{Title}}`

Uso recomendado:

- Use para evitar digitar manualmente o nome interno.
- Use quando estiver montando um modelo personalizado.

##### Remover da ordem

O botão de remover exclui o campo da lista de campos usados na junção.

Ele não exclui o campo do formulário nem da lista, apenas remove esse campo da ação de junção.

##### Separador

O campo `Separador` define o texto usado entre os campos quando o `Modelo de texto` está vazio.

Exemplos:

- Espaço: ` `
- Hífen: ` - `
- Barra: ` / `
- Vírgula: `, `

Exemplo:

Campos:

- `Numero`
- `Title`
- `Area`

Separador:

` / `

Resultado:

`123 / Solicitação de férias / RH`

Ponto de atenção:

O separador é usado apenas quando o `Modelo de texto` está vazio.

Se houver um modelo preenchido, o modelo tem prioridade.

#### Tipo Pedido HTTP

O tipo `Pedido HTTP` envia um pedido web a partir do **browser do utilizador**, no momento em que o botão é executado (na sequência de **Ações por ordem**). É o análogo, em termos de configuração, a uma ação «HTTP» em ferramentas como o Power Automate: define-se método, URL, cabeçalhos, parâmetros de consulta e corpo.

**Para que serve**

- Chamar APIs externas ou internas (REST) no clique do botão.
- Obter tokens ou dados de uma resposta JSON e reutilizá-los **nas ações seguintes** (por exemplo, colocar `{{http1.token}}` num cabeçalho de um segundo pedido HTTP na mesma cadeia).
- Combinar com `Juntar vários campos` ou outros tipos de ação: valores dos campos do formulário podem entrar no URL, nos cabeçalhos ou no corpo.

**Execução e CORS**

O pedido corre no contexto da página SharePoint. O servidor de destino tem de permitir **CORS** (Cross-Origin Resource Sharing) para o domínio onde a WebPart está, caso o URL seja noutro domínio. Se a API não permitir CORS, o browser bloqueia a resposta; isso não é contornável só pela configuração do formulário.

**Identificador deste passo**

Cada ação `Pedido HTTP` precisa de um **identificador único** dentro da sequência de ações HTTP daquele botão (por exemplo `http1`, `auth`, `zapi`). Esse identificador:

- Serve para montar expressões do tipo `{{identificador.caminho}}`.
- Associa a **resposta** do pedido a esse nome. Se a resposta for JSON, `caminho` pode usar notação por pontos para propriedades aninhadas (ex.: `{{http1.data.access_token}}`).

Regras práticas: use apenas letras, números e sublinhado; não repita o mesmo identificador em dois pedidos HTTP do mesmo botão.

**Método, URL, cabeçalhos, consulta e corpo**

- **Método**: `GET`, `POST`, `PUT`, `PATCH` ou `DELETE`.
- **URL**: endereço absoluto (`https://…`) ou relativo ao site. Aceita placeholders `{{NomeInternoDoCampo}}` nos mesmos moldes das outras ações.
- **Cabeçalhos**: lista de pares chave / valor; valores podem incluir placeholders de campos e de respostas de pedidos anteriores (`{{http1.token}}`, etc.).
- **Consulta (query string)**: parâmetros que o sistema junta ao URL na execução (útil quando o URL base é fixo e os parâmetros variam).
- **Corpo**: texto livre (em geral JSON). Em `GET` ou `DELETE`, se o corpo estiver vazio, não é enviado. Placeholders funcionam como no URL e nos cabeçalhos.

**Resposta e erros**

- Se a resposta tiver corpo e for JSON válido, é interpretada como objeto para efeitos de `{{identificador.propriedade}}`.
- Se não for JSON, o resultado tratado como valor único ainda pode ser referenciado de forma coerente com o que a interface de placeholders permite.
- Códigos HTTP de erro (4xx, 5xx) interrompem a sequência e são mostrados ao utilizador; as ações seguintes não correm.

**Ordem na cadeia**

Os pedidos HTTP são **assíncronos** e respeitam a ordem das ações: o pedido da «Ação 2» só corre depois de o da «Ação 1» concluir com sucesso. Isto permite encadear, por exemplo: (1) obter token; (2) chamar outro endpoint com `Authorization: Bearer {{http1.token}}`.

**Importar a partir de cURL**

Na configuração da ação há uma área para **colar um comando cURL** (por exemplo exportado do Chrome, Postman ou terminal) e um botão para **aplicar** esse comando aos campos.

Comportamento esperado:

- Lê continuações de linha com `\` seguido de nova linha (como em comandos multilinha).
- Interpreta argumentos entre aspas simples ou duplas.
- Suporta de forma comum: `-X` / `--request`, `-H` / `--header` / `--header=…`, `-b` / `--cookie`, `-d` / `--data` / `--data-raw` / `--data-binary` (e formas `--data=…`), `--url`, e o URL como argumento posicional.
- Separa automaticamente **URL base** e **parâmetros de consulta** quando o URL traz `?key=value&…`. Se o URL contiver placeholders `{{…}}`, a divisão automática da query pode não ser aplicada (mantém-se o URL completo no campo URL).
- **Não suporta** corpo a partir de ficheiro (`-d @arquivo`); nesse caso é necessário colar o corpo em texto.

O identificador do passo **não** é alterado pela importação de cURL (mantém-se o valor que já estava configurado, por exemplo `http1`).

**Segurança e boas práticas**

- Cabeçalhos com segredos (API keys, tokens fixos) ficam gravados na configuração do formulário; restrinja quem pode editar o gestor de formulário.
- Prefira tokens obtidos por um primeiro pedido HTTP e referenciados com `{{passo.campo}}` a credenciais fixas sempre que possível.
- Teste em ambiente de homologação e confirme CORS e contratos da API antes de produção.

**Interface e linha do tempo**

Quando o botão usa confirmação modal antes de executar, a linha do tempo de execução pode listar **um passo por ação**; para `Pedido HTTP`, o texto indica o método e o identificador do passo.

#### Condição por ação

Cada ação pode ter uma condição própria.

A opção `Só executar esta ação se` permite definir se aquela ação específica deve ou não rodar.

Essa condição avalia os valores já alterados pelas ações anteriores.

Isso significa que a ordem das ações pode influenciar o resultado da condição.

Exemplo:

- Ação 1 define `Status` como `Em análise`.
- Ação 2 só executa se `Status` for igual a `Em análise`.

Nesse caso, a segunda ação pode considerar o valor definido pela primeira.

##### Como combinar as condições

Quando a condição da ação é ativada, é possível definir como as condições serão combinadas.

Opções disponíveis:

- `Uma condição`: usa apenas uma regra.
- `Todas têm de ser verdade (E)`: todas as condições precisam ser atendidas.
- `Pelo menos uma verdadeira (OU)`: basta uma condição ser atendida.

##### Campo, operador, comparar com e valor

Cada condição da ação usa a mesma lógica das condições de exibição do botão.

Ela possui:

- `Campo`: campo avaliado.
- `Operador`: comparação aplicada.
- `Comparar com`: texto fixo, outro campo ou token.
- `Valor`: valor usado na comparação, quando necessário.

Uso recomendado:

- Use condições por ação quando apenas parte do fluxo deve ser executada.
- Evite duplicar condições se o botão inteiro já possui a mesma regra.
- Use com cuidado quando ações anteriores alteram valores usados depois.

#### Exemplo completo de ações por ordem

Cenário: botão `Enviar para análise`.

Configuração:

- Ação 1: `Definir valor de um campo`
- Campo: `Status`
- Valor: `Em análise`

- Ação 2: `Definir valor de um campo`
- Campo: `DataEnvio`
- Valor: `[today]`

- Ação 3: `Juntar vários campos num campo`
- Campo destino: `Resumo`
- Modelo de texto: `str:Solicitação {{Title}} enviada por [myName] em [today]`

Resultado:

Ao clicar no botão, o formulário altera o status, registra a data de envio e monta um resumo automático da solicitação.

#### Boas práticas para ações por ordem

- Use nomes claros nos botões para indicar o que a sequência faz.
- Mantenha as ações na ordem lógica do processo.
- Teste cada ação separadamente antes de criar fluxos maiores.
- Use condições por ação apenas quando necessário.
- Revise campos ocultos antes de usar `Mostrar campos`.
- Cuidado ao usar `Ocultar campos` para não esconder informações importantes.
- Em `Definir valor de um campo`, confirme se o valor informado é compatível com o tipo do campo.
- Em `Juntar vários campos`, prefira placeholders selecionados pela interface para evitar erro no nome interno.
- Em `Pedido HTTP`, valide CORS, identificadores únicos e ordem dos pedidos; use importação cURL para reduzir erros de cópia.

### Auditoria e versões

A aba `Auditoria e versões` reúne configurações para registrar ações executadas no formulário e, quando configurado, consultar versões do item.

Essa aba é usada quando o cliente precisa acompanhar o histórico de ações, decisões, alterações ou movimentações realizadas em um registro.

#### Lista de logs

A seção `Lista de logs` configura onde os registros de log serão gravados.

O log funciona como um histórico textual das ações executadas pelos botões do formulário.

Exemplos de uso:

- Registrar que uma solicitação foi aprovada.
- Registrar que uma solicitação foi reprovada.
- Registrar que um item foi enviado para análise.
- Registrar que um usuário abriu o histórico.
- Registrar alterações feitas por um botão de atualização.

#### Lista de registo e captação

O collapse `Lista de registo e captação` define a lista onde os logs serão armazenados e habilita a captura dos registros.

Essa configuração é necessária para que o formulário consiga gravar informações de auditoria em uma lista do SharePoint.

##### Lista para registos de log

O campo `Lista para registos de log` define qual lista do SharePoint receberá os registros de auditoria.

Essa lista deve ser preparada para armazenar os textos de log e se relacionar com a lista principal do formulário.

Exemplo:

Um formulário de solicitações pode usar uma lista chamada `Logs de Solicitações`.

Cada vez que um botão importante for executado, um novo registro pode ser gravado nessa lista.

Uso recomendado:

- Use uma lista separada para armazenar logs.
- Dê um nome claro para a lista, como `Logs de Solicitações` ou `Histórico de Aprovações`.
- Evite misturar logs de processos diferentes na mesma lista sem necessidade.
- Garanta que a lista esteja disponível no mesmo contexto em que o formulário será usado.

##### Carregamento de listas

Enquanto o painel busca as listas disponíveis, pode aparecer uma mensagem de carregamento.

Exemplo:

`A carregar listas...`

Se ocorrer erro ao carregar as listas, o painel pode mostrar uma mensagem de erro.

Uso recomendado:

- Aguarde o carregamento antes de selecionar a lista.
- Se houver erro, revise permissões e disponibilidade das listas no site.

##### Campo para guardar a ação

O campo `Campo para guardar a ação` define em qual coluna da lista de logs o texto do registro será gravado.

Esse campo deve ser uma coluna de `várias linhas de texto`.

Ele é responsável por armazenar a descrição da ação executada.

Exemplo:

A lista `Logs de Solicitações` pode ter uma coluna chamada `Descrição do log`.

Quando o botão `Aprovar` for executado, essa coluna pode receber um texto informando a ação realizada.

Uso recomendado:

- Crie uma coluna de várias linhas de texto para receber o conteúdo do log.
- Use um nome claro, como `Descrição`, `Log`, `Registro` ou `Detalhes da ação`.
- Evite usar colunas de texto curto quando o log pode ter mensagens maiores.

Ponto de atenção:

Se a lista escolhida não tiver colunas de várias linhas de texto visíveis, o painel informa que será necessário criar uma coluna desse tipo.

##### Lookup para a lista principal

O campo `Lookup para a lista principal (vínculo ao item)` define qual coluna da lista de logs liga o registro de auditoria ao item principal do formulário.

Esse vínculo permite saber a qual solicitação, cadastro ou item cada log pertence.

Na prática, a lista de logs precisa ter uma coluna de lookup apontando para a lista principal do formulário.

Exemplo:

A lista principal é `Solicitações`.

A lista de logs é `Logs de Solicitações`.

Na lista de logs, existe uma coluna lookup chamada `Solicitação`, apontando para a lista `Solicitações`.

Assim, cada log fica relacionado ao item correto.

Uso recomendado:

- Configure esse campo sempre que habilitar logs.
- Garanta que a lista de logs tenha uma coluna lookup para a lista principal.
- Use nomes claros, como `Solicitação`, `Item relacionado` ou `Registro principal`.

Ponto de atenção:

Se não existir uma coluna de lookup na lista de logs apontando para a lista principal, o painel exibirá um aviso indicando que essa coluna precisa ser criada.

##### Mensagens sobre a lista principal

Durante a configuração, o painel pode tentar resolver a lista principal do formulário.

Podem aparecer mensagens como:

- Informar o título da lista principal na origem dos dados.
- Avisar que a lista principal não foi encontrada.
- Carregar a lista principal para identificar os campos de vínculo.

Essas mensagens indicam se o formulário conseguiu identificar corretamente a lista principal que será relacionada aos logs.

Uso recomendado:

- Verifique se a origem dos dados do formulário está configurada.
- Confirme se o título da lista principal está correto.
- Corrija a configuração antes de habilitar a captação de logs.

##### Habilitar captação de logs

A opção `Habilitar captação de logs` liga ou desliga a gravação dos registros de auditoria.

Quando está ativa, o formulário pode gravar logs conforme as ações configuradas.

Quando está inativa, os logs não são capturados.

Essa opção só fica disponível quando os campos obrigatórios para a captação estão configurados.

Para habilitar, é necessário definir:

- A lista de logs.
- O campo multilinhas onde a ação será gravada.
- O lookup que vincula o log ao item principal.

Uso recomendado:

- Ative somente depois de concluir a configuração da lista e dos campos.
- Use quando o processo precisa de rastreabilidade.
- Desative quando o formulário não precisa registrar ações.

Exemplo:

Ao ativar a captação, um botão `Aprovar` pode gravar um log informando que a solicitação foi aprovada por determinado usuário.

##### Mensagem de bloqueio da captação

Quando a captação ainda não pode ser ativada, o painel informa que é necessário definir a lista, o campo multilinhas e o lookup de vínculo.

Essa mensagem ajuda a identificar o que falta para liberar a funcionalidade.

Uso recomendado:

- Revise os campos obrigatórios da configuração.
- Crie as colunas necessárias na lista de logs.
- Só habilite a captação depois que o botão estiver disponível.

##### Alterações automáticas

A opção `Alterações automáticas (botões Atualizar)` registra automaticamente mudanças feitas quando um botão do tipo `Atualizar` grava o item.

Quando essa opção está ativa, o log pode incluir as diferenças efetivas entre o valor que existia ao abrir o item e o valor gravado.

Exemplo:

O item foi aberto com:

- `Status`: `Em análise`
- `Responsável`: `João`

Depois da atualização:

- `Status`: `Aprovado`
- `Responsável`: `Maria`

O log pode registrar que esses campos foram alterados.

Uso recomendado:

- Use quando o cliente precisa rastrear alterações feitas no item.
- Use em processos com aprovação, revisão ou alteração de status.
- Use quando é importante saber o que mudou, e não apenas que o botão foi clicado.

Ponto de atenção:

Se o usuário alterar um campo e depois voltar para o mesmo valor original, essa diferença não será registrada, pois o valor final não mudou.

##### Fluxo recomendado de configuração

Para configurar a lista de logs:

- Criar ou escolher uma lista para armazenar os logs.
- Criar uma coluna de várias linhas de texto para guardar a descrição da ação.
- Criar uma coluna lookup apontando para a lista principal do formulário.
- Selecionar a lista no campo `Lista para registos de log`.
- Selecionar o campo multilinhas no campo `Campo para guardar a ação`.
- Selecionar o lookup no campo `Lookup para a lista principal`.
- Ativar `Habilitar captação de logs`.
- Ativar `Alterações automáticas` se quiser registrar mudanças feitas por botões de atualização.

##### Exemplo completo

Cenário: registrar aprovações de solicitações.

Configuração:

- Lista principal do formulário: `Solicitações`.
- Lista de logs: `Logs de Solicitações`.
- Campo multilinhas na lista de logs: `Descrição do log`.
- Lookup na lista de logs: `Solicitação`.
- Captação de logs: ativa.
- Alterações automáticas: ativa.

Resultado:

Quando um botão de aprovação ou atualização for executado, o sistema poderá gravar um registro na lista de logs relacionado ao item principal da solicitação.

#### Textos de registo por botão

O collapse `Textos de registo por botão` permite configurar o texto que será gravado no log para cada botão do formulário.

Essa área ajuda a deixar os registros de auditoria mais claros para quem consulta o histórico.

Em vez de gravar apenas uma ação genérica, é possível descrever o que cada botão representa no processo.

##### Quando essa seção fica disponível

Essa seção depende da captação de logs estar configurada.

Ela também precisa ter pelo menos um botão configurado ou o botão de histórico ativo.

Se ainda não houver configuração suficiente, o painel pode exibir mensagens orientando o usuário.

Exemplos de mensagens:

- Ativar a captação na seção anterior.
- Ativar o histórico na aba `Componentes`.
- Configurar botões na aba `Botões`.

Uso recomendado:

- Configure primeiro a lista de logs.
- Ative a captação.
- Depois configure os textos de registro para cada botão.

##### Botão de histórico integrado

Quando o histórico está ativo na aba `Componentes`, a seção pode exibir o bloco `Botão de histórico (integrado)`.

Esse bloco permite configurar o registro de log relacionado à abertura ou uso do histórico.

Exemplo:

Quando o usuário abre o histórico do item, o log pode registrar uma mensagem indicando que o histórico foi consultado.

##### Botões personalizados

Para cada botão personalizado configurado na aba `Botões`, a seção exibe um bloco próprio.

Cada bloco mostra o nome do botão e seu identificador interno.

Exemplo:

Um botão chamado `Aprovar` pode aparecer com seu identificador interno ao lado.

Isso ajuda a diferenciar botões com nomes parecidos ou botões clonados.

##### Cor do registo

O campo `Cor do registo` define a cor visual usada para destacar o registro no histórico.

Essa cor segue o tema do site.

Uso recomendado:

- Use a cor principal para ações importantes.
- Use cores diferentes para separar tipos de ação.
- Use uma cor mais forte para ações críticas.
- Mantenha padrão visual entre ações parecidas.

Exemplo:

- `Aprovar`: cor principal ou positiva.
- `Reprovar`: cor de destaque mais crítica.
- `Enviar para análise`: cor neutra ou principal.

##### Texto gravado no registo de log

O editor de texto permite escrever a mensagem que será gravada no log quando aquele botão for executado.

Esse texto deve explicar o significado da ação no processo.

Exemplos:

- `Solicitação aprovada pelo gestor.`
- `Solicitação enviada para análise.`
- `Item atualizado pelo responsável.`
- `Solicitação reprovada com justificativa.`

Uso recomendado:

- Escreva mensagens claras.
- Informe o que a ação representa.
- Evite textos genéricos demais.
- Use linguagem que faça sentido para o cliente.
- Padronize mensagens entre botões parecidos.

##### Editor rico

O texto do registro é configurado em um editor rico.

Isso permite escrever uma descrição mais elaborada do que um texto simples, conforme as opções disponíveis no editor.

Uso recomendado:

- Use formatação apenas quando ajudar na leitura.
- Evite textos longos demais.
- Prefira mensagens objetivas, porque o histórico pode acumular muitos registros.

##### Exemplo de configuração

Cenário: formulário de aprovação.

Botões configurados:

- `Enviar para análise`
- `Aprovar`
- `Reprovar`

Textos de registro:

- `Enviar para análise`: `Solicitação enviada para análise.`
- `Aprovar`: `Solicitação aprovada pelo responsável.`
- `Reprovar`: `Solicitação reprovada. Consulte a justificativa informada.`

Resultado:

Quando cada botão for executado, o histórico terá registros mais claros, facilitando a auditoria do processo.

##### Boas práticas

- Configure texto para todos os botões importantes.
- Use mensagens diferentes para ações diferentes.
- Mantenha o texto curto e compreensível.
- Revise os textos quando clonar ou renomear botões.
- Teste a geração do log em homologação antes de publicar.

#### Versionamento do item

O collapse `Versionamento do item (SharePoint)` configura a exibição das versões nativas do item dentro do painel de histórico.

Essa funcionalidade usa o versionamento da própria lista principal do SharePoint.

Ela permite consultar versões anteriores do item e visualizar valores gravados em cada versão, conforme os campos configurados.

Exemplos de uso:

- Consultar como o item estava antes de uma alteração.
- Ver quando uma versão foi criada.
- Identificar a versão atual.
- Comparar valores registrados em versões anteriores.

Ponto de atenção:

Para essa funcionalidade funcionar, o versionamento precisa estar ativo nas configurações da lista principal do SharePoint.

#### No painel de histórico

O collapse `No painel de histórico` define se as versões do item serão exibidas dentro do histórico do formulário.

##### Mostrar versões do item no painel de histórico

A opção `Mostrar versões do item no painel de histórico` ativa ou desativa a exibição das versões do SharePoint no painel de histórico.

Quando marcada, o painel de histórico passa a incluir a seção de versionamento do item.

Quando desmarcada, o painel não mostra as versões nativas da lista.

Uso recomendado:

- Ative quando o cliente precisa consultar versões anteriores do item.
- Ative quando alterações de campos precisam ser rastreadas ao longo do tempo.
- Desative quando o processo só precisa dos logs por botão.
- Desative quando a lista principal não usa versionamento.

Exemplo:

Em uma solicitação que passa por várias alterações, o histórico pode mostrar a versão atual e versões anteriores do item.

##### Dependência do SharePoint

Essa opção usa o histórico de versões nativo da lista principal.

Por isso, o versionamento deve estar habilitado na lista do SharePoint.

Caminho conceitual:

- Acessar as configurações da lista.
- Abrir as configurações de versão.
- Ativar o controle de versões da lista.

Uso recomendado:

- Confirme com o administrador do SharePoint se o versionamento está ativo.
- Valide em homologação antes de liberar para o cliente.
- Combine com logs por botão quando quiser histórico técnico e histórico funcional.

#### Campos ao expandir uma versão

O collapse `Campos ao expandir uma versão` define quais campos serão carregados quando o usuário expandir uma versão no painel de histórico.

Quando uma versão é expandida, o sistema busca os valores daquela versão para os campos configurados.

Essa configuração ajuda a controlar quais informações serão exibidas e evita carregar campos desnecessários.

##### Lista principal necessária

Para listar os campos disponíveis, o formulário precisa saber qual é a lista principal.

Se a lista principal ainda não estiver definida na origem dos dados, o painel informa que é necessário indicar essa lista.

Uso recomendado:

- Configure corretamente a lista principal na origem dos dados.
- Verifique se a lista principal é a mesma usada pelo formulário.
- Corrija a origem antes de configurar os campos de versionamento.

##### Carregamento dos campos da lista principal

Enquanto o painel busca os campos da lista principal, pode aparecer uma mensagem de carregamento.

Exemplo:

`A carregar campos da lista principal...`

Se houver erro, o painel exibe uma mensagem de erro.

Uso recomendado:

- Aguarde o carregamento dos campos.
- Se houver erro, revise permissões, título da lista e configuração da origem dos dados.

##### Todos os campos compatíveis

A opção `Todos os campos compatíveis com o pedido REST das versões` usa automaticamente os campos visíveis que podem ser consultados nas versões.

Essa é a opção padrão.

Com essa opção marcada, o sistema escolhe os campos compatíveis para exibir quando uma versão for expandida.

Uso recomendado:

- Use quando não há necessidade de escolher campo por campo.
- Use para configuração mais rápida.
- Use quando o cliente quer uma visão geral das versões.

Ponto de atenção:

Nem todos os tipos de campo aparecem nessa consulta.

Campos como lookup, pessoa e taxonomia são excluídos da lista de campos elegíveis para esse pedido de versões.

##### Escolher campos individuais

Ao desmarcar `Todos os campos compatíveis`, o painel permite selecionar campos específicos.

Essa opção é útil quando o cliente quer ver apenas campos relevantes ao expandir uma versão.

Exemplos de campos úteis:

- `Status`
- `Responsável`
- `Data de aprovação`
- `Valor`
- `Observações`

Uso recomendado:

- Use quando há muitos campos na lista.
- Use quando o histórico deve mostrar apenas dados importantes.
- Use para facilitar a leitura do painel de versões.
- Evite selecionar campos que não ajudam na análise do histórico.

##### Marcar todos

O link `Marcar todos` seleciona todos os campos elegíveis disponíveis.

Use quando quiser partir de uma seleção completa e depois remover apenas alguns campos.

##### Limpar

O link `Limpar` remove a seleção manual de campos.

Use quando quiser recomeçar a escolha dos campos ou deixar a lista vazia antes de selecionar novamente.

##### Lista de campos

Quando a seleção manual está ativa, o painel exibe uma lista de campos com caixas de seleção.

Cada item mostra o título do campo e o nome interno.

Exemplo:

`Status (Status)`

`Data de aprovação (DataAprovacao)`

Uso recomendado:

- Marque apenas campos que o usuário precisa consultar em versões anteriores.
- Prefira campos de status, datas, valores e observações relevantes.
- Evite excesso de campos para não deixar a expansão da versão pesada ou difícil de ler.

##### Campos não elegíveis

Alguns tipos de campo não aparecem na lista de seleção por limitação do pedido de versões.

Tipos excluídos:

- Lookup.
- Lookup múltiplo.
- Pessoa.
- Pessoa múltipla.
- Taxonomia.
- Taxonomia múltipla.

Também são ignorados campos ocultos ou campos cujo nome interno não é compatível com o padrão necessário para consulta.

Se não houver campos elegíveis, o painel informa que nenhum campo pode ser selecionado.

##### Limite de campos

Ao selecionar campos individuais, existe um limite de campos usados na consulta de versões.

O sistema limita a quantidade para evitar consultas muito grandes.

Uso recomendado:

- Selecione apenas os campos realmente importantes.
- Evite tentar transformar o versionamento em uma cópia completa do item.
- Use logs de auditoria para registrar descrições funcionais e versões para consultar valores principais.

#### Como aparece no histórico

Quando o versionamento está ativo, o painel de histórico pode exibir uma seção `Versionamento do item (SharePoint)`.

Nessa seção, o usuário pode ver as versões disponíveis.

Cada versão pode mostrar:

- Número ou rótulo da versão.
- Data de criação da versão.
- Indicação da versão atual.
- Campos carregados ao expandir a versão.

Ao expandir uma versão, o painel carrega os campos daquela versão e mostra os valores registrados.

Exemplo:

O usuário expande a versão `3.0` e vê que naquela versão o campo `Status` estava como `Em análise`.

Depois, expande a versão atual e vê que o `Status` passou para `Aprovado`.

#### Diferença entre logs e versionamento

Logs e versionamento têm finalidades diferentes.

Logs registram ações funcionais, como:

- Quem aprovou.
- Quem reprovou.
- Qual botão foi executado.
- Qual mensagem foi gravada no histórico.

Versionamento mostra versões do item na lista principal, como:

- Valores salvos em versões anteriores.
- Data de criação de cada versão.
- Estado do item em determinado momento.

Uso recomendado:

- Use logs para explicar o significado das ações.
- Use versionamento para consultar valores históricos do item.
- Use os dois juntos quando o processo precisa de auditoria mais completa.

#### Exemplo completo

Cenário: acompanhar alterações em uma solicitação.

Configuração:

- Ativar `Mostrar versões do item no painel de histórico`.
- Manter `Todos os campos compatíveis` ativo ou selecionar campos específicos.
- Garantir que o versionamento da lista principal esteja ativo no SharePoint.

Campos selecionados:

- `Status`
- `Data de aprovação`
- `Valor`
- `Observações`

Resultado:

No painel de histórico, o usuário consegue consultar logs funcionais e também expandir versões do item para ver como os campos estavam em versões anteriores.

#### Boas práticas

- Ative versionamento apenas quando houver necessidade real de consultar versões.
- Confirme se o versionamento está habilitado na lista do SharePoint.
- Selecione poucos campos quando a lista tiver muitos dados.
- Use campos relevantes para auditoria, como status, datas, valores e observações.
- Combine com `Textos de registo por botão` para ter histórico funcional claro.
- Teste a expansão de versões em homologação antes de liberar para o cliente.

### Listas vinculadas

A aba `Listas vinculadas` permite configurar listas secundárias dentro do formulário principal.

Essas listas secundárias funcionam como blocos de registros relacionados ao item principal.

Exemplos de uso:

- Uma solicitação principal com várias despesas vinculadas.
- Um pedido principal com vários itens de produto.
- Um cadastro principal com várias linhas de dependentes.
- Um processo principal com várias tarefas associadas.

Na prática, o formulário principal grava o item principal e também permite trabalhar com linhas de outra lista SharePoint que estejam ligadas a esse item.

#### Antes de configurar

Para usar listas vinculadas, é necessário ter:

- Uma lista principal configurada no formulário.
- Uma lista filha no SharePoint.
- Um campo Lookup na lista filha apontando para a lista principal.

Esse campo Lookup é o que cria a ligação entre o item principal e os itens filhos.

Exemplo:

Lista principal: `Solicitações`

Lista filha: `Itens da solicitação`

Campo Lookup na lista filha: `Solicitação`

Resultado:

Cada item da lista `Itens da solicitação` fica associado a uma solicitação específica.

#### Avisos da aba

Se a lista principal ainda não estiver definida na configuração da origem de dados, a aba mostra um aviso pedindo para configurar primeiro a lista principal.

Isso acontece porque a ligação da lista filha depende da lista principal.

Também pode aparecer um aviso quando a lista principal informada não é encontrada no site.

Nesse caso, os campos Lookup podem não ser filtrados corretamente, porque o sistema não consegue confirmar qual lista deve ser usada como referência.

Uso recomendado:

- Configure primeiro a origem de dados do formulário principal.
- Confirme se o nome da lista principal está correto.
- Confirme se a lista principal existe no mesmo contexto esperado.
- Só depois configure as listas vinculadas.

#### Botão Adicionar lista vinculada

O botão `Adicionar lista vinculada` cria um novo bloco de lista filha.

Cada bloco representa uma lista secundária que poderá aparecer dentro do formulário principal.

Ao adicionar, o sistema cria um bloco inicialmente sem lista selecionada.

Depois disso, é necessário abrir o bloco e configurar a lista filha e o campo Lookup de ligação.

Uso recomendado:

- Adicione uma lista vinculada para cada conjunto de registros filhos.
- Use quando a informação precisa ter várias linhas relacionadas ao mesmo item principal.
- Evite usar lista vinculada para informações simples que poderiam ser apenas campos do formulário principal.

Exemplo:

Em um formulário de pedido, os dados gerais ficam na lista principal.

Os produtos do pedido ficam em uma lista vinculada, porque um mesmo pedido pode ter vários produtos.

#### Bloco de lista vinculada

Cada lista vinculada aparece como um bloco com o título:

`Lista vinculada 1: (sem título)`

Depois que uma lista filha é escolhida, o título passa a mostrar o nome da lista.

Exemplo:

`Lista vinculada 1: Itens da solicitação`

O bloco pode ser aberto para mostrar suas configurações.

Também existem ações no topo do bloco:

- `Mover bloco para cima`.
- `Mover bloco para baixo`.
- `Remover bloco`.

Essas ações controlam a ordem dos blocos de listas vinculadas dentro da configuração.

##### Mover bloco para cima

O botão `Mover bloco para cima` sobe o bloco da lista vinculada uma posição na lista de blocos configurados.

Use quando aquela lista vinculada deve aparecer antes de outra.

Exemplo:

Se o formulário possui os blocos `Despesas` e `Participantes`, mas o cliente precisa preencher primeiro os participantes, use `Mover bloco para cima` no bloco `Participantes`.

Quando o bloco já está na primeira posição, o botão fica indisponível, porque não há posição acima para mover.

##### Mover bloco para baixo

O botão `Mover bloco para baixo` desce o bloco da lista vinculada uma posição na lista de blocos configurados.

Use quando aquela lista vinculada deve aparecer depois de outra.

Exemplo:

Se o bloco `Análises` deve ser preenchido somente depois de `Itens da solicitação`, use `Mover bloco para baixo` até chegar na posição desejada.

Quando o bloco já está na última posição, o botão fica indisponível, porque não há posição abaixo para mover.

##### Como a movimentação se comporta

A movimentação altera apenas a ordem dos blocos de listas vinculadas na configuração.

Ela não altera:

- A lista SharePoint vinculada.
- O campo Lookup de ligação.
- Os campos configurados dentro do bloco.
- As regras condicionais da lista vinculada.
- Os itens já cadastrados na lista filha.

Uso recomendado:

- Organize os blocos na ordem em que o usuário deve preencher as informações.
- Coloque primeiro as listas mais importantes para o processo.
- Deixe listas complementares ou opcionais mais abaixo.
- Revise a ordem depois de adicionar novas listas vinculadas.

##### Exemplo de ordem dos blocos

Cenário: formulário de solicitação de compra.

Ordem recomendada:

- `Itens da compra`.
- `Cotações`.
- `Aprovações complementares`.

Resultado:

O usuário primeiro informa os itens, depois adiciona cotações e, por último, consulta ou preenche informações complementares de aprovação.

##### Remover bloco

O botão `Remover bloco` exclui a configuração daquela lista vinculada do formulário.

Use quando a lista filha não deve mais aparecer no formulário principal.

Ao remover o bloco, o formulário deixa de exibir aquela lista vinculada e suas configurações deixam de fazer parte da configuração do formulário.

Isso inclui:

- A lista filha selecionada no bloco.
- O campo Lookup configurado para ligação ao principal.
- A apresentação configurada para aquele bloco.
- Os campos adicionados na etapa geral.
- As regras dos campos da lista vinculada.
- As regras condicionais daquele bloco.
- As configurações de anexos por linha daquele bloco.

Ponto de atenção:

Remover o bloco remove a configuração da lista vinculada no formulário, mas não significa apagar automaticamente a lista SharePoint nem os itens já existentes nela.

Ou seja, a lista SharePoint continua existindo e os dados já gravados nela também continuam existindo.

O que muda é que o formulário deixa de usar aquele bloco.

Uso recomendado:

- Remova o bloco apenas quando tiver certeza de que aquela lista vinculada não será mais usada no formulário.
- Antes de remover, confirme se os dados existentes na lista filha não precisam mais ser exibidos nesse formulário.
- Se a remoção for feita por engano, será necessário configurar novamente a lista vinculada.
- Revise essa alteração em homologação antes de aplicar em produção.

Exemplo:

Um formulário tinha a lista vinculada `Cotações`, mas o processo mudou e as cotações passaram a ser registradas em outro sistema.

Nesse caso, o botão `Remover bloco` pode ser usado para retirar essa lista vinculada do formulário.

#### Lista e ligação ao principal

O collapse `Lista e ligação ao principal` é a primeira configuração essencial de uma lista vinculada.

Ele define qual lista SharePoint será usada como lista filha e qual campo cria a relação com a lista principal.

Sem essa ligação, o formulário não sabe quais linhas pertencem ao item principal.

##### Lista filha (SharePoint)

O campo `Lista filha (SharePoint)` permite escolher a lista secundária que será vinculada ao formulário principal.

Essa lista será usada para armazenar os registros filhos.

Exemplos:

- `Itens da solicitação`
- `Despesas`
- `Participantes`
- `Tarefas do processo`

Quando a lista filha é selecionada, o sistema carrega os campos dessa lista.

Esse carregamento é necessário para encontrar os campos disponíveis e permitir escolher o Lookup que aponta para a lista principal.

Ponto de atenção:

Ao trocar a lista filha, o campo Lookup de ligação é limpo.

Isso acontece porque cada lista pode ter campos diferentes, então a ligação precisa ser escolhida novamente.

Uso recomendado:

- Escolha uma lista criada especificamente para guardar os registros filhos daquele processo.
- Confirme se a lista filha tem um campo Lookup para a lista principal.
- Evite reutilizar listas que tenham dados de processos diferentes sem uma separação clara.

##### Carregamento dos campos da lista

Depois de escolher a lista filha, pode aparecer a mensagem:

`A carregar campos da lista...`

Enquanto essa mensagem aparece, o sistema está buscando os campos da lista filha.

Após o carregamento, o campo de Lookup fica disponível para seleção.

Se a lista ainda não foi escolhida, o sistema indica que é necessário escolher primeiro a lista filha.

##### Campo Lookup para a lista principal

O campo `Campo Lookup para a lista principal` define qual campo da lista filha aponta para o item da lista principal.

Esse é o campo mais importante da ligação.

Ele garante que cada linha filha fique associada ao item principal correto.

Exemplo:

Lista principal: `Solicitações`

Lista filha: `Itens da solicitação`

Campo Lookup na lista filha: `Solicitação`

Quando o usuário abre a solicitação número 25, o formulário consegue carregar apenas os itens da lista filha cujo campo `Solicitação` aponta para essa solicitação.

##### Quando não aparece nenhum Lookup

Se a lista filha não tiver um campo Lookup apontando para a lista principal, o sistema pode mostrar a indicação de que não existe campo Lookup para a lista principal.

Nesse caso, a lista filha precisa ser ajustada no SharePoint.

Uso recomendado:

- Crie na lista filha um campo Lookup para a lista principal.
- Confirme se o Lookup aponta para a lista correta.
- Depois volte ao configurador e selecione novamente a lista filha.

##### Campo Lookup salvo na configuração

Quando uma configuração antiga ou importada por JSON possui um campo Lookup salvo, o sistema tenta manter esse campo mesmo que ele não apareça imediatamente na lista filtrada.

Isso ajuda em cenários de migração entre ambientes.

Exemplo:

Uma configuração feita em homologação é copiada para produção via JSON.

Se o campo Lookup tiver o mesmo nome interno, a configuração pode ser preservada.

Ponto de atenção:

Mesmo quando o campo aparece como salvo na configuração, é recomendado revisar a ligação no ambiente final.

##### Rótulo JSON legado

Em configurações antigas, pode existir um rótulo salvo para a lista vinculada.

Quando isso acontece, o painel mostra:

`Rótulo JSON legado`

Também aparece a opção `Usar só o nome da lista`.

Essa opção remove o rótulo antigo e passa a usar apenas o nome da lista como identificação do bloco.

Uso recomendado:

- Use o nome da lista quando ele já for claro para o usuário.
- Remova rótulos antigos quando eles não fizerem mais sentido.
- Revise essa informação depois de importar configurações por JSON.

#### Apresentação no formulário

O collapse `Apresentação no formulário` define como a lista vinculada aparece para o usuário dentro do formulário principal.

Essa configuração não muda a ligação entre as listas.

Ela controla a ordem do bloco, a quantidade permitida de linhas e o formato visual usado para mostrar os registros filhos.

Exemplos de uso:

- Mostrar despesas em formato de tabela.
- Mostrar participantes em blocos.
- Mostrar tarefas em cartões.
- Limitar a quantidade de itens que podem ser cadastrados.

##### Ordem de exibição

O campo `Ordem de exibição` define a posição da lista vinculada em relação a outras listas vinculadas.

Quando está em `Automático (0)`, o sistema usa a ordem padrão dos blocos.

Também é possível escolher um número para controlar a ordem manualmente.

Quanto menor o número, mais cedo o bloco tende a aparecer.

Exemplo:

- `0`: dados complementares.
- `1`: itens da solicitação.
- `2`: anexos ou despesas relacionadas.

Uso recomendado:

- Use `Automático (0)` quando houver apenas uma lista vinculada.
- Use números quando houver várias listas vinculadas e a ordem for importante para o processo.
- Mantenha uma sequência simples para facilitar manutenção.

##### Mínimo de linhas

O campo `Mínimo de linhas` define a quantidade mínima de registros filhos que o usuário deve ter naquele bloco.

Quando está em `0 (sem mínimo obrigatório)`, o usuário não é obrigado a cadastrar linhas nessa lista vinculada.

Quando recebe outro valor, o formulário passa a exigir pelo menos essa quantidade de linhas.

Exemplo:

Se `Mínimo de linhas` for `1`, o usuário precisa cadastrar pelo menos uma linha filha.

Uso recomendado:

- Use `0` quando a lista vinculada for opcional.
- Use `1` quando o processo sempre precisa de pelo menos um item.
- Use valores maiores apenas quando houver uma regra de negócio clara.

Exemplo prático:

Em uma solicitação de compra, pode ser obrigatório cadastrar pelo menos um item da compra.

Nesse caso, configure `Mínimo de linhas` como `1`.

##### Máximo de linhas

O campo `Máximo de linhas` define a quantidade máxima de registros filhos permitidos naquele bloco.

Quando está em `Sem limite`, o usuário pode adicionar quantas linhas forem necessárias.

Quando recebe um número, o formulário limita a quantidade de linhas daquela lista vinculada.

Exemplo:

Se `Máximo de linhas` for `5`, o usuário poderá cadastrar até cinco registros filhos.

Uso recomendado:

- Use `Sem limite` quando não houver restrição de quantidade.
- Defina um limite quando o processo tiver uma regra clara.
- Use limite para evitar cadastros excessivos quando o bloco deve ter poucas linhas.

Exemplo prático:

Em um cadastro de dependentes, a empresa pode permitir no máximo `4` dependentes.

Nesse caso, configure `Máximo de linhas` como `4`.

##### Apresentação das linhas

O campo `Apresentação das linhas` define o formato visual dos registros filhos no formulário.

Opções disponíveis:

- `Blocos (em coluna)`.
- `Tabela`.
- `Compacto`.
- `Cartões`.

Cada formato atende melhor a um tipo de uso.

##### Blocos em coluna

A opção `Blocos (em coluna)` mostra cada linha filha como um bloco separado, um abaixo do outro.

É uma apresentação boa quando cada registro tem vários campos ou precisa de mais espaço para leitura.

Uso recomendado:

- Use quando os registros filhos têm muitos campos.
- Use quando o usuário precisa preencher cada linha com atenção.
- Use quando o formulário deve ficar mais parecido com um mini-formulário por item.

Exemplo:

Cadastro de participantes, onde cada participante possui nome, documento, telefone e observações.

##### Tabela

A opção `Tabela` mostra os registros filhos em formato de linhas e colunas.

É útil quando os dados são mais objetivos e precisam ser comparados rapidamente.

Uso recomendado:

- Use para itens de pedido.
- Use para despesas.
- Use quando os campos principais cabem bem em colunas.
- Use quando o usuário precisa visualizar várias linhas ao mesmo tempo.

Exemplo:

Lista de produtos com `Produto`, `Quantidade`, `Valor unitário` e `Total`.

##### Compacto

A opção `Compacto` mostra as linhas de forma mais reduzida.

Ela ajuda quando o bloco precisa ocupar menos espaço na tela.

Uso recomendado:

- Use quando os registros filhos têm poucos campos.
- Use quando o usuário precisa apenas de uma visão resumida.
- Use quando o formulário principal já tem muitas informações.

Exemplo:

Lista simples de contatos adicionais com nome e e-mail.

##### Cartões

A opção `Cartões` mostra cada registro filho com aparência de cartão.

É útil quando cada linha representa uma unidade visual importante, como uma tarefa, etapa ou item de análise.

Uso recomendado:

- Use quando a leitura individual de cada item é importante.
- Use quando deseja uma visualização mais destacada.
- Use para listas com poucos ou médios registros.

Exemplo:

Tarefas de um processo, onde cada cartão mostra título, responsável, status e prazo.

##### Pré-visualização

Abaixo da seleção de apresentação, o painel mostra uma pré-visualização do formato escolhido.

Essa pré-visualização ajuda a entender como as linhas vinculadas ficarão no formulário antes de salvar a configuração.

Uso recomendado:

- Troque entre os formatos e observe a pré-visualização.
- Escolha o formato mais fácil para o cliente preencher e consultar.
- Valide a apresentação com campos reais da lista filha.

#### Exemplo de apresentação

Cenário: pedido de compra com itens vinculados.

Configuração recomendada:

- `Ordem de exibição`: `1`.
- `Mínimo de linhas`: `1`.
- `Máximo de linhas`: `Sem limite`.
- `Apresentação das linhas`: `Tabela`.

Resultado:

O formulário exige pelo menos um item de compra e mostra os itens em tabela, facilitando a comparação entre produto, quantidade e valor.

#### Boas práticas para apresentação

- Use tabela para listas com dados curtos e comparáveis.
- Use blocos quando cada linha tem muitos campos.
- Use compacto para economizar espaço.
- Use cartões quando cada registro precisa de destaque visual.
- Defina mínimo e máximo apenas quando houver regra de negócio.
- Teste a apresentação em tela pequena e tela grande antes de liberar.

#### Anexos e biblioteca por linha

O collapse `Anexos e biblioteca (por linha)` define se cada registro da lista vinculada poderá ter arquivos próprios.

Essa configuração é independente dos anexos do item principal.

Ela serve para cenários em que cada linha filha precisa ter seus próprios documentos.

Exemplos de uso:

- Cada despesa possui seu próprio comprovante.
- Cada item de compra possui sua própria cotação.
- Cada tarefa possui seus próprios arquivos de evidência.
- Cada participante possui documentos específicos.

##### Modo

O campo `Modo` define como os anexos da linha filha serão tratados.

Opções disponíveis:

- `Sem anexos neste bloco`.
- `Anexos nativos (lista filha)`.
- `Biblioteca da aba Anexos (pastas herdadas; Lookup à lista filha)`.
- `Outra biblioteca (estrutura própria)`.

##### Sem anexos neste bloco

A opção `Sem anexos neste bloco` desativa anexos para as linhas da lista vinculada.

Use quando os registros filhos não precisam receber arquivos.

Exemplo:

Uma lista vinculada de participantes pode ter apenas nome, cargo e e-mail, sem necessidade de anexos.

Uso recomendado:

- Use quando os arquivos ficam apenas no item principal.
- Use quando a lista filha armazena somente dados simples.
- Use para manter o formulário mais limpo.

##### Anexos nativos da lista filha

A opção `Anexos nativos (lista filha)` usa o recurso padrão de anexos do SharePoint na própria lista filha.

Nesse modo, cada item filho pode ter seus arquivos anexados diretamente nele.

Exemplo:

Na lista filha `Despesas`, cada despesa pode ter um comprovante anexado ao próprio item da despesa.

Uso recomendado:

- Use quando a lista filha precisa guardar arquivos simples por linha.
- Use quando não há necessidade de organizar arquivos em biblioteca com pastas.
- Use quando o processo aceita o comportamento padrão de anexos do SharePoint.

Ponto de atenção:

Esse modo grava os arquivos como anexos do item da lista filha, não como documentos em uma biblioteca.

##### Biblioteca da aba Anexos

A opção `Biblioteca da aba Anexos (pastas herdadas; Lookup à lista filha)` usa a mesma biblioteca configurada na aba `Anexos` do formulário principal.

Ela permite reaproveitar a biblioteca e a estrutura de pastas já configuradas para o formulário.

Nesse modo, os arquivos das linhas filhas ficam na biblioteca de documentos, mas precisam de um campo Lookup que aponte para a lista filha.

Exemplo:

A aba `Anexos` principal usa a biblioteca `Documentos do processo`.

A lista vinculada `Itens da compra` também pode gravar arquivos nessa mesma biblioteca, desde que exista um Lookup na biblioteca apontando para `Itens da compra`.

##### Quando essa opção aparece

A opção de herdar a biblioteca da aba `Anexos` só fica disponível quando a aba `Anexos` do formulário principal está configurada com:

- Destino em `Biblioteca de documentos`.
- Título da biblioteca preenchido.
- Campo Lookup para a lista principal configurado.

Se isso não estiver configurado, o painel mostra um aviso informando que é necessário configurar a biblioteca na aba `Anexos`.

Uso recomendado:

- Use quando o projeto quer centralizar arquivos em uma única biblioteca.
- Use quando os anexos do item principal e das linhas filhas devem seguir uma organização parecida.
- Use quando a estrutura de pastas da aba `Anexos` também faz sentido para as listas vinculadas.

##### Campo Lookup na biblioteca de Anexos

Quando o modo herdado é usado, aparece o campo `Campo Lookup na biblioteca de Anexos (referência à lista filha)`.

Esse campo define qual coluna da biblioteca de documentos aponta para a lista filha.

Ele é necessário para que cada arquivo seja ligado à linha filha correta.

Exemplo:

Biblioteca: `Documentos do processo`

Lista filha: `Itens da compra`

Campo Lookup na biblioteca: `Item da compra`

Resultado:

O arquivo enviado em uma linha específica da lista vinculada fica relacionado àquela linha, e não apenas ao item principal.

Ponto de atenção:

O Lookup da biblioteca para a lista principal não substitui o Lookup para a lista filha.

Para anexos por linha, a biblioteca precisa conseguir identificar a linha filha.

##### Outra biblioteca

A opção `Outra biblioteca (estrutura própria)` permite usar uma biblioteca diferente para os anexos daquela lista vinculada.

Esse modo é útil quando os arquivos da lista filha precisam ficar separados dos arquivos do item principal.

Exemplo:

Os anexos principais ficam na biblioteca `Documentos do processo`.

Os comprovantes das despesas ficam na biblioteca `Comprovantes de despesas`.

Uso recomendado:

- Use quando os arquivos da lista filha têm regra de organização diferente.
- Use quando precisam ficar em uma biblioteca separada.
- Use quando a lista vinculada tem uma estrutura própria de pastas.

##### Biblioteca de documentos

Quando o modo `Outra biblioteca` é escolhido, aparece o campo `Biblioteca de documentos`.

Esse campo define em qual biblioteca os arquivos da lista vinculada serão guardados.

Ao trocar a biblioteca, o campo Lookup da biblioteca é limpo, porque cada biblioteca pode ter colunas diferentes.

Uso recomendado:

- Escolha a biblioteca correta antes de configurar o Lookup.
- Confirme se a biblioteca existe no site.
- Evite trocar a biblioteca depois que a configuração já estiver em uso.

##### Campo Lookup na biblioteca

No modo `Outra biblioteca`, aparece o campo `Campo Lookup na biblioteca (referência à lista filha)`.

Esse campo define qual coluna da biblioteca aponta para a lista filha.

Sem esse campo, o sistema não consegue saber a qual linha filha o arquivo pertence.

Exemplo:

Biblioteca: `Comprovantes de despesas`

Lista filha: `Despesas`

Campo Lookup na biblioteca: `Despesa`

Resultado:

Cada comprovante enviado fica ligado à despesa correspondente.

##### Pastas

No modo `Outra biblioteca`, também é possível configurar a estrutura de pastas da biblioteca.

Essa estrutura define onde os arquivos serão organizados dentro da biblioteca.

Quando a lista vinculada tem mais de uma etapa configurada, pode aparecer a escolha de etapa para a pasta.

Isso permite organizar arquivos de acordo com partes específicas do formulário da lista filha.

Uso recomendado:

- Use pastas quando houver muitos arquivos.
- Use nomes claros para facilitar localização.
- Evite criar estruturas profundas demais.
- Configure pastas conforme a forma como o cliente procura os documentos.

##### Exemplo com anexos nativos

Cenário: lista vinculada de despesas.

Configuração:

- `Modo`: `Anexos nativos (lista filha)`.

Resultado:

Cada despesa pode receber um comprovante diretamente como anexo do item da lista filha.

##### Exemplo com biblioteca herdada

Cenário: processo com biblioteca central de documentos.

Configuração:

- Aba `Anexos` principal usando biblioteca de documentos.
- Lista vinculada usando `Biblioteca da aba Anexos`.
- Campo Lookup na biblioteca apontando para a lista filha.

Resultado:

Os arquivos do item principal e os arquivos das linhas filhas ficam na mesma biblioteca, mantendo uma organização centralizada.

##### Exemplo com outra biblioteca

Cenário: despesas com comprovantes separados.

Configuração:

- `Modo`: `Outra biblioteca (estrutura própria)`.
- `Biblioteca de documentos`: `Comprovantes de despesas`.
- `Campo Lookup na biblioteca`: `Despesa`.

Resultado:

Os comprovantes ficam separados em uma biblioteca própria, ligados diretamente à despesa correspondente.

#### Boas práticas para anexos em listas vinculadas

- Use `Sem anexos neste bloco` quando não houver necessidade real de arquivos por linha.
- Use anexos nativos para cenários simples.
- Use biblioteca quando precisar organizar documentos com pastas.
- Garanta que a biblioteca tenha Lookup para a lista filha.
- Não confunda o Lookup para a lista principal com o Lookup para a lista filha.
- Teste o envio de arquivos em uma linha filha antes de publicar.
- Revise essa configuração ao copiar JSON entre ambientes.

#### Campos na etapa Geral

O collapse `Campos na etapa «Geral» (ordem)` define quais campos da lista filha aparecem no bloco da lista vinculada e em qual ordem.

Essa configuração funciona como a montagem do mini-formulário da lista filha.

Ela permite escolher os campos que o usuário irá preencher ou visualizar em cada linha vinculada.

Exemplos de uso:

- Em `Itens da compra`, mostrar produto, quantidade, valor e observação.
- Em `Despesas`, mostrar tipo, data, valor e comprovante.
- Em `Participantes`, mostrar nome, e-mail, cargo e telefone.

##### Nenhum campo

Quando ainda não há campos configurados, o painel mostra a mensagem:

`Nenhum campo. Adicione abaixo.`

Isso indica que a lista vinculada já existe, mas ainda não possui campos visíveis configurados para a etapa geral.

Uso recomendado:

- Depois de escolher a lista filha, adicione os campos que o usuário precisa preencher.
- Evite deixar a lista vinculada sem campos, porque o bloco não terá utilidade prática no formulário.

##### Lista de campos adicionados

Cada campo adicionado aparece em uma linha dentro do collapse.

A linha mostra:

- O nome amigável do campo.
- O nome interno do campo.
- Botões para ordenar, remover e configurar regras.

Exemplo:

`Produto (Produto)`

`Quantidade (Quantidade)`

`Valor unitário (ValorUnitario)`

O nome interno ajuda a identificar o campo correto, principalmente quando existem campos com nomes parecidos.

##### Subir

O botão `Subir` move o campo uma posição para cima.

Use quando o campo deve aparecer antes dos demais no formulário da lista vinculada.

Exemplo:

Se `Quantidade` deve aparecer antes de `Valor unitário`, use `Subir` até chegar na posição desejada.

##### Descer

O botão `Descer` move o campo uma posição para baixo.

Use quando o campo deve aparecer depois dos demais.

Exemplo:

O campo `Observações` normalmente pode ficar no final do bloco, depois dos campos principais.

##### Remover

O botão `Remover` tira o campo da etapa geral da lista vinculada.

Isso remove o campo da configuração visual do formulário, mas não apaga a coluna da lista SharePoint.

Uso recomendado:

- Remova campos que não precisam aparecer para o usuário.
- Remova campos técnicos ou preenchidos automaticamente.
- Tenha cuidado para não remover campos importantes para o processo.

Ponto de atenção:

Ao remover um campo, as configurações específicas dele dentro daquela lista vinculada também podem deixar de ser usadas.

##### Regras

O botão `Regras...` abre o painel de regras daquele campo da lista vinculada.

Esse painel permite configurar comportamentos do campo, como exibição, validação, valor padrão, transformação e outras regras conforme o tipo de campo.

Ele funciona de forma parecida com as regras dos campos do formulário principal, mas aplicado ao campo da lista filha.

Exemplos:

- Tornar `Quantidade` obrigatória.
- Definir valor padrão para `Status`.
- Ocultar um campo conforme uma condição.
- Aplicar regra de validação em um campo de data.

Uso recomendado:

- Configure regras apenas quando houver uma necessidade clara.
- Use regras para facilitar o preenchimento da linha filha.
- Evite excesso de regras para não dificultar a manutenção.

##### Adicionar campo à etapa Geral

O seletor `Adicionar campo à etapa Geral...` permite incluir novos campos da lista filha no bloco.

Ao selecionar um campo, ele é adicionado ao final da lista de campos configurados.

Depois disso, é possível reorganizar usando `Subir` e `Descer`.

Campos que já foram adicionados não aparecem novamente para seleção.

Também não são oferecidos campos internos ou campos que não devem ser configurados diretamente, como campos técnicos do SharePoint e o campo usado para ligar a lista filha à lista principal.

Uso recomendado:

- Adicione somente campos necessários para o usuário.
- Organize os campos na mesma ordem em que devem ser preenchidos.
- Coloque campos principais no início.
- Deixe observações e campos complementares no final.

##### Sem campos disponíveis

Quando não há campos disponíveis para adicionar, o seletor mostra:

`— sem campos disponíveis —`

Isso pode acontecer quando todos os campos elegíveis já foram adicionados ou quando a lista filha não possui campos configuráveis disponíveis.

Uso recomendado:

- Verifique se a lista filha correta foi selecionada.
- Verifique se a lista filha possui colunas além dos campos técnicos.
- Confirme se o campo Lookup de ligação não está sendo esperado como campo visual.

#### Exemplo de campos na etapa

Cenário: lista vinculada `Itens da compra`.

Campos adicionados:

- `Produto`.
- `Quantidade`.
- `Valor unitário`.
- `Total`.
- `Observações`.

Ordem recomendada:

- Primeiro os campos que identificam o item.
- Depois quantidade e valores.
- Por último observações.

Resultado:

Cada linha da lista vinculada apresenta apenas os campos necessários para o usuário preencher os itens da compra.

#### Boas práticas para campos na etapa

- Adicione apenas campos úteis para o processo.
- Ordene os campos conforme a sequência natural de preenchimento.
- Remova campos técnicos da visualização.
- Use `Regras...` para validar campos importantes.
- Revise os campos após trocar a lista filha.
- Teste a criação e edição de linhas vinculadas em homologação.

#### Regras condicionais

O collapse `Regras condicionais` permite criar regras aplicadas somente aos campos da lista filha.

Essas regras controlam o comportamento dos campos dentro das linhas da lista vinculada.

Elas são úteis para mostrar, ocultar, obrigar, desativar ou exibir mensagens conforme o valor preenchido em outro campo da própria lista filha.

Exemplos de uso:

- Mostrar `Justificativa` quando `Status` for `Reprovado`.
- Tornar `Comprovante` obrigatório quando `Tipo de despesa` for `Reembolso`.
- Desativar `Valor aprovado` para usuários que não pertencem ao grupo de aprovadores.
- Exibir uma mensagem de aviso quando `Valor` for maior que um limite.

##### Regras condicionais só da lista filha

As regras desse collapse trabalham apenas com campos da lista filha configurada naquele bloco.

Elas não são regras gerais do formulário principal.

Isso significa que uma regra criada na lista vinculada `Despesas` afeta os campos das despesas, não os campos da solicitação principal.

Uso recomendado:

- Use para regras específicas das linhas filhas.
- Use quando o comportamento muda dentro de cada item vinculado.
- Não use para controlar campos do formulário principal.

##### Nova regra

O botão `Nova regra` cria uma regra condicional em branco.

Uma regra é formada por duas partes:

- `Quando`: define a condição.
- `Então`: define o que acontece quando a condição for verdadeira.

Exemplo:

Quando `Status` for igual a `Reprovado`, então mostrar o campo `Justificativa`.

##### Modelos prontos

O painel oferece modelos para criar regras comuns mais rapidamente.

Modelos disponíveis:

- `Modelo: mostrar B quando A = valor`.
- `Modelo: mostrar B quando A contém texto`.
- `Modelo: mostrar B quando A ≠ valor`.
- `Modelo: mostrar B quando A > número`.
- `Modelo: obrigar B quando A = valor`.

Esses modelos já criam uma estrutura inicial de regra.

Depois de aplicar o modelo, o usuário deve ajustar os campos e valores conforme o processo.

Uso recomendado:

- Use modelos para acelerar configurações simples.
- Revise sempre os campos `A`, `B` e o valor comparado.
- Ajuste o efeito se o modelo não representar exatamente a regra desejada.

##### Cartão da regra

Cada regra aparece como um cartão.

No topo do cartão, o painel mostra um resumo da condição configurada.

Também aparecem os botões:

- `Duplicar`.
- `Excluir`.

##### Duplicar

O botão `Duplicar` cria uma cópia da regra.

Use quando precisar criar uma regra parecida, mudando apenas o campo, valor ou efeito.

Exemplo:

Uma regra mostra `Justificativa` quando `Status` for `Reprovado`.

Você pode duplicar e ajustar para mostrar outro campo quando `Status` for `Cancelado`.

##### Excluir

O botão `Excluir` remove a regra condicional.

Use quando a regra não deve mais ser aplicada.

Ponto de atenção:

Excluir a regra remove o comportamento configurado, mas não apaga os campos da lista filha.

##### Quando

A seção `Quando` define a condição que precisa ser verdadeira para a regra ser aplicada.

Ela possui os campos:

- `Campo`.
- `Operador`.
- `Comparar com`.
- `Valor`.

##### Campo

O campo `Campo` define qual campo da lista filha será analisado.

Exemplo:

`Status`

`Tipo de despesa`

`Valor`

##### Operador

O campo `Operador` define como a comparação será feita.

Opções disponíveis:

- `é igual a`.
- `é diferente de`.
- `contém`.
- `não contém`.
- `começa com`.
- `termina com`.
- `maior que`.
- `maior ou igual a`.
- `menor que`.
- `menor ou igual a`.
- `está vazio`.
- `não está vazio`.
- `é verdadeiro`.
- `é falso`.

Uso recomendado:

- Use `é igual a` para status, escolhas e valores exatos.
- Use `contém` para textos livres.
- Use `maior que` ou `menor que` para números e valores.
- Use `está vazio` ou `não está vazio` para validar preenchimento.

##### Comparar com

O campo `Comparar com` define de onde vem o valor usado na comparação.

Opções disponíveis:

- `Texto fixo`.
- `Outro campo`.
- `Token`.

##### Texto fixo

Use `Texto fixo` quando a comparação será feita com um valor digitado diretamente.

Exemplo:

Campo `Status` é igual a `Reprovado`.

##### Outro campo

Use `Outro campo` quando o valor de um campo deve ser comparado com o valor de outro campo da mesma lista filha.

Exemplo:

Campo `Data final` menor que `Data inicial`.

##### Token

Use `Token` quando a comparação precisa usar um valor dinâmico.

Exemplo:

Comparar uma data com um token de data atual, quando aplicável ao comportamento configurado.

##### Valor

O campo `Valor` recebe o valor usado na comparação.

Ele pode ficar desativado em operadores que não precisam de valor, como:

- `está vazio`.
- `não está vazio`.
- `é verdadeiro`.
- `é falso`.

Nesses casos, a própria condição já é suficiente.

##### Incluir grupos SharePoint

O campo `Incluir: grupos SharePoint (títulos, vírgula)` limita a aplicação da regra a usuários de determinados grupos.

Quando fica vazio, a regra pode ser aplicada para qualquer usuário.

Quando preenchido, a regra só vale para usuários que pertencem a pelo menos um dos grupos informados.

Exemplo:

`Aprovadores, Gestores`

Uso recomendado:

- Use quando uma regra deve valer apenas para aprovadores, gestores ou equipes específicas.
- Separe os grupos por vírgula.
- Use exatamente o título do grupo no SharePoint.

##### Excluir grupos SharePoint

O campo `Excluir: grupos SharePoint (títulos, vírgula)` impede que a regra seja aplicada para usuários de determinados grupos.

Quando fica vazio, nenhum grupo é excluído.

Quando preenchido, usuários pertencentes aos grupos informados não recebem a regra.

Exemplo:

`Administradores`

Uso recomendado:

- Use para liberar exceções.
- Use quando administradores ou responsáveis não devem sofrer a mesma limitação dos demais usuários.

##### Então

A seção `Então` define o que acontece quando a condição for verdadeira.

Cada regra pode ter um ou mais efeitos.

O painel mostra uma linha por combinação de efeito e campo.

##### Efeito

O campo `Efeito` define a ação aplicada pela regra.

Opções disponíveis:

- `Mostrar campo`.
- `Ocultar campo`.
- `Tornar obrigatório`.
- `Tornar opcional`.
- `Desativar campo`.
- `Ativar campo`.
- `Somente leitura`.
- `Permitir edição`.
- `Exibir mensagem`.

##### Campo alvo

Quando o efeito atua sobre um campo, aparece o campo `Campo alvo`.

Ele define qual campo da lista filha será afetado.

Exemplo:

Quando `Status` for `Reprovado`, o campo alvo pode ser `Justificativa`.

##### Mostrar campo

O efeito `Mostrar campo` exibe o campo alvo quando a condição for verdadeira.

Use quando um campo só deve aparecer em determinadas situações.

##### Ocultar campo

O efeito `Ocultar campo` esconde o campo alvo quando a condição for verdadeira.

Use para simplificar o formulário e esconder campos que não fazem sentido naquele caso.

##### Tornar obrigatório

O efeito `Tornar obrigatório` exige o preenchimento do campo alvo quando a condição for verdadeira.

Exemplo:

Se `Tipo de despesa` for `Reembolso`, tornar `Comprovante` obrigatório.

##### Tornar opcional

O efeito `Tornar opcional` remove a obrigatoriedade do campo alvo quando a condição for verdadeira.

Use quando um campo só deve ser obrigatório em alguns cenários.

##### Desativar campo

O efeito `Desativar campo` impede a edição do campo alvo.

Use quando o usuário deve ver o campo, mas não deve alterá-lo.

##### Ativar campo

O efeito `Ativar campo` libera a edição do campo alvo.

Use quando um campo começa bloqueado e só deve ser editado em determinada condição.

##### Somente leitura

O efeito `Somente leitura` deixa o campo visível, mas sem permitir alteração.

Use quando a informação precisa ser consultada sem ser modificada.

##### Permitir edição

O efeito `Permitir edição` libera o campo para edição.

Use em conjunto com regras que bloqueiam campos em alguns cenários.

##### Exibir mensagem

O efeito `Exibir mensagem` mostra uma mensagem ao usuário quando a condição for verdadeira.

Nesse caso, em vez de `Campo alvo`, aparecem:

- `Tipo`.
- `Texto`.

##### Tipo da mensagem

O campo `Tipo` define o estilo da mensagem.

Opções disponíveis:

- `Info`.
- `Aviso`.
- `Erro`.

Uso recomendado:

- Use `Info` para orientação.
- Use `Aviso` para atenção.
- Use `Erro` para situações que impedem ou indicam preenchimento incorreto.

##### Texto da mensagem

O campo `Texto` define a mensagem exibida ao usuário.

Exemplo:

`Informe a justificativa para continuar.`

Uso recomendado:

- Escreva mensagens curtas.
- Explique o que o usuário precisa fazer.
- Evite textos técnicos.

##### Adicionar efeito

O botão `Adicionar efeito` adiciona mais uma ação à mesma condição.

Isso permite que uma única regra execute mais de um comportamento.

Exemplo:

Quando `Status` for `Reprovado`:

- Mostrar `Justificativa`.
- Tornar `Justificativa` obrigatório.
- Exibir mensagem de aviso.

##### Remover efeito

O botão `Remover efeito` remove uma ação da regra.

Use quando uma regra deve continuar existindo, mas um dos comportamentos não é mais necessário.

##### Prévia

A linha `Prévia` mostra quantas regras internas serão geradas a partir da configuração visual.

Essa informação ajuda a entender se a regra está gerando um ou mais comportamentos no motor do formulário.

##### Nenhuma regra condicional

Quando não existe nenhuma regra criada, o painel mostra:

`Nenhuma regra condicional nesta lista vinculada.`

Isso significa que os campos da lista filha seguirão o comportamento padrão, sem condições adicionais nesse bloco.

##### Regras só no motor

Pode aparecer a seção `Regras só no motor (não editadas por esta UI)`.

Ela mostra regras que existem na configuração, mas que não são editadas por essa interface visual.

Uso recomendado:

- Revise essas regras ao importar JSON de outro ambiente.
- Tenha cuidado antes de alterar manualmente configurações avançadas.
- Se a regra não for compreendida pela interface, valide o comportamento em homologação.

#### Exemplo de regras condicionais

Cenário: lista vinculada `Despesas`.

Regra:

- Quando `Tipo de despesa` for igual a `Reembolso`.
- Então mostrar `Comprovante`.
- Então tornar `Comprovante` obrigatório.

Resultado:

O campo de comprovante só ganha destaque quando a despesa exige comprovação.

#### Boas práticas para regras condicionais

- Use regras simples e fáceis de entender.
- Prefira uma condição clara por regra.
- Use modelos prontos quando possível.
- Teste regras com diferentes valores preenchidos.
- Valide regras com usuários de grupos SharePoint diferentes.
- Evite criar muitas regras sobre o mesmo campo sem necessidade.

#### Exemplo completo

Cenário: solicitação de compra com vários itens.

Configuração:

- Lista principal: `Solicitações de compra`.
- Lista filha: `Itens da compra`.
- Campo Lookup para a lista principal: `Solicitação de compra`.

Resultado:

Ao criar ou editar uma solicitação de compra, o formulário pode exibir um bloco para cadastrar os itens daquela compra.

Cada item cadastrado na lista filha fica ligado à solicitação principal pelo campo Lookup.

#### Boas práticas

- Sempre crie o Lookup na lista filha antes de configurar a lista vinculada.
- Use nomes claros para listas e campos.
- Valide a ligação em homologação criando um item principal com linhas filhas.
- Evite configurar listas vinculadas sem uma relação real de pai e filho.
- Revise a ligação depois de copiar JSON entre ambientes.

### Quebra de permissões

A aba `Quebra de permissões` permite configurar permissões específicas para os itens criados ou atualizados pelo formulário.

No SharePoint, normalmente um item herda as permissões da lista onde está armazenado.

Com essa funcionalidade, o formulário pode quebrar essa herança e aplicar permissões próprias no item.

Exemplos de uso:

- Permitir que somente o autor e os aprovadores vejam uma solicitação.
- Dar leitura para um grupo e edição para outro.
- Aplicar permissões também em itens de listas vinculadas.
- Aplicar permissões em arquivos enviados para bibliotecas.
- Dar acesso automaticamente para pessoas escolhidas em campos do formulário.

Ponto de atenção:

Essa é uma configuração sensível.

Quando usada incorretamente, pode impedir usuários de acessar itens ou arquivos.

Use sempre primeiro em homologação.

#### Quando a quebra é aplicada

A quebra de permissões é aplicada depois que o formulário grava o item.

Em fluxos de criação ou atualização, ela entra depois da gravação dos dados, sincronização de anexos e listas vinculadas, conforme o que estiver configurado no formulário.

Na prática, o sistema precisa primeiro ter o item criado ou atualizado para depois aplicar permissões nele.

Uso recomendado:

- Configure e teste com itens reais em homologação.
- Valide com usuários de perfis diferentes.
- Confirme se os usuários corretos conseguem acessar o item depois da gravação.

#### Ativar

O botão `Ativar` liga ou desliga toda a configuração de quebra de permissões.

Quando está desligado, a aba mostra apenas uma mensagem informando que é necessário ativar para configurar alvos e principais.

Quando está ligado, aparecem as configurações de:

- Comportamento da quebra.
- Alvos.
- Principais e níveis.

Uso recomendado:

- Ative somente quando o processo realmente precisa de permissões específicas.
- Deixe desativado quando a herança normal da lista SharePoint for suficiente.

#### Copiar permissões herdadas ao quebrar

A opção `Copiar permissões herdadas ao quebrar (primeira vez)` controla como o SharePoint inicia a quebra de herança quando o item ainda herda permissões.

Quando ativada, o SharePoint inicia a quebra copiando permissões herdadas.

Quando desativada, a quebra começa sem depender das permissões herdadas.

Mesmo assim, o objetivo final da configuração é aplicar os principais definidos na aba.

Uso recomendado:

- Mantenha desativada na maioria dos casos.
- Ative apenas se houver orientação clara de manter o comportamento inicial do SharePoint ao quebrar.
- Teste sempre, porque permissões herdadas podem trazer acessos além do esperado.

#### Manter autor

A opção `Manter autor (Created By) com nível abaixo` define se o autor do item continuará recebendo uma permissão explícita após a quebra.

Quando ativada, o usuário que criou o item recebe o nível definido em `Nível do autor`.

Quando desativada, o autor não recebe permissão automaticamente por essa opção.

Uso recomendado:

- Mantenha ativada quando o autor precisa acompanhar ou editar sua solicitação.
- Desative apenas quando o processo exige que o autor perca acesso após enviar.
- Valide esse comportamento com o cliente antes de publicar.

#### Nível do autor

O campo `Nível do autor` define qual permissão será dada ao autor quando a opção `Manter autor` estiver ativa.

Opções disponíveis:

- `Leitura`.
- `Contribuir`.
- `Editar`.
- `Controlo total`.

Uso recomendado:

- Use `Leitura` quando o autor só precisa acompanhar.
- Use `Contribuir` quando o autor precisa editar ou complementar informações.
- Use `Editar` apenas quando o processo permitir alterações mais amplas.
- Evite `Controlo total`, salvo quando houver necessidade administrativa.

#### Alvos

A seção `Alvos` define onde a quebra de permissões será aplicada.

Os alvos podem incluir:

- Item da lista principal.
- Itens filhos das listas vinculadas.
- Ficheiros na biblioteca de anexos do item principal.
- Ficheiros em bibliotecas usadas por listas vinculadas.

#### Item na lista principal

A opção `Item na lista principal` aplica a quebra de permissões no item principal do formulário.

Esse é o alvo mais comum.

Exemplo:

Em uma lista `Solicitações`, a permissão será aplicada diretamente na solicitação criada ou atualizada.

Uso recomendado:

- Mantenha ativado quando o objetivo é proteger o item principal.
- Desative apenas se a quebra deve ser aplicada somente em itens filhos ou arquivos.

Ponto de atenção:

Quando os anexos do item principal usam anexos nativos da lista, eles acompanham a aplicação ligada ao item principal.

#### Listas vinculadas

Quando existem listas vinculadas configuradas, aparece a seção `Listas vinculadas (itens filhos)`.

Ela permite escolher em quais listas filhas a quebra será aplicada.

Se todas estiverem marcadas, a quebra pode ser aplicada nos itens filhos dessas listas.

Se uma lista vinculada for desmarcada, os itens daquela lista não recebem a quebra por essa configuração.

Uso recomendado:

- Marque listas vinculadas que contêm dados sensíveis.
- Desmarque listas vinculadas que não precisam de controle específico.
- Revise essa seção sempre que adicionar ou remover listas vinculadas.

Exemplo:

O item principal tem duas listas vinculadas:

- `Itens da compra`.
- `Cotações`.

Se apenas `Cotações` tiver documentos sensíveis, é possível aplicar quebra somente nessa lista vinculada.

#### Ficheiros na biblioteca de anexos

A opção `Ficheiros na biblioteca de anexos (lookup ao item principal)` aparece quando os anexos do formulário principal estão configurados para uma biblioteca de documentos.

Quando marcada, a quebra também é aplicada nos arquivos da biblioteca relacionados ao item principal.

Exemplo:

O item principal está na lista `Solicitações`.

Os anexos estão na biblioteca `Documentos do processo`, ligados ao item pelo campo Lookup.

Ao marcar essa opção, os arquivos ligados à solicitação também recebem permissões específicas.

Uso recomendado:

- Marque quando os arquivos precisam ter a mesma proteção do item principal.
- Use quando anexos contêm informações sensíveis.
- Teste se os arquivos continuam acessíveis para os usuários corretos.

#### Ficheiros na biblioteca por lista vinculada

Quando uma lista vinculada usa anexos em biblioteca de documentos, aparece a seção `Ficheiros na biblioteca por lista vinculada (lookup à linha filha)`.

Ela permite aplicar quebra nos arquivos ligados a cada linha filha.

Exemplo:

A lista vinculada `Despesas` usa a biblioteca `Comprovantes de despesas`.

Cada comprovante fica ligado a uma despesa por Lookup.

Ao marcar essa lista vinculada nessa seção, os comprovantes também recebem permissões específicas.

Uso recomendado:

- Marque quando os anexos da linha filha também precisam ser protegidos.
- Confirme se a biblioteca possui Lookup correto para a lista filha.
- Revise após configurar `Anexos e biblioteca (por linha)`.

#### Principais e níveis

A seção `Principais e níveis` define quem receberá permissão e qual nível será aplicado.

Cada linha representa uma concessão de permissão.

Uma configuração pode ter várias linhas.

Exemplos:

- Grupo `Aprovadores` com nível `Editar`.
- Grupo `Solicitantes` com nível `Leitura`.
- Pessoa específica com nível `Controlo total`.
- Pessoas selecionadas em um campo do formulário com nível `Contribuir`.

#### Filtrar grupos por nome

O campo `Filtrar grupos por nome (dropdown «Grupo»)` ajuda a localizar grupos do site.

Ao digitar parte do nome, a lista do campo `Grupo` fica filtrada.

Uso recomendado:

- Use quando o site possui muitos grupos.
- Digite parte do nome do grupo.
- Confira se o grupo selecionado é o grupo correto do SharePoint.

#### Adicionar principal

O botão `Adicionar principal` cria uma nova linha de permissão.

Por padrão, a nova linha começa como `Grupo do site` com nível `Leitura`, quando houver grupo disponível.

Depois de adicionar, é possível alterar:

- Tipo.
- Nível.
- Grupo, pessoa ou campo.

Uso recomendado:

- Adicione uma linha para cada grupo, pessoa ou campo que deve receber acesso.
- Evite adicionar permissões duplicadas sem necessidade.
- Revise todas as linhas antes de salvar a configuração.

#### Tipo

O campo `Tipo` define de onde virá o principal que receberá a permissão.

Opções disponíveis:

- `Grupo do site`.
- `Pessoa do Site`.
- `Campo (Pessoa / Vários)`.

#### Grupo do site

O tipo `Grupo do site` permite escolher um grupo SharePoint do site.

Quando selecionado, aparece o campo `Grupo`.

Esse grupo receberá o nível de permissão escolhido na linha.

Exemplo:

Grupo `Aprovadores de Compras` com nível `Editar`.

Uso recomendado:

- Use grupos sempre que possível.
- Prefira grupos a pessoas individuais para facilitar manutenção.
- Garanta que o grupo tenha os membros corretos no SharePoint.

#### Pessoa do Site

O tipo `Pessoa do Site` permite escolher uma pessoa específica.

Quando selecionado, aparece o campo `Pessoa do Site (pesquisar)`.

A pesquisa começa quando o texto digitado tem pelo menos dois caracteres.

Depois de selecionar uma pessoa, o painel mostra quem foi selecionado.

Uso recomendado:

- Use para exceções pontuais.
- Evite usar muitas pessoas individuais quando um grupo atender melhor.
- Revise se a pessoa selecionada é a correta antes de salvar.

#### Campo Pessoa ou Vários

O tipo `Campo (Pessoa / Vários)` usa o valor de um campo Pessoa do formulário para definir quem receberá acesso.

Esse campo pode estar:

- Na lista principal.
- Em uma lista vinculada.

O sistema usa os usuários preenchidos nesse campo para conceder permissão.

Exemplo:

Campo `Responsável` com nível `Contribuir`.

Quando o item for salvo, o usuário informado em `Responsável` recebe permissão no alvo configurado.

Uso recomendado:

- Use quando a permissão depende de quem foi selecionado no formulário.
- Use para responsáveis, aprovadores, gestores ou participantes.
- Garanta que o campo esteja preenchido antes da gravação.

#### Lista

Quando o tipo é `Campo (Pessoa / Vários)`, aparece o campo `Lista`.

Ele define onde está o campo Pessoa que será usado.

Opções possíveis:

- `Lista principal`.
- Uma das listas vinculadas configuradas.

Se `Lista principal` for selecionada, o campo Pessoa será buscado na lista principal.

Se uma lista vinculada for selecionada, o campo Pessoa será buscado naquela lista filha.

Uso recomendado:

- Use `Lista principal` para responsáveis gerais do item.
- Use listas vinculadas quando cada linha filha possui seu próprio responsável.

#### Campo

O campo `Campo` aparece quando o tipo é `Campo (Pessoa / Vários)`.

Ele permite escolher um campo de pessoa ou pessoa múltipla.

Somente campos compatíveis aparecem nessa lista.

Exemplos:

- `Responsável`.
- `Aprovadores`.
- `Participantes`.

Ponto de atenção:

Se a lista vinculada ainda estiver carregando os campos, pode aparecer um carregamento antes da lista ficar disponível.

#### Nível

O campo `Nível` define qual permissão será concedida para aquela linha.

Opções disponíveis:

- `Leitura`.
- `Contribuir`.
- `Editar`.
- `Controlo total`.

Uso recomendado:

- `Leitura`: consultar o item ou arquivo.
- `Contribuir`: colaborar com o item, conforme permissões SharePoint.
- `Editar`: editar com mais liberdade.
- `Controlo total`: administrar permissões e configurações, usar com muita cautela.

#### Remover linha de permissão

O botão com ícone de lixeira remove uma linha de permissão da seção `Principais e níveis`.

Use quando aquele grupo, pessoa ou campo não deve mais receber permissão.

Ponto de atenção:

Remover a linha não apaga o grupo, usuário ou campo no SharePoint.

Remove apenas aquela concessão da configuração da quebra de permissões.

#### Como a permissão final é formada

Ao aplicar a quebra, o sistema monta a permissão final com base em:

- Autor do item, se `Manter autor` estiver ativo.
- Linhas configuradas em `Principais e níveis`.
- Alvos selecionados na seção `Alvos`.

Depois, aplica permissões únicas nos itens e arquivos selecionados.

Uso recomendado:

- Sempre configure pelo menos um grupo ou pessoa com acesso suficiente.
- Evite deixar apenas o autor com acesso se outras pessoas precisam atuar no processo.
- Garanta que administradores ou equipe de suporte mantenham acesso quando necessário.

#### Exemplo com item principal

Cenário: solicitação sigilosa.

Configuração:

- `Ativar`: ligado.
- `Item na lista principal`: marcado.
- `Manter autor`: ligado.
- `Nível do autor`: `Leitura`.
- Grupo `Aprovadores` com nível `Editar`.

Resultado:

O autor consegue consultar a solicitação e os aprovadores conseguem trabalhar nela.

#### Exemplo com campo Pessoa

Cenário: cada solicitação possui um responsável.

Configuração:

- Tipo: `Campo (Pessoa / Vários)`.
- Lista: `Lista principal`.
- Campo: `Responsável`.
- Nível: `Contribuir`.

Resultado:

A pessoa preenchida no campo `Responsável` recebe permissão automaticamente no item.

#### Exemplo com listas vinculadas

Cenário: uma solicitação possui despesas vinculadas.

Configuração:

- Alvo `Listas vinculadas`: marcar `Despesas`.
- Tipo: `Grupo do site`.
- Grupo: `Financeiro`.
- Nível: `Editar`.

Resultado:

Os itens da lista vinculada `Despesas` recebem permissões específicas para o grupo financeiro.

#### Exemplo com arquivos em biblioteca

Cenário: anexos do processo ficam em biblioteca de documentos.

Configuração:

- `Ficheiros na biblioteca de anexos`: marcado.
- Grupo `Aprovadores` com nível `Leitura`.
- Grupo `Gestores` com nível `Editar`.

Resultado:

Os arquivos relacionados ao item principal também passam a respeitar permissões específicas.

#### Boas práticas para quebra de permissões

- Teste sempre em homologação antes de usar em produção.
- Comece com poucos alvos e poucas linhas de permissão.
- Prefira grupos SharePoint em vez de pessoas individuais.
- Use campos Pessoa quando o acesso depende do responsável preenchido no formulário.
- Evite `Controlo total` salvo em cenários administrativos.
- Confirme se anexos e listas vinculadas realmente precisam de quebra.
- Garanta que pelo menos um grupo de suporte ou administração mantenha acesso.
- Revise a configuração ao copiar JSON entre ambientes.


