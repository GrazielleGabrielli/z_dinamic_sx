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

