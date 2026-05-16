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

