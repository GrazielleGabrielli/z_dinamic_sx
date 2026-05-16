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

