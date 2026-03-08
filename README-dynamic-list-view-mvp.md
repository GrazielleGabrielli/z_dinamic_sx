# Dynamic View Engine for SharePoint — MVP (Modo Lista)

## Visão Geral

Este projeto tem como objetivo criar uma **webpart SPFx dinâmica** para SharePoint, orientada por **configuração em JSON**, capaz de se conectar a uma **lista ou biblioteca** e renderizar, no MVP, um **modo lista** com:

- dashboard opcional
- listagem dinâmica
- paginação performática
- configuração guiada por tela amigável
- JSON final persistido na configuração da webpart

A proposta é que o usuário **não precise editar propriedades técnicas diretamente** no painel clássico da webpart. Em vez disso, a webpart abrirá uma **experiência guiada em steps**, e ao final gerará um **JSON de configuração** que será salvo como base da renderização.

---

## Diretriz principal do MVP

No MVP, a webpart terá **3 modos previstos em arquitetura**:

- `list`
- `projectManagement`
- `formManager`

Porém, **somente o modo `list` ficará habilitado para seleção e uso**.

Isso permite preparar a base técnica para evolução futura sem complicar a primeira entrega.

---

## Estratégia de configuração

### Princípio

A configuração da webpart será baseada em **JSON**, não em listas de configuração no SharePoint.

### Motivos

- evita criar listas auxiliares para configuração
- facilita exportação/importação de layouts
- simplifica versionamento da estrutura da view
- permite criar uma **tela user friendly** que monta o JSON internamente
- facilita evoluções futuras para outros modos

### Fluxo esperado

1. usuário adiciona a webpart na página
2. a webpart ainda não possui configuração
3. ao entrar em edição/configuração, o usuário verá um **wizard por steps**
4. esse wizard coleta as decisões do usuário
5. ao concluir, a webpart gera e salva um **JSON de configuração principal**
6. a renderização passa a obedecer esse JSON

---

## Regras do MVP para configuração inicial

Quando a webpart for implantada pela primeira vez, ela **não terá nenhuma informação pronta no painel de edição tradicional**.

A configuração inicial será feita em uma tela guiada por steps.

### Step 1 — Fonte de dados
Definir:
- se a origem é **lista** ou **biblioteca**
- qual o nome da lista/biblioteca

O sistema deverá:
- validar se a origem existe
- identificar metadados básicos
- descobrir automaticamente se é lista ou biblioteca
- carregar campos e views disponíveis para uso futuro

### Step 2 — Modo da webpart
Exibir os 3 modos planejados:
- modo lista
- gestão de projeto / kanban
- formulário + gestor

No MVP:
- somente `modo lista` pode ser selecionado
- os demais devem aparecer como “em breve” ou desabilitados

### Step 3 — Dashboard
Definir:
- se deseja dashboard ou não
- se sim, quantos cards deseja

No MVP, essa etapa pode configurar a estrutura inicial dos cards, mesmo que o detalhamento fino de cada card seja refinado depois.

### Step 4 — Paginação
Definir:
- quantidade de itens por página
- opções permitidas para troca de tamanho da página

Faixa prevista no MVP:
- `[5, 10, 20, 50, 100]`

---

## Faz sentido esse fluxo?

Sim. Esse fluxo faz total sentido para o MVP porque:

- reduz a complexidade inicial
- evita lotar o property pane com dezenas de campos técnicos
- prepara uma experiência mais profissional e comercializável
- centraliza a lógica de configuração num JSON único
- deixa a webpart pronta para ganhar telas de configuração mais ricas depois

---

## Escopo funcional do MVP — Modo Lista

### Entradas principais
- seleção da origem de dados
- identificação de lista ou biblioteca
- carregamento de fields
- carregamento de views
- JSON de configuração

### Saídas principais
- dashboard opcional
- tabela dinâmica
- paginação server-side
- renderização de campos compatíveis com SharePoint

### Tipos de campo previstos no MVP
- texto
- múltiplas linhas
- número
- moeda
- data
- booleano
- choice
- lookup simples
- pessoa/usuário
- hyperlink
- nome do arquivo
- metadados básicos de biblioteca

---

## Arquitetura conceitual do MVP

### 1. Configuration Wizard
Responsável por:
- guiar a configuração inicial por etapas
- validar entradas
- gerar o JSON final

### 2. Config Engine
Responsável por:
- interpretar o JSON salvo
- validar estrutura
- fornecer configuração tipada para os componentes

### 3. Metadata Engine
Responsável por:
- ler metadados da lista/biblioteca
- identificar tipo da origem
- carregar fields e views
- mapear tipos do SharePoint para tipos internos

### 4. Data Engine
Responsável por:
- montar consultas
- aplicar select, expand, filter, orderby
- executar paginação server-side
- normalizar os itens retornados

### 5. Dashboard Engine
Responsável por:
- calcular e renderizar os cards configurados

### 6. List View Engine
Responsável por:
- montar a tabela dinâmica
- renderizar colunas configuradas
- integrar filtros, busca e paginação

### 7. Pagination Engine
Responsável por:
- controlar paginação server-side
- armazenar estado de navegação
- reagir a alterações de filtro/busca

---

## Estrutura inicial sugerida do JSON de configuração

```json
{
  "version": 1,
  "dataSource": {
    "type": "library",
    "siteUrl": "/sites/MeuSite",
    "listTitle": "Documentos"
  },
  "viewObjective": "list",
  "dashboard": {
    "enabled": true,
    "cardCount": 3,
    "cards": []
  },
  "listView": {
    "enabled": true,
    "columns": []
  },
  "pagination": {
    "enabled": true,
    "pageSize": 20,
    "pageSizeOptions": [5, 10, 20, 50, 100]
  }
}
```

---

## Estrutura evoluída esperada do JSON

```json
{
  "version": 1,
  "dataSource": {
    "type": "list",
    "siteUrl": "/sites/Projeto",
    "listTitle": "Tarefas",
    "detectedBaseType": "list",
    "selectedViewId": "<view-id-opcional>"
  },
  "viewObjective": "list",
  "wizard": {
    "isConfigured": true,
    "configuredAt": "2026-03-07T00:00:00Z"
  },
  "dashboard": {
    "enabled": true,
    "cardCount": 4,
    "cards": [
      {
        "id": "total",
        "title": "Total",
        "aggregate": "count"
      },
      {
        "id": "pendentes",
        "title": "Pendentes",
        "aggregate": "count",
        "filter": {
          "field": "Status",
          "operator": "eq",
          "value": "Pendente"
        }
      }
    ]
  },
  "listView": {
    "enabled": true,
    "columns": [
      {
        "internalName": "Title",
        "label": "Título",
        "visible": true,
        "sortable": true
      },
      {
        "internalName": "Status",
        "label": "Status",
        "visible": true,
        "sortable": true
      }
    ]
  },
  "pagination": {
    "enabled": true,
    "pageSize": 20,
    "pageSizeOptions": [5, 10, 20, 50, 100]
  }
}
```

---

## Ordem recomendada de implementação

A ordem abaixo considera a base do MVP e a forma como você quer construir por partes.

### Etapa 0 — Configuração inicial da webpart
Implementar primeiro:
- estado vazio da webpart
- tela de onboarding/configuração
- wizard por steps
- geração do JSON principal
- persistência do JSON na propriedade da webpart

### Etapa 1 — Dashboard
Depois que o JSON básico existir:
- habilitar dashboard opcional
- renderizar cards simples
- conectar aos dados da lista/biblioteca

### Etapa 2 — Listagem
Depois:
- montar tabela dinâmica
- suportar colunas definidas no JSON
- renderizar tipos de campo básicos

### Etapa 3 — Paginação
Por fim nesta primeira trilha:
- paginação server-side
- controle de page size
- navegação próxima/anterior
- integração com listagem

---

# Prompt 1 — Configuração inicial da webpart (Wizard + JSON)

Use este prompt para pedir ao Cursor a primeira base do projeto.

```text
Quero implementar a primeira etapa do meu projeto SPFx: a configuração inicial da webpart dinâmica.

Contexto do projeto:
- É uma webpart SPFx moderna.
- O projeto usará configuração em JSON persistida na propriedade da webpart.
- Não quero criar listas auxiliares de configuração no SharePoint.
- O usuário configurará a webpart por uma tela amigável em steps.
- O MVP terá 3 modos previstos em arquitetura: list, projectManagement e formManager.
- No MVP, somente o modo list poderá ser selecionado; os outros devem aparecer desabilitados ou como “em breve”.

Objetivo desta etapa:
Criar a experiência inicial da webpart quando ainda não há configuração salva.

Requisitos funcionais:
1. Quando a webpart não possuir configuração, ela deve renderizar uma tela de onboarding/configuração.
2. Essa tela deve funcionar como um wizard por steps.
3. Steps obrigatórios do MVP:
   - Step 1: escolher se a origem será lista ou biblioteca e informar/selecionar o nome da lista/biblioteca.
   - Step 2: escolher o modo da webpart, exibindo list, projectManagement e formManager, mas permitindo selecionar somente list.
   - Step 3: definir se haverá dashboard e, se sim, quantos cards inicialmente.
   - Step 4: definir paginação, com pageSize e pageSizeOptions usando os valores [5, 10, 20, 50, 100].
4. Ao concluir o wizard, deve ser gerado um JSON principal de configuração.
5. Esse JSON deve ser salvo na propriedade da webpart.
6. A estrutura do JSON deve ser tipada e validada.
7. O código deve ser organizado para permitir futura expansão.

Requisitos técnicos:
- Usar TypeScript.
- Criar interfaces fortes para a configuração.
- Separar componentes do wizard, tipos, utilitários e camada de persistência da webpart.
- Preparar uma função de validação do JSON.
- Preparar a estrutura para evolução futura dos modos projectManagement e formManager.
- Não implementar ainda dashboard real, listagem real nem paginação real nesta etapa; apenas a configuração inicial e persistência do JSON.

Entregáveis esperados:
- Estrutura de tipos da configuração.
- Componente principal que detecta se a webpart já foi configurada.
- Wizard com steps.
- Geração do JSON final.
- Persistência na propriedade da webpart.
- Código limpo e escalável.
```

---

# Prompt 2 — Dashboard (Modo Lista)

```text
Agora quero implementar a etapa de dashboard do meu projeto SPFx de view dinâmica para SharePoint.

Contexto já existente:
- A webpart já possui um JSON de configuração salvo.
- O JSON possui os blocos dataSource, viewObjective, dashboard, listView e pagination.
- O único modo habilitado no MVP é viewObjective = "list".
- A origem pode ser uma lista ou biblioteca do SharePoint.

Objetivo desta etapa:
Implementar o dashboard configurável do modo lista.

Requisitos funcionais:
1. O dashboard deve ser opcional, controlado por dashboard.enabled.
2. Deve respeitar dashboard.cardCount.
3. Deve permitir renderizar cards com métricas simples no MVP.
4. Tipos de agregação iniciais:
   - count
   - sum
5. Cada card deve suportar:
   - id
   - title
   - aggregate
   - field opcional
   - filtro opcional
6. O dashboard deve funcionar tanto para lista quanto para biblioteca.
7. Deve ser fácil evoluir depois para cards clicáveis, filtros aplicáveis e layouts mais ricos.

Requisitos técnicos:
- Criar um DashboardEngine ou serviço responsável por calcular os dados dos cards.
- Integrar com a camada de dados da lista/biblioteca.
- Criar tipos fortes para os cards.
- Separar componente visual de cálculo de dados.
- Tratar loading e empty state.
- Evitar acoplamento excessivo com a futura listagem.

Entregáveis esperados:
- Tipos/interfaces dos cards.
- Componente de dashboard.
- Serviço/engine para cálculo das métricas.
- Integração com o JSON de configuração.
- Estrutura escalável para futuras evoluções.
```

---

# Prompt 3 — Listagem dinâmica (Modo Lista)

```text
Agora quero implementar a etapa de listagem dinâmica do meu projeto SPFx de view para SharePoint.

Contexto já existente:
- A webpart já possui JSON de configuração.
- A origem é uma lista ou biblioteca.
- O modo habilitado é "list".
- O dashboard já pode existir, mas a listagem deve funcionar independentemente dele.

Objetivo desta etapa:
Criar uma tabela dinâmica configurável para exibir os itens da lista ou biblioteca.

Requisitos funcionais:
1. A tabela deve ser baseada em listView.columns.
2. Cada coluna deve suportar ao menos:
   - internalName
   - label
   - visible
   - sortable
3. A tabela deve suportar os tipos de campo do MVP:
   - texto
   - múltiplas linhas
   - número
   - moeda
   - data
   - booleano
   - choice
   - lookup simples
   - usuário/pessoa
   - hyperlink
   - nome do arquivo
4. Para bibliotecas, a listagem deve tratar também o nome do documento e metadados básicos.
5. A listagem deve ser preparada para integração com paginação server-side.
6. A listagem deve ser preparada para ordenação futura.
7. O sistema deve ter um registry ou estratégia de renderização por tipo de campo.

Requisitos técnicos:
- Criar tipos fortes para colunas configuráveis.
- Criar um FieldRendererRegistry ou estratégia equivalente.
- Separar a obtenção/normalização dos itens da renderização da tabela.
- Criar estados de loading, empty e erro.
- Manter o código escalável para evolução futura em kanban e outras views.

Entregáveis esperados:
- Componente de tabela dinâmica.
- Renderizadores por tipo de campo.
- Normalização dos itens retornados.
- Integração com o JSON da webpart.
- Estrutura escalável.
```

---

# Prompt 4 — Paginação (Modo Lista)

```text
Agora quero implementar a etapa de paginação do meu projeto SPFx de view dinâmica para SharePoint.

Contexto já existente:
- A webpart já possui JSON de configuração.
- A listagem dinâmica já existe.
- O modo habilitado é "list".
- A paginação será usada tanto para listas quanto para bibliotecas.

Objetivo desta etapa:
Implementar paginação server-side performática integrada à listagem dinâmica.

Requisitos funcionais:
1. A paginação deve ser opcional, controlada por pagination.enabled.
2. Deve respeitar pagination.pageSize.
3. Deve permitir pageSizeOptions configuráveis, inicialmente [5, 10, 20, 50, 100].
4. Deve funcionar com listas e bibliotecas grandes.
5. Deve usar abordagem server-side, evitando carregar tudo no cliente.
6. Deve suportar ao menos:
   - próxima página
   - página anterior quando possível
   - alteração do tamanho da página
7. Deve ser projetada para integrar futuramente com filtros e busca.

Requisitos técnicos:
- Criar um PaginationEngine ou serviço equivalente.
- Trabalhar com skip token, next link ou estratégia equivalente do SharePoint/PnP.
- Não acoplar a paginação diretamente ao componente visual da tabela.
- Criar estado tipado da paginação.
- Tratar reset de paginação ao alterar pageSize no futuro.

Entregáveis esperados:
- Tipos/interfaces da paginação.
- Componente visual de paginação.
- Serviço/engine da paginação server-side.
- Integração com a listagem dinâmica.
- Estrutura preparada para filtros e busca futuros.
```

---

## Prompt 5 — Metadata Engine (recomendado antes da listagem real)

Embora você tenha pedido a trilha dashboard → listagem → paginação, tecnicamente vale muito a pena implementar a leitura de metadados logo no início.

```text
Quero implementar a camada de metadados do meu projeto SPFx de view dinâmica para SharePoint.

Objetivo:
Criar um Metadata Engine capaz de descobrir e normalizar a estrutura da origem de dados configurada na webpart.

Requisitos funcionais:
1. Detectar se a origem configurada é lista ou biblioteca.
2. Carregar informações básicas da origem.
3. Carregar fields disponíveis.
4. Carregar views disponíveis.
5. Mapear os tipos nativos do SharePoint para tipos internos do sistema.
6. Expor uma estrutura normalizada para as próximas etapas do projeto.

Requisitos técnicos:
- Criar tipos fortes para fields, views e origem.
- Isolar essa lógica num serviço próprio.
- Preparar a base para lookup, user fields e bibliotecas.
- Não misturar metadados com renderização.

Entregáveis esperados:
- Serviço MetadataEngine.
- Tipos/interfaces.
- Funções de normalização.
- Estrutura reutilizável nas próximas etapas.
```

---

## Estrutura de pastas sugerida

```text
src/
  core/
    config/
      types/
      validators/
      builders/
    metadata/
    dashboard/
    list/
    pagination/
    data/
    utils/

  components/
    Wizard/
    Dashboard/
    DynamicTable/
    Pagination/
    States/

  services/
    sharepoint/

  webparts/
    dynamicView/
```

---

## Decisões fechadas do MVP

### Confirmadas
- configuração baseada em JSON
- sem listas auxiliares de configuração
- onboarding por steps
- modo lista como único habilitado
- dashboard opcional
- listagem dinâmica
- paginação server-side
- opções de page size: `[5, 10, 20, 50, 100]`

### Preparadas para o futuro
- modo `projectManagement`
- modo `formManager`
- filtros avançados
- busca
- ordenação mais rica
- ações por linha
- cards clicáveis
- kanban
- step manager

---

## Próximo passo recomendado

Seguir nesta ordem prática:

1. configuração inicial da webpart
2. metadata engine
3. dashboard
4. listagem
5. paginação

Mesmo com os prompts separados do jeito solicitado, a camada de metadados tende a facilitar muito as etapas seguintes.

---

## Nome sugerido do módulo

- Dynamic View Engine
- Smart List View
- FlexView
- MetaView

Sugestão principal para o MVP:

**FlexView**

