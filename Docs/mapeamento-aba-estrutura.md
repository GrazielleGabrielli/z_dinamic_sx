# Mapeamento funcional — Aba Estrutura

> **Escopo:** painel lateral **"Configurar formulário e regras"** → aba **Estrutura**.
> Documento gerado a partir da análise estática do código-fonte (`FormManagerConfigPanel.tsx`).
> Ordem: visual, de cima para baixo e da esquerda para direita.

---

## 1. Navegação superior da configuração

### 1.1 Título do painel
- **Tipo:** Texto de cabeçalho fixo (painel lateral grande — `PanelType.large`)
- **Onde aparece:** Topo do painel ao abrir "Configurar formulário e regras"
- **Observação:** Sempre visível independente da aba ativa.

### 1.2 Link "JSON (ver / colar)"
- **Tipo:** Link clicável
- **Onde aparece:** Imediatamente abaixo do cabeçalho, lado esquerdo, antes das abas
- **Observação:** Abre painel lateral de edição de JSON bruto da configuração — **possível painel lateral** ("Configuração em JSON").

### 1.3 Barra de abas (Pivot)
- **Tipo:** Componente Pivot (abas tabuladas)
- **Onde aparece:** Abaixo do link JSON
- **Abas disponíveis:**
  1. Estrutura *(foco deste documento)*
  2. Regras dos campos
  3. Componentes
  4. Anexos
  5. Botões
  6. Auditoria e versões
  7. Listas vinculadas
  8. Quebra de permissões

---

## 2. Ações gerais da aba Estrutura

### 2.1 MessageBar de aviso — Campos obrigatórios ausentes
- **Tipo:** Barra de aviso condicional (`MessageBarType.warning`)
- **Onde aparece:** Topo da aba Estrutura, quando há campos marcados como obrigatórios na lista SharePoint ainda não incluídos em nenhuma etapa
- **Observação:** Lista os nomes e nomes internos dos campos faltantes. Não bloqueia, apenas avisa.

### 2.2 Barra amarela de seleção múltipla
- **Tipo:** Barra de ação condicional (fundo `#fff4ce`, borda cinza)
- **Onde aparece:** Aparece apenas quando um ou mais campos estão selecionados via checkbox de seleção nas linhas de campo
- **Itens internos:**
  - Texto: `X campo(s) selecionado(s)` (em negrito)
  - Texto: `Escolha a etapa de destino:`
  - **2.2.1** IconButton (ícone `Forward` com pulsação animada) → abre menu dropdown com lista de etapas para mover os campos selecionados — **possível dropdown de etapas**
  - **2.2.2** DefaultButton `Limpar seleção` → desmarca todos os campos selecionados

### 2.3 Botão "Nova etapa"
- **Tipo:** PrimaryButton
- **Onde aparece:** Logo abaixo das seções expansíveis, antes da lista de etapas
- **Possível função:** Cria uma nova etapa vazia no formulário
- **Observação:** Etapas criadas pelo usuário são arrastáveis.

---

## 3. Seções expansíveis da Estrutura

### 3.1 Layout do formulário (vista)
- **Tipo:** Seção recolhível (`FormManagerCollapseSection`)
- **Estado padrão:** Recolhida (abre ao clicar no cabeçalho)
- **Itens internos identificados:**
  - **3.1.1** Dropdown `Largura` — opções: *Percentagem da área disponível* / *Largura total (100%)*
  - **3.1.2** TextField `Percentagem da largura (1–100)` — visível **somente** quando modo "Percentagem" selecionado; descrição: "Ex.: 80 para ocupar 80% da largura disponível."
  - **3.1.3** Dropdown `Alinhamento horizontal` — opções: *Início (esquerda)* / *Centro* / *Fim (direita)*
  - **3.1.4** TextField `Padding (px)` — aceita 1–160; descrição: "Espaço interior em torno do conteúdo. Vazio = sem padding extra."

### 3.2 Navegação entre etapas (formulário)
- **Tipo:** Seção recolhível (`FormManagerCollapseSection`)
- **Estado padrão:** Recolhida
- **Itens internos identificados:**
  - **3.2.1** Toggle `Exigir obrigatórios preenchidos para avançar (Próximo / etapa à frente)` — sempre habilitado
  - **3.2.2** Toggle `Ao avançar, aplicar todas as regras de validação nos campos da etapa (não só obrigatório)` — **desabilitado** se o toggle 3.2.1 estiver desativado
  - **3.2.3** Toggle `Permitir voltar etapa sem validar a atual` — **desabilitado** se o toggle 3.2.1 estiver desativado

---

## 4. Etapas do formulário

> As etapas seguem ordem fixa para as especiais: **Ocultos** (sempre primeira), **Fixos** (sempre segunda), depois as etapas criadas pelo usuário na ordem configurada.

---

### 4.1 Etapa especial: Ocultos

- **Tipo da etapa:** Reservatório interno — não entra no passador de etapas do formulário
- **Posição:** Sempre fixada na primeira posição (não arrastável)
- **Observações visíveis:** Texto "Não entra no passador (reserva de campos)"
- **Ações disponíveis na etapa:**
  - **4.1.1** IconButton `ChevronRight/Down` — expandir / recolher o conteúdo da etapa
  - **4.1.2** Ícone `GripperBarVertical` — cursor `default` (não arrastável)
  - **4.1.3** TextField `Título da etapa (ocultos)` — editável
  - **4.1.4** Quando recolhida: texto auxiliar com contagem `X campo(s)`
- **Sem botão "Configurar", sem botão "Configurar Colunas", sem botão "Remover etapa"**
- **Conteúdo quando expandida:**
  - Lista de campos já atribuídos à etapa Ocultos (com ações por campo — ver seção 5)
  - Zona de drop (área pontilhada): *"Soltar aqui para colocar no fim desta etapa"*
  - Sub-seção: **Campos fora do formulário** (ver 4.1.5 a 4.1.9)
    - **4.1.5** Texto explicativo: arrastar para etapa ou usar barra amarela
    - **4.1.6** Campo virtual `Anexos ao item` (se ainda não colocado em nenhuma etapa):
      - Grip arrastável
      - Checkbox (selecionar para mover em lote)
      - Label "Anexos ao item (controlo de ficheiros)"
      - Texto: `anexos`
      - Texto auxiliar: instrução de uso e referência à aba Anexos
    - **4.1.7** DefaultButton `Adicionar alerta` — cria campo virtual de alerta condicional
    - **4.1.8** DefaultButton `Adicionar banner` — cria campo virtual de banner por URL
    - **4.1.9** Lista de campos da lista SharePoint ainda não colocados no formulário (ordenados por obrigatório > nome):
      - Grip arrastável para etapa
      - Checkbox com label `Título (InternalName)[*]` — `*` se obrigatório na lista
      - Texto: `MappedType` · `obrig. lista` (se aplicável)

---

### 4.2 Etapa especial: Fixos

- **Tipo da etapa:** Zona fixa — campos aparecem fixados no topo ou rodapé do formulário
- **Posição:** Sempre segunda (logo após Ocultos, não arrastável)
- **Observações visíveis:**
  - Badge `Topo ou rodapé fixo`
  - Texto de modos ativos (ex: "Criar, Editar, Ver") e resumo de condição `showStepWhen` se configurado
- **Ações disponíveis na etapa:**
  - **4.2.1** IconButton `ChevronRight/Down` — expandir / recolher
  - **4.2.2** Ícone `GripperBarVertical` — cursor `default` (não arrastável)
  - **4.2.3** TextField `Título da etapa (fixos)` — editável
  - **4.2.4** DefaultButton `Configurar` → abre **painel lateral "Visibilidade da etapa"** (modos + condição `when`)
  - **4.2.5** DefaultButton `Configurar Colunas` → abre **Modal "Configurar Colunas"** (editor em lote)
- **Sem botão "Remover etapa"**
- **Conteúdo quando expandida:**
  - Lista de campos já atribuídos (com ações de campo — ver seção 5; campos Banner/Alert têm opções de zona fixa)
  - Zona de drop: *"Soltar aqui para colocar no fim desta etapa"*
  - Sub-seção: **Incluir em Fixos**
    - **4.2.6** Texto explicativo
    - **4.2.7** DefaultButton `Adicionar alerta`
    - **4.2.8** DefaultButton `Adicionar banner`
    - **4.2.9** Lista de campos disponíveis (mesmo padrão de 4.1.9)

---

### 4.3 Etapas criadas pelo usuário (etapas normais)

- **Tipo da etapa:** Etapa do passador — aparece na navegação do formulário
- **Posição:** Arrastável (após Ocultos e Fixos)
- **Observações visíveis:**
  - Texto de modos: ex. `Criar, Editar` (se restrito) ou resumo de condição `showStepWhen`
- **Ações disponíveis na etapa:**
  - **4.3.1** IconButton `ChevronRight/Down` — expandir / recolher
  - **4.3.2** Ícone `GripperBarVertical` — cursor `grab` (arrastável para reordenar entre etapas)
  - **4.3.3** TextField `Título da etapa (id)` — editável; label mostra o ID interno
  - **4.3.4** Quando recolhida: texto `X campo(s)` ao lado
  - **4.3.5** DefaultButton `Configurar` → abre **painel lateral "Visibilidade da etapa"**
  - **4.3.6** DefaultButton `Configurar Colunas` → abre **Modal "Configurar Colunas"** (editor em lote)
  - **4.3.7** DefaultButton `Remover etapa` → remove a etapa (os campos retornam ao pool)
- **Conteúdo quando expandida:**
  - Lista de campos já atribuídos (com ações de campo — ver seção 5)
  - Zona de drop: *"Soltar aqui para colocar no fim desta etapa"*

---

## 5. Campos dentro das etapas

> Três tipos de campo com layouts distintos.

---

### 5.1 Campo tipo Alerta (`fieldKind === 'alert'`)

- **Nome interno:** Prefixo virtual — não corresponde a coluna da lista
- **Tipo:** Alerta condicional (info / sucesso / aviso / erro)
- **Status visual:** Fundo e borda seguem `requiredFieldRowStyles` (se obrigatório: fundo amarelo/avermelhado)
- **Ações disponíveis (linha principal — sempre visíveis):**
  - **5.1.1** Ícone `GripperBarVertical` — arrastar campo dentro da etapa (reordenar)
  - **5.1.2** Checkbox — selecionar para mover em lote (barra amarela)
  - **5.1.3** IconButton `ChevronRight/Down` — expandir / recolher configurações inline
  - **5.1.4** Texto bold `Alerta`
  - **5.1.5** Texto descritivo: `internalName · configurações visíveis / clique para configurar`
  - **5.1.6** DefaultButton `Remover` — remove o campo
- **Configurações expandidas (quando ChevronDown ativo):**
  - **5.1.7** TextField `Título`
  - **5.1.8** TextField `Mensagem` (multiline, 3 linhas)
  - **5.1.9** Dropdown `Campos no alerta` (multiSelect — lista campos do formulário)
  - **5.1.10** Dropdown `Tipo` — Informação / Sucesso / Aviso / Erro
  - **5.1.11** Checkbox `Mostrar só quando a condição abaixo for verdadeira`
  - **5.1.12** Se condição ativa:
    - Dropdown `Campo`
    - Dropdown `Operador` (eq / neq / isEmpty / isFilled / isTrue / isFalse / etc.)
    - Dropdown `Comparar com` — Texto fixo / Outro campo / Token
    - TextField `Valor` (desabilitado para operadores sem valor, ex. isEmpty/isFilled)
  - **5.1.13** TextField `Ícone` — nome de ícone Fluent UI (opcional)
  - **5.1.14** Checkbox `Destacar visualmente`
  - **5.1.15** Checkbox `Fechável`
  - **5.1.16** Dropdown `Posição no formulário` — Na etapa / Topo fixo / Rodapé fixo
  - **5.1.17** Se posição ≠ "Na etapa":
    - Dropdown `Zona fixa` — Topo / Rodapé
    - Dropdown `Posicionamento` — Sticky / Absoluto ao contentor / Fluxo normal

---

### 5.2 Campo tipo Banner (`fieldKind === 'banner'`)

- **Nome interno:** Prefixo virtual — não corresponde a coluna da lista
- **Tipo:** Imagem por URL
- **Status visual:** Fundo/borda via `requiredFieldRowStyles`
- **Ações disponíveis (linha principal — sempre visíveis):**
  - **5.2.1** Ícone `GripperBarVertical` — arrastar (reordenar)
  - **5.2.2** Checkbox — selecionar para mover em lote
  - **5.2.3** IconButton `ChevronRight/Down` — expandir / recolher configurações inline
  - **5.2.4** Texto bold `Banner`
  - **5.2.5** Texto descritivo: `internalName · configurações visíveis / clique para configurar`
  - **5.2.6** DefaultButton `Remover`
- **Configurações expandidas:**
  - **5.2.7** TextField `URL da imagem`
  - **5.2.8** TextField `Largura (%)` — 1–100; descrição: "Largura da imagem em % da área do formulário."
  - **5.2.9** TextField `Altura (px)` — opcional; 40–2000
  - **5.2.10** Se etapa = Fixos:
    - Dropdown `Zona fixa` — Topo / Rodapé
    - Dropdown `Posicionamento` — Sticky / Absoluto / Fluxo normal
  - **5.2.11** Se etapa ≠ Fixos:
    - Dropdown `Posição no formulário` — Na etapa / Topo fixo / Rodapé fixo
    - Se posição ≠ "Na etapa": Dropdown `Posicionamento`

---

### 5.3 Campo normal (coluna SharePoint ou campo virtual de Anexos)

- **Nome interno:** InternalName da coluna SharePoint (ou `__dinamic_attachments` para Anexos)
- **Tipo:** Depende da coluna — text, multiline, number, datetime, lookup, user, boolean, etc.
- **Obrigatório:** Indicado com `· obrigatório na lista` na linha de descrição
- **Status visual:**
  - Fundo/borda padrão: `#faf9f8` com borda `#edebe9`
  - Campos obrigatórios da lista sem etapa: fundo/borda de aviso (amarelo/vermelho — `requiredFieldRowStyles`)
  - Campos de sistema: nota adicional `· sistema: só leitura no formulário (aba Regras não aplica)`
- **Ações disponíveis:**
  - **5.3.1** Ícone `GripperBarVertical` — arrastar para reordenar dentro da etapa ou entre etapas
  - **5.3.2** Checkbox — selecionar para mover em lote (barra amarela)
  - **5.3.3** Texto bold — Título do campo (ou InternalName se sem metadado)
  - **5.3.4** Texto descritivo — `InternalName · MappedType [· obrigatório na lista] [· sistema: só leitura...]`
    - Para Anexos: `campo virtual · etapa definida aqui; destino da gravação na aba Anexos`
  - **5.3.5** DefaultButton `Colunas` → abre **Modal "Colunas na grelha"** (individual por campo) — **possível modal**
  - **5.3.6** Texto info do span atual — ex.: `12`, `N12 · V6 · E12`, `12 · resp.` (resumo responsivo)
  - **5.3.7** Se etapa = Fixos e campo tem `fcRow`:
    - Dropdown `Zona fixa` — Topo / Rodapé
    - Dropdown `Posicionamento` — Sticky / Absoluto / Fluxo normal
  - **5.3.8** DefaultButton `Remover`
    - **Desabilitado** se: campo obrigatório na lista AND há pelo menos um campo em alguma etapa (tooltip explica)

---

## 6. Seção: Listas vinculadas (etapa no passador)

### 6.1 Título de seção "Listas vinculadas (etapa no passador)"
- **Tipo:** Texto `medium` em negrito
- **Onde aparece:** Após o bloco de etapas, dentro da aba Estrutura
- **Visibilidade:** Condicional — só aparece se houver listas vinculadas configuradas na aba "Listas vinculadas"
- **Observação:** Permite configurar em qual etapa do formulário principal cada lista vinculada aparece.

### 6.2 MessageBar info — "Adicione pelo menos uma etapa"
- **Tipo:** Barra de informação condicional
- **Onde aparece:** No lugar da lista de vínculos, quando não há etapas normais disponíveis para posicionamento
- **Observação:** Apenas informativa; não há ação direta aqui.

### 6.3 Linha de posicionamento de lista vinculada
- **Tipo:** Linha por lista vinculada (uma linha por item)
- **Componentes da linha:**
  - **6.3.1** Texto — nome/título da lista vinculada (até 120 caracteres)
  - **6.3.2** Dropdown `Etapa` — escolhe a etapa do formulário principal onde o bloco aparece

---

## 7. Botões e ações identificadas (modais e painéis)

### 7.1 Modal "Colunas na grelha" — individual por campo
- **Acesso:** Botão `Colunas` na linha de campo normal (5.3.5)
- **Possível função:** Configurar número de colunas (de 12) que o campo ocupa por faixa de viewport e por modo de formulário
- **Conteúdo do modal:**
  - Cabeçalho com nome e InternalName do campo
  - Texto explicativo: "Colunas (de 12) por faixa de largura e por modo. Sem valor na faixa: herda da faixa menor (mobile-first)."
  - Pivot (abas por breakpoint): **XS** (≥0px) / **S** (≥480px) / **M** (≥640px) / **L** (≥1024px) / **XL** (≥1366px) / **XXL** (≥1920px)
  - Para cada breakpoint, 3 cards de modo: **Novo** / **Ver** / **Editar**
  - Cada card: pills selecionáveis de span — **2 / 3 / 4 / 6 / 8 / 12**
  - DefaultButton `Fechar`

### 7.2 Modal "Configurar Colunas" — em lote por etapa
- **Acesso:** Botão `Configurar Colunas` no cabeçalho da etapa (etapas Fixos e etapas normais)
- **Possível função:** Configurar spans de todos os campos da etapa de uma só vez, sem abrir cada campo individualmente
- **Conteúdo do modal:**
  - Cabeçalho: "Configurar Colunas" + nome da etapa
  - Seletor de modo (botões toggle): **Novo** / **Editar** / **Ver**
  - Para cada campo da etapa: um **card** contendo:
    - Nome do campo (bold) + InternalName (monospace, secundário)
    - Linha de breakpoints com flex-wrap: **XS / S / M / L / XL / XXL** (cada um com rótulo e px mínimo)
    - Pills de span por breakpoint: **2 / 3 / 4 / 6 / 8 / 12** (selecionado = azul)
  - Nota: "Mobile-first: sem valor em uma faixa = herda da faixa menor."
  - DefaultButton `Fechar`
  - `×` (IconButton Cancel) no canto superior direito

### 7.3 Painel lateral "Visibilidade da etapa"
- **Acesso:** Botão `Configurar` no cabeçalho da etapa (etapas Fixos e etapas normais)
- **Possível função:** Controlar em quais modos de formulário a etapa aparece e configurar condição condicional (`showStepWhen`)
- **Conteúdo do painel:**
  - Nome e ID da etapa
  - Bloco "Modos de formulário": checkboxes **Criar / Editar / Ver**
  - Nota: "Todas marcadas ou nenhuma restrição = Criar, Editar e Ver."
  - Checkbox `Só mostrar esta etapa quando as condições abaixo forem verdadeiras`
  - Se condição ativa:
    - Dropdown `Lógica entre condições` — Todas (E) / Pelo menos uma (OU)
    - Linhas de condição (campo + operador + comparador + valor)
    - IconButton `Delete` para remover condição
    - DefaultButton `Adicionar condição`

### 7.4 Painel lateral "Configuração em JSON"
- **Acesso:** Link "JSON (ver / colar)" acima das abas
- **Possível função:** Ver e editar diretamente o JSON completo da configuração do formulário
- **Conteúdo do painel:**
  - TextField multiline (monospace) com JSON serializado
  - PrimaryButton `Aplicar JSON`
  - DefaultButton `Fechar`

### 7.5 Dropdown de etapas — mover campos em lote
- **Acesso:** IconButton `Forward` (pulsante) na barra amarela de seleção
- **Possível função:** Escolher etapa de destino para os campos selecionados via checkbox; move todos de uma vez

---

## 8. Estados visuais e comportamentos percebidos

### 8.1 Campo obrigatório na lista sem etapa
- **Onde aparece:** Linha do campo nas etapas E no pool de campos disponíveis
- **O que indica:** Fundo e borda de aviso (amarelo/vermelho); campo precisa ser colocado em uma etapa

### 8.2 Barra de seleção amarela (`#fff4ce`)
- **Onde aparece:** Topo da aba Estrutura, somente quando há campos selecionados
- **O que indica:** Modo de seleção múltipla ativo; aguardando destino de movimentação

### 8.3 Ícone `Forward` pulsante
- **Onde aparece:** Na barra amarela de seleção
- **O que indica:** Ação disponível de mover campos; chama atenção com animação

### 8.4 Etapa recolhida — contador de campos
- **Onde aparece:** Ao lado do TextField do título da etapa, quando expandida = false
- **O que indica:** Quantidade de campos naquela etapa sem precisar abrir

### 8.5 Texto de modos e condição na etapa
- **Onde aparece:** Abaixo do título da etapa (etapas normais e Fixos), sempre visível
- **O que indica:** Resumo de em quais modos a etapa aparece (ex: "Criar, Editar") e se há condição `when` ativa

### 8.6 Botão "Remover" desabilitado
- **Onde aparece:** Linha de campo obrigatório da lista quando há campos atribuídos no formulário
- **O que indica:** O campo não pode ser removido pois é obrigatório na fonte de dados; tooltip explica

### 8.7 Resumo de span na linha do campo (informativo)
- **Onde aparece:** Ao lado do botão `Colunas`, linha de campo normal
- **O que indica:** Configuração atual de span resumida — ex.: `12` (igual nos 3 modos), `N12 · V6 · E12` (por modo), `12 · resp.` (tem configuração responsiva)

### 8.8 Zona de drop pontilhada (drop zone)
- **Onde aparece:** Fim da lista de campos de cada etapa
- **O que indica:** Área de destino para soltar campos arrastados; aceita campos de outras etapas e do pool

### 8.9 Campo de sistema (metadados SharePoint)
- **Onde aparece:** Linha do campo na etapa
- **O que indica:** Nota `· sistema: só leitura no formulário (aba Regras não aplica)` — campo é read-only por natureza e as regras de edição não se aplicam

---

## 9. Rodapé do painel (sempre visível)

### 9.1 PrimaryButton `Salvar`
- **Local:** Rodapé fixo do painel
- **Possível função:** Persiste toda a configuração do formulário (todas as abas)
- **Observação:** Desabilitado enquanto `loading = true`

### 9.2 DefaultButton `Restaurar padrão (estrutura)`
- **Local:** Rodapé fixo do painel
- **Possível função:** Restaura etapas e campos para a configuração padrão do sistema — **afeta apenas Estrutura** (steps e fields); outras configurações permanecem

### 9.3 DefaultButton `Cancelar`
- **Local:** Rodapé fixo do painel
- **Possível função:** Fecha o painel sem salvar

---

## 10. Pendências para análise detalhada posterior

> Itens que precisam ser abertos / clicados para documentação completa:

1. **Botão `Configurar` da etapa** → painel "Visibilidade da etapa" — documentar toda a lógica de condições (`showStepWhen`)
2. **Botão `Configurar Colunas` da etapa** → modal em lote — documentar comportamento de herança mobile-first entre breakpoints
3. **Botão `Colunas` de cada campo** → modal individual — documentar como os spans por modo interagem entre si e com `columnSpanByMode` / `columnSpanByBreakpointByMode`
4. **Ícone de expandir do campo Alerta** → configurações inline — documentar operadores disponíveis e tokens aceitos
5. **Ícone de expandir do campo Banner** → configurações inline — documentar comportamento de posicionamento fixo vs. na etapa
6. **Barra amarela → IconButton Forward** → dropdown de etapas — documentar regras de movimentação (Ocultos, Fixos, etapas normais)
7. **Link "JSON (ver / colar)"** → painel JSON — documentar esquema completo aceito e comportamento de validação ao aplicar
8. **Aba "Regras dos campos"** — documentar editor de regras (visibilidade condicional, obrigatoriedade, desabilitar, valor padrão, expressões calculadas)
9. **Aba "Componentes"** — documentar configurações de UI do formulário (spinner, shimmers, etc.)
10. **Aba "Anexos"** — documentar storage de anexos (item SharePoint vs. biblioteca de documentos)
11. **Aba "Botões"** — documentar botões customizados, ações em cadeia, comportamento de redirecionamento
12. **Aba "Auditoria e versões"** — documentar log de ações e versionamento de itens
13. **Aba "Listas vinculadas"** — documentar configuração completa de formulários filhos vinculados
14. **Aba "Quebra de permissões"** — documentar configuração de permissões herdadas/quebradas por item
15. **Seção "Campos fora do formulário" (Ocultos)** — documentar pool de campos disponíveis, ordenação e comportamento de campos obrigatórios não alocados
16. **Seção "Incluir em Fixos"** — documentar diferença de comportamento entre campo no topo vs. rodapé vs. sticky vs. absoluto
17. **Seção "Listas vinculadas (etapa no passador)"** — documentar como o posicionamento de etapa afeta a renderização no formulário principal
18. **Estados de carregamento** — documentar `Spinner "Campos da lista..."` e `MessageBar` de erro ao falhar carregamento de metadados
