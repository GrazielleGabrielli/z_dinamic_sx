# Condicionais (campo Text) — decisões de produto, motor e QA

Documento de apoio à especificação funcional da aba **Regras → Condicionais** para campos **Text** (`MappedType: text`). Não altera o plano aprovado; fixa decisões pendentes e guia implementação futura.

---

## 1. Decisões de produto (validação negócio)

### 1.1 Baseline com mistura «mostrar» e «ocultar»

**Decisão:** manter o modelo da especificação.

- Quando coexistem grupos de visibilidade **mostrar** e **ocultar**, o **baseline** da camada condicional (nenhum grupo show/hide satisfeito) é **oculto**.
- Quando vários grupos satisfeitos competem, **ocultar prevalece sobre mostrar**.

### 1.2 Grupo sem condições

**Decisão (v1):** não permitir grupo com lista de condições vazia.

- Configuração inválida para gravação/publicação do formulário.
- Motivo: evita ambiguidade («sempre verdadeiro» implícito) e reforça intenção explícita do configurador.
- Para efeito «sempre em modo Ver», usar o comportamento global do modo **Ver** (só leitura) ou uma condição trivial explícita (ex.: **não vazio** num campo técnico sempre preenchido), conforme já recomendado na especificação.

### 1.3 Campo oculto por condicional e validação

**Decisão (v1):**

- Campos **ocultos** pelo resultado das condicionais **não** entram na validação de **obrigatoriedade** (não bloqueiam submissão por «campo obrigatório vazio» enquanto ocultos).
- O valor pode permanecer no estado interno do formulário; **não** limpar automaticamente ao ocultar (evita perda de dados ao alternar condições), salvo regra de produto futura explícita.
- Se o campo voltar a **visível** e estiver **obrigatório** na config estática, aí sim aplica-se validação de obrigatoriedade normalmente.

### 1.4 Modo Ver e grupos só em Criar/Editar

**Decisão:** se nenhum grupo de **ocultar** incluir **Ver**, o campo pode permanecer visível em Ver mesmo oculto em Criar/Editar — comportamento intencional; o configurador deve marcar **Ver** no grupo quando quiser o mesmo efeito em todos os modos.

---

## 2. Mapeamento ao motor existente (`TFormRule` / `buildFormDerivedState`)

### 2.1 Princípio

- Os grupos do bloco **Condicionais** do campo alvo são **compiláveis** para regras já suportadas pelo motor (`when` em `TFormConditionNode`, `modes`, `setVisibility`, `setReadOnly`), **sem** obrigar alteração do modelo na primeira entrega de UI (rascunho: anexo lógico ou lista derivada até existir schema no campo).
- Regras geradas a partir deste bloco devem usar **ids estáveis e prefixados** para não colidir com regras manuais em `cfg.rules` (padrão análogo a `ui_f_*` / `ui_card_*` em [`formManagerVisualModel.ts`](../../src/webparts/dinamicApp/core/formManager/formManagerVisualModel.ts)).

### 2.2 Condições do grupo → `when`

- Operador de grupo **E** → nó `{ kind: 'all', children: [...] }` com um filho `leaf` por condição.
- Operador de grupo **OU** → nó `{ kind: 'any', children: [...] }`.
- Cada condição → `leaf` com `field` = internal name do campo de origem, `op` alinhado a [`TFormConditionOp`](../../src/webparts/dinamicApp/core/config/types/formManager.ts), `compare` com `kind: 'literal'` e `value` quando o operador exige valor.
- Operadores funcionais na especificação vs motor hoje:
  - igual a → `eq`
  - diferente de → `ne`
  - contém → `contains`
  - não contém → **ausente hoje** em `TFormConditionOp`; na implementação: **estender** o motor com operador dedicado **ou** representar como subárvore negada (ex.: `all` com um único filho negado — só se o avaliador suportar negação; caso contrário, estender `compareResolved` / tipo de op).
  - vazio → `isEmpty`
  - não vazio → `isFilled`

### 2.3 Ação do grupo → regras

- **mostrar:** `action: 'setVisibility'`, `targetKind: 'field'`, `targetId: <campo alvo>`, `visibility: 'show'`, mesmo `when` e `modes` do grupo.
- **ocultar:** idem com `visibility: 'hide'`.
- **readonly:** `action: 'setReadOnly'`, `field: <campo alvo>`, `readOnly: true`, mesmo `when` e `modes`.

**Nota:** o motor atual aplica `setReadOnly` apenas se `formMode !== 'view'` ([`formRuleEngine.ts`](../../src/webparts/dinamicApp/core/formManager/formRuleEngine.ts)); em **Ver**, o readonly já vem do modo. Grupos só com **Ver** e ação readonly são redundantes para UX mas podem existir para consistência de configuração.

### 2.4 Ordem e conflitos vs lista linear de `cfg.rules`

- A especificação define **precedência semântica** (hide > show; baseline; OR para readonly) que **não** coincide com «última regra na lista ganha» para visibilidade.
- **Opções de implementação (escolher na fase de código):**
  1. **Pré-processar** condicionais do campo num único passo antes de `buildFormDerivedState` e escrever apenas o **resultado** em flags efémeras por campo (extensão do estado derivado), **sem** emitir múltiplas `setVisibility` conflituosas na lista global; ou
  2. Emitir regras **ordenadas** e alterar o motor para uma fase dedicada «regras por campo Text condicional» (mais invasivo).

Recomendação documental: **(1)** reduz risco de regressão nas regras globais existentes.

### 2.5 Compatibilidade

- Regras manuais existentes em `cfg.rules` continuam válidas.
- Condicionais do campo Text não devem remover nem alterar regras que não sejam das séries de ids geridas por este bloco.

---

## 3. Matriz QA — campo origem × operadores (v1, alvo Text)

Legenda: **S** = suportado na v1 com coerção para string definida abaixo; **N** = não disponível no seletor de origem na v1; **P** = suportado só para igual / diferente / vazio / não vazio (sem contains textual sensato).

| Origem (`FieldMappedType`) | Coerção para comparação textual | igual | diferente | contém | não contém | vazio | não vazio |
| -------------------------- | -------------------------------- | ----- | ---------- | ------ | ----------- | ----- | ---------- |
| `text` | Valor do campo como string (trim onde aplicável ao motor) | S | S | S | S | S | S |
| `multiline` | Texto plano quando possível; se nota rich text, usar texto extraído pelo pipeline do formulário (mesma fonte que o motor usa para valor) | S | S | S | S | S | S |
| `choice` | Valor escolhido como string (valor guardado no item) | S | S | S | S | S | S |
| `multichoice` | Serialização estável (ex.: `;#` ou lista juntada) igual ao usado em `values` no runtime | S | S | P | P | S | S |
| `number` | Representação decimal string coerente com o motor (ex.: `coerceNumber`/stringificação) | S | S | P | P | S | S |
| `currency` | Idem número + símbolo/formato conforme valor em runtime | S | S | P | P | S | S |
| `boolean` | `true`/`false` ou equivalentes aceites pelo motor | S | S | P | P | S | S |
| `datetime` | ISO ou string apresentada conforme valor em `values` | S | S | P | P | S | S |
| `url` | URL como string (Descrição + URL: alinhar ao formato já serializado no item) | S | S | S | S | S | S |
| `calculated` | Valor calculado como string no runtime | S | S | P | P | S | S |
| `lookup` | **N** em v1 (objeto `{ Id, … }`; exige regra de display única) | N | N | N | N | N | N |
| `lookupmulti` | **N** | N | N | N | N | N | N |
| `user` | **N** | N | N | N | N | N | N |
| `usermulti` | **N** | N | N | N | N | N | N |
| `taxonomy` | **N** | N | N | N | N | N | N |
| `taxonomymulti` | **N** | N | N | N | N | N | N |
| `unknown` | **N** | N | N | N | N | N | N |

**Notas de teste:**

- **contains / não contém** com **P:** validar apenas cenários onde o literal faz sentido (ex.: número `123` contém `2`); não é obrigatório expor na UI para tipos **P** se o produto optar por ocultar esses operadores por tipo de origem.
- **Vazio / não vazio:** alinhar à função `isEmptyish` do motor ([`formRuleEngine.ts`](../../src/webparts/dinamicApp/core/formManager/formRuleEngine.ts)) para consistência.
- Casos-limite obrigatórios em QA: `null`/`undefined`, string só espaços, maiúsculas vs minúsculas em **contains**, valor de lookup vazio (quando v2 permitir lookup).

---

*Última atualização: documento gerado para fechar os itens de trabalho de especificação; revisão de negócio pode ajustar células **N**/**P** em versões posteriores.*
