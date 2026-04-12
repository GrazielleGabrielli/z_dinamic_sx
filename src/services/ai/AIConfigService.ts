import type { IFieldMetadata } from '../shared/types';

export interface IAIStepConfig {
  id: string;
  title: string;
  fieldNames: string[];
}

export interface IAIStructureOutput {
  steps: IAIStepConfig[];
}

const SYSTEM_FIELDS = new Set([
  'ID', 'Id', 'Author', 'Editor', 'Created', 'Modified',
  'Attachments', 'ContentType', 'ContentTypeId', 'FileSystemObjectType',
  '_UIVersionString', 'GUID', 'UniqueId',
]);

const SYSTEM_PROMPT = `Você é um assistente de configuração de formulários SharePoint SPFx.

Sua tarefa é organizar campos de uma lista SharePoint em etapas lógicas de um formulário.

Retorne APENAS JSON válido, sem markdown, sem explicações adicionais, no seguinte formato:
{
  "steps": [
    { "id": "ocultos", "title": "Ocultos", "fieldNames": [] },
    { "id": "fixos", "title": "Fixos", "fieldNames": [] },
    { "id": string, "title": string, "fieldNames": string[] }
  ]
}

Regras obrigatórias:
1. Os dois primeiros steps DEVEM ser exatamente { "id": "ocultos", "title": "Ocultos", "fieldNames": [] } e { "id": "fixos", "title": "Fixos", "fieldNames": [] }
2. Crie entre 1 e 5 steps adicionais agrupando os campos de forma lógica para o sistema descrito
3. Use APENAS os internalName exatamente como fornecidos na lista de campos disponíveis
4. Distribua TODOS os campos fornecidos (exceto os de sistema indicados) entre os steps adicionais
5. Cada step adicional deve ter um id único em snake_case (ex: "dados_gerais") e título em português
6. Não repita o mesmo campo em mais de um step
7. Campos de sistema (ID, Author, Editor, Created, Modified, Attachments, ContentType, ContentTypeId) já foram excluídos e NÃO devem aparecer nos fieldNames`;

function buildUserMessage(input: { description: string; listTitle: string; fields: IFieldMetadata[] }): string {
  const filteredFields = input.fields.filter((f) => !SYSTEM_FIELDS.has(f.InternalName));
  const fieldLines = filteredFields
    .map((f) => `- internalName: "${f.InternalName}" | label: "${f.Title}" | tipo: "${f.MappedType}"`)
    .join('\n');

  return `Descrição do sistema: ${input.description}

Lista SharePoint: "${input.listTitle}"

Campos disponíveis (${filteredFields.length} campos):
${fieldLines}

Organize esses campos em etapas lógicas para o formulário descrito.`;
}

export async function generateFormStructure(
  apiKey: string,
  input: { description: string; listTitle: string; fields: IFieldMetadata[] }
): Promise<IAIStructureOutput> {
  const content = await callOpenAI(apiKey, SYSTEM_PROMPT, buildUserMessage(input));

  let parsed: unknown;
  try {
    parsed = JSON.parse(content);
  } catch {
    throw new Error('A IA retornou JSON inválido.');
  }

  const out = parsed as IAIStructureOutput;
  if (!Array.isArray(out?.steps)) {
    throw new Error('Resposta da IA não contém steps válidos.');
  }

  for (const s of out.steps) {
    if (typeof s.id !== 'string' || typeof s.title !== 'string' || !Array.isArray(s.fieldNames)) {
      throw new Error('Step inválido na resposta da IA.');
    }
  }

  return out;
}

export interface IAIButtonConfig {
  id: string;
  label: string;
  appearance?: 'primary' | 'default';
  behavior?: 'actionsOnly' | 'draft' | 'submit' | 'close';
  operation?: 'legacy' | 'redirect';
  redirectUrlTemplate?: string;
  modes?: ('create' | 'edit' | 'view')[];
}

export interface IAIButtonsOutput {
  buttons: IAIButtonConfig[];
}

const BUTTONS_SYSTEM_PROMPT = `Você é um assistente de configuração de formulários SharePoint SPFx.

Sua tarefa é gerar botões para um formulário com base na descrição do usuário.

Retorne APENAS JSON válido, sem markdown, sem explicações, no seguinte formato:
{
  "buttons": [
    {
      "id": string,
      "label": string,
      "appearance": "primary" | "default",
      "behavior": "submit" | "draft" | "close" | "actionsOnly",
      "operation": "legacy" | "redirect",
      "redirectUrlTemplate": string (só quando operation === "redirect"),
      "modes": ["create"] | ["edit"] | ["view"] | ["create","edit"] | ["create","edit","view"] | etc
    }
  ]
}

Regras:
1. "id" deve ser único em snake_case (ex: "btn_enviar")
2. "label" é o texto visível do botão em português
3. "appearance": "primary" para o botão principal de ação, "default" para os demais
4. "behavior":
   - "submit" → salva e envia o formulário
   - "draft" → salva como rascunho
   - "close" → fecha sem salvar
   - "actionsOnly" → só executa ações (usar com operation redirect)
5. "operation":
   - "legacy" → comportamento padrão (usar com submit/draft/close)
   - "redirect" → redireciona para uma URL (use {{FormID}} para o ID do item, {{Form}} para o modo)
6. "redirectUrlTemplate" → obrigatório quando operation é "redirect"; pode conter placeholders {{FormID}} e {{Form}}
7. "modes" → em quais modos o botão aparece: "create", "edit", "view"
8. Cada botão deve ter APENAS as propriedades listadas acima`;

function buildButtonsUserMessage(description: string, systemDescription: string): string {
  return `Sistema: ${systemDescription}

Descrição dos botões desejados: ${description}

Gere os botões para o formulário.`;
}

async function callOpenAI(apiKey: string, systemPrompt: string, userMessage: string): Promise<string> {
  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMessage },
      ],
      temperature: 0.2,
      response_format: { type: 'json_object' },
    }),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Erro na API OpenAI (${response.status}): ${text}`);
  }

  const data = (await response.json()) as {
    choices?: { message?: { content?: string } }[];
  };

  const content = data.choices?.[0]?.message?.content;
  if (!content) throw new Error('Resposta vazia da API OpenAI.');
  return content;
}

export async function generateFormButtons(
  apiKey: string,
  input: { description: string; systemDescription: string }
): Promise<IAIButtonsOutput> {
  const content = await callOpenAI(
    apiKey,
    BUTTONS_SYSTEM_PROMPT,
    buildButtonsUserMessage(input.description, input.systemDescription)
  );

  let parsed: unknown;
  try {
    parsed = JSON.parse(content);
  } catch {
    throw new Error('A IA retornou JSON inválido.');
  }

  const out = parsed as IAIButtonsOutput;
  if (!Array.isArray(out?.buttons)) {
    throw new Error('Resposta da IA não contém buttons válidos.');
  }

  for (const b of out.buttons) {
    if (typeof b.id !== 'string' || typeof b.label !== 'string') {
      throw new Error('Botão inválido na resposta da IA.');
    }
  }

  return out;
}
