import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

interface ShapeItem {
  id: number;
  text?: string;
  paragraphs?: string[];
}

interface Classification {
  id: number;
  category: string;
  label?: string;
  rewrite?: string;
  paragraphRewrites?: string[];
}

Deno.serve(async (req: Request) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const openaiApiKey = Deno.env.get("OPENAI_API_KEY");
    if (!openaiApiKey) {
      throw new Error("OPENAI_API_KEY not configured");
    }

    const { shapes } = await req.json() as { shapes: ShapeItem[] };

    if (!shapes || shapes.length === 0) {
      return new Response(
        JSON.stringify({ classifications: [] }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Prepare shapes for GPT
    const shapeEntries = shapes.map(s => {
      const content = s.paragraphs
        ? `"paragraphs": ${JSON.stringify(s.paragraphs)}`
        : `"text": ${JSON.stringify(s.text)}`;
      const tablePart = (s as any).isTableCell ? `, "isTableCell": true` : '';
      return `{ "id": ${s.id}, ${content}${tablePart} }`;
    }).join(",\n");

    const systemPrompt = `You classify text from PowerPoint shapes in consulting presentations and generate FULL REPLACEMENT TEXT that preserves the visual layout of the slide.

The goal is to ANONYMIZE the presentation — all identifying details (company names, person names, numbers, industry-specific information, project names, dates, locations, confidential data) are replaced with generic but realistic alternatives. The replacement text must read like real consulting text, NOT like a meta-description.

CRITICAL RULE: The replacement text MUST be written in the SAME LANGUAGE as the original text. If the original is in English, write the replacement in English. If the original is in Swedish, write in Swedish. NEVER translate.

Categories:
- title: Main heading on a slide
- body: Descriptive text, bullet points, activities, KPIs, business goals, strategies
- name: Person names or role references
- label_value: "Label: Value" pattern (put the label part in the "label" field)
- table_header: Structural column/row headers (e.g. "Activity", "Status", "Responsible")
- keep: ONLY completely generic single words like "Purpose", "Goals", "Agenda"

IMPORTANT: Use keep SPARINGLY. Most shapes contain specific content.

## Rewrite rules

For each shape (except keep and table_header), generate a "rewrite" field with FULL REPLACEMENT TEXT:

- **title**: Write a complete heading of approximately the same length. Replace specific details with generic alternatives. Example: "Volvo Q3 Strategy Review" → "Quarterly Strategy Review for the Organization"
- **body** (short text, 1 sentence): Write a complete sentence of similar length. Replace identifying details but keep the meaning and tone.
- **body** (longer text): Write complete replacement text with APPROXIMATELY THE SAME WORD COUNT as the original. Keep the same structure (bullet points stay as bullet points, paragraphs stay as paragraphs). Replace all identifying details with generic but realistic consulting language.
- **name**: A generic role title in the SAME LANGUAGE as the original, e.g. "[Project Manager]" for English text, "[Projektledare]" for Swedish text, "[Konsult]", "[Avdelningschef]"
- **label_value**: Write a realistic generic value. Example: "45 MSEK" → "[X] MSEK", "Erik Svensson" → "[Project Manager]"
- **Table cells** (have isTableCell: true): Write realistic generic cell content of similar length. Example: "Volvo Trucks" → "Business Unit A", "Q3 2024" → "[Period]"

IMPORTANT: Do NOT write meta-descriptions like "Paragraph about how technology has grown organically". Write ACTUAL TEXT that could appear on a consulting slide. The anonymized slide must look visually identical to the original — same amount of text, same structure, same tone.

## Multi-paragraph text (paragraphs)

If a shape has "paragraphs" (array of paragraphs), generate "paragraphRewrites" — an array with full replacement text for each paragraph.
- Empty paragraphs in the original → empty string in paragraphRewrites
- Each non-empty paragraph → full replacement text of similar length in the same language

## JSON format

{ "classifications": [
  { "id": 0, "category": "title", "rewrite": "Centralizing Core Capabilities to Balance Speed and Governance" },
  { "id": 1, "category": "body", "paragraphRewrites": ["Technology has emerged organically across the organization in pockets of high maturity, such as specific business units or operational functions. However, this fragmented growth has created isolated thinking and inefficient workarounds.", "", "A centralized platform approach is now needed to reduce duplication and enable scalable delivery across the enterprise."] },
  { "id": 2, "category": "name", "rewrite": "[Project Manager]" },
  { "id": 3, "category": "label_value", "label": "Driver", "rewrite": "[Team Lead]" },
  { "id": 4, "category": "keep" },
  { "id": 5, "category": "table_header" }
] }`;

    const userPrompt = `Classify the following shapes and generate full replacement text (in the same language as each original):

[${shapeEntries}]`;

    // Call OpenAI API
    const openaiResponse = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${openaiApiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt }
        ],
        temperature: 0,
        response_format: { type: "json_object" }
      }),
    });

    if (!openaiResponse.ok) {
      const error = await openaiResponse.text();
      console.error("OpenAI API error:", error);
      throw new Error("OpenAI API request failed");
    }

    const openaiData = await openaiResponse.json();
    const content = openaiData.choices[0]?.message?.content;

    if (!content) {
      throw new Error("No response from OpenAI");
    }

    const parsed = JSON.parse(content);
    const classifications: Classification[] = (parsed.classifications || []).map((c: Classification) => {
      const item: Classification = {
        id: c.id,
        category: c.category,
      };
      if (c.label) item.label = c.label;
      if (c.rewrite) item.rewrite = c.rewrite;
      if (c.paragraphRewrites) item.paragraphRewrites = c.paragraphRewrites;
      return item;
    });

    return new Response(
      JSON.stringify({ classifications }),
      { headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );

  } catch (error) {
    console.error("Error:", error);
    return new Response(
      JSON.stringify({ error: error.message }),
      {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      }
    );
  }
});
