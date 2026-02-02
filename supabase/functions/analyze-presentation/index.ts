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

    const systemPrompt = `Du klassificerar text från PowerPoint-shapes i konsultpresentationer och genererar beskrivande mall-texter.
Syftet är att GENERALISERA presentationen — allt specifikt innehåll ersätts med en kort beskrivning av vad texten handlar om. Resultatet ska vara sökbart och återanvändbart som mall.

Kategorier:
- title: Huvudrubrik på en slide
- body: Beskrivande text, punktlistor, aktiviteter, KPI:er, affärsmål, strategier
- name: Personnamn eller rollreferenser
- label_value: "Etikett: Värde"-mönster (ange label-delen i "label"-fältet)
- table_header: Strukturella kolumn-/radrubrik (t.ex. "Aktivitet", "Status", "Ansvarig")
- keep: ENBART helt generiska enstaka ord som "Syfte", "Mål", "Agenda"

VIKTIGT: Använd keep SPARSAMT. De flesta shapes innehåller specifikt innehåll.

## Rewrite-regler

För varje shape (utom keep och table_header), generera ett "rewrite"-fält med en BESKRIVANDE MALL-TEXT:

- **title**: "Rubrik om [ämne och syfte i 5-10 ord]"
- **body** (kort text, 1 mening): "Mening om [ämne]"
- **body** (längre text): "Stycke om [ämne, poäng och syfte i 10-20 ord]"
- **name**: En generisk rolltitel, t.ex. "[Projektledare]", "[Konsult]", "[Avdelningschef]"
- **label_value**: Beskriv bara värde-delen: t.ex. "beskrivning av metriken" eller "namn på ansvarig person"
- **Tabellceller** (har isTableCell: true): Kort beskrivning av cellens innehåll i 2-5 ord

Skriv beskrivningen på SAMMA SPRÅK som originaltexten.
Fånga ämne och syfte men ta bort alla specifika namn, siffror, och detaljer.

## Flerraderstext (paragraphs)

Om en shape har "paragraphs" (array av stycken), generera "paragraphRewrites" — en array med en beskrivning per stycke.
- Tomma stycken i originalet → tom sträng i paragraphRewrites
- Varje icke-tomt stycke → beskrivande mall-text

## JSON-format

{ "classifications": [
  { "id": 0, "category": "title", "rewrite": "Rubrik om centralisering av kärnkompetenser för att balansera snabbhet och styrning" },
  { "id": 1, "category": "body", "paragraphRewrites": ["Stycke om hur teknik vuxit organiskt och skapat siloeffekter", "", "Mening om varför centralisering av plattformar nu behövs"] },
  { "id": 2, "category": "name", "rewrite": "[Projektledare]" },
  { "id": 3, "category": "label_value", "label": "Driver", "rewrite": "namn på ansvarig person" },
  { "id": 4, "category": "keep" },
  { "id": 5, "category": "table_header" }
] }`;

    const userPrompt = `Klassificera följande shapes:

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
