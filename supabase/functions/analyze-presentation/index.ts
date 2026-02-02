import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

interface ShapeItem {
  slideIndex: number;
  shapeIndex: number;
  text?: string;
  paragraphs?: string[];
}

interface Classification {
  slideIndex: number;
  shapeIndex: number;
  category: string;
  label?: string;
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
      return `{ "slideIndex": ${s.slideIndex}, "shapeIndex": ${s.shapeIndex}, ${content} }`;
    }).join(",\n");

    const systemPrompt = `Du klassificerar text från PowerPoint-shapes i konsultpresentationer.

Kategorier:
- title: Huvudrubrik på en slide
- body: Beskrivande text, punktlistor, aktiviteter
- name: Personnamn eller rollreferenser
- label_value: "Etikett: Värde"-mönster (ange label-delen i "label"-fältet)
- table_header: Kolumn-/radrubrik
- keep: Redan generisk text som inte behöver ändras (t.ex. "Syfte", "Mål", "Agenda")

Svara med JSON: { "classifications": [{ "slideIndex": 0, "shapeIndex": 0, "category": "title" }] }
För label_value, inkludera "label"-fält: { "slideIndex": 0, "shapeIndex": 5, "category": "label_value", "label": "Driver" }

Klassificera BARA, generera INGEN ny text.`;

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
        slideIndex: c.slideIndex,
        shapeIndex: c.shapeIndex,
        category: c.category,
      };
      if (c.label) item.label = c.label;
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
