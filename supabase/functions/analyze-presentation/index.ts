import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

interface TextItem {
  slideIndex: number;
  shapeIndex: number;
  text: string;
}

interface RewriteItem {
  slideIndex: number;
  shapeIndex: number;
  rewrittenText: string;
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

    const { texts } = await req.json() as { texts: TextItem[] };

    if (!texts || texts.length === 0) {
      return new Response(
        JSON.stringify({ rewrites: [] }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Prepare context for GPT
    const textEntries = texts.map(t =>
      `{ "slideIndex": ${t.slideIndex}, "shapeIndex": ${t.shapeIndex}, "text": ${JSON.stringify(t.text)} }`
    ).join(",\n");

    const systemPrompt = `Du är en expert på att anonymisera konsultpresentationer.

Du får text från enskilda shapes i en PowerPoint-presentation.
Skriv om VARJE text så att ALL identifierande information ersätts med generiska platshållare.

Regler:
- Företagsnamn → [Företag A], [Företag B] etc. (konsekvent genom hela presentationen)
- Personnamn → [Projektledare], [Kontaktperson], [Konsult] etc.
- E-postadresser → [email]
- Telefonnummer → [telefon]
- Specifika belopp → [X MSEK], [belopp] etc.
- Specifika procent → [X]%
- Datum → [datum] eller [kvartal] etc.
- Adresser → [adress]
- Projektnamn/kodnamn → [Projekt A], [Projekt B] etc.

KRITISKT:
- Behåll EXAKT samma antal rader (radbrytningar) i varje text
- Behåll samma struktur, ton och längd
- Text som redan är generisk eller inte innehåller känslig info ska returneras oförändrad
- Var konsekvent — samma företag ska alltid bli samma platshållare

Svara ENDAST med giltig JSON i följande format:
{
  "rewrites": [
    { "slideIndex": 0, "shapeIndex": 0, "rewrittenText": "..." }
  ]
}`;

    const userPrompt = `Anonymisera följande texter från en presentation. Returnera en omskrivning per shape:

[${textEntries}]`;

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
        temperature: 0.1,
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
    const rewrites: RewriteItem[] = (parsed.rewrites || []).map((r: RewriteItem) => ({
      slideIndex: r.slideIndex,
      shapeIndex: r.shapeIndex,
      rewrittenText: r.rewrittenText,
    }));

    return new Response(
      JSON.stringify({ rewrites }),
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
