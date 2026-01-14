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

interface Finding {
  type: string;
  original: string;
  suggestion: string;
  occurrences: { slideIndex: number; shapeIndex: number }[];
  confidence: number;
}

interface AnalysisResponse {
  findings: Finding[];
  warnings: { type: string; slideIndex?: number; message: string }[];
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
        JSON.stringify({ findings: [], warnings: [] }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Prepare context for GPT
    const slideTexts = texts.map(t =>
      `[Slide ${t.slideIndex + 1}, Shape ${t.shapeIndex}]: ${t.text}`
    ).join("\n");

    const systemPrompt = `Du är en expert på att identifiera känslig affärsinformation i konsultpresentationer.

Din uppgift är att analysera text från PowerPoint-slides och identifiera information som bör anonymiseras innan presentationen delas externt.

Identifiera följande typer av känslig information:
- **company**: Kundnamn, företagsnamn, organisationer (ej generiska termer som "kunden" eller "företaget")
- **person**: Personnamn (för- och efternamn)
- **email**: E-postadresser
- **phone**: Telefonnummer
- **financial**: Specifika belopp, intäkter, kostnader (t.ex. "45 MSEK", "2,3 miljoner")
- **percentage**: Specifika procentsatser som kan vara känsliga (marknadsandelar, tillväxt)
- **date**: Specifika datum som kan identifiera projekt
- **project**: Projektnamn eller kodnamn
- **address**: Fysiska adresser
- **other**: Annan känslig information

För varje fynd, föreslå en lämplig anonym ersättning:
- Företag → [Klient A], [Klient B], etc.
- Personer → [Projektledare], [Kontaktperson], [Konsult], etc.
- Belopp → [X MSEK], [belopp], etc.
- E-post → [email borttagen]

Svara ENDAST med giltig JSON i följande format:
{
  "findings": [
    {
      "type": "company",
      "original": "Volvo",
      "suggestion": "[Klient A]",
      "confidence": 0.95
    }
  ]
}

Var noggrann:
- Inkludera INTE generiska termer som redan är anonyma
- Inkludera INTE vanliga ord som råkar matcha företagsnamn i fel kontext
- Var konsekvent - samma företag ska alltid bli [Klient A], samma person ska alltid få samma ersättning
- confidence ska vara 0.0-1.0 baserat på hur säker du är`;

    const userPrompt = `Analysera följande text från en presentation och identifiera känslig information:

${slideTexts}`;

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
    const findings: Finding[] = [];

    // Map findings back to slide/shape locations
    for (const finding of parsed.findings || []) {
      const occurrences: { slideIndex: number; shapeIndex: number }[] = [];

      // Find all occurrences of this text in the original data
      for (const textItem of texts) {
        if (textItem.text.includes(finding.original)) {
          occurrences.push({
            slideIndex: textItem.slideIndex,
            shapeIndex: textItem.shapeIndex,
          });
        }
      }

      if (occurrences.length > 0) {
        findings.push({
          type: finding.type,
          original: finding.original,
          suggestion: finding.suggestion,
          confidence: finding.confidence || 0.8,
          occurrences,
        });
      }
    }

    const response: AnalysisResponse = {
      findings,
      warnings: [],
    };

    return new Response(
      JSON.stringify(response),
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
