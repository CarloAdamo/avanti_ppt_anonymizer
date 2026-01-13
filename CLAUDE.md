# Avanti PPT Anonymizer

## Projektöversikt

PowerPoint-tillägg som hjälper konsulter att automatiskt anonymisera presentationer innan de delas eller laddas upp till företagets slide-bibliotek.

**Syfte:** När en konsult blir klar med ett projekt och vill dela sina bästa slides, måste känslig kundinformation tas bort först. Detta plugin automatiserar den processen med hjälp av AI.

## Relaterat projekt

- **avanti_ppt_template** - Slide-bibliotek där anonymiserade slides laddas upp
- **avanti-slide-pipeline** - Pipeline som processar uppladdade slides

## Funktionalitet

### Vad ska anonymiseras?

**Text:**
- Kundnamn/företagsnamn (t.ex. "Volvo" → "[Klient A]")
- Personnamn (t.ex. "Erik Svensson" → "[Projektledare]")
- Specifika siffror (intäkter, marknadsandelar, etc.)
- Projektnamn, datum, platser
- Konfidentiella affärsdata

**Visuellt:**
- Kundlogotyper (identifieras, användaren väljer åtgärd)
- Bilder på personer
- Screenshots av kundsystem

**Metadata:**
- Författare
- Kommentarer
- Revisionshistorik

### Användarflöde

```
┌─────────────────────────────────────────────────────────────────┐
│  1. Konsult öppnar presentation i PowerPoint                    │
│  2. Klickar "Anonymisera" i task pane                           │
│  3. Plugin scannar alla slides och extraherar text              │
│  4. AI (GPT-4) analyserar och identifierar känslig info         │
│  5. Användaren ser lista: "Hittade potentiellt känslig info"    │
│     ┌─────────────────────────────────────────────────────┐     │
│     │ ☑ "Volvo" (12 förekomster) → [Klient A]             │     │
│     │ ☑ "Erik Svensson" (3 förekomster) → [Projektledare] │     │
│     │ ☐ "45 MSEK" (2 förekomster) → [X MSEK]              │     │
│     │ ⚠ Logotyp hittad på slide 3 (manuell åtgärd)        │     │
│     └─────────────────────────────────────────────────────┘     │
│  6. Användaren granskar, justerar och godkänner                 │
│  7. Plugin ersätter text programmatiskt via Office.js           │
│  8. Klar! Presentationen kan nu delas/laddas upp                │
└─────────────────────────────────────────────────────────────────┘
```

## Teknisk arkitektur

```
┌─────────────────────────────────────────────────────┐
│  PowerPoint                                         │
│  ┌───────────────────────────────────────────────┐  │
│  │  Task Pane (Add-in)                           │  │
│  │  - "Scanna presentation"-knapp                │  │
│  │  - Lista med hittad känslig info              │  │
│  │  - Ersättningsförslag (redigerbara)           │  │
│  │  - "Anonymisera"-knapp                        │  │
│  └───────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────┘
            │
            ▼
┌─────────────────────────────────────────────────────┐
│  Supabase Edge Function                             │
│  ┌───────────────────────────────────────────────┐  │
│  │  analyze-presentation                         │  │
│  │  - Tar emot extraherad text från alla slides  │  │
│  │  - Anropar OpenAI GPT-4 för NER-analys        │  │
│  │  - Returnerar lista med känslig info +        │  │
│  │    föreslagna ersättningar                    │  │
│  └───────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────┘
            │
            ▼
┌─────────────────────────────────────────────────────┐
│  OpenAI API                                         │
│  - GPT-4 för Named Entity Recognition (NER)         │
│  - Identifierar: företag, personer, siffror, etc.   │
│  - Genererar kontextuella ersättningar              │
└─────────────────────────────────────────────────────┘
```

## Office.js API:er som används

### Läsa text från shapes
```javascript
await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        for (const shape of shapes.items) {
            shape.textFrame.textRange.load("text");
            await context.sync();
            console.log(shape.textFrame.textRange.text);
        }
    }
});
```

### Ersätta text i shapes
```javascript
shape.textFrame.textRange.text = text.replace(/Volvo/g, "[Klient A]");
await context.sync();
```

### Ta bort shapes (t.ex. logotyper)
```javascript
shape.delete();
await context.sync();
```

## Filstruktur

```
/
├── manifest.xml          # Office Add-in manifest
├── taskpane.html         # UI för anonymisering
├── taskpane.js           # Huvudlogik
├── taskpane.css          # Styles
├── assets/               # Ikoner
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── supabase/
    └── functions/
        └── analyze-presentation/
            └── index.ts  # Edge Function för AI-analys
```

## Teknisk stack

- **Frontend:** Vanilla JS, Office.js API
- **Hosting:** GitHub Pages
- **Backend:** Supabase Edge Functions
- **AI:** OpenAI GPT-4 (för NER och ersättningsförslag)

## API-design

### analyze-presentation Edge Function

**Request:**
```json
POST /functions/v1/analyze-presentation
{
    "slides": [
        {
            "index": 1,
            "texts": [
                "Volvo Q3 Rapport 2024",
                "Sammanställd av Erik Svensson"
            ]
        },
        {
            "index": 2,
            "texts": [
                "Omsättning: 45 MSEK",
                "Kontakt: erik.svensson@volvo.com"
            ]
        }
    ]
}
```

**Response:**
```json
{
    "findings": [
        {
            "type": "company",
            "original": "Volvo",
            "suggestion": "[Klient A]",
            "occurrences": [
                { "slideIndex": 1, "shapeIndex": 0 },
                { "slideIndex": 2, "shapeIndex": 1 }
            ],
            "confidence": 0.95
        },
        {
            "type": "person",
            "original": "Erik Svensson",
            "suggestion": "[Projektledare]",
            "occurrences": [
                { "slideIndex": 1, "shapeIndex": 1 }
            ],
            "confidence": 0.92
        },
        {
            "type": "email",
            "original": "erik.svensson@volvo.com",
            "suggestion": "[email borttagen]",
            "occurrences": [
                { "slideIndex": 2, "shapeIndex": 1 }
            ],
            "confidence": 0.99
        },
        {
            "type": "financial",
            "original": "45 MSEK",
            "suggestion": "[X MSEK]",
            "occurrences": [
                { "slideIndex": 2, "shapeIndex": 0 }
            ],
            "confidence": 0.88
        }
    ],
    "warnings": [
        {
            "type": "potential_logo",
            "slideIndex": 1,
            "message": "Möjlig logotyp detekterad - granska manuellt"
        }
    ]
}
```

## Deployment

1. **Frontend:** GitHub Pages (samma som avanti_ppt_template)
2. **Edge Function:** Supabase (samma projekt som slide-biblioteket)
3. **Manifest:** Separat manifest.xml för detta tillägg

## Utvecklingsplan

### Fas 1: Grundläggande struktur
- [ ] Sätt upp projekt med manifest.xml
- [ ] Skapa basic taskpane UI
- [ ] Implementera text-extraktion från alla slides

### Fas 2: AI-integration
- [ ] Skapa Edge Function för analys
- [ ] Integrera OpenAI GPT-4
- [ ] Implementera NER-prompt

### Fas 3: Anonymisering
- [ ] Visa hittade entiteter i UI
- [ ] Låt användaren redigera ersättningar
- [ ] Implementera faktisk text-ersättning

### Fas 4: Polish
- [ ] Hantera edge cases (tabeller, diagram)
- [ ] Förbättra UI/UX
- [ ] Testa med riktiga presentationer

## URLs

- **GitHub Pages:** https://carloadamo.github.io/avanti_ppt_anonymizer/
- **Repo:** https://github.com/CarloAdamo/avanti_ppt_anonymizer
- **Supabase:** https://vnjcwffdhywckwnjothu.supabase.co (samma som template-projektet)
