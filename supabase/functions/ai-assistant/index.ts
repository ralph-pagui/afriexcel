// Supabase Edge Function: ai-assistant
// Moteur d'IA propulsé par Google Gemini 3.6 Flash avec rotation automatique des 7 clés

import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

// Clés Gemini 3.6 Flash lues depuis les variables d'environnement Deno Supabase
const envKeys = Deno.env.get("GEMINI_API_KEYS") || Deno.env.get("GEMINI_API_KEY") || "";
const GEMINI_KEYS = envKeys
  ? envKeys.split(",").map(k => k.trim()).filter(Boolean)
  : ["DEMO_GEMINI_KEY"];

let keyIndex = 0;

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const { prompt, country, excelContext } = await req.json();

    const systemInstruction = `Tu es l'Assistant IA Excel Universel de classe mondiale pour Microsoft Excel (2016 à 365), spécialisé pour les entreprises, directions financières et cabinets comptables en Afrique Francophone et espace OHADA.
Tu es capable d'exécuter N'IMPORTE QUELLE demande professionnelle dans Excel avec une rigueur absolue :
1. Calculs & Formules : Formules simples ou complexes en français (RECHERCHEV, INDEX/EQUIV, SOMME.SI.ENS, NB.SI.ENS, SI.CONDITIONS, TRIER, FILTRE).
2. Modélisation financière & SYSCOHADA : Balances, Bilans Actif/Passif, Compte de Résultat (Marge, VA, EBE), amortissements, budgets, cash-flow.
3. Rapprochement bancaire & Mobile Money : Lettrage automatique des comptes 521/571 contre Wave, Orange Money, MTN MoMo, agios au compte 631.
4. Paie locale & Fiscalité : Cotisations, impôts sur salaires, TVA.
5. Visualisation & Tableaux : Graphiques en colonnes, barres, lignes ou camemberts avec titres, tableaux structurés avec filtres.
6. Design & Finition Professionnelle : En-têtes corporate vert Excel (#107C41) ou bleu marine (#1E3A8A), séparateurs de milliers, bordures d'audit comptable.

CATALOGUE DES ACTIONS DISPONIBLES :
- {"type": "WRITE_RANGE", "address": "A1:E10", "values": [[...]], "formatAsCurrency": true, "hasHeader": true}
- {"type": "APPLY_FORMULA", "targetCell": "E2", "formula": "=SOMME.SI.ENS(...)", "explanation": "..."}
- {"type": "FORMAT_RANGE", "address": "A1:E1", "fill": "#107C41", "fontColor": "#FFFFFF", "bold": true, "alignment": "Center"}
- {"type": "FORMAT_RANGE", "address": "A10:E10", "bold": true, "border": "double-bottom"}
- {"type": "CREATE_CHART", "chartType": "ColumnClustered"|"BarClustered"|"Line"|"Pie", "sourceRange": "A1:D10", "title": "Titre du Graphique"}
- {"type": "CREATE_TABLE", "rangeAddress": "A1:E10", "tableName": "NomTableau", "hasHeaders": true}
- {"type": "SORT_RANGE", "rangeAddress": "A2:E20", "columnIndex": 2, "ascending": false}
- {"type": "ADD_SHEET", "sheetName": "NomFeuille", "activate": true}
- {"type": "AUTOFIT", "rangeAddress": "A:E"}
- {"type": "CLEAN_DATA"}
- {"type": "EXECUTE_CODE", "code": "...", "description": "..."}

FORMAT DE SORTIE STRICT JSON :
{
  "message": "Explication claire, professionnelle et valorisante en français",
  "actions": [...],
  "suggestedFollowUp": "Une suggestion d'approfondissement pertinent pour l'utilisateur"
}`;

    const userContent = `Pays sélectionné : ${country.name} (${country.currencySymbol} - ${country.currency})
Données actuelles de la feuille Excel :
Adresse : ${excelContext?.address || "Non spécifiée"}
Échantillon : ${JSON.stringify(excelContext?.sampleValues || [])}

Demande de l'utilisateur :
"${prompt}"`;

    let aiResult = null;
    let lastError = null;

    // Rotation sur les 7 clés Gemini
    for (let i = 0; i < GEMINI_KEYS.length; i++) {
      const currentKey = GEMINI_KEYS[(keyIndex + i) % GEMINI_KEYS.length];

      try {
        const geminiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${currentKey}`;

        const requestBody = {
          contents: [{ role: "user", parts: [{ text: userContent }] }],
          systemInstruction: { parts: [{ text: systemInstruction }] },
          generationConfig: {
            responseMimeType: "application/json",
            temperature: 0.2
          }
        };

        const response = await fetch(geminiEndpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(requestBody)
        });

        if (response.ok) {
          const geminiData = await response.json();
          const rawText = geminiData.candidates?.[0]?.content?.parts?.[0]?.text;
          if (rawText) {
            aiResult = JSON.parse(rawText);
            keyIndex = (keyIndex + i + 1) % GEMINI_KEYS.length;
            break;
          }
        }
      } catch (err: any) {
        lastError = err;
      }
    }

    // Fallback de secours sur Groq (openai/gpt-oss-120b) si Gemini est saturé
    if (!aiResult) {
      const rawGroq = Deno.env.get("GROQ_API_KEYS") || Deno.env.get("GROQ_API_KEY") || "";
      const GROQ_KEYS = rawGroq
        ? rawGroq.split(",").map(k => k.trim()).filter(Boolean)
        : ["DEMO_GROQ_KEY"];

      for (let g = 0; g < GROQ_KEYS.length; g++) {
        try {
          const groqRes = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${GROQ_KEYS[g]}`,
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              model: "openai/gpt-oss-120b",
              messages: [
                { role: "system", content: systemInstruction },
                { role: "user", content: userContent }
              ],
              temperature: 0.15,
              max_tokens: 4096,
              reasoning_effort: "low",
              response_format: { type: "json_object" }
            })
          });

          if (groqRes.ok) {
            const groqData = await groqRes.json();
            const content = groqData.choices?.[0]?.message?.content;
            if (content) {
              aiResult = JSON.parse(content);
              break;
            }
          }
        } catch (e) {
          // Continuer sur la clé suivante
        }
      }
    }

    if (!aiResult) {
      throw lastError || new Error("Toutes les clés Gemini et Groq en rotation sont indisponibles.");
    }

    return new Response(JSON.stringify(aiResult), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
      status: 200
    });
  } catch (error: any) {
    return new Response(JSON.stringify({ error: error.message }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
      status: 500
    });
  }
});
