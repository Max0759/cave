// Relais Cloudflare Worker : reçoit la photo d'une étiquette depuis l'appli,
// la fait lire par Claude et renvoie les informations du vin en JSON.
// La clé API reste ici (secret Cloudflare) et n'est jamais envoyée au navigateur.
import Anthropic from "@anthropic-ai/sdk";

const MODELE = "claude-opus-5-5";
const TAILLE_MAX = 6 * 1024 * 1024; // image base64 max ~6 Mo

const SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["estEtiquette", "domaine", "cuvee", "appellation", "region", "couleur", "millesime", "format", "cepages", "gardeDebut", "gardeFin", "prixIndicatif", "texteLu", "remarque"],
  properties: {
    estEtiquette: { type: "boolean", description: "true si la photo montre bien une bouteille ou une étiquette de vin" },
    domaine: { type: "string", description: "Domaine, château ou producteur. Chaîne vide si illisible." },
    cuvee: { type: "string", description: "Nom de la cuvée s'il y en a une, sinon chaîne vide." },
    appellation: { type: "string", description: "Appellation (AOC/AOP/IGP), ex. « Côte-Rôtie »." },
    region: { type: "string", enum: ["Alsace", "Beaujolais", "Bordeaux", "Bourgogne", "Champagne", "Jura", "Languedoc", "Loire", "Provence", "Rhône", "Savoie", "Sud-Ouest", "Étranger", ""] },
    couleur: { type: "string", enum: ["Rouge", "Blanc", "Rosé", "Jaune", "Effervescent", "Liquoreux", ""] },
    millesime: { type: ["integer", "null"], description: "Année lue sur l'étiquette, null si absente ou illisible." },
    format: { type: "string", enum: ["Bouteille", "Demi-bouteille", "Magnum", "Jéroboam"] },
    cepages: { type: "string", description: "Cépages habituels de ce vin (connaissance générale), ex. « Syrah, Viognier »." },
    gardeDebut: { type: ["integer", "null"], description: "Année à partir de laquelle boire, estimée selon le millésime et le vin." },
    gardeFin: { type: ["integer", "null"], description: "Année limite conseillée." },
    prixIndicatif: { type: ["number", "null"], description: "Prix public indicatif en euros par bouteille si raisonnablement connu, sinon null." },
    texteLu: { type: "string", description: "Principaux mots lus sur l'étiquette, tels quels." },
    remarque: { type: "string", description: "Courte remarque en français (incertitude, mention particulière), ou chaîne vide." }
  }
};

const CONSIGNE = `Tu aides à gérer une cave à vin personnelle. On te montre la photo d'une bouteille ou de son étiquette.
Lis l'étiquette et identifie le vin. Ne devine pas le millésime : s'il n'est pas lisible, mets null.
Pour les cépages, la période de garde et le prix indicatif, utilise tes connaissances générales sur ce vin ; laisse vide ou null si tu n'en sais rien.
Réponds en français.`;

function cors(env, origine) {
  const ok = origine && (origine === env.ORIGINE_AUTORISEE || origine.startsWith("http://localhost"));
  return {
    "Access-Control-Allow-Origin": ok ? origine : env.ORIGINE_AUTORISEE,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Cave-Code",
    "Vary": "Origin",
  };
}
const json = (corps, status, entetes) =>
  new Response(JSON.stringify(corps), { status, headers: { "Content-Type": "application/json", ...entetes } });

export default {
  async fetch(request, env) {
    const h = cors(env, request.headers.get("Origin"));
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: h });
    if (request.method !== "POST") return json({ erreur: "Méthode non autorisée" }, 405, h);
    if (!env.CAVE_CODE || request.headers.get("X-Cave-Code") !== env.CAVE_CODE) return json({ erreur: "Code d'accès incorrect" }, 401, h);

    let corps;
    try { corps = await request.json(); } catch { return json({ erreur: "Requête illisible" }, 400, h); }
    const image = String(corps.image || "");
    const type = ["image/jpeg", "image/png", "image/webp"].includes(corps.type) ? corps.type : "image/jpeg";
    if (!image || image.length > TAILLE_MAX) return json({ erreur: "Image absente ou trop lourde" }, 400, h);

    const client = new Anthropic({ apiKey: env.ANTHROPIC_API_KEY });
    try {
      const reponse = await client.beta.messages.create({
        model: MODELE,
        max_tokens: 4000,
        betas: ["server-side-fallback-2026-07-01"],
        fallbacks: "default",
        output_config: { effort: "low", format: { type: "json_schema", schema: SCHEMA } },
        system: CONSIGNE,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: type, data: image } },
            { type: "text", text: `Année en cours : ${new Date().getFullYear()}. Identifie ce vin.` },
          ],
        }],
      });
      if (reponse.stop_reason === "refusal") return json({ erreur: "Lecture refusée pour cette image" }, 422, h);
      if (reponse.stop_reason === "max_tokens") return json({ erreur: "Réponse incomplète, réessayez" }, 502, h);
      const texte = reponse.content.filter(b => b.type === "text").map(b => b.text).join("");
      return json({ vin: JSON.parse(texte) }, 200, h);
    } catch (e) {
      if (e instanceof Anthropic.RateLimitError) return json({ erreur: "Trop de demandes, réessayez dans un instant" }, 429, h);
      if (e instanceof Anthropic.AuthenticationError) return json({ erreur: "Clé API Claude invalide (voir le relais)" }, 500, h);
      if (e instanceof Anthropic.APIError) return json({ erreur: "Service IA indisponible (" + e.status + ")" }, 502, h);
      if (e instanceof SyntaxError) return json({ erreur: "Réponse IA illisible, réessayez" }, 502, h);
      return json({ erreur: "Erreur du relais" }, 500, h);
    }
  },
};
