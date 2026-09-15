# Piano: risolvere paragraphIndex per gli anchor dei commenti Word

**Stato**: implementato. `handleGetComments` in `word-commands.ts` ora usa
`resolveParagraphIndicesForComments` (bucketing a due fasi) per arricchire l'output con
`anchorStart`/`anchorEnd`/`startParagraphIndex`/`endParagraphIndex`/`localStart`/`localEnd`,
ed è stato corretto il bug del doppio `c.getRange()`. Test aggiunti in
`word-commands.test.ts` (incluso il caso limite multi-paragrafo). `TOOLS.md` aggiornato.

## Problema

`word_get_comments` (handler `handleGetComments`, `word-commands.ts:2144`) espone oggi
solo `anchorText` (primi 80 caratteri del testo su cui è ancorato il commento). Non
espone alcuna posizione (offset o paragrafo). Un chiamante che vuole *modificare* il
testo commentato (es. via `word_replace_text`, che richiede `paragraphIndex` +
`oldText`) non ha modo di ottenere `paragraphIndex` dal commento — deve indovinarlo o
fare ricerche testuali separate e fragili (`anchorText` potrebbe non essere univoco nel
documento).

## Bug esistente da correggere nello stesso intervento

In `handleGetComments`, righe 2166 e 2174, `c.getRange()` viene chiamato **due volte**
come due oggetti proxy Office.js distinti e tracciati separatamente:

```ts
// riga 2166 — .load() applicato a QUESTO oggetto range
c.getRange().load("text");
// ...
// riga 2174 — c.getRange() qui è una SECONDA chiamata, un oggetto diverso, mai .load()-ato
try { return String(c.getRange().text ?? "").slice(0, 80); }
```

Risultato: `anchorText` è sempre vuoto in produzione (confermato empiricamente:
14 commenti reali nel documento di test hanno tutti restituito `anchorText: ""`).
Il fix è banale — salvare il riferimento una sola volta:

```ts
const ranges = new Map<string, Word.Range>();
for (const c of comments.items) {
	for (const r of c.replies.items) r.load("id,content,authorName,authorEmail,creationDate");
	const range = c.getRange();
	range.load("text,start,end");
	ranges.set(c.id, range);
}
await ctx.sync();
// più avanti: usare ranges.get(c.id)!.text / .start / .end — mai richiamare c.getRange() di nuovo
```

Nota: `c.id` non è ancora caricato al momento in cui si crea la Map nello stesso loop
in cui si chiama `c.getRange()` — verificare l'ordine di `.load()`/`.sync()` esistente
(le proprietà scalari di `c` sono già caricate in un sync precedente, righe 2154-2158,
quindi `c.id` è già popolato quando si arriva al loop di riga 2160 — nessun problema, ma
va verificato in fase di implementazione che l'ordine dei sync non cambi).

## Requisito: `start`/`end` richiedono WordApiDesktop 1.4

Confermato empiricamente (test via Safari devtools su Word per Mac desktop) che
`Word.Range.start`/`Word.Range.end` funzionano sulla piattaforma target. Sono offset di
carattere **assoluti dall'inizio del documento**, senza alcun legame esposto con la
struttura OOXML (`<w:p>`/`<w:r>`) — Office.js non fornisce un metodo per risalire da un
offset numerico a un `Range` o a un indice di paragrafo. Va quindi calcolato
manualmente, con un algoritmo a due fasi per restare efficienti (vedi sotto).

## Perché non un algoritmo naive (full-scan o binary search adattiva)

- **Full-scan** (`body.paragraphs.items[i].getRange().load("start,end,text")` per
  *tutti* i paragrafi, poi test di overlap in JS) — funziona ma è O(n) di payload
  trasferito per ogni chiamata a `word_get_comments`, su documenti grandi diventa
  sprecato.
- **Binary search adattiva** (una probe = una comparazione = un `context.sync()`) — in
  Office.js il costo dominante è il numero di round-trip `sync()`, non i byte
  trasferiti (raccomandazione ufficiale: minimizzare le sync, non i dati). Una binary
  search su un documento con ~2000 paragrafi farebbe ~11 sync *sequenziali* (non
  batchabili, ognuna dipende dal risultato della precedente) — più lenta del full-scan
  a singola sync in un contesto locale add-in↔host.

## Algoritmo scelto: bucketing a due fasi, batchato su TUTTI i commenti insieme

Punto critico: le 2 sync vanno fatte **una volta per l'intera chiamata**
`word_get_comments`, non per singolo commento. Con N commenti, il costo totale deve
restare O(1) sync, non O(N).

### Fase 1 — una sync, campione sparso di paragrafi

Carica solo `start` (proprietà scalare più leggera, non `end`/`text`) per un
sottoinsieme campionato dei paragrafi con passo fisso `k` (es. `k = Math.ceil(Math.sqrt(n))`
per ottimalità asintotica, o un valore fisso come 32 se si preferisce semplicità):

```ts
const paragraphs = ctx.document.body.paragraphs;
paragraphs.load("items");
await ctx.sync();

const n = paragraphs.items.length;
const k = Math.max(1, Math.ceil(Math.sqrt(n)));
const sampleIndices: number[] = [];
for (let i = 0; i < n; i += k) sampleIndices.push(i);
if (sampleIndices[sampleIndices.length - 1] !== n - 1) sampleIndices.push(n - 1);

const sampleRanges = sampleIndices.map((i) => paragraphs.items[i].getRange());
sampleRanges.forEach((r) => r.load("start"));
await ctx.sync(); // sync #1
```

In JS locale (zero sync), per ogni commento e per **entrambi** `anchorStart` e
`anchorEnd` di tutti i commenti insieme, trova il bucket `[sampleIndices[j], sampleIndices[j+1])`
che contiene l'offset (scansione lineare sul campione, O(√n) per lookup, trascurabile).

### Fase 2 — una sola sync, unione di tutti i bucket necessari

Prima di caricare, calcola l'**unione** di tutti i range di paragrafi necessari per
risolvere tutti gli offset (start+end) di tutti i commenti — probabile sovrapposizione
alta se i commenti sono vicini nel documento:

```ts
const neededParagraphIndices = new Set<number>();
// per ogni bucket trovato in Fase 1, aggiungi tutti gli indici [bucketStart, bucketEnd)
// ...

const bucketRanges = [...neededParagraphIndices].map((i) => ({
	i,
	range: paragraphs.items[i].getRange(),
}));
bucketRanges.forEach(({ range }) => range.load("start,end,text"));
await ctx.sync(); // sync #2
```

Poi, per ogni commento, in JS locale:

1. Trova il paragrafo di `anchorStart` tra i `bucketRanges` caricati →
   `(anchorStart, startParagraphIndex, localStart = anchorStart - paragraph.start)`.
2. Controlla se `anchorEnd <= startParagraph.end` (già disponibile, zero sync
   aggiuntive) → se sì, stesso paragrafo: `endParagraphIndex = startParagraphIndex`,
   `localEnd = anchorEnd - startParagraph.start`.
3. Solo se il commento attraversa il confine del paragrafo (raro), risolvi
   `anchorEnd` separatamente usando gli stessi `bucketRanges` già caricati in Fase 2
   (nella grande maggioranza dei casi il bucket di Fase 2 già contiene anche il
   paragrafo di `anchorEnd`, perché i bucket sono contigui e i commenti multi-paragrafo
   sono rari — ma va gestito il caso in cui non lo contenga, con un fallback che
   ricarica il bucket corretto).

**Totale: 2 sync per l'intera chiamata `word_get_comments`, indipendentemente dal
numero di commenti o dalla dimensione del documento.**

### Nomenclatura output

Usare `startParagraphIndex`/`endParagraphIndex` (non un singolo `paragraphIndex`)
perché nel caso raro multi-paragrafo differiscono — un solo campo `paragraphIndex`
introdurrebbe ambiguità silenziosa in quel caso limite.

Struttura per commento nel risultato:

```ts
{
	// ...campi esistenti (id, author, email, date, text, resolved, anchorText, replies)...
	anchorStart: number,
	anchorEnd: number,
	startParagraphIndex: number,
	endParagraphIndex: number,
	localStart: number,   // offset dentro startParagraphIndex
	localEnd: number,     // offset dentro endParagraphIndex
}
```

## Integrazione: NESSUN nuovo tool MCP

Non introdurre un tool generico `word_get_text_range(start, end)`. Motivo: al momento
`comment.getRange()` è l'unica fonte di offset assoluti nel codebase — un tool pubblico
per risolverli avrebbe un solo chiamante ipotetico e costringerebbe il modello a due
round-trip (chiamare `word_get_comments` per leggere gli offset, poi richiamare un
secondo tool per risolverli) per un'informazione che può essere già presente nella
prima risposta.

Implementare invece la risoluzione come **funzione interna** (non esportata come tool),
es. `resolveParagraphIndicesForComments(ctx, comments)` dentro `word-commands.ts`,
chiamata direttamente da `handleGetComments` per arricchire il risultato. Zero tool
nuovi, zero round-trip aggiuntivi per il modello — l'output di `word_get_comments`
diventa immediatamente utilizzabile con `word_replace_text(startParagraphIndex,
oldText: anchorText, newText: ...)` senza passaggi intermedi.

Se in futuro emerge un secondo consumatore di offset assoluti (es. un altro tool che
restituisce anch'esso `start`/`end`), estrarre la funzione interna in un tool pubblico
dedicato sarà semplice — diventerebbe il body del nuovo handler, senza dover disfare
nulla di questo lavoro.

## Nota su cache/staleness (se in futuro si aggiunge una cache locale)

Gli offset `start`/`end` di ogni paragrafo successivo a un punto di modifica cambiano ad
ogni insert/delete nel documento. Se in futuro si introduce una cache locale di
paragrafi risolti (per servire più chiamate successive senza ripetere la Fase 1/2), la
cache deve essere invalidata con una verifica economica — confrontare un breve prefisso
del `text` cachato col `text` live del paragrafo — prima di fidarsi degli offset
cachati. Non implementare questa cache nella prima iterazione: aggiunge complessità e
un vettore di bug di correttezza silenziosi senza un beneficio misurato ancora.

## Fix da applicare insieme (stesso intervento)

1. Correggere il bug del doppio `c.getRange()` (vedi sezione "Bug esistente").
2. Aggiungere `anchorStart`/`anchorEnd`/`startParagraphIndex`/`endParagraphIndex`/
   `localStart`/`localEnd` al risultato di `handleGetComments`, tramite la nuova
   funzione interna di risoluzione a due fasi.
3. Aggiornare `TOOLS.md` (sezione Word → Comments, riga `word_get_comments`) per
   documentare i nuovi campi.
4. Aggiornare/estendere i test in `word-commands.test.ts` per la fixture di
   `word_get_comments` (verificare offset e paragraphIndex risolti su un documento
   mock multi-paragrafo con almeno un commento che attraversa un confine di paragrafo,
   per validare il caso limite).

## Non fare in questa iterazione

- Nessun nuovo tool MCP pubblico per risoluzione offset→paragrafo.
- Nessuna cache locale con invalidazione — rivalutare solo se emerge un bisogno reale
  di chiamate ripetute a distanza di tempo sullo stesso documento.
- Nessuna modifica a `McpToolEngine.cs` a parte, se necessario, l'aggiornamento delle
  description string dello schema tool per i nuovi campi opzionali in output (i tool
  output non hanno schema rigido lato dotnet, quindi probabilmente non serve toccare
  nulla lì — verificare in fase di implementazione).
