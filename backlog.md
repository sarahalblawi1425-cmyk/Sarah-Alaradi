# Mirsah (مرساة): Project Brief & Backlog

> National marine-insurance and route-risk platform for Saudi Arabia.
> Naqlthon 5 (نقل ثون 5), Track 4: Innovation in Logistics Services & Supply Chains.
> Status: **concept + clickable front-end prototype**. Every number and insurer name in the prototype is illustrative.

---

## 1. Project summary

### The problem
- **Geopolitical tension raised war-risk insurance sharply.** A ship crossing a security-tension zone (Red Sea, Bab el-Mandeb, Strait of Hormuz) pays an extra *war-risk premium*.
  - Before the tension: about **$100K** per tanker trip.
  - After: about **$3M** from Yanbu, and about **$7M** from southern Saudi ports or via Hormuz.
  - Yanbu tanker premiums went from **<1% to ~3%** of hull value, and can reach **7%** for ports south of Yanbu such as Jazan.
  - Source: Reuters via Insurance Journal, Sep 2026.
- **Premiums are volatile.** They usually cover 7 days and are repriced every 24 hours.
- **One foreign reference dominates.** The Joint War Committee (JWC) in London effectively monopolises risk-zone assessment, and most countries and insurers follow it. Saudi Arabia has no national risk index for its own trade routes.
- **No price transparency.** Agreed prices aren't public, and the figures that circulate are broker quotes.
- **Shippers decide blind.** They contact each insurer separately and can't see how the premium changes if they switch port or route.
- **Weak local insurers.** Marine, aviation, transport and liability together are only ~4.5% of Saudi written premiums (GlobalData, 2023), so premiums flow to foreign insurers.

### The solution
Mirsah has two parts:
1. **A Saudi maritime risk index (مؤشر مرساة) that we build and own.** It uses our own methodology and national data. ⚠️ **It does not depend on JWC assessments.** This is a core decision.
2. **One website** that gives the shipper risk per route, an estimated premium, a neutral comparison of Saudi insurers, and alerts on the best time to insure.

The final price is always confirmed by the insurer.

### Strategy
1. **National data layer.** The Kingdom owns it. It combines port data, security incidents, vessel tracking (AIS) and local insurer data, and grows in value with every voyage.
2. **Independent methodology.** It is transparent and Saudi-made, and is not based on the JWC.
3. **Regional, then global.** Start with the waters around the Kingdom, then expand into a reference other countries can use.

**Why:** Vision 2030 positions KSA as a global logistics hub, so it should have its own risk measure. This draws on the DIKW pyramid (Data → Information → Knowledge → Wisdom): whoever owns the data owns the decision.

### Alignment
- **Vision 2030:**
  - The target is 40M TEU/year through Saudi ports by 2030 (8.3M TEU in 2025, +10.58%).
  - Clearer insurance costs encourage shippers to choose Saudi ports.
- **National War-Risk Insurance Pool (الوعاء الوطني لتأمين مخاطر الحرب):**
  - Approved Sep 2026 and led by the Saudi Reinsurance Company.
  - Mirsah is its **digital front door**: it guides shippers to participating insurers. It complements the pool and does not compete with it.
- **Track 4 question ("How can delivery cost and time be reduced?"):**
  - Total cost per route is known before deciding.
  - No more contacting insurers one by one.
  - Smarter timing through alerts.
  - Premiums stay in the Saudi economy by empowering local insurers.

### Beneficiaries
1. Exporters, importers and shipping companies (primary).
2. Saudi marine insurers.
3. The national pool.
4. Regulators: the Insurance Authority (هيئة التأمين) and the Transport General Authority (TGA, الهيئة العامة للنقل).

### Proposed owner
A joint project of the Insurance Authority and TGA under the Ministry of Transport & Logistic Services, coordinated with Mawani (موانئ) and Saudi Re.

### Implementation plan (5 years)

| # | Phase | Years |
|---|---|---|
| 1 | Build the Saudi risk index (ports, incidents, AIS; transparent method) | 1–2 |
| 2 | Agreements with at least 2 Saudi insurers | 2 |
| 3 | Build and secure the platform | 2–3 |
| 4 | Pilot with a limited set of companies | 3 |
| 5 | Official launch; adopt the index locally, then regionally | 4–5 |

### Key risks and mitigations

| Risk | Mitigation |
|---|---|
| **JWC monopoly (biggest challenge)** | Start with our own market and routes; earn trust with a transparent method and independent review; then expand, leaning on the Kingdom's logistics position |
| Insurers hesitant to join | Position the platform as a channel for new clients; start with limited pilots |
| Prices change fast | Show prices as estimates with timestamp and validity; the insurer confirms |
| Data sensitivity and privacy | Restricted access; governance under the Personal Data Protection Law (PDPL) |
| Recommendation bias | Show why each item is ranked where it is; no paid ranking |

### Competitors / existing solutions

| Solution | What it does | What Mirsah adds |
|---|---|---|
| ShipWRITE (Canopius) | Instant war-risk quote and issue, global vessel tracking | Compares several Saudi insurers instead of one |
| WarRates.com | Indicative hull war-risk rates | Links price to the route, the two ports and the cargo type |
| Insurwave + Ambrey | Global maritime threat data | National focus, with local port and insurer data |
| DP World cargo war-risk | One door-to-door policy | Neutral comparison, not a single product |
| Local insurance comparison platforms | Licensed e-brokers | A Saudi risk index, with premiums estimated from it |
| National war-risk pool | National mechanism (Saudi Re) | A complementary digital front door |

### KPIs
- Number of participating insurers.
- Ports and routes covered.
- Time to get a unified comparison.
- Monthly active users and saved shipments.
- Quote requests sent to local insurers.

---

## 2. The prototype

| | |
|---|---|
| **File** | `presentation/prototype/mirsah-prototype.html`: a single self-contained HTML file with vanilla JS and no build step. Open it directly in a browser. |
| **Live page** | https://claude.ai/artifact/Q8Ujte34o184SVyzLyZfK3 (private; the owner must share it) |
| **Language / direction** | Arabic, RTL (`<html lang="ar" dir="rtl">`) |
| **Font** | IBM Plex Sans Arabic (Google Fonts), falling back to Tahoma / system |
| **Banner on every page** | «نموذج أولي لشرح الفكرة»: figures are illustrative, prices are estimates reviewed every 48h, and the insurer gives final confirmation |

### Design tokens (from the site; reuse these)

| Token | Hex | Use |
|---|---|---|
| Navy | `#0B2545` | Header, primary dark surfaces, logo background |
| Deep navy | `#0B1D33` | Radar/map panel |
| Teal | `#0E7C86` | Primary buttons, accents |
| Teal dark | `#0B5C66` | Links, tags text |
| Ice | `#8FD3DA` | Logo anchor, highlights on dark |
| Ice background | `#DDF2F3` | Tags, soft highlights |
| Page background | `#F4F6F9` | |
| Border | `#DDE3EB` | Cards |
| Muted text | `#4A5568` | |

**Risk scale (0–10):**

| Band | Score | Map colour | Text colour |
|---|---|---|---|
| Low (منخفض) | 0–3 | `#3FA36B` | `#2E7D4F` |
| Medium (متوسط) | 3–5 | `#E0A526` | `#A86A0C` |
| High (مرتفع) | 5–7 | `#EA6A2A` | `#C2410C` |
| Critical (حرج) | 7–10 | `#E5383B` | `#B42318` |

**Logo:** a stroked anchor icon in Ice on Navy, with the word «مرساة» in white.

### Screens

Navigation (header) has 5 buttons. Only the first 3 work.

#### 2.1 البحث عن المسارات: Route search + Risk radar (`#search`)
- **Search form:** origin port (Shanghai / Singapore / Mumbai), Saudi destination ("any port – compare all", Jeddah, Dammam, Yanbu, Jazan), cargo type, sailing date, and a search button.
- **Route list:** 3 sample routes from Shanghai. Each card shows max risk score + band, a risk bar, sailing days, distance and the count of risk zones. Clicking a card selects that route.
  1. To Jeddah via Bab el-Mandeb: risk 7.4, 18 days, 6,600 nm.
  2. To Dammam via Hormuz: risk 5.8, 16 days, 5,900 nm. Tagged "best balance".
  3. To Jeddah via Cape of Good Hope + Suez: risk 3.2, 33 days, 14,500 nm.
- **Radar map:** a hand-drawn SVG of the region with:
  - glowing risk zones;
  - zone badges (Bab el-Mandeb 7.4, Hormuz 5.8, North Red Sea 2.1);
  - an animated radar sweep;
  - animated ships: white ones on the selected route, ice-coloured ones for general AIS traffic;
  - a "live · 214 ships" pill and a legend.
- **Segments:** the zones the selected route crosses, each with its score. A CTA goes to Prices.

#### 2.2 تفقد الأسعار: Prices (`#prices`)
- **Inputs:** route chips (synced with the search screen), cargo value in SAR (live input), and policy type.
- **Tiles:** lowest premium, average rate %, 14-day change %, and route risk score.
- **14-day premium-rate line chart** (SVG).
- **AI timing recommendation card:** text + confidence %, plus a «كيف حُسبت؟» link.
- **Price alert:** a target premium is pre-filled at 90% of the lowest; a "save & activate" button.
- **Insurer offers table:** 4 illustrative insurers (أ/ب/ج/د) with policy type, rate %, estimated premium, quote validity, "national pool participant" label and a «اطلب عرضاً رسمياً» button. The note says there is no paid ranking.

#### 2.3 شركات التأمين: Insurers (`#insurers`), added last
- A banner explaining the National War-Risk Pool.
- **Filter chips:** All (6), Pool participants (4), Containers, Oil, Vehicles, Refrigerated.
- **6 insurer cards (أ–و)** showing:
  - user rating;
  - a "✓ مشارك في الوعاء الوطني" badge;
  - cargo types;
  - policy types;
  - indicative premium range (% of cargo value);
  - a «اطلب عرضاً رسمياً» button.

#### 2.4 مشاريعي (My projects) and المنهجية والمصادر (Methodology & sources)
- Nav buttons exist, but **there are no pages yet.**

### Code map (inside `<script>`)

| Item | What it is |
|---|---|
| `ROUTES[]` | Per route: `id, chip, name, via, risk, days, nm, rate, rec, segs[[name,score]], trend[14], advice, conf, pts[[x,y]]` |
| `INSURERS[]` | Offers table: `[name, inPool, policyLabel, rateMultiplier, validity]` |
| `INS_CARDS[]` | Insurer cards: `[name, inPool, cargoTypes[], policyTypes[], rangeText, rating]` |
| `sel` | Currently selected route id (default 2) |
| `renderSearch()` / `renderPrices()` / `renderAll()` | Re-render the views from state |
| `show(view)` | Toggles `.view.on` and the nav state |
| `color(r)` / `label(r)` | Risk score → colour / Arabic band label |
| Radar animation | `SHIPS[]`, `at(pts, f, rev)` (position along a polyline) and `frame()` (`requestAnimationFrame` loop) |

**How the numbers are calculated (all mock):**
- Lowest premium = `value × rate × 0.94 / 100`
- Average rate = `rate × 1.04`
- Each offer = `rate × multiplier`
- Trend change = `(trend[13] − trend[0]) / trend[0]`

### What works vs. what is mocked
- ✅ **Works:**
  - Tab switching.
  - Route selection, synced across screens and the map.
  - Live premium recalculation from the cargo value.
  - Chart redraw.
  - Radar and ship animation.
- ⚠️ **Static / not wired:**
  - The search form (selects and the button do nothing; only Shanghai data exists).
  - The insurer filter chips.
  - The «اطلب عرضاً رسمياً» buttons.
  - The notification bell.
  - The «كيف حُسبت؟» link.
  - The policy-type select.
  - "Save alert" (only changes the button label).
  - مشاريعي and المنهجية والمصادر.
- ⚠️ **All data is hard-coded.** Risk scores, prices, ship count and insurer names are illustrative. The map is a schematic, not geographic.

---

## 3. Product rules (don't break these)
1. **The risk index is Saudi and independent.** Never present JWC lists as an input or a dependency. JWC may be mentioned only as context, the incumbent or the competitor.
2. **AI is never used to decide risk zones.** Zones come from the index and official data. AI only does three things: timing alerts with a confidence score, insurer recommendations with the reason shown, and an assistant that answers questions.
3. **All prices are estimates.** Show a timestamp and validity, and say that the insurer confirms the final price.
4. **No paid ranking.** Always show why something ranks where it does.
5. **Use placeholder insurer names (مؤمّن أ…و).** Don't use real company names or branding in the prototype.
6. **Arabic-first, RTL.** Keep the existing tokens and the risk colour scale.

---

## 4. Backlog

Priority: **P0** = before the demo, **P1** = strengthens the pitch, **P2** = later / production.

### P0: make the demo feel real

| ID | Feature | Acceptance criteria |
|---|---|---|
| P0-1 | **Wire the search form** | Origin / destination / cargo / date drive the route list. Add mock routes for Singapore and Mumbai. "Any Saudi port" shows routes to all 4 ports, sorted by risk and premium. The cargo type changes the rate (e.g. oil ×1.2, refrigerated ×1.1). |
| P0-2 | **Insurer filters** | The chips filter the cards (pool participants and cargo types) and the count updates. Show an empty state when nothing matches. |
| P0-3 | **Request-a-quote flow** | Clicking «اطلب عرضاً رسمياً» opens an in-page modal pre-filled with route, cargo, value and insurer, plus contact fields. Submit shows a confirmation ("تم إرسال الطلب" + reference no.). There is no real network call. |
| P0-4 | **مشاريعي page** | Saved shipments list (route, value, current premium, target, status), seeded with 2 examples. "Save alert" on Prices adds a project here, using `localStorage` wrapped in try/catch. |
| P0-5 | **المنهجية والمصادر page** | Explains the index inputs (incidents, AIS density, port data, claims/premiums, national alerts), illustrative weights, the 0–10 bands, update frequency, the "independent of JWC" stance, data sources and disclaimers. |
| P0-6 | **«كيف حُسبت؟» panel** | Explains the timing recommendation: 14-day trend, volatility, risk changes, confidence. Must state "AI recommends; it doesn't set prices or risk zones". |

### P1: strengthen the story

| ID | Feature | Acceptance criteria |
|---|---|---|
| P1-1 | **Radar browse mode** | Usable without entering a trip. Click or hover a zone to see score, band, main drivers and last update. Layer toggles: risk zones / ships / routes. |
| P1-2 | **Notifications panel** | The bell opens a list of price-change and risk-change alerts tied to مشاريعي. |
| P1-3 | **AI assistant (mock)** | A chat drawer with 4–5 scripted Q&As (e.g. "why is Hormuz 5.8?", "cheapest route to Jeddah?"). |
| P1-4 | **Insurer detail view** | Coverage, policies, pool participation, indicative ranges, rating and a request-quote button. |
| P1-5 | **Policy-type select** | Changes the offers table (single trip / annual open / hull). |
| P1-6 | **Mobile and accessibility pass** | Works at 400px width, keyboard focus states, `prefers-reduced-motion` stops the radar animation, chart alt text. |
| P1-7 | **Total-cost view** | Per route: freight (mock) + war-risk premium + delay cost per extra day, to show the "total cost, not freight alone" argument. |

### P2: toward production

| ID | Feature | Notes |
|---|---|---|
| P2-1 | Split data from UI | Move `ROUTES` / `INSURERS` / `INS_CARDS` to JSON, and add a small data layer. |
| P2-2 | Real geographic map | A real map library with real coordinates and AIS feed integration (licensed provider at first, a national network later). |
| P2-3 | Risk-index engine | A scoring service that implements the methodology and supports versioning and independent review. |
| P2-4 | Insurer integration | APIs or periodic upload for rates and terms, under signed agreements. |
| P2-5 | Accounts and alerts backend | Authentication, saved projects, email/SMS/push alerts. |
| P2-6 | Governance and security | PDPL compliance, access control, audit log, penetration testing (implementation plan phase 3). |
| P2-7 | English UI toggle | For regional and international expansion. |

---

## 5. Related files in this repo

| Path | What it is |
|---|---|
| `presentation/prototype/mirsah-prototype.html` | The prototype (source of truth) |
| `presentation/Mirsah-Pitch.pptx` / `Mirsah-Pitch-preview.pdf` | The pitch deck (19 slides, Arabic) |
| `presentation/screenshots/` | Prototype screenshots used in the deck |
| `presentation/build-deck.js` | Script that generates the deck (pptxgenjs) |

## 6. Open items
- [ ] Team names and roles (the deck's team slide still says «الاسم»).
- [ ] Decide whether to keep the deck sentence noting the JWC expanded the Red Sea high-risk zone (Jul 2026). It's used only as context for the price rise.
- [ ] Confirm the illustrative index weights before building P0-5.

## 7. Sources (accessed Nov 2026)
- Reuters via *Insurance Journal*: "Saudi Oil Export Strategy Hits New Hurdle as Red Sea Insurance Costs Soar" (Sep 25, 2026).
- *Insurance Journal*: "London Insurers Expand Red Sea High-Risk Zone" (Jul 30, 2026).
- *The Saudi Gazette* via *Hellenic Shipping News*: "Saudi Arabia approves marine war-risk insurance pool" (Sep 2026).
- *Arab News* (40M TEU target); *Container News*: "Mawani reports 10.58% growth in container throughput in 2025".
- Insurance Authority: Q1 2026 market report (SAR 29.8B written premiums, +14.4%).
- GlobalData via *Arab News* (marine/aviation/transport/liability ≈ 4.5% of 2023 premiums).
- Goharbakhsh et al., *Sustainability* 18, 9534 (2026): Hormuz disruption, total-cost comparison.
- Klopott, M. (2024), *European Research Studies Journal* 27(S3): cargo insurance in conflict zones.
