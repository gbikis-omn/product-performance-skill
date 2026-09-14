---
name: product_performance
description: >
  Builds a full product/item-level performance report by joining GA4 e-commerce data
  (views, add-to-cart, purchases, revenue per item ID) with Google Ads spend per item ID
  and Meta Ads spend per product ID, then produces Brand Analysis, Category Analysis and
  Brand & Category Analysis breakdowns (with AOV, buy-to-detail rate, cart-to-detail rate,
  cart-to-purchase rate, % of spend, % of revenue) as an xlsx workbook. Use this whenever
  the user wants to see which products/brands/categories are getting ad spend but not
  converting, wants a cross-channel (GA4 + Google Ads + Meta) product performance report,
  asks to merge item-level GA4 data with ad spend by product ID, or wants a "product
  performance" / "brand & category analysis" / "spend vs revenue per product" report.
  Triggers on: product performance analysis, brand & category analysis, spend vs revenue
  per item, GA4 + Google Ads + Meta merge by item ID, which products get spend but no
  sales, product-level ROAS by category, item performance report.
---

# Product Performance Analysis

Ενοποιεί δεδομένα από GA4 (views/cart/purchases/revenue ανά item ID), Google Ads
(spend ανά item ID) και Meta Ads (spend ανά product ID), και παράγει ένα xlsx με
Brand Analysis / Category Analysis / Brand & Category Analysis — ίδιας ακριβώς
λογικής με ένα πραγματικό reference workbook που επαληθεύτηκε αναλυτικά (τύπος
προς τύπο, subtotal δομή, grand total) πριν γραφτεί αυτό το skill.

**Μη μαντεύεις πεδία/resources.** Το `references/mcp_pull_reference.md` έχει τα
ακριβή, επιβεβαιωμένα ονόματα για κάθε MCP call. Αν κάποιο call αποτύχει με
"field not found" ή αντίστοιχο, ξαναέλεγξε με το σχετικό metadata/search tool
πριν δοκιμάσεις εναλλακτικό όνομα — μην το μαντεύεις.

**Ο πυρήνας των υπολογισμών (aggregation, subtotals, %, AOV, rates) ζει σε
`scripts/build_report.py`, όχι εδώ.** Κάλεσέ το αφού έχεις τα 3(-4) raw CSVs —
μην ξαναγράφεις τη λογική του manually, έχει ήδη επαληθευτεί γραμμή-γραμμή
έναντι πραγματικού workbook.

---

## Βήμα 1 — Ζήτα τους λογαριασμούς

Ρώτα τον χρήστη (και τα τρία μαζί, σε ένα μήνυμα):
> "Για ποιο account θέλεις την ανάλυση; Χρειάζομαι:
> 1. GA4 Property ID
> 2. Google Ads Account ID (χωρίς παύλες)
> 3. Meta Ads Account ID (χωρίς το `act_`)"

Αν έχεις πρόσβαση σε λίστα λογαριασμών μέσω των MCP (π.χ. `meta_find_account`),
χρησιμοποίησέ τη για να επιβεβαιώσεις το σωστό account και δείξε το όνομά του
στον χρήστη πριν προχωρήσεις.

## Βήμα 2 — Ζήτα το χρονικό διάστημα

> "Για ποιο χρονικό διάστημα να τραβήξω τα δεδομένα;
> - Τελευταίες 7 μέρες
> - Τελευταίες 14 μέρες
> - Τελευταίες 30 μέρες
> - Από την 1η του τρέχοντος μήνα μέχρι χθες
> - Custom (δώσε μου ημερομηνίες)"

Μετέτρεψε την επιλογή σε **συγκεκριμένες ημερομηνίες** πριν συνεχίσεις (βλ.
πίνακα στο τέλος του `references/mcp_pull_reference.md`) — και χρησιμοποίησε
**ακριβώς το ίδιο διάστημα** και στα 3 platforms. Διαφορετικό διάστημα ανά
πηγή παράγει αναξιόπιστη σύγκριση spend/revenue.

## Βήμα 3 — Ζήτα το βάθος κατηγορίας και αν θέλει Brand Analysis

Ρώτα και τα δύο μαζί:
> "Δύο ακόμη πράγματα πριν τραβήξω δεδομένα:
> 1. Μέχρι ποιο επίπεδο item category να συμπεριλάβω στο 'Category' (1 έως 5 —
>    π.χ. 3 σημαίνει Item category > Item category 2 > Item category 3);
> 2. Θέλεις και ανάλυση ανά Brand (Brand Analysis + Brand & Category Analysis),
>    ή μόνο Category Analysis;"

## Βήμα 4 — Ρώτα για προαιρετικό feed field

> "Θέλεις να προσθέσω και κάποιο πεδίο από το product feed σου (π.χ. Gender,
> Custom Label, Product Type, Google Product Category); Αν ναι, πες μου ποιο
> πεδίο και πώς μπορώ να έχω πρόσβαση στο feed (Merchant Center account ID,
> ή link feed αρχείου)."

Αν ναι, θα προστεθεί σαν επιπλέον στήλη στο master table + ένα ακόμη
"`<Πεδίο>` Analysis" tab (ίδιας λογικής με Category Analysis) — αυτό γίνεται
πάντα, ανεξάρτητα από την επόμενη ερώτηση.

Ρώτα **και** αυτό, ξεχωριστά, μόνο αν ο χρήστης θέλει feed field:
> "Θέλεις αυτό το πεδίο (π.χ. Gender) να μπει ΚΑΙ μέσα στα υπόλοιπα tabs σαν
> επιπλέον επίπεδο ανάλυσης (Category Analysis γίνεται Category→[πεδίο],
> Brand Analysis γίνεται Brand→[πεδίο], Brand & Category Analysis γίνεται
> Brand→Category→[πεδίο]), ή προτιμάς μόνο το ξεχωριστό tab του;"

Αν **ναι** (nested): ρώτα ένα ακόμη πράγμα:
> "Σε αυτά τα tabs με 2+ επίπεδα, θέλεις να συμπεριλάβω AOV/Buy To Detail
> Rate/Cart To Detail Rate/Cart To Purchase Rate σε κάθε επίπεδο, ή να μείνει
> σαν το Brand & Category Analysis (μόνο τα 5 βασικά νούμερα + %Spend/%Revenue,
> χωρίς αυτά); Και οι δύο επιλογές είναι λογικές — απλά πες μου τι προτιμάς."

Και οι δύο απαντήσεις γίνονται flags στο script (Βήμα 6): `--nest-feed-field`
και `--include-rates-in-nested`. Default (αν δεν πει τίποτα ο χρήστης): ΧΩΡΙΣ
nesting, ΧΩΡΙΣ rates σε nested tabs — δηλαδή τα δύο flags παραλείπονται.

## Βήμα 5 — Τράβηξε τα δεδομένα

Ακολούθα ακριβώς το `references/mcp_pull_reference.md` για κάθε πηγή (GA4,
Google Ads, Meta, προαιρετικό feed). Σημεία προσοχής:

- Χρησιμοποίησε **το ίδιο resolved date range** και στα 3 calls (Βήμα 2).
- Το GA4 dimensions set πρέπει να έχει ακριβώς όσα category levels ζήτησε ο
  χρήστης στο Βήμα 3 (`itemCategory` έως `itemCategory<N>`), plus `itemId`,
  `itemBrand`.
- Αν κάποιο MCP δεν είναι συνδεδεμένο/authenticated, πες καθαρά στον χρήστη
  ποιο λείπει και σταμάτα — μην προχωρήσεις με μερικά δεδομένα σιωπηλά.
- Αν κάποιο query επιστρέψει 0 γραμμές, ενημέρωσε τον χρήστη αμέσως (μπορεί
  να σημαίνει ότι δεν τρέχουν Shopping/PMax καμπάνιες ή catalog ads σε αυτό
  το διάστημα) αντί να συνεχίσεις σαν να ήταν αναμενόμενο.
- Μετάτρεψε κάθε πηγή σε CSV με ακριβώς τα columns που περιμένει το script
  (βλ. reference file, ενότητα ανά πηγή) και αποθήκευσέ τα στο scratchpad
  directory.

## Βήμα 6 — Τρέξε το script

```bash
python3 ~/.claude/skills/product_performance/scripts/build_report.py \
  --ga4-csv <path>/ga4_items.csv \
  --google-ads-csv <path>/google_ads_spend.csv \
  --meta-csv <path>/meta_spend.csv \
  --category-levels <N απο Βήμα 3> \
  [--include-brand-analysis] \
  [--feed-csv <path>/feed.csv --feed-field-name "<όνομα πεδίου>"] \
  [--nest-feed-field] \
  [--include-rates-in-nested] \
  --output <path>/product_performance_<account>_<περίοδος>.xlsx
```

`--nest-feed-field` και `--include-rates-in-nested` περνάνε ΜΟΝΟ αν ο χρήστης
το ζήτησε ρητά στο Βήμα 4 — δεν είναι defaults.

Το script:
1. Ομαδοποιεί το GA4 pull ανά `item_id` (χειρίζεται σωστά περιπτώσεις όπου το
   ίδιο item_id εμφανίστηκε με 2+ διαφορετικά brand/category labels μέσα στο
   διάστημα — βλ. σχόλια στο script γιατί αυτό είναι σημαντικό, βρέθηκε σαν
   πραγματικό bug pattern κατά την επαλήθευση).
2. Κάνει LEFT JOIN spend πάνω στο GA4 σύνολο item IDs (fillna 0) — δηλαδή
   προϊόντα με spend αλλά καθόλου GA4 δραστηριότητα **δεν** εμφανίζονται στο
   report. Αυτό είναι σκόπιμο (ίδιο design με το reference workbook), αλλά
   **πρέπει να το αναφέρεις στον χρήστη** μαζί με το πόσα τέτοια item IDs
   βρέθηκαν (το script τα μετράει και τα τυπώνει).
3. Χτίζει Category Analysis, (προαιρετικά) Brand Analysis, (προαιρετικά)
   Brand & Category Analysis, (προαιρετικά) `<feed field>` Analysis, + raw
   tabs (Meta Data, Google Ads Data, Feed Data, Master Data) για audit.
4. **Κάνει reconciliation self-check** μετά από κάθε tab (πχ. ότι το άθροισμα
   revenue στο Category Analysis == grand total == Master Data) και **σταματά
   με σφάλμα** αν κάτι δεν βγαίνει ακριβές. Αν το script πετάξει
   `AssertionError: RECONCILIATION FAILED`, **μην το προσπεράσεις και μην
   στείλεις το αρχείο στον χρήστη** — σταμάτα, διάβασε το μήνυμα, εντόπισε
   το πρόβλημα στα input CSVs (συνηθέστερη αιτία: λάθος/ασυνεπές date range
   μεταξύ πηγών, ή προϊόν με λάθος τύπο item_id — string vs number — που
   έσπασε το merge).

## Βήμα 7 — Επιβεβαίωσε πριν παραδώσεις

Πριν πεις στον χρήστη ότι τελείωσες:

1. Διάβασε το stdout του script — grand totals (views/cart/purchases/revenue/
   spend), αριθμό conflicts (item_id με πολλαπλά labels), αριθμό unmatched
   spend items, και το `case-folded N id value(s)` NOTE (αν N Google Ads
   ή Meta ids φαίνεται κοντά στο σύνολο των ids εκείνης της πηγής, καλό
   σημάδι — σημαίνει ότι το case-fold δούλεψε στα σωστά ids· βλ. §"CRITICAL
   bug" πιο κάτω για το γιατί αυτός ο έλεγχος υπάρχει).
   **Πάντα**, ανεξάρτητα από το αν κάτι φαίνεται ύποπτο, έλεγξε το unmatched
   Google Ads/Meta ποσό σαν % του αντίστοιχου raw pull total (Βήμα 5) — αν
   είναι πάνω από ~20-30%, κάνε τον έλεγχο set-intersection πριν/μετά
   uppercase (βλ. μεθοδολογία στο §"CRITICAL bug" πιο κάτω) πριν πεις στον
   χρήστη ότι το spend "απλά δεν είχε GA4 δραστηριότητα" — μπορεί να είναι
   ξανά ID mismatch, όχι απουσία δραστηριότητας.
2. Άνοιξε το xlsx (π.χ. με openpyxl) και έλεγξε ότι:
   - Το Grand Total row σε κάθε analysis tab έχει τα ίδια views/cart/purchases/
     revenue/spend σε ΟΛΑ τα tabs (πρέπει να ταυτίζονται — αν το script
     πέρασε τα reconciliation checks του, θα ταυτίζονται ήδη, αλλά κάνε ένα
     οπτικό sanity check).
   - Τα top 3-5 brands/categories βγάζουν νόημα (όχι π.χ. αρνητικά νούμερα,
     όχι AOV εξωφρενικά υψηλό λόγω λάθος currency conversion).
3. Ανέφερε στον χρήστη με απλά λόγια:
   - Πόσα items μπήκαν στην ανάλυση, grand totals.
   - **Πόσα item IDs είχαν spend αλλά καθόλου GA4 δραστηριότητα** (και άρα
     εξαιρέθηκαν) — αυτό είναι σημαντικό εύρημα από μόνο του, όχι απλά ένα
     τεχνικό detail.
   - Αν βρέθηκαν item IDs με ασυνεπή brand/category labels μέσα στο διάστημα.
   - Το path του τελικού xlsx.

## Βασικές φόρμουλες (για αναφορά — ήδη υλοποιημένες στο script)

**Category Analysis / Brand Analysis** (μία γραμμή ανά κατηγορία/brand):
- `% Spend = row Total Spend / GRAND TOTAL Total Spend`
- `% Revenue = row Item Revenue / GRAND TOTAL Item Revenue`
- `AOV = Item Revenue / Items Purchased` (κενό αν purchases = 0)
- `Buy To Detail Rate = Items Purchased / Items Viewed`
- `Cart To Detail Rate = Items Added To Cart / Items Viewed`
- `Cart To Purchase Rate = Items Purchased / Items Added To Cart` (κενό αν cart = 0)
- Ταξινόμηση: φθίνουσα κατά Item Revenue. Τελευταία γραμμή: Grand Total.

**Brand & Category Analysis** (2 επίπεδα: Brand έξω, Category μέσα σε κάθε brand),
**και γενικά κάθε N-επίπεδο tab** (π.χ. αν προστεθεί nested feed field):
- Μόνο 5 metric columns + `% Spend` + `% Revenue` ανά default — AOV/rates
  προαιρετικά μόνο αν το ζήτησε ρητά ο χρήστης (`--include-rates-in-nested`,
  βλ. Βήμα 4). Tabs με 1 μόνο επίπεδο (Category Analysis, Brand Analysis,
  standalone `<feed field>` Analysis) έχουν ΠΑΝΤΑ AOV/rates, ανεξάρτητα.
- `% Spend` / `% Revenue` κάθε γραμμής (leaf ή subtotal) = % επί του **subtotal
  του ΑΜΕΣΟΥ γονέα της**, ΟΧΙ επί του grand total — επιβεβαιωμένο έναντι live
  formula στο reference workbook (2 επίπεδα) και γενικευμένο/επαληθευμένο
  χειροκίνητα σε 3 επίπεδα (leaf rows αθροίζουν στο 100% του γονέα τους).
- **Κάθε subtotal γραμμή (σε ΚΑΘΕ επίπεδο, όχι μόνο το εξωτερικό) δείχνει
  πάντα 100% (`% Spend`=`% Revenue`=1.0), self-referential** — είναι το
  σύνολο του ίδιου της του group, όχι ποσοστό επί κάποιου γονέα.
- Κάθε group σε κάθε επίπεδο (εκτός του πιο εσωτερικού) κλείνει με μια
  `"<value> Total"` subtotal γραμμή· το κείμενο μπαίνει στην ΕΠΟΜΕΝΗ στήλη
  (πιο εσωτερικό dimension), όχι στη δική του — η δική της στήλη μένει κενή,
  οι πιο εξωτερικές στήλες κρατάνε τις πραγματικές τιμές τους.
- Τελευταία γραμμή όλου του πίνακα: `Grand Total` (ίδια σύμβαση θέσης
  κειμένου με τα subtotals).

## Γνωστά όρια / πράγματα που πρέπει να πεις στον χρήστη αν συμβούν

- Ένα item_id με spend σε Meta ή Google Ads αλλά μηδέν GA4 views/cart/
  purchases δεν εμφανίζεται πουθενά στο report (by design — matching το
  reference workbook). Αν ο χρήστης θέλει να δει ΚΑΙ αυτά, χρειάζεται
  διαφορετικό base join (right/full join), πες του το ρητά και ρώτα αν
  το θέλει πριν αλλάξεις προσέγγιση.
- Το Meta `product_id` breakdown ταιριάζει με το GA4/Google Ads `item_id`
  μόνο αν το Meta catalog feed και το Merchant Center feed του πελάτη
  χρησιμοποιούν το ίδιο ID scheme. Αν μετά το merge βλέπεις πολύ χαμηλό
  match rate Meta spend → GA4 items, ανέφερέ το — πιθανό ID mismatch, όχι
  απαραίτητα απλά "δεν υπάρχει spend".
- Ζωντανά sheets/λογαριασμοί αλλάζουν σε real time — αν ξανατρέξεις το ίδιο
  pull λίγα λεπτά αργότερα και τα νούμερα διαφέρουν ελαφρώς, δεν είναι bug.
- **Case-sensitivity στο item_id join (διορθωμένο 2026-09-14, βλ. live test
  παρακάτω)**: `normalize_id_series()` πλέον κάνει uppercase ΚΑΘΕ id πριν το
  merge, γιατί βρέθηκε live ότι το Google Ads `segments.product_item_id`
  μπορεί να επιστρέφει lowercase ids ενώ το GA4 item_id/feed id για το ΙΔΙΟ
  προϊόν είναι uppercase — χωρίς αυτό το fix, το merge αποτυγχάνει 100%
  σιωπηλά (καμία exception, απλά "0 spend" και λάθος NOTE μήνυμα ότι το
  item "δεν είχε GA4 δραστηριότητα"). Αν δεις στο stdout πολλά
  `case-folded N id value(s)`, είναι αναμενόμενο/θετικό σημάδι ότι το fix
  δούλεψε — όχι κάτι για ανησυχία.

## Τι έχει πραγματικά δοκιμαστεί live (2026-09-11, Kanellopoulos)

- **Live, με πραγματικό pull, χωρίς feed**: `--category-levels 2
  --include-brand-analysis` (Category Analysis + Brand Analysis + Brand &
  Category Analysis) — 100% reconciled, cell-by-cell verified.
- **Live, με πραγματικό feed pull (raw XML feed URL, όχι mc-mcp)**:
  `--feed-csv --feed-field-name gender --nest-feed-field
  --include-rates-in-nested` πάνω σε πραγματικό Merchant Center-style feed
  (`<g:id>`/`<g:gender>`, 7.791 προϊόντα, 84% match rate με τα GA4 items).
  Επιβεβαιώθηκε: Category→gender, Brand→gender (2-level) και
  Brand→Category→gender (3-level, πραγματικό βάθος 3) όλα σωστά δομημένα
  (leaf rows αθροίζουν στο 100% του άμεσου γονέα τους, κάθε subtotal σε κάθε
  βάθος 100% self-referential, "<value> Total" στη σωστή στήλη), standalone
  `gender Analysis` tab σωστό, όλα τα Grand Total rows ταυτίζονται σε όλα τα
  tabs, όλα τα (ενισχυμένα, leaf-sum-based) reconciliation checks πέρασαν.
  **Σημαντικό λάθος που έγινε ΚΑΤΑ το live test (όχι bug του script)**: το
  Meta export CSV δόθηκε στο script στην raw μορφή του
  (`spend,date_start,date_stop,product_id`) αντί να μετατραπεί πρώτα σε
  `item_id,spend` — έσκασε με `KeyError: 'item_id'` αμέσως, καθαρό error,
  όχι σιωπηλό πρόβλημα. Θυμήσου πάντα να μετατρέπεις κάθε πηγή στη μορφή
  που περιμένει το script (Βήμα 5) πριν το τρέξιμο.
- **Δεν έχει δοκιμαστεί ακόμα live**: το "μόνο Category Analysis, χωρίς
  Brand" path (Βήμα 3, δεύτερη επιλογή), και το feed field μέσω `mc-mcp`
  (χρησιμοποιήθηκε raw feed URL και στα δύο live tests μέχρι τώρα).

## CRITICAL bug βρέθηκε + διορθώθηκε live (2026-09-14, ProteinMax.gr)

- **Τι συνέβη**: 2 reports παραδόθηκαν στον χρήστη (7ήμερο 2026-09-07/13 και
  11ήμερο 2026-08-17/27, `--category-levels 2 --include-brand-analysis`,
  χωρίς feed) με το **100% του Google Ads spend σιωπηλά εξαφανισμένο** από
  Master Data / Category Analysis / Brand Analysis / Grand Total. Ο χρήστης
  ζήτησε ρητά αναλυτικό έλεγχο ("είσαι 100% σίγουρος;") — χωρίς αυτή την
  ερώτηση το bug θα περνούσε απαρατήρητο, γιατί το script δεν πετάει error.
- **Root cause**: Google Ads `shopping_performance_view.segments.product_item_id`
  επέστρεψε lowercase ids (π.χ. `im4966`) ενώ το GA4 item_id/Merchant Center
  feed για το ΙΔΙΟ προϊόν είναι uppercase (`IM4966`). Το `normalize_id_series()`
  ΔΕΝ έκανε case folding πριν το fix, το pandas `merge(on="item_id")` είναι
  case-sensitive → **0 exact matches / 1.476-1.155 Google Ads item_ids και
  στα δύο reports**, ενώ 94-97% θα ταίριαζαν uppercased. Το script μάλιστα
  τύπωνε ένα εντελώς λάθος NOTE ("αυτά τα items δεν είχαν GA4 δραστηριότητα")
  — η πραγματική αιτία ήταν αποκλειστικά case mismatch, όχι έλλειψη δραστηριότητας.
- **Πώς εντοπίστηκε (μεθοδολογία επαλήθευσης — επανέλαβέ την όποτε ο χρήστης
  ζητήσει έλεγχο)**: (1) sum raw pull == sum converted CSV, ανά πηγή — ΟΚ και
  στις δύο πλευρές. (2) `set(ga4 item_ids) & set(google_ads item_ids)` →
  0 exact matches, αλλά 94-97% matches αν γίνει `.upper()` σε ένα από τα δύο
  σύνολα — αυτό απέδειξε το case bug άμεσα, όχι απλά "λίγο spend λείπει".
  (3) Ξαναϋπολογίστηκε χειροκίνητα το Total Spend ανά brand από το Master
  Data tab και συγκρίθηκε με το Brand Analysis tab — 0 mismatches σε 121
  brands, που απέδειξε ότι η ΑΘΡΟΙΣΗ ήταν πάντα σωστή, το πρόβλημα ήταν
  αποκλειστικά στο ποια δεδομένα μπήκαν στο merge.
- **Fix**: `normalize_id_series()` κάνει τώρα `.str.upper()` σε ΚΑΘΕ id από
  ΚΑΘΕ πηγή (GA4, Google Ads, Meta, feed) πριν οποιοδήποτε merge — μόνιμο
  fix στο ίδιο το script, όχι μια φορά μόνο. Επαληθεύτηκε live στο ίδιο
  11ήμερο dataset: unmatched Google Ads έπεσε από 1.476 items (€3.540,55)
  σε 42 items (€13,40) — δηλαδή 97% matched σωστά μετά το fix. Grand Total
  spend άλλαξε από €2.137,44 (μόνο Meta, λάθος) σε €5.664,60 (Meta + Google
  Ads, σωστό). Brand-level recompute check ξανατρέχτηκε μετά το fix: 121/121
  brands exact match, 0 mismatches.
- **Συμπέρασμα για μελλοντικά reports**: πάντα, μετά από κάθε report, κάνε
  τον έλεγχο #2 παραπάνω (set intersection πριν/μετά uppercase) σαν μέρος
  του Βήματος 7 sanity check — ειδικά αν δεις "unmatched Google Ads" ή
  "unmatched Meta" ποσά που φαίνονται μεγάλα σε σχέση με το raw pull total.
