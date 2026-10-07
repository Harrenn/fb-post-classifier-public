# User Guide: How to Read the Facebook Post Classifier Output

This guide explains the output in practical terms: what each sheet means, how
the sheets connect, and what the classifier is expected to do in common
“what if ganito?” situations.

For the formal definition of every column, use `data_dictionary.md`. This guide
focuses on interpretation and examples.

## 1. What the classifier is trying to answer

For every Facebook post, the classifier tries to answer four separate business
questions:

1. What is the post mainly about?
2. Is there a real financial promotion?
3. Is there motorcycle pricing that can be connected to a particular unit?
4. Is there a free item, service, or bundled benefit?

It reads the Facebook caption and every non-empty OCR column together. A detail
can come from either source:

- The caption may contain promo dates, eligibility, participating branches, or
  instructions to message the page.
- The images may contain motorcycle models, downpayments, installment amounts,
  discounts, and freebies.

The script does not require all information to appear in one image. It does,
however, avoid connecting facts across images when the relationship is unclear.

## 2. The four-sheet mental model

One source post can produce several related output rows:

```text
Posts: one row for the whole Facebook post
  |
  +-- Offers: zero or more campaign rows
  |     |
  |     +-- PricePoints: zero or more pricing rows
  |
  +-- Addons: zero or more free-benefit rows
```

The sheets answer different questions:

| Sheet | Row meaning | Main question |
|---|---|---|
| `Posts` | One Facebook post | What is the post generally about? |
| `Offers` | One campaign and eligibility scope | What is the deal, and who or what qualifies? |
| `PricePoints` | One safely aligned amount | What amount was advertised, and what does it belong to? |
| `Addons` | One free item or benefit | What free benefit is being offered, and under what conditions? |

Do not expect the four sheets to have the same number of rows. For example, one
post can produce one `Posts` row, one `Offers` row, four `PricePoints` rows, and
two `Addons` rows.

## 3. How to inspect one post

Use this order when reviewing a classified post:

1. Find the Facebook permalink in `Posts.url`.
2. Read `post_detail` for the short summary.
3. Check `primary_post_type` and `content_themes`.
4. Check the three commercial flags:
   - `has_financial_promo`
   - `has_advertised_pricing`
   - `has_addon`
5. Copy the `post_id`.
6. Filter `Offers.post_id`, `PricePoints.post_id`, and `Addons.post_id` using
   that value.
7. If something looks surprising, read `llm_context`. It contains the exact
   labeled caption and OCR context supplied to the model.
8. Check `processing_error`. A non-empty value means the row experienced an
   operational failure and should not be treated as a successful classification.

The `url` is repeated in the detail sheets so analysts can open the source post
without performing a join first.

## 4. Reading the `Posts` sheet

### `primary_post_type`

This is the single dominant purpose of the post.

| Value | Plain-language meaning |
|---|---|
| `commercial_offer` | At least one validated financial promo, advertised motorcycle price, or add-on exists. |
| `product_showcase` | A motorcycle is highlighted, but no concrete deal, aligned price, or free benefit is established. |
| `brand_lifestyle` | Brand awareness, lifestyle content, greetings, or generic branch marketing. |
| `event_community` | A ride, event, celebration, or community activity. |
| `customer_service_advisory` | Registration, plate, process, service, or customer-support information. |
| `corporate_employer` | Hiring, employee recognition, incentive trips, or employer content. |
| `other` | Useful content that does not fit the defined categories. |
| `unclear` | There is not enough usable caption/OCR text to understand the post. |

`primary_post_type` is single-label, but `content_themes` is multi-label. A post
can therefore be a `commercial_offer` with themes such as
`financing; pricing; product_showcase`.

### The three commercial flags

The flags are independent. They are not competing categories.

| Flag | `yes` means |
|---|---|
| `has_financial_promo` | A special financial mechanic exists, such as piso or zero downpayment, reduced downpayment, discount, rebate, special financing, trade-in, or reservation discount. |
| `has_advertised_pricing` | An identifiable motorcycle is safely connected to a price or financing amount. |
| `has_addon` | At least one free item, service, accessory, subscription, or bundled benefit exists. |

All combinations are allowed. `offer_bundle_type` simply summarizes the three
flags:

| Promo | Pricing | Add-on | `offer_bundle_type` |
|---|---|---|---|
| no | no | no | `no_offer` |
| yes | no | no | `promo` |
| no | yes | no | `pricing` |
| no | no | yes | `addon` |
| yes | yes | no | `promo_and_pricing` |
| yes | no | yes | `promo_and_addon` |
| no | yes | yes | `pricing_and_addon` |
| yes | yes | yes | `promo_and_pricing_and_addon` |

Important: `has_financial_promo = no` does not mean the post is unimportant. It
may still contain ordinary pricing, a product showcase, an event, an advisory,
or a free add-on.

## 5. Reading the `Offers` sheet

An offer row describes a campaign and its eligibility scope. It does not try to
store every price in the same row.

### `promo_mechanics`

This tells you why the deal qualifies as a financial promotion. Multiple values
are separated with `; `.

Examples:

- `piso_downpayment`
- `zero_downpayment`
- `reduced_downpayment`
- `cash_discount`
- `rebate`
- `reduced_monthly_payment`
- `special_financing`
- `trade_in`
- `reservation_discount`

A normal cash price or ordinary downpayment is not automatically a promo
mechanic.

### `offer_scope`

This field prevents pictured motorcycles from being incorrectly treated as the
complete eligibility list.

| Scope | Meaning |
|---|---|
| `specific_model` | Exactly one model is explicitly stated as eligible. |
| `specific_models` | Several explicitly named models are grouped under the same campaign. |
| `participating_models` | The post says selected, participating, qualified, or eligible models, but the complete list is not given. |
| `brand_wide` | The post explicitly says all or any models of one brand qualify. |
| `all_models` | The post explicitly says all motorcycles or all models qualify. |
| `unspecified` | The eligibility scope cannot be established safely. |

Use these columns together:

- `eligible_models`: models explicitly confirmed as eligible.
- `featured_models`: models shown or mentioned as examples without proof that
  they are the complete eligible set.
- `scope_details`: why the classifier selected that scope.
- `eligibility`: what the customer must do or satisfy.
- `promo_details`: fuller explanation of the mechanics, comparisons, dates,
  branch restrictions, and important conditions.

## 6. Reading the `PricePoints` sheet

Each row should represent one amount or one coherent set of related amounts.
This is how the output avoids mismatching models and prices.

### `pricing_role`

| Role | Meaning |
|---|---|
| `eligible_model` | The exact motorcycle is explicitly eligible for the campaign, and the amount belongs to it. |
| `featured_example` | The model and amount are aligned in the creative, but the model is only a featured example—not proven to define campaign eligibility. |
| `campaign_term` | The amount is shared by the campaign; `unit_name` is normally blank. |
| `unassigned` | The amount is visibly motorcycle-related, but no safe model or exact relationship can be established. |

### Monetary fields are not interchangeable

| Field | Use it for |
|---|---|
| `previous_downpayment` | Earlier downpayment in a comparison. |
| `current_downpayment` | Current/final downpayment, including piso = `1` and zero = `0`. |
| `cash_price` | Explicit full cash price. |
| `previous_price` | Previous full motorcycle price. |
| `current_price` | Current/sale full motorcycle price. |
| `observed_price` | Price-like motorcycle amount whose exact role is unclear. |
| `payment_amount` | Repeating installment amount; read with `payment_frequency`. |
| `reservation_amount` | Explicit reservation payment. |
| `discount_amount` | Explicit discount amount. |
| `rebate_amount` | Explicit rebate amount. |

Blank does not mean zero. Zero is stored only when the source explicitly says
zero, such as zero downpayment.

## 7. Reading the `Addons` sheet

Each distinct free benefit receives its own row. Standardized `addon_type`
values include:

- `helmet`
- `riding_gear`
- `registration`
- `insurance`
- `maintenance`
- `delivery`
- `transfer_assistance`
- `accessory`
- `subscription`
- `service`
- `other`

Read `addon_name` for the source-grounded wording and `addon_details` for the
explanation. `linked_offer_id` is filled only when the benefit can be safely
connected to a particular offer. It stays blank rather than guessing.

## 8. Scenario guide: “What if ganito?”

### Scenario 1: Generic branch or brand marketing

Source meaning:

> Visit our branch and discover your dream ride today!

No model, price, special deal, or free benefit is stated.

Expected result:

```text
Posts.primary_post_type = brand_lifestyle
Posts.has_financial_promo = no
Posts.has_advertised_pricing = no
Posts.has_addon = no
Posts.offer_bundle_type = no_offer
Offers rows = 0
PricePoints rows = 0
Addons rows = 0
```

The classifier does not set a commercial flag merely because the post says
“visit our branch.”

### Scenario 2: Motorcycle showcase without pricing

Source meaning:

> Meet the new Honda PCX160. Stylish, comfortable, and ready for your next ride.

Expected result:

```text
primary_post_type = product_showcase
content_themes includes product_showcase
has_financial_promo = no
has_advertised_pricing = no
has_addon = no
```

A motorcycle name by itself does not create advertised pricing.

### Scenario 3: Identifiable motorcycle with a normal cash price

Source meaning:

> Honda Click 125 cash price ₱80,000.

Expected result:

```text
has_financial_promo = no
has_advertised_pricing = yes
has_addon = no
offer_bundle_type = pricing

PricePoints.brand = Honda
PricePoints.unit_name = Click 125
PricePoints.cash_price = 80000
```

A normal price is pricing, not automatically a promo.

### Scenario 4: Normal downpayment without special wording

Source meaning:

> Yamaha Mio Gear, downpayment ₱8,000.

Expected result:

```text
has_financial_promo = no
has_advertised_pricing = yes
current_downpayment = 8000
```

An ordinary downpayment becomes a price point, but it is not a financial promo
unless the source presents it as discounted, reduced, special, zero, piso, or a
limited-time deal.

### Scenario 5: Piso downpayment for a named motorcycle

Source meaning:

> Get the Honda PCX160 for piso downpayment.

Expected result:

```text
has_financial_promo = yes
has_advertised_pricing = yes
promo_mechanics = piso_downpayment
offer_scope = specific_model
eligible_models = PCX160

PricePoints.unit_name = PCX160
PricePoints.current_downpayment = 1
PricePoints.pricing_role = eligible_model
```

“Piso downpayment” is stored numerically as `1`, not as the words “piso deal.”

### Scenario 6: Before/now downpayment

Source meaning:

> BEFORE ₱8,000, NOW ₱4,000 downpayment for Honda Click 125.

Expected result:

```text
promo_mechanics = reduced_downpayment
previous_downpayment = 8000
current_downpayment = 4000
```

The usable current downpayment is `4000`. The complete comparison remains in
`price_details` and `promo_details`, for example:

```text
Downpayment was reduced from ₱8,000 to ₱4,000 for Honda Click 125.
```

The values are not placed in `previous_price` or `current_price` because the
source specifically labels them as downpayment.

### Scenario 7: Zero downpayment plus free helmet

Source meaning:

> Zero downpayment plus free helmet for qualified buyers.

Expected result:

```text
has_financial_promo = yes
has_addon = yes
promo_mechanics = zero_downpayment
current_downpayment = 0
addon_type = helmet
offer_bundle_type = promo_and_addon
```

If a safely aligned motorcycle price is also present, then
`has_advertised_pricing = yes` and the bundle becomes
`promo_and_pricing_and_addon`.

### Scenario 8: Free helmet and free registration only

Source meaning:

> Buy this month and get a free helmet and free registration.

Expected result:

```text
has_financial_promo = no
has_addon = yes
Addons rows = 2
```

Example `Addons` rows:

| addon_type | addon_name |
|---|---|
| `helmet` | `Free helmet` |
| `registration` | `Free registration` |

Freebies are not financial promos by default.

The benefit must still be identifiable and guaranteed. Wording such as
“exciting freebies,” “exclusive freebies,” or “raffle prizes” without a named
guaranteed item creates no `Addons` row. “Same-day release” is fulfillment
speed, not a free service add-on.

### Scenario 9: Two motorcycles are pictured, but the campaign says participating models

Source meaning:

> Get a Kymco motorcycle for piso downpayment. Message us for participating models.

The image shows Agility ECD and Skytown 150Fi.

Expected result:

```text
Offers.offer_scope = participating_models
Offers.eligible_models = blank
Offers.featured_models = Agility ECD; Skytown 150Fi

PricePoints.pricing_role = campaign_term
PricePoints.unit_name = blank
PricePoints.current_downpayment = 1
```

The pictured motorcycles are treated as creative examples. The classifier does
not claim that the promotion is limited to those two models.

### Scenario 10: The post explicitly says the offer is only for PCX160

Source meaning:

> Promo applicable only to Honda PCX160.

Expected result:

```text
offer_scope = specific_model
eligible_models = PCX160
featured_models may be blank
pricing_role = eligible_model
unit_name = PCX160
```

The word “only” makes the model restriction explicit.

### Scenario 11: Four motorcycles with four separate prices

Suppose one post clearly shows:

```text
Honda Click 125 — ₱80,000
Yamaha Mio Gear — ₱79,000
Suzuki Burgman Street — ₱92,000
Kymco Like 150i — ₱135,000
```

Expected result:

```text
Posts rows = 1
PricePoints rows = 4
has_advertised_pricing = yes
```

Each `PricePoints` row contains exactly one model-price relationship:

| unit_name | current_price or cash_price |
|---|---:|
| Click 125 | 80000 |
| Mio Gear | 79000 |
| Burgman Street | 92000 |
| Like 150i | 135000 |

If the post also says “ask for participating models,” the four rows can use
`pricing_role = featured_example`. Their visible model-price relationships are
still useful, but they do not define the complete promo eligibility list.

The same caution applies to a broader multi-image campaign whose caption does
not confirm eligibility: an image-only model-price pair remains a
`featured_example`, and the campaign scope can stay `unspecified`. A trim such
as `ABS/TCS` remains part of one model name rather than becoming two models.

### Scenario 12: Four motorcycles are pictured, but only one shared piso amount exists

Source meaning:

> Piso downpayment for participating models.

Four motorcycles are shown, but no model-specific amount is printed.

Expected result:

```text
PricePoints rows for the shared piso term = 1
pricing_role = campaign_term
unit_name = blank
current_downpayment = 1
```

The script does not repeat `1` four times and does not randomly assign it to a
pictured motorcycle.

### Scenario 13: A readable price-like amount has no clear label

Source meaning:

> Repo motorcycle — ₱64,352

The amount is clearly commercial and motorcycle-related, but the source does
not say whether it is a cash price, outstanding balance, or another price type.

Expected result:

```text
pricing_role = unassigned
observed_price = 64352
cash_price = blank
current_price = blank
current_downpayment = blank
```

The number is preserved without inventing its meaning.

### Scenario 14: A peso amount is unrelated to an identifiable motorcycle

Source meaning:

> Event registration fee ₱500. Join our weekend ride!

Expected result:

```text
has_advertised_pricing = no
```

The amount belongs to the event, not to a motorcycle purchase.

### Scenario 15: Phone numbers, dates, displacement, and mileage

Source text may contain:

```text
0917-123-4567
July 31, 2026
160cc
12,000 km
```

Expected behavior:

- None of these values becomes motorcycle pricing.
- `160` is not stored as `payment_amount` or `current_price` merely because it
  appears beside a model.
- `12000` is not treated as downpayment when it is labeled as mileage.

### Scenario 16: Caption and OCR contain complementary details

Caption:

> Promo available at participating branches until July 31. Message us to apply.

OCR:

```text
Honda Click 125
Reduced DP: ₱3,000
Free helmet
```

Expected behavior:

- The offer date, branch limitation, and call to action come from the caption.
- The model, downpayment, and helmet come from OCR.
- The result can contain promo, pricing, and add-on information at the same time.

Neither the caption nor OCR is treated as secondary.

### Scenario 17: Caption and OCR conflict

Caption:

> As low as ₱150 daily.

OCR:

> As low as ₱156 daily.

If the conflict cannot be resolved safely:

```text
payment_amount = blank
price_details explains that the source contains conflicting daily amounts
```

The classifier keeps the useful context but avoids choosing one amount without
support.

Even without a conflict, a heading such as “Daily Ride Deals” plus an ordinary
“as low as ₱100 daily” rate is not `special_financing` by itself. A separate
special or reduced financing mechanic must be stated before
`has_financial_promo` becomes `yes`.

### Scenario 18: OCR typo such as `PI50`

If the surrounding panel clearly says “daily payment” and the OCR reads
`PI50`, the script may normalize the common OCR confusion to `P150` and store:

```text
payment_amount = 150
payment_frequency = daily
```

It does not apply this correction to an isolated code or identifier without
clear payment context.

### Scenario 19: “All models” versus “participating models”

These are intentionally different:

| Source wording | Expected scope |
|---|---|
| `All Honda models qualify` | `brand_wide` |
| `Any motorcycle model qualifies` | `all_models` |
| `Available on participating models` | `participating_models` |
| `Honda motorcycles available now` | Usually `unspecified`; plural wording alone does not prove all models qualify |

Universal scope requires explicit wording such as “all,” “any,” “every,”
“lahat,” or “anumang.”

### Scenario 20: Several explicitly eligible models

Source meaning:

> Promo valid for Honda Click 125, Beat, and PCX160 only.

Expected result:

```text
offer_scope = specific_models
eligible_models = Click 125; Beat; PCX160
```

Model-specific prices are retained as separate `PricePoints` rows so the
amounts stay aligned.

### Scenario 21: A generic inventory category has a price

Source meaning:

> Repo motorcycles from ₱30,000.

“Repo motorcycles” is an inventory category, not an identifiable model.

Expected behavior:

- The post can still be summarized as repo/pricing content.
- The amount may be retained as a campaign-level or unassigned observation when
  it is clearly commercial.
- `unit_name` stays blank.
- `has_advertised_pricing` does not become `yes` solely from a generic category
  and amount because no identifiable motorcycle-price pair exists.

### Scenario 22: One post contains several promos

Source meaning:

> Piso downpayment, ₱500 rebate, and cash discount for participating models.

Expected result:

```text
promo_mechanics = piso_downpayment; rebate; cash_discount
has_financial_promo = yes
```

The comprehensive terms remain in `promo_details`. Separate price-point rows
may be used when the monetary mechanics cannot be safely combined into one
model-specific observation.

When the same anonymous zero-downpayment, rebate, or discount mechanics are
repeated across creative panels, identical compatible terms consolidate into
one `campaign_term`. A distinct model-aligned before/now price panel remains a
separate `featured_example`.

### Scenario 23: Add-on applies to one model only

Source meaning:

> Free helmet with every PCX160 purchase.

Expected `Addons` row:

```text
unit_name = PCX160
addon_type = helmet
addon_name = Free helmet
```

If the source only says “free helmet with every purchase,” `unit_name` stays
blank because the benefit is not model-specific.

### Scenario 24: Add-on relationship is unclear in a multi-offer post

If a post contains two campaigns and one free registration mention, but the
source does not say which campaign receives it:

```text
linked_offer_id = blank
```

The add-on remains in the data because it is real; only the uncertain campaign
link is omitted.

### Scenario 25: Customer-service advisory

Source meaning:

> Plate release schedule and registration claiming procedure.

Expected result:

```text
primary_post_type = customer_service_advisory
content_themes may include registration; customer_support
```

It is not forced into generic marketing just because no commercial flags are
present.

### Scenario 26: Hiring or employee-recognition post

Expected result:

```text
primary_post_type = corporate_employer
has_financial_promo = no
has_advertised_pricing = no
has_addon = no
```

### Scenario 27: Caption is empty but OCR exists

Expected behavior:

- The model receives the available OCR sections.
- Classification proceeds normally.
- `llm_context` contains only the non-empty image sections.

### Scenario 28: OCR is empty but the caption exists

Expected behavior:

- The model receives the caption.
- Classification proceeds normally.
- Empty OCR values are not included in `llm_context`.

### Scenario 29: Caption and all OCR fields are empty

Expected result:

```text
Gemini API request = skipped
primary_post_type = unclear
has_financial_promo = no
has_advertised_pricing = no
has_addon = no
post_detail = No usable text was available.
Offers rows = 0
PricePoints rows = 0
Addons rows = 0
llm_input_tokens = 0
estimated_api_cost_usd = 0
```

### Scenario 30: One API row fails

Expected behavior:

- The script records a sanitized message in `processing_error`.
- The API key is removed from the error text.
- Other rows and other CSV files continue processing.
- A persistent authentication, quota, or model-access failure can stop the run
  because continuing would cause the remaining rows to fail for the same reason.

## 9. Example of one post across all sheets

Source meaning:

> Honda PCX160 promo: downpayment reduced from ₱8,000 to ₱4,000, plus free helmet for approved buyers until July 31.

Possible simplified output:

### `Posts`

| post_id | primary_post_type | has_financial_promo | has_advertised_pricing | has_addon | offer_bundle_type |
|---|---|---|---|---|---|
| `POST123` | `commercial_offer` | `yes` | `yes` | `yes` | `promo_and_pricing_and_addon` |

### `Offers`

| offer_id | post_id | promo_mechanics | offer_scope | eligible_models | eligibility | valid_until |
|---|---|---|---|---|---|---|
| `POST123-O01` | `POST123` | `reduced_downpayment` | `specific_model` | `PCX160` | `Approved buyers` | `2026-07-31` |

### `PricePoints`

| price_point_id | offer_id | unit_name | pricing_role | previous_downpayment | current_downpayment |
|---|---|---|---|---:|---:|
| `POST123-P01` | `POST123-O01` | `PCX160` | `eligible_model` | 8000 | 4000 |

### `Addons`

| addon_id | linked_offer_id | unit_name | addon_type | addon_name |
|---|---|---|---|---|
| `POST123-A01` | `POST123-O01` | `PCX160` | `helmet` | `Free helmet` |

## 10. Common mistakes when analyzing the workbook

### Mistake: Counting `PricePoints` rows as posts

Use `Posts` for post counts. `PricePoints` intentionally has more than one row
when a post contains multiple motorcycles or amounts.

### Mistake: Treating `featured_models` as eligible models

`featured_models` means “shown or mentioned.” Use `eligible_models` and
`offer_scope` for campaign eligibility.

### Mistake: Treating every downpayment as a promo

Use `has_financial_promo` or `promo_mechanics`. An ordinary downpayment can be
pricing without being promotional.

### Mistake: Treating every peso amount as motorcycle pricing

Check the corresponding price field, `unit_name`, and `pricing_role`. Event
fees, phone numbers, dates, displacement, and mileage are not motorcycle
prices.

### Mistake: Reading blank as zero

Blank means unavailable or uncertain. Explicit zero downpayment is stored as
the numeric value `0`.

### Mistake: Assuming a freebie is a financial promo

Use `has_addon` for freebies. A financial promotion requires a separate special
financial mechanic.

### Mistake: Summing `estimated_api_cost_usd` without checking errors

The cost column is an estimate based on returned usage and configured rates.
Rows that fail before usage metadata is returned may not contain a complete
allocation. The provider billing console remains the final billing record.

## 11. Useful analytics starting points

### Content mix

Use `Posts.primary_post_type` to compare how much competitors publish about:

- commercial offers
- product showcases
- lifestyle/brand content
- events/community
- customer-service advisories
- corporate/employer topics

### Offer mix

Use `Offers.promo_mechanics` to compare piso deals, zero downpayment, reduced
downpayment, rebates, discounts, and special financing.

### Eligibility strategy

Use `Offers.offer_scope` to compare specific-model promotions against broader
participating-model or brand-wide campaigns.

### Model-price monitoring

Use `PricePoints.brand`, `unit_name`, `pricing_role`, and the monetary fields.
Filter out `unassigned` rows when an analysis requires confirmed model-price
alignment.

### Free-benefit strategy

Use `Addons.addon_type` to compare helmets, registration, insurance,
maintenance, delivery, and other benefits. Count distinct `addon_id` values,
not only `Posts.has_addon`.

### Engagement comparison

Use `Posts.likes`, `comments`, and `shares` at the post grain. When joining to
detail sheets, avoid duplicating engagement totals across multiple offer,
price, or add-on rows.

## 12. Quick FAQ

### Why are some model or brand fields blank?

Because the source did not provide the value clearly or the relationship was
uncertain. Blank is preferred over a guessed value.

### Why can a post have `has_financial_promo = yes` but `has_advertised_pricing = no`?

A campaign can advertise a shared financial mechanic for participating models
without safely connecting the amount to one identifiable motorcycle.

### Why can a post have `has_advertised_pricing = yes` but `has_financial_promo = no`?

An ordinary cash price, installment, or downpayment is valid pricing but not a
special promotion.

### Why is one post repeated several times in a detail sheet?

It is not duplicated accidentally. Each detail row represents a separate
campaign, amount, or add-on from the same source post.

### Why is `unit_name` blank for piso downpayment?

If piso downpayment applies to unresolved participating models, it is a shared
`campaign_term`. Assigning it to a pictured model would overstate the evidence.

### Why does `current_downpayment` contain `1`?

The source said piso downpayment. The numeric field stores the usable amount,
while the description fields preserve the promotional wording.

### Where do I check exactly what the model read?

Open `Posts.llm_context`. It contains the labeled Facebook caption and every
non-empty OCR section supplied for that row.

### Where is the complete column reference?

Use `data_dictionary.md` for types, allowed values, population rules, and the
formal definition of every output column.
