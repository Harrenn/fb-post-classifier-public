# Competitor Facebook Post Classification — Data Dictionary

## Purpose

This workbook is an experimental analytics layer for studying what competing
motorcycle sellers publish on Facebook. It separates four different analytical
grains:

1. `Posts`: one row per Facebook post.
2. `Offers`: one row per campaign and eligibility scope found in a post.
3. `PricePoints`: one row per reliably aligned model-price observation or
   campaign-level financial term.
4. `Addons`: one row per free item, service, or bundled benefit.

The separation is intentional. A campaign can mention participating models,
feature four motorcycles with four different prices, and include two freebies.
That should be represented as one post, one campaign, four price points, and two
add-on rows—not as one very wide row with unrelated values joined together.

## Relationships and keys

```text
Posts (post_id)
  ├── Offers (post_id, offer_id)
  │     ├── PricePoints (offer_id, price_point_id)
  │     └── Addons (linked_offer_id, addon_id) [when the relationship is clear]
  ├── PricePoints (post_id)
  └── Addons (post_id)
```

- `post_id` is the common join key across all sheets.
- `offer_id` joins a campaign in `Offers` to its observations in `PricePoints`.
- `linked_offer_id` is populated only when an add-on can be safely tied to one
  campaign. It remains blank rather than guessing.
- All generated IDs are unique within one output workbook.

## Global representation rules

- Blank means the source did not provide the value, the relationship was
  unclear, or the value could not be safely extracted. Blank does not mean zero.
- Lists stored in one cell use `; ` as the separator.
- Boolean fields use `yes` or `no`.
- Monetary fields are numeric. They do not contain peso signs, commas, or words;
  Excel formatting displays them as Philippine peso amounts where appropriate.
- Dates use `YYYY-MM-DD` only when a complete date is supported.
- The LLM may use the caption and every non-empty OCR field, but it may not use
  outside knowledge to fill missing facts.
- A brand or model must be readable in the supplied caption/OCR. Dealer names,
  such as Motortrade, are not treated as motorcycle brands.
- A phone number, date, engine displacement, mileage, or unrelated peso amount
  must not be stored as motorcycle pricing.
- A clearly price-like motorcycle amount whose exact role cannot be established
  is retained in `observed_price`; it is not forced into cash price,
  downpayment, discount, or another more specific field.
- API cost is an estimate based on the configured paid standard-tier token
  rates. It is not an invoice and may differ from the actual charge because of
  free-tier eligibility, pricing changes, billing adjustments, or responses
  that fail before token metadata is returned.
- Commercial posts receive a second Gemini completeness pass by default. The
  second pass re-reads the caption and every OCR section, audits the first
  structured result, and returns a corrected result. Its tokens and cost are
  included in the same post-level usage columns. Set
  `GEMINI_REVIEW_COMMERCIAL=0` to disable this pass.
- Requests are spaced by at least 4.2 seconds by default. Set
  `GEMINI_MIN_REQUEST_INTERVAL_SECONDS` to change the interval. Temporary 429
  responses use Gemini's returned retry delay; a row-level failure is recorded
  only after the retry budget is exhausted.

## Sheet: `Posts`

**Grain:** one row per source Facebook post.

**Primary key:** `post_id`.

| Column | Type | Population | Example | Definition and logic |
|---|---|---:|---|---|
| `post_id` | text identifier | Always | `pfbid0wXyv74...` | Stable row identifier derived from the Facebook permalink when possible; otherwise generated from the file name and row number. |
| `competitor_page` | text | When derivable | `MotortradePH` | First path component of the Facebook URL, used as the posting-page identifier. |
| `url` | URL text | Source value | `https://www.facebook.com/MotortradePH/posts/...` | Original Facebook post permalink retained from the input CSV. |
| `post_detail` | text | Always | `Promotes a piso-downpayment campaign for participating Kymco models.` | Concise, source-grounded overall description of the combined caption and OCR. |
| `primary_post_type` | enum | Always | `commercial_offer` | Exactly one dominant content type. Allowed values are documented under “Controlled values.” |
| `content_themes` | enum list | Always; may be blank | `financing; pricing; product_showcase` | All applicable topical themes, separated by semicolons. These are descriptive themes, not mutually exclusive classifications. |
| `offer_bundle_type` | derived enum | Always | `promo_and_pricing_and_addon` | Derived from the three commercial booleans. Values range from `no_offer` to combinations such as `promo_and_pricing`, `pricing_and_addon`, or `promo_and_pricing_and_addon`. |
| `has_financial_promo` | yes/no | Always | `yes` | `yes` when at least one validated offer contains a special financial mechanic such as piso/zero/reduced downpayment, cash discount, rebate, special financing, trade-in, or reservation discount. A normal price or ordinary downpayment alone is not a promo. |
| `has_advertised_pricing` | yes/no | Always | `yes` | `yes` when an identifiable motorcycle is reliably connected to a price or financing amount. Four featured models with four separately aligned prices qualify. One shared campaign amount with no reliable model relationship does not. The count of featured models alone never establishes a price relationship. |
| `has_addon` | yes/no | Always | `yes` | `yes` when at least one specific, identifiable, guaranteed free item, service, accessory, subscription, or benefit exists in `Addons`. Generic “freebies,” raffle prizes, and fulfillment-speed promises do not qualify. A real freebie alone does not make `has_financial_promo` yes. |
| `call_to_action` | enum list | Always | `visit_branch; message_page` | All explicit actions requested by the post. Uses semicolon-separated controlled values. `none` is used when no call to action is present. |
| `likes` | number | When parseable | `124` | Source engagement count. Abbreviations are expanded when parseable, for example `1.2K` becomes `1200`. |
| `comments` | number | When parseable | `18` | Source comment count, retained as a number. |
| `shares` | number | When parseable | `7` | Source share count, retained as a number. |
| `llm_context` | text | When source text exists | `[FACEBOOK POST CAPTION] ... [IMAGE 1 OCR] ...` | Exact labeled caption-plus-OCR context sent to Gemini. Empty caption/OCR values are omitted. |
| `llm_model` | text | When an API call is attempted | `gemini-3.5-flash-lite` | Exact Gemini model requested for the row. Blank when no usable context existed and Gemini was skipped. |
| `llm_input_tokens` | integer | When usage metadata is returned | `4000` | Sum of prompt and tool-use prompt tokens across all Gemini responses for the row, including responses that later fail structured-output validation and cause a retry. Zero when Gemini is intentionally skipped. |
| `llm_output_tokens` | integer | When usage metadata is returned | `600` | Sum of candidate/output tokens across all Gemini responses for the row. Thinking tokens are stored separately. |
| `llm_thinking_tokens` | integer | When usage metadata is returned | `100` | Sum of Gemini thinking tokens across all responses for the row. These tokens are included in output pricing. |
| `llm_total_tokens` | integer | When usage metadata is returned | `4700` | Sum of Gemini’s reported total tokens across all responses for the row. It is retained for operational monitoring and is not used alone for cost because input and output rates differ. |
| `estimated_api_cost_usd` | USD number | When usage and rates are available | `0.00295000` | Paid-tier estimate: `(llm_input_tokens × input rate + (llm_output_tokens + llm_thinking_tokens) × output rate) ÷ 1,000,000`. Zero when Gemini is skipped; blank when usage metadata or model pricing is unavailable. |
| `processing_error` | text | Only on row failure | `503 UNAVAILABLE: service temporarily unavailable` | Sanitized operational error for a failed row. It never contains the API key. Other rows continue processing. |

### API cost basis

Built-in standard paid-tier rates as of 2026-07-30 are:

| Model | Input per 1M tokens | Output/thinking per 1M tokens |
|---|---:|---:|
| `gemini-3.5-flash-lite` | US$0.30 | US$2.50 |
| `gemini-3.6-flash` | US$1.50 | US$7.50 |
| `gemini-3.5-flash` | US$1.50 | US$9.00 |
| `gemini-3.1-flash-lite` | US$0.25 | US$1.50 |

Pricing source: https://ai.google.dev/gemini-api/docs/pricing

Usage-metadata reference: https://googleapis.github.io/python-genai/

For another model or a later price, set both environment variables before
running:

```text
GEMINI_INPUT_USD_PER_MILLION
GEMINI_OUTPUT_USD_PER_MILLION
```

If the selected model has no built-in rate and no overrides are supplied, the
token columns are still populated but `estimated_api_cost_usd` remains blank.
When a generated response fails JSON validation and is retried, its returned
usage is included because it may still be billable. API/transport failures that
return no usage metadata cannot be allocated precisely to an individual row;
the Google billing console remains the final billing record.

## Sheet: `Offers`

**Grain:** one row per campaign and eligibility scope within a post.

**Primary key:** `offer_id`.

**Foreign key:** `post_id` → `Posts.post_id`.

Pricing values are deliberately not stored here. This sheet answers: “What is
the campaign, and who or what does it apply to?”

| Column | Type | Population | Example | Definition and logic |
|---|---|---:|---|---|
| `offer_id` | text identifier | Always | `pfbid0wXyv74-O01` | Generated campaign identifier, unique in the workbook. |
| `post_id` | text identifier | Always | `pfbid0wXyv74...` | Parent post identifier. |
| `url` | URL text | Source value | `https://www.facebook.com/...` | Parent Facebook post permalink, duplicated for convenient filtering and opening. |
| `promo_name` | text | When stated | `Kymco Piso Downpayment Promo` | Source-grounded campaign name or a concise label. It is not invented when no meaningful label is supported. |
| `promo_mechanics` | enum list | When promotional | `piso_downpayment; special_financing` | All validated special financial mechanics. Ordinary pricing is excluded. |
| `brand` | text | When established | `Kymco` | Motorcycle manufacturer governing the campaign when stated. Blank for mixed-brand or unsupported scope; model-level brands remain available in `PricePoints`. |
| `offer_scope` | enum | Always | `participating_models` | Eligibility scope, independent of what appears in the creative. See controlled values below. |
| `eligible_models` | text list | Only when explicit | `PCX160` | Models explicitly stated as eligible. Multiple explicit models use semicolons. Pictured examples are not placed here unless the wording confirms eligibility. |
| `featured_brands` | text list | When present | `Honda; Yamaha; Suzuki` | Motorcycle brands shown or mentioned as creative examples but not established as the campaign’s eligibility constraint. |
| `featured_models` | text list | When present | `Agility ECD; Skytown 150Fi ABS/TCS` | Readable models pictured or used as examples without proof that they form the complete eligible set. |
| `unit_condition` | enum | Always | `repo` | Source-supported inventory condition: `new`, `pre_owned`, `repo`, or `unspecified`. |
| `scope_details` | text | When useful | `Customers are instructed to ask about participating Kymco models.` | Concise explanation of the source wording that determined `offer_scope`, especially when featured examples are not confirmed eligible units. |
| `eligibility` | text | When stated | `Available to approved financing applicants at participating branches.` | Customer action, qualification, or constraint required to receive the campaign. |
| `valid_from` | ISO date | When complete | `2026-07-01` | Campaign start date. Partial or ambiguous dates remain blank. |
| `valid_until` | ISO date | When complete | `2026-07-31` | Campaign end date. Partial or ambiguous dates remain blank. |
| `participating_branches` | text | When stated | `Participating Motortrade branches` | Branch, dealer, or location restrictions stated in the source. |
| `promo_details` | text | When available | `Piso downpayment means ₱1; customers should ask for participating models and complete terms.` | Comprehensive campaign explanation, including financial comparisons, important conditions, dates, locations, and limitations. |

Scope is source-enforced after Gemini responds. `all_models` and `brand_wide`
require explicit universal wording such as `all`, `any`, `every`, `lahat`, or
`anumang`. Merely saying `pre-owned motorcycles`, `repo motorcycles`, or a
brand’s motorcycles does not establish universal eligibility.

Featured examples are normalized from general to specific. If one creative
contains both `Suzuki Raider` and a supported `Raider X150`, the generic label
is not retained as a second model. This prevents one model from appearing to be
two observations. Slash-containing trim names such as `ABS/TCS` remain one
model. A readable model-price pair in one image panel proves price alignment,
but does not by itself prove that the model defines eligibility for a wider
multi-image campaign.

## Sheet: `PricePoints`

**Grain:** one row per reliably aligned model-price observation or one
campaign-level financial term.

**Primary key:** `price_point_id`.

**Foreign keys:** `post_id` → `Posts.post_id`; `offer_id` →
`Offers.offer_id`.

This sheet answers: “What exact amount was advertised, and what unit or campaign
does that amount belong to?”

| Column | Type | Population | Example | Definition and logic |
|---|---|---:|---|---|
| `price_point_id` | text identifier | Always | `pfbid0wXyv74-P01` | Generated identifier for one pricing observation. |
| `offer_id` | text identifier | Always | `pfbid0wXyv74-O01` | Parent campaign identifier. |
| `post_id` | text identifier | Always | `pfbid0wXyv74...` | Parent post identifier. |
| `url` | URL text | Source value | `https://www.facebook.com/...` | Parent Facebook post permalink. |
| `brand` | text | When stated | `Honda` | Manufacturer tied to this observation. Blank rather than inferred. |
| `unit_name` | text | When safely aligned | `Click 125` | One identifiable model tied to this price point. Never contains multiple motorcycles. Blank for shared campaign terms or unresolved pricing. |
| `unit_condition` | enum | Always | `new` | Condition of the priced inventory: `new`, `pre_owned`, `repo`, or `unspecified`. |
| `pricing_role` | enum | Always | `featured_example` | Describes how the amount relates to campaign eligibility. See controlled values below. |
| `previous_downpayment` | integer | When stated | `8000` | Earlier downpayment in a before/now comparison. |
| `current_downpayment` | integer | When stated | `4000` | Current or final downpayment. `piso downpayment` is stored as `1`; zero downpayment is stored as `0`. For `BEFORE ₱8,000 NOW ₱4,000`, this field is `4000`. |
| `cash_price` | peso number | When stated | `80000` | Explicit cash price of the motorcycle. |
| `previous_price` | peso number | When stated | `85000` | Previous full motorcycle price, not a previous downpayment. |
| `current_price` | peso number | When stated | `79000` | Current or sale motorcycle price. |
| `observed_price` | peso number | Only when the amount is clearly price-like but its role is unclear | `159597.40` | Safely preserves a readable motorcycle-related amount that cannot be classified as cash price, downpayment, installment, reservation, discount, or rebate. It remains blank when the source does not establish that the number is commercial pricing. Permit numbers, phone numbers, dates, displacement, and mileage are never stored here. |
| `payment_amount` | peso number | When stated | `125` | Recurring installment amount. Interpret together with `payment_frequency`. |
| `payment_frequency` | enum | Always | `daily` | Frequency for `payment_amount`: `daily`, `weekly`, `monthly`, `other`, or `unspecified`. It remains `unspecified` when no recurring payment exists. |
| `loan_term_months` | integer | When stated | `36` | Financing term in whole months. |
| `reservation_amount` | peso number | When stated | `1000` | Amount explicitly labeled as a reservation fee or reservation payment. |
| `discount_amount` | peso number | When stated | `6000` | Explicit discount amount. It is not duplicated into a price field. |
| `rebate_amount` | peso number | When stated | `500` | Explicit rebate amount. |
| `price_details` | text | When available | `Downpayment changed from ₱8,000 to ₱4,000 for this model.` | Comprehensive wording that preserves comparisons, conditions, and context for this pricing observation. |

An unlabeled `observed_price` is separated from shared campaign mechanics. For
example, a RepoNalo post containing zero downpayment, a rebate, a discount, and
an unlabeled `₱64,352` produces one `campaign_term` row for the explicit
mechanics and one `unassigned` row containing only `observed_price = 64352`.
Multiple unlabeled amounts in one panel produce multiple `unassigned` rows.

Ordinary “as low as ₱203 daily” wording is advertised pricing when it is tied
to a model, but it does not create a financial-promo mechanic unless the source
also states a special or reduced deal. A heading such as “Daily Ride Deals” is
not enough to create `special_financing`. When caption and OCR give
contradictory generic daily amounts, both values remain explained in the
campaign details while the structured `payment_amount` stays blank. OCR strings
such as `PI50` are interpreted as `P150` only when surrounding text clearly
establishes a daily peso rate.

## Sheet: `Addons`

**Grain:** one row per distinct, identifiable, guaranteed free item, service,
accessory, subscription, or bundled benefit.

**Primary key:** `addon_id`.

**Foreign keys:** `post_id` → `Posts.post_id`; optional `linked_offer_id` →
`Offers.offer_id`.

If one post offers a free helmet and free registration, it creates two rows in
this sheet. `has_addon` remains one post-level `yes` in `Posts`.

Generic wording such as “exciting freebies” does not create a row unless a
specific benefit is named. Raffle prizes are contingent rather than guaranteed
add-ons. Same-day, fast, or express unit release describes fulfillment speed
and is not an add-on.

| Column | Type | Population | Example | Definition and logic |
|---|---|---:|---|---|
| `addon_id` | text identifier | Always | `pfbid0wXyv74-A01` | Generated identifier for one distinct add-on. |
| `post_id` | text identifier | Always | `pfbid0wXyv74...` | Parent post identifier. |
| `linked_offer_id` | text identifier | When safely linkable | `pfbid0wXyv74-O01` | Campaign receiving the add-on. Blank when the post has multiple campaigns and the relationship is unclear. |
| `url` | URL text | Source value | `https://www.facebook.com/...` | Parent Facebook post permalink. |
| `brand` | text | When explicitly tied | `Honda` | Manufacturer to which the add-on applies. Blank if the benefit is campaign-wide or unassigned. |
| `unit_name` | text | When explicitly tied | `PCX160` | Specific model receiving the add-on. It is not inferred from a nearby image. |
| `addon_type` | enum | Always | `helmet` | Standardized analytical category for the benefit. |
| `addon_name` | text | Always | `Free helmet` | Source-grounded benefit name, retaining useful wording from the post. |
| `addon_eligibility` | text | When stated | `Free with every financed purchase during the campaign.` | Customer action or requirement for receiving the add-on. |
| `addon_details` | text | When available | `One standard helmet is included at no additional cost.` | Concise source-grounded explanation, including relevant limitations. |

## Controlled values

### `primary_post_type`

| Value | Logic |
|---|---|
| `commercial_offer` | Contains a validated financial promo, identifiable advertised pricing, or add-on. |
| `product_showcase` | Highlights a motorcycle/model without a concrete commercial offer. |
| `brand_lifestyle` | Brand awareness, lifestyle, holiday, greeting, or generic branch marketing. |
| `event_community` | Ride, event, celebration, or community activity. |
| `customer_service_advisory` | Registration, plate, procedure, service, or support information. |
| `corporate_employer` | Hiring, employee recognition, incentive trip, or corporate/employer content. |
| `other` | Usable content that does not fit the defined types. |
| `unclear` | Insufficient usable caption/OCR context. |

### `content_themes`

Allowed values: `financing`, `pricing`, `product_showcase`, `product_launch`,
`pre_owned`, `registration`, `after_sales`, `safety`, `branch_promotion`,
`event`, `community`, `lifestyle`, `holiday`, `corporate`, `customer_support`,
and `other`.

### `call_to_action`

Allowed values: `apply_online`, `visit_branch`, `message_page`, `call`,
`visit_website`, `none`, and `other`.

### `promo_mechanics`

| Value | Logic |
|---|---|
| `piso_downpayment` | Promotional downpayment of exactly ₱1. |
| `zero_downpayment` | Promotional downpayment of ₱0. |
| `reduced_downpayment` | Downpayment is explicitly reduced or has a supported before/now comparison. |
| `cash_discount` | Explicit cash discount or discount amount. |
| `rebate` | Explicit rebate. |
| `reduced_monthly_payment` | Explicitly promotional reduced monthly installment. |
| `special_financing` | Concrete special financing terms not more precisely represented above. |
| `trade_in` | Trade-in-based deal. |
| `reservation_discount` | Promotional reservation-based discount. |
| `other` | Source-supported financial mechanic outside the standard list. |

### `offer_scope`

| Value | Logic |
|---|---|
| `specific_model` | Exactly one model is explicitly eligible. |
| `specific_models` | Multiple models are explicitly eligible under the same grouped campaign. This is a derived workbook value. |
| `participating_models` | The post refers to participating/selected/eligible models but does not provide a confirmed complete list. |
| `brand_wide` | The source explicitly says all or any models of the stated brand qualify. |
| `all_models` | The source explicitly says all motorcycles/models qualify without a brand limitation. |
| `unspecified` | The eligibility scope cannot be established. |

### `pricing_role`

| Value | Logic |
|---|---|
| `eligible_model` | The exact `unit_name` is explicitly eligible for the parent campaign. |
| `featured_example` | The creative shows one model with its own reliably aligned price, but that does not prove the model defines campaign eligibility. |
| `campaign_term` | The amount is a shared campaign mechanic, such as piso downpayment for unresolved participating models; `unit_name` is blank. |
| `unassigned` | A price-like amount is readable, but no safe model relationship can be established; `unit_name` is blank. |

Eligibility scope and pricing role are evaluated separately. A
`participating_models` campaign can therefore have a `campaign_term`, one or
more `featured_example` price points, or both. A single extracted featured model
does not automatically receive a shared campaign price.

### `unit_condition`

Allowed values: `new`, `pre_owned`, `repo`, and `unspecified`.

### `payment_frequency`

Allowed values: `daily`, `weekly`, `monthly`, `other`, and `unspecified`.

### `addon_type`

Allowed values: `helmet`, `riding_gear`, `registration`, `insurance`,
`maintenance`, `delivery`, `transfer_assistance`, `accessory`, `subscription`,
`service`, and `other`.

## Key classification scenarios

### Participating-model campaign with pictured examples

Source meaning: “Get a Kymco motorcycle for piso downpayment. Message us for
participating models,” while the image shows Agility ECD and Skytown 150Fi.

- `Offers`: one row with `offer_scope = participating_models`,
  `eligible_models` blank, and both models in `featured_models`.
- `PricePoints`: one row with `pricing_role = campaign_term`, `unit_name` blank,
  and `current_downpayment = 1`.
- The two pictured models are not falsely reported as the only eligible models.

### Explicit single-model offer

Source meaning: “This offer is applicable only to PCX160.”

- `Offers`: `offer_scope = specific_model` and `eligible_models = PCX160`.
- `PricePoints`: the PCX160 amounts use `pricing_role = eligible_model`.

### Four featured motorcycles with separate prices

If one post shows four motorcycle panels and each panel clearly pairs one model
with its own price:

- `Posts`: one row.
- `Offers`: normally one grouped campaign row.
- `PricePoints`: four rows, one per model-price pair.
- Each price point has one `unit_name`; values are never stored as four parallel
  semicolon lists that could become misaligned.
- If the caption says “participating models,” the four price rows can use
  `pricing_role = featured_example` while the campaign scope remains
  `participating_models`.
- `has_advertised_pricing = yes` because each model-price relationship is clear,
  even though eligibility remains broader or unresolved.

### One amount shared across four pictured motorcycles

If a post advertises one piso downpayment for participating models and merely
shows four motorcycles:

- Create one `campaign_term` price point with `current_downpayment = 1`.
- Keep `unit_name` blank.
- Do not repeat ₱1 four times or claim that the four pictured models are the
  complete eligible list.

### Repeated anonymous campaign mechanics across images

If several panels repeat the same campaign-level zero downpayment and ₱54,000
discount, those repetitions consolidate into one `campaign_term`. A separate
panel that clearly aligns a model with its own before/now motorcycle prices
remains a distinct `featured_example`; it is not promoted to an eligible model
without explicit eligibility wording.

### Before/now downpayment

For `BEFORE ₱8,000 NOW ₱4,000` when the wording is clearly about downpayment:

- `previous_downpayment = 8000`
- `current_downpayment = 4000`
- `price_details` and `promo_details` retain the full comparison.
- The value is not placed in `cash_price`, `previous_price`, or `current_price`.

### Multiple add-ons

For “Free helmet and free registration”:

- Create two `Addons` rows: one `helmet`, one `registration`.
- Set `Posts.has_addon = yes`.
- Do not turn the freebies into a financial promotion unless a separate
  financial mechanic is also present.
- Do not create an add-on from unnamed “freebies,” raffle prizes, or same-day
  release.

### No usable text

When both caption and every OCR value are empty:

- Skip the Gemini request.
- Create one `Posts` row with `primary_post_type = unclear`,
  all commercial booleans `no`, and
  `post_detail = No usable text was available.`
- Create no `Offers`, `PricePoints`, or `Addons` rows.
