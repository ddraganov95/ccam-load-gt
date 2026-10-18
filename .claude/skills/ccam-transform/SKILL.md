---
name: ccam-transform
description: Transform a CNAM CCAM Excel workbook into the exact semicolon-delimited GraphTalk loading format by reconstructing hierarchical Code, Activity, and Phase records, applying Code-only deferred substitution, and validating the result automatically.
---

# CCAM to GraphTalk Normalizer

## Purpose

Transform a CNAM CCAM Excel workbook into the exact text structure expected by GraphTalk.

Input: a CNAM CCAM Excel workbook.

Output: one logical GraphTalk record per physical line, with exactly nine semicolon-separated fields:

```text
Code;Actvty;Phase;Label;Date_str;Base_tariff_in;Tariff_out_cas;Group_code;Tail
```

Do not treat each physical Excel row as an independent output record. Reconstruct the logical hierarchy of Code, Activity, and Phase across rows.

Claude owns the transformation, deterministic correction, and validation. The human approves the finished artifact but must not be asked to repeat checks Claude can perform.

## Required execution behavior

Claude must:

1. Read the full workbook, including merged cells and the header area.
2. Process source rows in their original order.
3. Reconstruct every logical CCAM record, including records whose values are distributed across several rows.
4. Generate the complete output artifact.
5. Parse the generated artifact again.
6. Run every validation defined in this skill.
7. Correct deterministic failures.
8. Regenerate and revalidate until the result passes or a genuinely blocking contradiction remains.

A human approval gate does not mean the human performs Claude's data-quality work.

## Core record-building rule

Generate one output record for each distinct logical combination:

```text
(Code, Activity, Phase)
```

The Excel is hierarchical:

- Code is the parent.
- Activity belongs to Code.
- Phase belongs to Activity.
- If one Code has several activities, repeat the Code for every activity record.
- If one Activity has several phases, repeat both Code and Activity for every phase record.
- Values shown once at Code or Activity level must be inherited by the corresponding child records.
- A logical output record does not need to originate from one physical row containing every field.

Do not generate output records from section headings, chapter headings, `Comprend`, `Facturation`, exclusions, indications, notes, parenthetical references, or other descriptive rows.

Do not deduplicate using Code alone. `(Code, Activity, Phase)` must be unique in the generated output.

## Stateful source parsing

Process the workbook sequentially. Reconstruct logical Code blocks and preserve source-row values unless a rule explicitly changes one field.

### Logical Code blocks

A logical Code block begins when a non-empty `CodeValue` is encountered and continues until the next non-empty `CodeValue`.

The Code does not need to be repeated on every Activity or Phase row. Blank Code cells inside the block belong to the Code that opened the block.

For each block, collect its main label, Activity rows, Phase rows, child descriptions, tariffs, Regroupement values, and other structural values. Generate one record for each valid `(Code, Activity, Phase)` combination.

Example: if `FERP001` opens a block containing Activity 1 with Phases 1, 2, and 3, all three child rows belong to `FERP001` even though the Code is written only once.

### Deferred Code substitution when the CodeValue row has no Regroupement

When a row contains `CodeValue` and its main label but has no Regroupement on that same CodeValue row:

1. Store the new CodeValue and label as a pending source act.
2. Scan forward only until either:
   - another non-empty CodeValue is encountered; or
   - a qualifying row containing Regroupement together with Activity and Phase is encountered.

If another CodeValue is encountered first:

- discard the pending act;
- emit no output record for it.

If a qualifying Regroupement + Activity + Phase row is encountered first:

- generate the pending act's logical Activity/Phase record;
- substitute only the output Code with the previous valid output Code;
- preserve the pending act's own label;
- preserve Activity and Phase from the current qualifying row;
- preserve both tariffs from the current qualifying row exactly;
- preserve Group_code from the current qualifying row;
- preserve every other current-record value, including Tail.

Do not inherit or overwrite label, Activity, Phase, tariffs, Group_code, date, or Tail. Only the output Code is substituted.

Known V80 example:

- source act `ZZMP012` has no Regroupement on its CodeValue row;
- its qualifying Phase rows contain Activity, Phase, and Group_code `ATM`;
- output Code is substituted with previous valid Code `ZZMP018`;
- the phase labels, tariffs, Group_code, and other values remain those of the `ZZMP012` source rows.

```text
ZZMP018;1;1;<ZZMP012 source label> Phase 1 : <phase description>;05/11/2025;;;ATM;
ZZMP018;1;2;<ZZMP012 source label> Phase 2 : <phase description>;05/11/2025;668;668;ATM;
```

This is Code-only substitution. It is not full-record inheritance.

## Field mappings

### 1. Code

Source: Excel column `Code` / `CodeValue`, resolved through logical-block parsing and the deferred Code-substitution rule.

Normally, use the current block's source Code. When a CodeValue row has no Regroupement and a qualifying Regroupement + Activity + Phase row is found before the next CodeValue, substitute only the output Code with the previous valid output Code. Preserve all other fields from the current source act and qualifying row.

### 2. Actvty

Source: Excel column `Activité`.

Copy the activity code for the logical record.

Activity 4 represents anesthesia. Activity 5 can represent circulation extracorporelle `[CEC]`. The activity-label mechanism is general and is not limited to these values.

### 3. Phase

Source: Excel column `Phase`.

Phase belongs to Activity. One Activity can contain several phases.

For an ordinary non-phased activity, use source Phase 0. For a phased activity, generate one record per phase and repeat Code and Activity.

Example:

```text
FERP001;1;1;...
FERP001;1;2;...
FERP001;1;3;...
```

### 4. Label

The Label describes the specific `(Code, Activity, Phase)` output record.

#### Ordinary record

When there is no activity-specific or phase-specific child description:

```text
Label = main CCAM label
```

#### Activity-specific record

When an activity has its own textual description, combine the main label with only the description belonging to the current Activity.

```text
Activity 4 ordinary anesthesia:
<main label> - anesthésie

Activity 5 when labelled CEC:
<main label> - circulation extracorporelle [CEC]

Other explicitly named activities:
<main label> Activité N : <activity description>
```

Do not add a dash before `Activité N :`.

#### Anesthesia ownership across multiple activities

Determine anesthesia ownership first from source-row sequence inside the same Code and Phase.

If the sequence is:

```text
Activity 1
anesthésie
Activity 2
```

anesthesia belongs only to Activity 1. Do not propagate it to Activity 2.

If the sequence is:

```text
Activity 1
Activity 2
anesthésie
```

and the activities share the same Phase, use semantic judgment to classify anesthesia as either GLOBAL or ACTIVITY-SPECIFIC.

- If the label meaning clearly limits anesthesia to one or more activities, classify it as ACTIVITY-SPECIFIC.
- If the meaning is inconclusive, default to GLOBAL across the preceding activities in the same Phase.
- Never propagate anesthesia across a Phase boundary.
- Always record this ambiguous-case decision in the human-review report.

Output label by decision:

```text
GLOBAL:
<main label> - anesthésie

ACTIVITY-SPECIFIC:
<main label> Activité N : <activity description> - anesthésie
```

A GLOBAL anesthesia label must not contain `Activité N :`. An ACTIVITY-SPECIFIC label must identify the relevant Activity and preserve its description.

For each reviewed decision include Code, Phase, activities considered, source sequence, relevant labels, decision, activities receiving anesthesia, concise reason, and affected output record.

Do not fail solely because a reviewed anesthesia decision exists.

#### Phase-specific record

When an activity is divided into phases:

```text
<label> Phase N : <phase description>
```

Do not add a dash before `Phase N :`. Preserve the complete Phase number and description.

Example:

```text
Photochimiothérapie extracorporelle des cellules sanguines mononucléées Phase 1 : collecte des leucocytes par cytaphérèse [leucaphérèse]
```

#### Anesthesia associated with a phase

When anesthesia belongs to a specific phase, preserve the phase-specific label and append ` - anesthésie`:

```text
<main label> Phase N : <phase description> - anesthésie
```

Do not add a dash before `Phase N :`.

#### Label ownership and multiline text

Build every record's label only from text belonging to the current logical act.

- Never take text from the preceding or following act as the current act's label.
- Never carry a previous Activity description into the next Activity.
- Never carry a previous Phase description into the next Phase.
- Never stack several child descriptions into one Label.
- Never reproduce label leakage or stacked-label corruption from an older reference output.

For a multiline label cell, treat the first line as the main-label candidate and evaluate subsequent lines semantically in the context of the logical Code block and sibling acts.

Retain a subsequent line only when it is a short, direct continuation or qualification that makes sense as part of the act's name. Examples may include a concise qualification such as `Avec ou sans recouvrement céramique` when it directly completes the act label.

Exclude lines that function as:

- indications;
- billing or eligibility conditions;
- recommendations;
- exclusions;
- references to other acts;
- frequency conditions;
- long clinical context;
- lists of circumstances, indicators, or criteria;
- explanatory notes rather than the act name.

When text is retained, join it to the main label with one space. Remove embedded CR/LF characters, trim leading and trailing whitespace, and collapse accidental repeated whitespace.

If inclusion is genuinely ambiguous, make the most plausible semantic decision, generate the output, and add the case to the human-review report. Include Code, original multiline cell, retained text, excluded text, and concise reasoning. Do not fail solely because of this reviewed ambiguity.

#### Displaced activity text

The workbook can contain visually displaced child descriptions. If `anesthésie` is within the same logical Code block and is followed by Activity 4 structural values, associate it with Activity 4. Apply the same structural reasoning to other displaced child descriptions.

Do not cross a new `CodeValue` boundary when resolving displaced text.

#### Semicolons inside labels

Semicolon is the output delimiter. A literal semicolon inside a Label must not create an extra output field.

Before serialization:

- replace semicolons inside Label with periods;
- preserve the remaining label text;
- do not truncate the label.

After serialization, parse the line again and confirm that it contains exactly nine fields.

### 5. Date_str

Extract the applicability date from a workbook header matching the semantic pattern:

```text
CNAM - CCAM version <version> applicable <DD/MM/YYYY>
```

Use the same extracted applicability date for every generated record.

Do not hard-code a workbook cell. Search the header and merged header cells semantically. If exactly one applicability date is found, use it. If none or conflicting dates are found, fail with a precise blocking error. Do not invent or substitute a date.

### 6. Base_tariff_in

Source:

```text
Tarif Secteur 1 / adhérent OPTAM/OPTAM-ACO (en euro)
```

This is the CAS tariff, also called tarif in.

### 7. Tariff_out_cas

Source:

```text
Tarif Hors secteur 1 / hors adhérent OPTAM/OPTAM-ACO (en euro)
```

This is the Hors CAS tariff, also called tarif out.

#### Tariff handling

For numeric tariffs:

- preserve the Excel value exactly, including every decimal digit represented by the source value;
- convert only the decimal separator from `.` to `,` when needed;
- do not cap precision at seven decimal places;
- do not round;
- do not replace current-row tariffs through Code substitution;
- do not pad with additional trailing zeroes;
- write integers without a decimal separator.

Examples:

```text
126.92 -> 126,92
634.6 -> 634,6
209 -> 209
677.769888559025 -> 677,769888559025
```

If the tariff is blank or marked `Non pris en charge`, output an empty field. Do not output `0` or `Non pris en charge`.

### 8. Group_code

Source: Excel column `Regroupement`.

Copy Group_code from the current qualifying Activity/Phase row. Under deferred Code substitution, Group_code remains the current row's Regroupement and must not be inherited from the previous output record.

### 9. Tail

Tail is the ninth field. It is optional, but the ninth field must always exist.

```text
Empty Tail:     ...;Group_code;
Populated Tail: ...;Group_code;29
```

Tail is record-specific. Never inherit Tail from the main procedural activity into anesthesia or another secondary activity.

#### Tail meaning

Tail represents a discrete quantity that defines the principal dental act or a validated quantity-based act family.

Do not extract every number appearing in a dental label. Determine what the number counts and whether changing that number defines a different act in the same logical family.

Qualifying counted objects include:

- teeth, including named tooth types such as incisors, canines, premolars, and molars;
- dental roots;
- implants in qualifying direct dental/intrabuccal implant procedures.

The word `implant` alone is insufficient. Implant Tail applies only to validated dental/intrabuccal implant families. Cranial, facial, orthopedic, and other non-intrabuccal implants must have empty Tail.

#### Exact quantity

When a validated family uses one exact quantity greater than 1 as its discriminator, populate Tail with that integer.

Examples:

```text
Avulsion de 29 dents ...                   -> Tail 29
Avulsion de 2 racines ...                  -> Tail 2
Pose de 5 implants ...                     -> Tail 5
Avulsion de 3 incisives ...                -> Tail 3
Pose d'une prothèse ..., comportant 8 dents -> Tail 8
```

An explicit Tail value of 1 does not require special semantic extraction when downstream processing treats empty Tail as 1. Do not allow singular-language interpretation to introduce false positives.

#### Principal-act semantics

Use the quantity when it defines the principal act itself.

Examples of qualifying semantics include:

- direct avulsion, placement, removal, reimplantation, treatment, or another direct procedure on N qualifying objects;
- `Dégagement et activation de N implants`;
- a dental prosthesis whose defining size is `comportant N dents`;
- a quantity-based supplement whose own family is explicitly differentiated by `pose de N dents contreplaquées` or `pose de N dents contreplaquées ou massives`.

#### Quantity ranges

If the relevant quantity is expressed as a closed range `N à M`, Tail must be empty.

Examples:

```text
1 à 3 dents    -> empty Tail
5 à 6 dents    -> empty Tail
8 dents à 10 dents -> empty Tail
29 dents à 32 dents -> empty Tail
```

Apply this automatically.

#### Open-ended lower boundaries

`N ou plus` is not a closed `N à M` range.

When a validated numbered family uses the lower boundary as its discriminator, use N.

Example:

```text
11 implants ou plus -> Tail 11
```

Do not generalize this to every threshold expression. Use the lower boundary only when sibling acts establish a quantity-based family.

#### Secondary and incidental quantities

Keep Tail empty when the number does not define the principal act and instead refers to:

- a secondary object or relationship;
- implants connected by a bar;
- faces or surfaces;
- angles;
- arcades;
- prosthetic elements that do not define a validated tooth-count prosthesis family;
- ordinal elements such as `1er élément` or `2e élément`;
- the quantity in an underlying procedure merely referenced by a supplement;
- unrelated numbered characteristics.

Important negative examples:

```text
Pose de barre de conjonction entre 2 implants intrabuccaux -> empty Tail
Supplément pour avulsion de 4 troisièmes molaires ...      -> empty Tail
```

The first number describes the implants connected by the bar. The second number describes the underlying avulsion referenced by the supplement. Neither quantity is the discriminator of the principal output act.

#### Anesthesia and secondary activities

Activity 4 anesthesia must always have empty Tail, regardless of any quantity appearing in the label.

```text
Pose de 7 implants ...; Activity 1 -> Tail 7
Pose de 7 implants ...; Activity 4 -> empty Tail
```

Do not copy the main activity's Tail into anesthesia.

#### Tail ambiguity handling

Resolve normal Tail cases by semantic family analysis. Inspect sibling acts with the same wording pattern and apply one rule consistently to the entire family.

Do not produce a record-by-record discretionary review list. Only fail a family when the workbook, sibling acts, written rules, and available reference output contain a genuine unresolved contradiction.

## Reference normalized output

Use an older normalized output to establish GraphTalk structure, formatting conventions, and validated family behavior.

Do not blindly reproduce defects from an older output, including:

- text leaking from a preceding or following act;
- stacked child labels;
- truncated labels;
- malformed field counts;
- broken physical records caused by embedded line breaks;
- incorrect activity inheritance;
- Code or label association that violates Code qualification by Regroupement;
- missing Tail values that contradict the validated Tail rules.

When the old output differs from the workbook, first determine whether the difference results from a valid hierarchical or deferred Code-substitution rule. Do not classify a Code as wrong until the stateful source parsing has been applied.

## Expected-output comparison normalization

When comparing generated output with an older expected/reference output, exclude formatting-only differences from the difference summary, issue count, PASS/FAIL result, and human-review report.

For comparison only, normalize both sides by:

1. replacing CR, LF, and tabs with spaces;
2. collapsing consecutive whitespace to one space;
3. trimming leading and trailing whitespace;
4. treating ` - Phase N :` and ` Phase N :` as equivalent;
5. treating ` - Activité N :` and ` Activité N :` as equivalent;
6. ignoring CRLF-versus-LF file-line-ending differences.

Do not report differences caused only by leading/trailing spaces, repeated spaces, tabs, embedded line breaks, labels split across physical lines in the reference, or dashes before `Phase N :` / `Activité N :`.

This normalization applies only to reference comparison. The generated artifact must still follow the current output rules: no dash before `Phase N :`, no dash before `Activité N :`, normalized whitespace, and no embedded line breaks.

## Output format

Produce one logical record per physical line.

- Encoding: Windows-1252 / CP1252.
- Byte order mark: none.
- Line ending: Windows CRLF (`\r\n`).
- Field separator: semicolon (`;`).
- Header row: none unless explicitly requested.
- Embedded line breaks: forbidden in serialized records.
- Every record must contain exactly nine fields in this order:

```text
Code;Actvty;Phase;Label;Date_str;Base_tariff_in;Tariff_out_cas;Group_code;Tail
```

If Tail is empty, the physical line ends with `;`.

If Tail is present, the physical line ends with the numeric Tail value and has no additional trailing semicolon.

## Mandatory validation loop

Before completion, parse and validate the generated artifact itself.

Validate all of the following:

- Every physical line contains exactly one logical record.
- Every line has exactly nine semicolon-separated fields.
- Every line has Code, Activity, Phase, Label, Date, and Group_code.
- The file is Windows-1252 without BOM.
- All record separators are CRLF.
- No embedded CR or LF remains inside a record.
- One applicability date is used consistently.
- Every `(Code, Activity, Phase)` combination is unique.
- Multiple activities repeat the resolved Code correctly.
- Multiple phases repeat the resolved Code and Activity correctly.
- Every logical Code block is reconstructed from its opening non-empty `CodeValue` through the row before the next non-empty `CodeValue`.
- Blank Code cells inside a logical block inherit the Code that opened the block.
- For a CodeValue row without Regroupement, look ahead only until the next CodeValue or a qualifying Regroupement + Activity + Phase row.
- If the next CodeValue appears first, the pending act is skipped.
- If the qualifying row appears first, only output Code is substituted with the previous valid output Code.
- Under Code substitution, label, Activity, Phase, tariffs, Group_code, date, and Tail remain from the current source act/current qualifying row.
- The V80 regression case outputs the `ZZMP012` phase records under Code `ZZMP018`, while preserving Phase 1 empty tariffs, Phase 2 tariffs `668/668`, and Group_code `ATM`.
- Regroupement qualification is evaluated across the complete logical Code block, not only on the `CodeValue` row.
- The V80 regression case outputs `ZZMP012` label/phase information under Code `ZZMP018` and Group_code `ATM`; `ZZMP012` does not appear in output.
- A pending act is never carried across a new `CodeValue` boundary.
- A `CodeValue` row without Regroupement is not emitted unless a qualifying subsequent row supplies Regroupement and Activity/Phase.
- No headings, notes, exclusions, indications, or billing text became records.
- No label text leaked from an adjacent act.
- Multiline labels were normalized to one physical line.
- Blank and `Non pris en charge` tariffs became empty fields.
- Tariffs use comma decimal separators without unnecessary trailing zeroes.
- Tariffs preserve every decimal digit represented by Excel and are never capped at seven decimals.
- Empty-Tail records end with `;`.
- Populated-Tail records end with the Tail number and no trailing `;`.
- Exact-quantity Tail families are populated consistently.
- Closed `N à M` ranges have empty Tail.
- Tail is not copied to anesthesia.
- Secondary and incidental quantities do not become Tail.
- Child labels do not leak or stack across records.
- GLOBAL anesthesia labels use `<main label> - anesthésie` and contain no Activity number.
- ACTIVITY-SPECIFIC anesthesia labels use `<main label> Activité N : <description> - anesthésie`.
- Every ambiguous multi-activity anesthesia decision is present in the human-review report.
- Phase labels contain no dash before `Phase N :`.
- Activity 5 CEC labels are constructed correctly when present.
- Labels contain no delimiter semicolon after serialization.
- No label was truncated during sanitization.

Correct deterministic failures and rerun validation before returning the artifact.

## Blocking failure definition

FAIL is allowed only when a genuine source contradiction remains after applying the full hierarchy, logical-block and Regroupement-qualification rules, semantic family analysis, and automated validation.

A blocking failure must identify:

- Code or family;
- conflicting source evidence;
- exact reason no deterministic result can be produced.

Do not classify ordinary transformation decisions as blocking failures.

## Required final response

Return:

- generated output file;
- status `PASS` or `FAIL`;
- total output record count;
- unique Code count;
- multi-activity Code count;
- phased-record count;
- populated-Tail count;
- empty-tariff record count;
- human-review report containing every ambiguous anesthesia decision and ambiguous multiline-label decision, if any;
- if `FAIL`, only blocking Codes/families with exact reasons.

Do not provide a general manual checklist. The explicit anesthesia and multiline-label review entries are mandatory nonblocking exceptions.

## Human-in-the-loop boundary

The human reviews and approves the completed artifact before GraphTalk loading.

Claude remains responsible for transformation correctness, deterministic corrections, and automated validation. Human approval must not be used as a substitute for checks defined in this skill.
