# Engineering note: the region-merge bug

Kept deliberately. This is the single most useful thing in the repository,
because it shows the difference between code that runs and code that is correct.

## What happened

The first version of the entity resolver normalised every value to a
punctuation-free lowercase "fingerprint" and merged any two fingerprints whose
edit distance was ≤ 3. That works beautifully for company names:

```
auroratrading     ~ auroratradingltd    distance 3  -> merged  OK
orionlogistcs     ~ orionlogistics      distance 1  -> merged  OK
```

So it was shipped. Then the region column came out like this:

```
===== orders.region: 14 raw distinct -> 2 groups
  G1: canon='South'   <-- MERGED
        - 'East' (n=13)
        - 'West' (n=13)
        - 'South' (n=13)
        - 'North' (n=7)
        - 'None' (n=3)
        - 'NULL' (n=3)
        - 'null' (n=1)
        - 'n/a' (n=1)
        ...
```

**Every region in the file had collapsed into a single value.** The chain:

```
'n/a'  ~ 'east'   distance 3
'east' ~ 'west'   distance 2
'north' ~ 'south' distance 2
```

Each individual link looked reasonable. Together they bridged four unrelated
regions, and because the canonical label was chosen by frequency, the surviving
value was whichever happened to be most common.

## Why it matters

This is not a cosmetic bug. Every report the client runs —

- revenue by region
- regional manager performance
- shipping cost by territory

— would have come back with one region holding 100% of the data. The failure is
**silent**: the script exits 0, the output parses, the numbers are internally
consistent. Nothing looks wrong unless you know there should be five regions.

It is also the reason "just use a fuzzy match library" is not a complete answer.
`fuzzywuzzy`/`rapidfuzz` with a default threshold reproduces exactly this.

## The fix

Three changes, all in `build_canonical_map` / `_similar`:

**1. Scale the edit-distance threshold by word length.**

```python
allowed = 1 if min(len(fa), len(fb)) < 6 else 2
```

Short strings are always close to each other. Two 4-character words are one
letter-swap apart by chance; two 12-character words are not. A flat threshold
cannot serve both.

**2. Prefer token overlap for multi-word values.** `Aurora Trading` and
`Aurora Trading Ltd` share every meaningful word, so they merge on evidence that
does not degrade for short strings.

**3. Remove placeholders from the vocabulary before clustering.** `n/a`, `null`,
`-`, `N/A` and `0` are not spellings of an entity. Leaving them in lets them act
as bridges — `n/a` was the link that connected `east` to everything else.

## Verified result after the fix

```
===== orders.region: 8 real spellings -> 5 entities (7 placeholders set aside)
  1. 'East'      <- 'West'?  NO
  2. 'South'
  3. 'Central'   <- 'CENTRAL'
  4. 'North'
  5. 'West'
```

North, South, East and West stay separate; case variants still merge; spelling
mistakes still merge. Customer resolution was unaffected (41 → 8) and the
customer/order join improved from **38 unmatched orders to 0**, because pooling
both files into one clustering pass now produces one label per company.

## The general lesson

Similarity thresholds are **domain-specific and must be tested against a case
where the right answer is known.** For this pipeline that meant asserting that
five regions come out as five regions. A cleaning script with no such assertion
will happily report success while destroying the client's ability to segment
their own data.
