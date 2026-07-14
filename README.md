# WB stock by supply batch

A Wildberries seller had a feeling that some stock seems to "go missing". Can we actually track this from the API?

Three things they wanted to know, per supply number:
1. how much was **received**
2. how much was **sold**
3. which batches are **just sitting there** — stock left, but no sales in ages

## Short answer

Yes for 1 and 2 - the supply number (`incomeId`) links receipts and sales cleanly. Number 3 is trickier: the Stocks endpoint doesn't carry the supply number, so per-batch stock isn't a thing you can just read. You derive it: `received − sold`. That gap, by the way, is exactly why stock feels like it "disappears" - nowhere is it broken down by batch

## What's here

- [`main.py`](main.py) — pulls Incomes + Sales, joins them, writes an xlsx (Batches / Diagnostics / Reconcile sheets). All config via env vars
- [`index.html`](index.html) — a dashboard mockup of the same table → **[live demo](https://larina-ch-daria.github.io/API-WB-checker/)**

## Heads up

The dashboard runs on made-up data — it's a layout prototype, not wired to the API. The script *is* real but has only been tested on synthetic input (no live token on hand), so field names and pagination want a check against actual API responses. Also: Sales only keeps 90 days, so for older batches "sold" undercounts — long history needs `reportDetailByPeriod` instead

## Run it
