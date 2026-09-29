# Aware Super Seminar Room AV Upgrade (Project 2604704)
## Cost vs sell analysis: Complete AV bill INV41565 vs Aware Super invoice INV-10596

Prepared 29 Sep 2026 from Xero (Corporate Technology Services Pty Ltd). All figures ex GST unless marked.

Sources
- Bill INV41565, Complete AV Solutions Pty Ltd, dated 1 Sep 26, due 26 Sep 26, status Authorised (unpaid). Total $161,926.60 ex GST / $178,119.25 inc GST. Xero note: "Hold off until payment received from Aware".
- Invoice INV-10596, Aware Super, dated 9 Sep 26, due 9 Oct 26, status Authorised (unpaid). Total $205,352.17 ex GST / $225,887.40 inc GST.
- PO-00016902 "Q-12140v1 Audio Visual Proposal (PO 1/2)", Complete AV, $177,457.98 ex GST, status Billed.
- PO-00016899 "Q-12140v1 Audio Visual Proposal (PO 2/2)", Complete AV, $17,700.52 ex GST, status Authorised (NOT yet billed).

---

## 1. Headline

| | Ex GST | Inc GST |
|---|---|---|
| Sell to Aware (INV-10596) | 205,352.17 | 225,887.40 |
| Cost billed so far (INV41565) | 161,926.60 | 178,119.25 |
| Gross profit on what is billed today | 43,425.57 (21.1% margin) | |
| Complete AV retention still to be billed | 4,848.68 | 5,333.55 |
| Gross profit after retention | 38,576.89 (18.8% margin) | |
| PO 2/2 to Complete AV, not yet billed, contents unknown | 17,700.52 | 19,470.57 |
| Gross profit if PO 2/2 is all project cost | 20,876.37 (10.2% margin) | |

Plain English: on paper this job makes about 21%. Once you include the $4,848.68 Complete AV has told you is still coming, it is 18.8%. If the second purchase order ($17,700.52) is also cost for this job, the real margin drops to about 10%. You need to confirm what PO 2/2 covers before you call this job a 20% job.

---

## 2. Your tracker figure of $136,163 reconciles

Your note says "Complete AV materials ($136,163) is going out this week". That is exactly the bill inc GST minus the four labour lines:

| | Ex GST | Inc GST |
|---|---|---|
| Bill total | 161,926.60 | 178,119.25 |
| Less labour lines (23,320 + 7,196 + 6,561.80 + 1,064) | 38,141.80 | 41,955.98 |
| Materials portion | 123,784.80 | 136,163.27 |

So the $136,163 is the equipment/freight/cables part of INV41565 inc GST. The remaining $41,956 inc GST of labour on the same bill will still be owed to Complete AV after this payment. Make sure whoever pays knows it is a part payment of one bill, not the whole bill.

---

## 3. Section by section: what Complete AV charged vs what Aware was charged

The supplier bill is in three blocks that add up exactly to PO 1/2 ($177,457.98), then a variations block, then a retention line. Mapping below is best fit to the Aware invoice headings.

| Aware invoice line | Sell to Aware | Complete AV cost (best-fit block) | Gross profit | Margin |
|---|---|---|---|---|
| Phase 1 Audit | 0.00 | 0.00 | 0.00 | n/a |
| Phase 2 Installation and Integration | 144,049.00 | 120,676.64 (block 1: core kit + labour 23,320) | 23,372.36 | 16.2% |
| Option A Advanced Camera Tracking | 45,162.00 | 45,761.88 (block 3: tracking software, PC, 4x Sennheiser TCC2 mics, Dante licence, labour 6,561.80) | -599.88 | -1.3% |
| Option B Additional Presenter Cameras | 13,436.00 | 11,019.46 (block 2: 2x NVX-E30 encoders, cables, labour 7,196) | 2,416.54 | 18.0% |
| Upgrade AW-UE40 to AW-UE80, 2 devices | 12,042.88 | 0.00 on this bill | 12,042.88 | see section 5 |
| Variations (net credit) | -9,337.71 | -10,682.70 | 1,344.99 | see section 4 |
| Total | 205,352.17 | 166,775.28 (incl. 4,848.68 retention) | 38,576.89 | 18.8% |

Caveat on the mapping: the 4x Sennheiser ceiling mics ($23,786.60) sit in block 3 on the supplier bill. If they actually belong to Phase 2 rather than Option A, then Option A is very profitable (about 53%) and Phase 2 is a small loss (about -0.8%). Either way one of those two sections is priced at or below cost. The total is not affected.

Option B check: Aware's own variation wording says the job added "(+2) PA AW-UE40KEJ" cameras, but the supplier bill only shows 4 cameras (block 1) and 2 extra encoders (block 2). The 2 extra cameras are either inside the "Removed Panasonic Cameras & Licenses" credit or replaced by the UE80 upgrade. Confirm with Complete AV's quote Q-12140v1.

---

## 4. Variations: the pricing is backwards

Every variation was passed to Aware at a fixed percentage of the supplier figure. Five lines at 85%, two lines at 77%.

| Variation | Complete AV | Aware | Aware as % of cost | Effect on CTS |
|---|---|---|---|---|
| Removed Panasonic cameras and licences | -7,281.07 | -6,188.91 | 85% | +1,092.16 |
| RMB hardware added (angled mount kit) | 321.50 | 273.28 | 85% | -48.22 |
| Removed control processor CP4N | -3,032.23 | -2,577.40 | 85% | +454.83 |
| Removed camera Crestron encoder/decoder | -4,600.20 | -3,910.17 | 85% | +690.03 |
| Cables (rack cabling) | 691.60 | 587.86 | 85% | -103.74 |
| Graphics card RTX 3060 | 2,153.70 | 1,658.35 | 77% | -495.35 |
| Lighting integration programming | 1,064.00 | 819.28 | 77% | -244.72 |
| Net | -10,682.70 | -9,337.71 | | +1,344.99 |

What this means
- On the four ADDED items (mount kit, cables, graphics card, lighting programming) Aware is being charged LESS than Complete AV charged you. You lose $892.03 on those four lines.
- On the three REMOVED items you are crediting Aware less than Complete AV credited you. You keep $2,237.02.
- Net effect is +$1,344.99 in your favour, but only because the removals happened to be bigger than the additions. The method is wrong: someone applied a 15% (and 23%) discount to supplier cost instead of a markup. Phase 2 was priced at roughly cost x 1.19. At that rate the added items should have been about $5,050 not $3,339, and the credits about $17,802 not $12,677.
- Because the figures are exact 0.85 and 0.77 multiples, this looks like a formula or template error, not a negotiated price. Worth checking whether the same template was used on other Complete AV jobs.

Recommendation: leave INV-10596 as issued (re-issuing a $225k invoice to Aware three weeks before due date will delay payment), but fix the variation pricing method before the next Aware claim and before the Complete AV retention/PO 2/2 gets passed through.

---

## 5. Items on Aware's invoice with no matching cost on INV41565

Upgrade AW-UE40 to AW-UE80, supply and removal of loan units, 2 devices: $12,042.88 charged to Aware. Nothing on INV41565 relates to UE80 cameras. The cost must be either:
- in PO-00016899 (PO 2/2, $17,700.52 ex GST, still Authorised and unbilled), or
- from another supplier not yet billed.

If it is PO 2/2, the upgrade is being sold for $12,042.88 against a cost of $17,700.52, a loss of $5,657.64 on that line, and the job margin falls to 10.2% overall. Pull up PO-00016899 in Xero and check its line items. This is the single biggest unknown in the job.

---

## 6. Full supplier line detail (INV41565), grouped

Block 1, core install, matches Phase 2. Total $120,676.64
- 6x Crestron DM-NVX-E30 encoder @ 1,345.56 = 8,073.36
- 2x Crestron DM-NVX-385 switcher @ 3,390.25 = 6,780.50
- 8x Crestron DM-NVX-360 encoder/decoder @ 2,144.14 = 17,153.12
- 4x Panasonic AW-UE40WEJ camera @ 4,063.96 = 16,255.84
- 2x Crestron UC-C100-T Teams kit @ 5,350.41 = 10,700.82
- 2x Q-SYS Core 24f DSP @ 8,295.67 = 16,591.34
- 2x Crestron DM-NAX-BTIO wall plate @ 815.49 = 1,630.98
- 2x Crestron FP-G1-B-T faceplate @ 21.00 = 42.00
- 2x Crestron TSS-880-B scheduling screen @ 1,223.23 = 2,446.46
- 2x Crestron TSW-1080-B touch screen @ 3,186.38 = 6,372.76
- 1x Crestron CP4N control system 3,032.23 (later credited back in variations)
- 2x Netgear M4250 PoE switch @ 3,179.58 = 6,359.16
- 2x Netgear AXC761 SFP @ 121.46 = 242.92
- Freight 100.00, e-waste 59.85, cables/fixings 1,515.30
- Labour and programming 23,320.00

Block 2, extra encoders, matches Option B. Total $11,019.46
- 2x Crestron DM-NVX-E30 encoder @ 1,517.06 = 3,034.12 (note: $171.50 per unit dearer than the same part in block 1)
- Freight 53.20, waste 39.90, cables/fixings 696.24
- Labour 7,196.00

Block 3, camera tracking and audio, matches Option A. Total $45,761.88
- Panasonic AW-SF200Z auto tracking 2,541.01
- Panasonic AW-SF300Z visual preset 1,921.47
- Panasonic AW-SF501Z auto framing 5,309.33
- Precision Computers 4U rack PC i7 3,883.95
- 4x Sennheiser TeamConnect Ceiling 2 mic @ 5,946.65 = 23,786.60
- Q-SYS Dante 16x16 licence 743.84
- Freight 59.85, waste 19.95, cables/fixings 934.08
- Labour 6,561.80

Variations. Net -$10,682.70 (detailed in section 4)

Retention. "Remaining claim balance $4,848.68 to be invoiced": -4,848.68

Check: 120,676.64 + 11,019.46 + 45,761.88 = 177,457.98 = PO 1/2 exactly. Less variations 10,682.70 = 166,775.28. Less retention 4,848.68 = 161,926.60 = bill. Everything ties.

Other observations on the bill
- The removed items do not reconcile to whole units. "Removed camera Crestron decoder" is $4,600.20, which is not a multiple of either E30 price (1,345.56 or 1,517.06) or the NVX-360 price (2,144.14). "Removed Panasonic cameras and licences" is $7,281.07, but 4 cameras cost 16,255.84 and the three licences 9,771.81. Ask Complete AV for the variation breakdown.
- Labour on the bill is $38,141.80 ex GST, about 23% of the supplier cost. Aware's invoice does not split labour out, so you cannot see labour margin separately.
- Coding: equipment to 51100, labour to 51325, freight to 58225 on the bill. Aware side: Phase 2 to 42300, options and variations to 42600. Phase 2 is mostly equipment on the cost side, so revenue and cost land in different categories. Not wrong, but it will make gross margin by category look odd in the P&L.

---

## 7. Cash flow timing

| | Date | Amount inc GST |
|---|---|---|
| Complete AV bill due | 26 Sep 26 (already 3 days overdue) | 178,119.25 |
| Planned materials payment "this week" | w/c 28 Sep 26 | 136,163.27 |
| Aware payment due | 9 Oct 26 | 225,887.40 |

You will be $136k out of pocket for roughly two weeks minimum, and Aware paid its last two invoices in 6 and 7 days, but its June invoice took the full 30. Aware uses a Converga inbox (awaresuper.invoices@converga.com.au), so confirm the invoice has been received and approved there now rather than waiting for the due date. The Xero bill note says hold payment until Aware pays. Paying the materials portion early is a business call, but it contradicts that note, so make sure it is a deliberate decision.

---

## 8. Actions, in order

1. Open PO-00016899 (PO 2/2, $17,700.52) and confirm what it covers. If it is the UE80 upgrade, the job is a 10% job, not 21%, and the upgrade line is sold below cost.
2. Confirm with Aware's Converga inbox that INV-10596 is received and in approval.
3. Ask Complete AV for the variation breakdown behind the three "Removed" lines so the credits can be verified to unit prices.
4. Fix the variation pricing method (cost x 0.85 / 0.77 should be cost x markup) before the retention claim and any further Aware variations are billed.
5. When paying the $136,163, record it in Xero as a part payment of INV41565 so the $41,956 labour balance stays visible.
6. Budget for the $4,848.68 retention and, if applicable, PO 2/2 when reporting this job's margin in the month-end pack.
