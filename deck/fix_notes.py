"""Replace the template's prompt notes with real speaking notes, and tidy wording."""
import os
from pptx import Presentation
from pptx.util import Emu

D = os.path.dirname(os.path.abspath(__file__))
f = f'{D}/CTS Monthly Business Review - September 2026.pptx'
prs = Presentation(f)
S = list(prs.slides)

NOTES = {
 1: "September 2026. Every figure traces to the Financial Controller Pack v19, the Xero general "
    "ledger dump inside it, or a Xero report named on the slide.",
 4: "Do not let the room celebrate the 148k beat. Lead with break even: we need 549k of revenue "
    "a month and we are running at 541k. The miss is margin, not cost, and overheads are 48k "
    "under budget. Expect the question on the full year: revenue lands, profit does not.",
 5: "The point is the overhead key, not the departments. Events carries 87.5 per cent of "
    "premises on 13 per cent of revenue. Say clearly that the key is fixed for the financial "
    "year and must not be changed mid year, but it should be tested for FY28.",
 6: "Two bars clear zero. If pushed on Integration, be straight that its bar is one project and "
    "the underlying unit is about 60k a quarter.",
 7: "43 per cent contracted. The 487k shortfall on Integration and Consulting non-contract is "
    "not recoverable, so ask for a reforecast rather than defending the budget. Flag the two "
    "unmapped budget lines yourself before anyone finds them.",
 9: "Concentration is the message, not the ranking. Deloitte is roughly the whole EBITDA base. "
    "Aware Super is a fit-out, not a relationship.",
 10: "Same data as the table. Use it to show how fast the book drops after the top five.",
 11: "This is the cash slide. 309k is sitting in debtors and net debt is 723k. That is the "
     "fastest lever available and it needs no trading improvement. Name the two old invoices "
     "and ask for a decision, chase or provide.",
 12: "Growth is three new names, not the base. If asked about ASX, it is down a third and that "
     "is a relationship question for sales, not a market one.",
 14: "Be straight that Support recording 11 non-chargeable hours is not credible. Do not defend "
     "the 99.5 per cent. The fix is reporting, and it belongs to Ops.",
 15: "Revenue per hour is not a charge rate, say so before anyone treats Events as one. The "
     "Integration line is the real finding: 235k billed on 9.5 hours means the job was bought "
     "in, which is why the margin is 10 per cent.",
 17: "Four asks, each with an owner and a date. Do not leave the room without them agreed.",
}
for n, txt in NOTES.items():
    tf = S[n-1].notes_slide.notes_text_frame
    if tf is None:
        print("  slide", n, "has no notes box, skipped")
        continue
    tf.text = txt

# "$0.83m", not "0.83m"
for sh in S[10].shapes:
    if sh.has_text_frame and Emu(sh.left).inches > 6.8 and Emu(sh.height).inches > 2:
        p0 = sh.text_frame.paragraphs[0]
        if p0.runs and p0.runs[0].text.startswith('Receivables are 0.83m'):
            p0.runs[0].text = p0.runs[0].text.replace('Receivables are 0.83m',
                                                      'Receivables are $0.83m')
prs.save(f)
print('notes replaced on', len(NOTES), 'slides')
