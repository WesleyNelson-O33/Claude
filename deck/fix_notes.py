"""Replace the template's prompt notes with real speaking notes, and tidy wording."""
import os
from pptx import Presentation
from pptx.util import Emu

D = os.path.dirname(os.path.abspath(__file__))
f = f'{D}/CTS Monthly Business Review - September 2026.pptx'
prs = Presentation(f)
S = list(prs.slides)

NOTES = {
 1: "September 2026 review. Every figure comes from the Financial Controller Pack v18, the "
    "Xero general ledger dump inside it, or a Xero report named on the slide.",
 4: "Lead with the beat, then take it away: the month is ahead, the year is not. The 148k beat "
    "is one Integration project. Overheads are on budget so do not let the conversation go there.",
 5: "Support is the only department covering its overhead. Expect a question on why Events "
    "carries 111k: it is the Office Dept key giving it 87.5 per cent of premises cost.",
 6: "Point at the two bars that clear zero. If asked about the Office Dept key, the answer is "
    "that it is fixed for the financial year and is reviewed at FY28.",
 7: "Half the book is non-contract. The two budget lines with no budget mapped are a pack fix, "
    "not a trading issue, but say so before someone finds it.",
 9: "Concentration is the message. Top three are 49 per cent. Aware Super is a project, not a "
    "customer relationship yet.",
 10: "Same data as the table, shown so the drop after the top five is obvious.",
 11: "DSO 47 days against a 30 day target. Two invoices over 90 days need a decision this month, "
     "chase or provide. Aware Super is current but it is one 226k invoice.",
 12: "Growth is three new customers, not the base. If they do not repeat, FY28 starts 368k behind.",
 14: "Be straight that Support showing 100 per cent is a recording gap. Video and Integration "
     "do not report hours at all.",
 15: "Revenue per hour is not a charge rate. The pack does not hold charge rates. Say that before "
     "anyone treats the Events figure as a rate.",
 17: "Agree owner and date on each line before the meeting closes.",
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
