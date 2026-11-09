"""The figures quoted in the five department pages' commentary.

The tables on those pages were re-rendered from the pack, so they already carry
the new allocation. The commentary beside them still quoted the old overhead and
the old result. Her wording is kept, only the figures move.

Support's allocated share of the admin pool is now 66,426 of its 79,527 total, and
Events' is 63,710 of its 103,146, which is where the shared overhead figures come
from.
"""
import sys
from pptx import Presentation

EDITS = {
    17: [('Support generated a $29k profit on $243k revenue',
          'Support generated a $22k profit on $243k revenue'),
         ('recorded a $54k loss after $111k in allocated overheads',
          'recorded a $46k loss after $103k in allocated overheads'),
         ('A 10% gross margin limited the final profit to $10k.',
          'A 10% gross margin limited the final profit to $11k.'),
         ('approximately $20k of overheads', 'approximately $21k of overheads')],
    20: [('After $72k of allocated overheads, Support recorded a $29k profit.',
          'After $80k of allocated overheads, Support recorded a $22k profit.'),
         ('Its allocated office wages were $42k',
          'Its allocated office wages were $52k')],
    21: [('Events recorded a $54k loss after $111k of allocated overheads.',
          'Events recorded a $46k loss after $103k of allocated overheads.'),
         ('Approximately $72k relates to shared overhead allocations.',
          'Approximately $64k relates to shared overhead allocations.')],
    22: [('approximately $13k of overheads', 'approximately $12k of overheads')],
    23: [('After $13k of overheads, Integration reported a $10k profit.',
          'After $13k of overheads, Integration reported an $11k profit.')],
    24: [('The department recorded a $20k loss after approximately $20k of overheads.',
          'The department recorded a $21k loss after approximately $21k of overheads.')],
}

p = Presentation(sys.argv[1])
total = 0
for page, pairs in EDITS.items():
    slide = p.slides[page - 1]
    for key in ('Text 5', 'Text 5b'):
        box = next((sh for sh in slide.shapes if sh.name == key), None)
        if box is None:
            continue
        for para in box.text_frame.paragraphs:
            if not para.runs:
                continue
            text = ''.join(r.text for r in para.runs)
            new = text
            for old, repl in pairs:
                new = new.replace(old, repl)
            if new != text:
                para.runs[0].text = new
                for r in list(para.runs)[1:]:
                    r._r.getparent().remove(r._r)
                total += 1
                print(f'  p{page}: {new[:76]}')
print(f'{total} bullets updated')
p.save(sys.argv[2])
