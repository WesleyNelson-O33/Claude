"""Round seven.

1. Delete the fourth quarter revenue and net profit slide, as asked.
2. The benchmark commentary says percentages as percentages. It read "nine
   points of cover", "nine points above" and "up six points". The figures come
   from the pack's own ratio block so the wording and the table agree.
3. Refresh the chart images so the reformatted benchmark table comes through.
"""
import sys, os, re
import openpyxl
from PIL import Image
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, fill_box

HERE = os.path.dirname(os.path.abspath(__file__))
SNIP = os.path.join(HERE, '..', 'snips')
PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
DROP = 'revenue and net profit, fourth quarter'
ROW = {'gp': 229, 'oh_gp': 230, 'ebitda': 231, 'util': 232, 'salary': 234}


def swap_images(pres):
    n = 0
    for slide in pres.slides:
        for sh in list(slide.shapes):
            if 'PICTURE' not in str(sh.shape_type) or not sh.name.startswith('snip '):
                continue
            png = os.path.join(SNIP, sh.name[5:] + '.png')
            if not os.path.exists(png):
                continue
            left, top, width, name = sh.left, sh.top, sh.width, sh.name
            tree = slide.shapes._spTree
            tree.remove(sh._element)
            im = Image.open(png)
            pic = slide.shapes.add_picture(png, left, top, width=width,
                                           height=int(width * im.height / im.width))
            pic.name = name
            tree.remove(pic._element)
            tree.insert(2, pic._element)
            n += 1
    return n


def drop_slide(pres, title_starts):
    for i, s in enumerate(pres.slides):
        if title_of(s).lower().startswith(title_starts):
            lst = pres.slides._sldIdLst
            entry = list(lst)[i]
            pres.part.drop_rel(entry.get(RID))
            lst.remove(entry)
            print('  slide %d removed: %s' % (i + 1, title_of(s)))
            return True
    print('  nothing matched %r, no slide removed' % title_starts)
    return False


def benchmarks(pres):
    ws = openpyxl.load_workbook(PACK, data_only=True)['Xero History']
    gp = ws.cell(ROW['gp'], 3).value * 100
    oh_gp = ws.cell(ROW['oh_gp'], 3).value * 100
    ebitda = ws.cell(ROW['ebitda'], 3).value * 100
    util = ws.cell(ROW['util'], 3).value * 100
    salary = ws.cell(ROW['salary'], 3).value * 100
    salary_up = ws.cell(ROW['salary'], 4).value
    chg = ws.cell(236, 2).value * 100, ws.cell(236, 3).value * 100
    # the benchmark averages are written into the benchmark cell, so read them
    # out of it rather than typing them in again
    pick = lambda r: [float(x) for x in re.findall(r'(\d+(?:\.\d+)?)%',
                                                   ws.cell(r, 5).value or '')]
    gp_bm = pick(ROW['gp'])[0]
    eb_bm = pick(ROW['ebitda'])
    util_bm = pick(ROW['util'])
    cover = 100 - oh_gp

    for i, s in enumerate(pres.slides, 1):
        if 'ratios against industry' not in title_of(s).lower():
            continue
        by = {sh.name: sh for sh in s.shapes}
        fill_box(by['Text 5'], [
            'Gross margin is %.1f%%, above the %.1f%% the integrator benchmark reports for 2025.'
            % (gp, gp_bm),
            'EBITDA margin is %.1f%% against a %.1f%% professional services average, and against '
            'the %.0f%% that benchmark treats as the expected level.' % (ebitda, eb_bm[0], eb_bm[1]),
            'Overheads take %.0f%% of gross profit. At 100%% the company breaks even, so there is '
            '%.1f%% of cover left.' % (oh_gp, cover),
        ])
        fill_box(by['Text 5b'], [
            'Utilisation is %.0f%%, which is the benchmark target and %.1f%% above the %.1f%% average.'
            % (util, util - util_bm[0], util_bm[0]),
            'Salary is %.0f%% of revenue, up %.0f%%, and non-chargeable labour has nearly doubled '
            'from %.0f%% to %.0f%%.' % (salary, salary_up, chg[0], chg[1]),
            'No published benchmark was found for overheads over gross profit or for salary to '
            'revenue on comparable definitions.',
        ])
        s.notes_slide.notes_text_frame.text = '\n'.join(
            ''.join(r.text for r in p.runs)
            for box in (by['Text 5'], by['Text 5b'])
            for p in box.text_frame.paragraphs)
        print('  slide %d benchmark commentary rewritten in %%' % i)
        return


def main(src, out):
    pres = Presentation(src)
    print('  %d chart images refreshed' % swap_images(pres))
    benchmarks(pres)
    drop_slide(pres, DROP)
    pres.save(out)
    print('  %d slides out' % len(pres.slides))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
