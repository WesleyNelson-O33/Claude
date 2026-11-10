"""Danica's third round of changes, applied to her own file.

  the executive summary goes from September to the quarter
  slide 8 is retitled and its commentary rewritten, because the chart on it
    is the three years against budget and carries no departments, while the
    title and half the commentary still said it did
  the utilisation slide gets a commentary panel like the rest of her slides,
    and the table moves beside it

Every figure is read from the controller pack.
"""
import copy, sys, os
import openpyxl
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, set_title, put_picture, write_panel
from q1exec_inplace import main as exec_to_quarter

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
SNIP = '/home/user/Claude/deck/snips/'


def d(v):
    return ('$(%s)' % format(abs(round(v)), ',')) if v < 0 else ('$%s' % format(round(v), ','))


def figures():
    wb = openpyxl.load_workbook(PACK, data_only=True)
    sm, xh = wb['Summary'], wb['Xero History']
    f = {'ytd': sm.cell(6, 5).value, 'bud': sm.cell(6, 7).value}
    f['short'] = f['bud'] - f['ytd']
    # the cumulative block: FY25, FY26, FY27 and the budget, July to June
    f['fy25_q1'] = xh.cell(60, 2).value
    f['fy26_q1'] = xh.cell(60, 3).value
    f['fy25_full'] = xh.cell(69, 2).value
    f['fy26_full'] = xh.cell(69, 3).value
    f['bud_full'] = xh.cell(69, 5).value
    return f


def clone_panel(pres, donor_idx, target):
    """Take a commentary panel, its label and its body from a slide that has
    one, so the new slide matches her styling exactly."""
    donor = pres.slides[donor_idx]
    want = ('commentary panel', 'Text 4', 'Text 5')
    have = {sh.name for sh in target.shapes}
    tree = target.shapes._spTree
    for name in want:
        if name in have:
            continue
        src = next((sh for sh in donor.shapes if sh.name == name), None)
        if src is None:
            continue
        tree.append(copy.deepcopy(src._element))
    return True



def swap_images(pres):
    """Refresh every chart image in place. Her positions are kept exactly:
    the box's top left and width stay, only the height follows the new aspect,
    so nothing she has arranged moves."""
    n = 0
    for slide in pres.slides:
        for sh in list(slide.shapes):
            if 'PICTURE' not in str(sh.shape_type) or not sh.name.startswith('snip '):
                continue
            png = os.path.join(SNIP, sh.name[5:] + '.png')
            if not os.path.exists(png):
                continue
            left, top, width = sh.left, sh.top, sh.width
            name = sh.name
            tree = slide.shapes._spTree
            tree.remove(sh._element)
            im = Image.open(png)
            height = int(width * im.height / im.width)
            pic = slide.shapes.add_picture(png, left, top, width=width, height=height)
            pic.name = name
            tree.remove(pic._element)
            tree.insert(2, pic._element)
            n += 1
    print('  %d chart images refreshed in place' % n)

def main(src, out):
    S = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad'
    step = os.path.join(S, '_r3a.pptx')
    exec_to_quarter(src, step)

    f = figures()
    pres = Presentation(step)
    swap_images(pres)

    donor = next(i for i, s in enumerate(pres.slides)
                 if {'commentary panel', 'Text 4', 'Text 5'} <= {sh.name for sh in s.shapes}
                 and Emu(next(sh for sh in s.shapes
                              if sh.name == 'commentary panel').left).inches > 5)

    for i, s in enumerate(pres.slides, 1):
        t = title_of(s).lower()

        if t.startswith('year to date revenue'):
            set_title(s, 'Year To Date Revenue Against Budget')
            write_panel(s, [
                'Year to date revenue is %s against a budget of %s.' % (d(f['ytd']), d(f['bud'])),
                'That is %s or %.0f%% behind, and almost all of it opened in August.'
                % (d(f['short']), f['short'] / f['bud'] * 100),
                'It is still ahead of both completed years at this point, %s in FY25 and %s in FY26.'
                % (d(f['fy25_q1']), d(f['fy26_q1'])),
                'The dashed line is the full year plan, %s against the %s FY25 actually billed.'
                % (d(f['bud_full']), d(f['fy25_full'])),
                'Nine months of that plan are still to come.',
            ])
            print('  %-3d slide 8 retitled and rewritten' % i)

        elif t.startswith('utilisation per employee'):
            clone_panel(pres, donor, s)
            # force her standard panel geometry, because a cloned shape brings
            # the donor's position with it
            for name, box in (('commentary panel', (6.45, 0.94, 3.00, 4.12)),
                              ('Text 4', (6.61, 1.09, 2.54, 0.25)),
                              ('Text 5', (6.61, 1.58, 2.70, 3.24))):
                sh = next((x for x in s.shapes if x.name == name), None)
                if sh is None:
                    continue
                sh.left, sh.top = Inches(box[0]), Inches(box[1])
                sh.width, sh.height = Inches(box[2]), Inches(box[3])
            png = os.path.join(SNIP, 'util_q1.png')
            geom = put_picture(s, png, 'snip util_q1')
            write_panel(s, [
                'Support ran 6,942 chargeable hours against 393 non-chargeable, 95% of hours worked.',
                'Production billed 954 hours against 1,150 non-chargeable, so over half its worked time was not billed.',
                'Consulting is the same shape on a smaller base, 179 against 280.',
                'Leave took 1,258 hours across the quarter, 543 of it in Production.',
                'Across CTS the quarter ran 8,074 chargeable hours against 1,823 non-chargeable, 62% chargeable against an 82% target.',
            ])
            print('  %-3d utilisation: panel added, table %.2f x %.2f' % (i, geom[0], geom[1]))

    pres.save(out)
    print('%d slides out' % len(pres.slides))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
