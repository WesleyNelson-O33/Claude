"""Embed files in an xlsx as OLE Package objects that open on double click.

Builds a minimal compound file (CFB) per file, holding an \\x01Ole10Native
stream in the classic packager layout, then wires it into the worksheet as an
oleObject with an icon image. Verified by reading the CFB back and comparing
the payload byte for byte with the source.
"""
import os, re, shutil, struct, zipfile

FREE, ENDCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
PACKAGE_CLSID = bytes.fromhex('0c000003000000000000c000000000046'.replace('046', '46'))
# 0003000C-0000-0000-C000-000000000046 little endian for the first three fields
PACKAGE_CLSID = (struct.pack('<IHH', 0x0003000C, 0x0000, 0x0000)
                 + bytes.fromhex('C000000000000046'))
SEC = 512


def ole10native(label, filename, data):
    """The packager payload: label, original path, then the raw file."""
    lab = label.encode('cp1252', 'replace') + b'\x00'
    fn = filename.encode('cp1252', 'replace') + b'\x00'
    temp = filename.encode('cp1252', 'replace') + b'\x00'
    body = (struct.pack('<H', 0x0002) + lab + fn
            + struct.pack('<HH', 0x0000, 0x0003)
            + struct.pack('<I', len(temp)) + temp
            + struct.pack('<I', len(data)) + data)
    return struct.pack('<I', len(body)) + body


def _dir_entry(name, etype, clsid=b'\x00' * 16, child=-1, start=ENDCHAIN, size=0):
    nm = name.encode('utf-16-le') + b'\x00\x00'
    assert len(nm) <= 64
    e = nm + b'\x00' * (64 - len(nm))
    e += struct.pack('<H', len(nm))          # name length including the null
    e += struct.pack('<BB', etype, 1)        # type, colour black
    e += struct.pack('<iii', -1, -1, child)  # left, right, child
    e += clsid
    e += struct.pack('<I', 0)                # state bits
    e += b'\x00' * 16                        # creation, modification time
    e += struct.pack('<I', start)
    e += struct.pack('<Q', size)
    assert len(e) == 128, len(e)
    return e


def build_cfb(payload):
    n_data = (len(payload) + SEC - 1) // SEC
    f = 1
    while True:
        total = f + 1 + n_data
        if (total + 127) // 128 <= f:
            break
        f += 1
    dir_sec = f
    data_start = f + 1

    fat = [FREE] * (f * 128)
    for i in range(f):
        fat[i] = FATSECT
    fat[dir_sec] = ENDCHAIN
    for i in range(n_data):
        fat[data_start + i] = ENDCHAIN if i == n_data - 1 else data_start + i + 1

    hdr = bytes.fromhex('d0cf11e0a1b11ae1') + b'\x00' * 16
    hdr += struct.pack('<HHH', 0x003E, 0x0003, 0xFFFE)
    hdr += struct.pack('<HH', 9, 6)
    hdr += b'\x00' * 6
    hdr += struct.pack('<I', 0)              # directory sector count (0 for v3)
    hdr += struct.pack('<I', f)              # FAT sector count
    hdr += struct.pack('<I', dir_sec)
    hdr += struct.pack('<I', 0)              # transaction signature
    hdr += struct.pack('<I', 4096)           # mini stream cutoff
    hdr += struct.pack('<I', ENDCHAIN) + struct.pack('<I', 0)   # mini FAT
    hdr += struct.pack('<I', ENDCHAIN) + struct.pack('<I', 0)   # DIFAT
    difat = [i if i < f else FREE for i in range(109)]
    hdr += b''.join(struct.pack('<I', x) for x in difat)
    assert len(hdr) == SEC, len(hdr)

    directory = (_dir_entry('Root Entry', 5, PACKAGE_CLSID, child=1)
                 + _dir_entry('\x01Ole10Native', 2, start=data_start,
                              size=len(payload))
                 + _dir_entry('', 0) + _dir_entry('', 0))
    body = payload + b'\x00' * (n_data * SEC - len(payload))
    out = hdr + b''.join(struct.pack('<I', x) for x in fat) + directory + body
    assert len(out) == SEC * (1 + f + 1 + n_data)
    return out


def read_back(blob):
    """Parse our own CFB and return the Ole10Native payload."""
    assert blob[:8] == bytes.fromhex('d0cf11e0a1b11ae1')
    nfat, dirsec = struct.unpack('<II', blob[44:52])
    difat = struct.unpack('<109I', blob[76:512])
    fat = []
    for s in difat[:nfat]:
        off = SEC + s * SEC
        fat += list(struct.unpack('<128I', blob[off:off + SEC]))
    doff = SEC + dirsec * SEC
    ent = blob[doff + 128:doff + 256]
    start, size = struct.unpack('<I', ent[116:120])[0], struct.unpack('<Q', ent[120:128])[0]
    chunks, s = [], start
    while s != ENDCHAIN:
        off = SEC + s * SEC
        chunks.append(blob[off:off + SEC])
        s = fat[s]
    return b''.join(chunks)[:size]


def payload_file_bytes(payload):
    """Pull the embedded file back out of a packager payload."""
    p = 6
    p = payload.index(b'\x00', p) + 1          # label
    p = payload.index(b'\x00', p) + 1          # filename
    p += 4                                      # flags2 + unknown
    tlen = struct.unpack('<I', payload[p:p + 4])[0]
    p += 4 + tlen
    dlen = struct.unpack('<I', payload[p:p + 4])[0]
    p += 4
    return payload[p:p + dlen]


def icon_png(path, label):
    from PIL import Image, ImageDraw, ImageFont
    W, H = 176, 124
    img = Image.new('RGBA', (W, H), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    x0, y0, w, h, fold = 62, 6, 52, 64, 16
    d.polygon([(x0, y0), (x0 + w - fold, y0), (x0 + w, y0 + fold),
               (x0 + w, y0 + h), (x0, y0 + h)], fill=(255, 255, 255),
              outline=(140, 148, 160))
    d.polygon([(x0 + w - fold, y0), (x0 + w, y0 + fold),
               (x0 + w - fold, y0 + fold)], fill=(226, 230, 236),
              outline=(140, 148, 160))
    d.rectangle([x0 + 4, y0 + h - 22, x0 + w - 4, y0 + h - 4],
                fill=(192, 28, 40))
    try:
        f = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', 11)
        fs = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 11)
    except OSError:
        f = fs = ImageFont.load_default()
    d.text((x0 + 10, y0 + h - 20), 'PDF', font=f, fill=(255, 255, 255))
    y = y0 + h + 6
    for line in label.split('\n'):
        tw = d.textbbox((0, 0), line, font=fs)[2]
        d.text(((W - tw) / 2, y), line, font=fs, fill=(36, 44, 57))
        y += 14
    img.save(path)
    return W, H


EMU = 9525  # one pixel at 96 dpi


def embed(xlsx, items, sheet='sheet1.xml'):
    """items: list of dicts with path, label, caption, row, col."""
    tmp = xlsx + '.build'
    shutil.copy(xlsx, tmp)
    with zipfile.ZipFile(tmp) as z:
        parts = {n: z.read(n) for n in z.namelist()}

    rels_name = 'xl/worksheets/_rels/%s.rels' % sheet
    rels = parts.get(rels_name) or (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        b'relationships"></Relationships>')
    rels = rels.decode()
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)] or [0]
    nid = max(used) + 1

    objs = []
    for i, it in enumerate(items, 1):
        raw = open(it['path'], 'rb').read()
        base = os.path.basename(it['path'])
        payload = ole10native(it['label'], base, raw)
        blob = build_cfb(payload)
        got = payload_file_bytes(read_back(blob))
        if got != raw:
            raise SystemExit('round trip failed for %s' % base)
        parts['xl/embeddings/oleObject%d.bin' % i] = blob

        ipath = '%s.icon%d.png' % (tmp, i)
        w, h = icon_png(ipath, it['caption'])
        parts['xl/media/oleimage%d.png' % i] = open(ipath, 'rb').read()
        os.remove(ipath)

        rid_obj, rid_img = 'rId%d' % nid, 'rId%d' % (nid + 1)
        nid += 2
        rels = rels.replace('</Relationships>',
            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/oleObject" Target="../embeddings/'
            'oleObject%d.bin"/>'
            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/image" Target="../media/'
            'oleimage%d.png"/></Relationships>' % (rid_obj, i, rid_img, i))
        objs.append((1024 + i, rid_obj, rid_img, it['row'], it['col'], w, h))

    parts[rels_name] = rels.encode()

    body = []
    for shape, rid_obj, rid_img, row, col, w, h in objs:
        body.append(
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/'
            'markup-compatibility/2006"><mc:Choice Requires="x14">'
            '<oleObject progId="Package" shapeId="%d" r:id="%s">'
            '<objectPr defaultSize="0" r:id="%s"><anchor moveWithCells="1">'
            '<from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></from>'
            '<to><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></to>'
            '</anchor></objectPr></oleObject></mc:Choice>'
            '<mc:Fallback><oleObject progId="Package" shapeId="%d" r:id="%s"/>'
            '</mc:Fallback></mc:AlternateContent>'
            % (shape, rid_obj, rid_img, col, row, col, w * EMU, row, h * EMU,
               shape, rid_obj))

    block = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/'
             'markup-compatibility/2006"><mc:Choice Requires="x14">'
             '<oleObjects>' + ''.join(body) + '</oleObjects></mc:Choice>'
             '<mc:Fallback/></mc:AlternateContent>')

    s = parts['xl/worksheets/%s' % sheet].decode()
    m = re.match(r'(.*?<worksheet\b[^>]*?)(/?>)', s, re.S)
    head = m.group(1)
    for ns, uri in (('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/'
                            'spreadsheetDrawing'),
                    ('mc', 'http://schemas.openxmlformats.org/markup-'
                           'compatibility/2006'),
                    ('x14', 'http://schemas.microsoft.com/office/spreadsheetml/'
                            '2009/9/main')):
        if 'xmlns:%s=' % ns not in head:
            head += ' xmlns:%s="%s"' % (ns, uri)
    if 'mc:Ignorable' not in head:
        head += ' mc:Ignorable="x14"'
    s = head + m.group(2) + s[m.end():]
    s = s.replace('</worksheet>', block + '</worksheet>')
    parts['xl/worksheets/%s' % sheet] = s.encode()

    ct = parts['[Content_Types].xml'].decode()
    for ext, mime in (('bin', 'application/vnd.openxmlformats-officedocument.'
                              'oleObject'),
                      ('png', 'image/png')):
        if 'Extension="%s"' % ext not in ct:
            ct = ct.replace('<Default', '<Default Extension="%s" ContentType="%s"/>'
                            '<Default' % (ext, mime), 1)
    parts['[Content_Types].xml'] = ct.encode()

    with zipfile.ZipFile(xlsx, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)
    os.remove(tmp)
    print('embedded %d object(s) in %s' % (len(items), xlsx))
