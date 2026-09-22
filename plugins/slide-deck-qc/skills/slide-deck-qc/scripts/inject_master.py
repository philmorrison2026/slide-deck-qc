#!/usr/bin/env python3
"""Copy a slide master (its layouts, theme and media) from a source deck
into a target deck that doesn't have it, so the Syndeio layouts appear in
PowerPoint's New Slide gallery."""
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

PR = "http://schemas.openxmlformats.org/package/2006/relationships"
OR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def rels_path(part):
    d, f = os.path.split(part)
    return "%s/_rels/%s.rels" % (d, f)


def read_rels(z, part):
    try:
        return etree.fromstring(z.read(rels_path(part)))
    except KeyError:
        return None


def resolve(base, target):
    return os.path.normpath(os.path.join(os.path.dirname(base),
                                         target)).replace("\\", "/")


def next_index(names, pattern):
    n = 0
    for name in names:
        m = re.match(pattern, name)
        if m:
            n = max(n, int(m.group(1)))
    return n + 1


def inject(src_path, master_part, tgt_path, out_path):
    zs = zipfile.ZipFile(src_path)
    zt = zipfile.ZipFile(tgt_path)
    tgt_names = set(zt.namelist())

    # ---- collect the source parts we need -------------------------------
    wanted = {master_part}
    mrels = read_rels(zs, master_part)
    layouts, theme = [], None
    for rel in mrels:
        kind = rel.get("Type").rsplit("/", 1)[-1]
        part = resolve(master_part, rel.get("Target"))
        if kind == "slideLayout":
            layouts.append(part)
        elif kind == "theme":
            theme = part
        wanted.add(part)

    media = set()
    for part in list(wanted):
        rels = read_rels(zs, part)
        if rels is None:
            continue
        for rel in rels:
            if rel.get("TargetMode") == "External":
                continue
            p = resolve(part, rel.get("Target"))
            if "/media/" in p:
                media.add(p)
    wanted |= media

    # ---- work out new names in the target -------------------------------
    rename = {}
    mi = next_index(tgt_names, r"ppt/slideMasters/slideMaster(\d+)\.xml")
    rename[master_part] = "ppt/slideMasters/slideMaster%d.xml" % mi

    li = next_index(tgt_names, r"ppt/slideLayouts/slideLayout(\d+)\.xml")
    for k, lay in enumerate(layouts):
        rename[lay] = "ppt/slideLayouts/slideLayout%d.xml" % (li + k)

    ti = next_index(tgt_names, r"ppt/theme/theme(\d+)\.xml")
    rename[theme] = "ppt/theme/theme%d.xml" % ti

    xi = next_index(tgt_names, r"ppt/media/image(\d+)\.\w+")
    for k, m in enumerate(sorted(media)):
        ext = m.rsplit(".", 1)[-1]
        rename[m] = "ppt/media/image%d.%s" % (xi + k, ext)

    # ---- build the output package ---------------------------------------
    tmp = out_path + ".tmp"
    zo = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)

    pres_rels_path = "ppt/_rels/presentation.xml.rels"
    skip = {pres_rels_path, "ppt/presentation.xml", "[Content_Types].xml"}

    for n in zt.namelist():
        if n not in skip:
            zo.writestr(n, zt.read(n))

    # copy the wanted source parts under their new names, fixing their rels
    for part in wanted:
        new = rename[part]
        data = zs.read(part)
        zo.writestr(new, data)

        rels = read_rels(zs, part)
        if rels is None:
            continue
        for rel in rels:
            if rel.get("TargetMode") == "External":
                continue
            old = resolve(part, rel.get("Target"))
            if old in rename:
                rel.set("Target", os.path.relpath(
                    rename[old], os.path.dirname(new)).replace("\\", "/"))
        zo.writestr(rels_path(new), etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True))

    # ---- content types ---------------------------------------------------
    ct = etree.fromstring(zt.read("[Content_Types].xml"))
    have = {o.get("PartName") for o in ct}
    types = {
        "slideMaster": "application/vnd.openxmlformats-officedocument."
                       "presentationml.slideMaster+xml",
        "slideLayout": "application/vnd.openxmlformats-officedocument."
                       "presentationml.slideLayout+xml",
        "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    }
    defaults = {d.get("Extension", "").lower() for d in ct
                if d.tag.endswith("Default")}
    for old, new in rename.items():
        if "/media/" in new:
            ext = new.rsplit(".", 1)[-1].lower()
            if ext not in defaults:
                el = etree.SubElement(ct, "{%s}Default" % CT)
                el.set("Extension", ext)
                el.set("ContentType", "image/%s"
                       % ("jpeg" if ext in ("jpg", "jpeg") else ext))
                defaults.add(ext)
            continue
        kind = ("slideMaster" if "slideMaster" in new
                else "slideLayout" if "slideLayout" in new else "theme")
        pn = "/" + new
        if pn not in have:
            el = etree.SubElement(ct, "{%s}Override" % CT)
            el.set("PartName", pn)
            el.set("ContentType", types[kind])
    zo.writestr("[Content_Types].xml", etree.tostring(
        ct, xml_declaration=True, encoding="UTF-8", standalone=True))

    # ---- presentation rels + sldMasterIdLst -------------------------------
    prels = etree.fromstring(zt.read(pres_rels_path))
    used = {int(r.get("Id")[3:]) for r in prels
            if r.get("Id", "").startswith("rId")
            and r.get("Id")[3:].isdigit()}
    new_rid = "rId%d" % (max(used) + 1)
    el = etree.SubElement(prels, "{%s}Relationship" % PR)
    el.set("Id", new_rid)
    el.set("Type", OR + "/slideMaster")
    el.set("Target", rename[master_part].replace("ppt/", ""))
    zo.writestr(pres_rels_path, etree.tostring(
        prels, xml_declaration=True, encoding="UTF-8", standalone=True))

    pres = etree.fromstring(zt.read("ppt/presentation.xml"))
    lst = pres.find("{%s}sldMasterIdLst" % P)
    # Master and layout IDs share one global namespace. Scan every master
    # and every layout in the package, or the new master's id can collide
    # with an existing layout id and PowerPoint rejects the file.
    ids = [int(m.get("id")) for m in lst]
    for nm in list(zt.namelist()) + list(rename.values()):
        if "slideMaster" in nm and nm.endswith(".xml") \
                and "_rels" not in nm:
            try:
                data = zt.read(nm) if nm in tgt_names else zs.read(
                    [k for k, v in rename.items() if v == nm][0])
                el = etree.fromstring(data)
                for lid in el.iter("{%s}sldLayoutId" % P):
                    ids.append(int(lid.get("id")))
            except Exception:
                pass
    m = etree.SubElement(lst, "{%s}sldMasterId" % P)
    m.set("id", str(max(ids) + 1))
    m.set("{%s}id" % R, new_rid)
    zo.writestr("ppt/presentation.xml", etree.tostring(
        pres, xml_declaration=True, encoding="UTF-8", standalone=True))

    zo.close()
    shutil.move(tmp, out_path)
    print("injected %s (+%d layouts, %d media) -> %s"
          % (master_part, len(layouts), len(media), out_path))


if __name__ == "__main__":
    inject(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
