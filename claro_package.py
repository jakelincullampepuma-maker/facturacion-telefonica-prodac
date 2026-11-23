"""Añade hojas al XLSX sin reescribir Factura Detallada ni sus recursos."""
import copy
import io
import posixpath as pp
import re
import zipfile
from xml.etree import ElementTree as ET

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'


def xml(root, original=b''):
    # iterparse devuelve (evento, (prefijo, URI)).
    namespaces = dict(pair for _, pair in ET.iterparse(io.BytesIO(original), events=('start-ns',))) if original else {}
    for prefix, uri in namespaces.items():
        if not re.fullmatch(r'ns\d+', prefix):
            ET.register_namespace(prefix, uri)
    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    # Conservar declaraciones referidas por mc:Ignorable/Requires aunque no
    # aparezcan en nombres de elementos; ElementTree las omitiría.
    additions = []
    for prefix, uri in namespaces.items():
        name = 'xmlns' + (':' + prefix if prefix else '')
        if (name + '=').encode() not in data:
            additions.append(f' {name}="{uri}"')
    if additions:
        start = data.index(b'>', data.index(b'?>') + 2)
        data = data[:start] + ''.join(additions).encode() + data[start:]
    return data


def relpath(part):
    return pp.join(pp.dirname(part), '_rels', pp.basename(part) + '.rels')


def target(part, value):
    return value.lstrip('/') if value.startswith('/') else pp.normpath(pp.join(pp.dirname(part), value))


def dependencies(parts, seeds):
    found = set()
    todo = list(seeds)
    while todo:
        part = todo.pop()
        if part in found or part not in parts:
            continue
        found.add(part)
        rel = relpath(part)
        if rel in parts:
            found.add(rel)
            for item in ET.fromstring(parts[rel]):
                if item.get('TargetMode') != 'External':
                    todo.append(target(part, item.get('Target')))
    return found


def merge_styles(original, generated):
    old, new = ET.fromstring(original), ET.fromstring(generated)
    maps = {}
    for section in ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs']:
        source = new.find(f'{{{S}}}{section}')
        dest = old.find(f'{{{S}}}{section}')
        if source is None:
            continue
        if dest is None:
            dest = ET.Element(f'{{{S}}}{section}')
            old.insert(0 if section == 'numFmts' else len(old), dest)
        mapping = {}
        for index, child in enumerate(source):
            child = copy.deepcopy(child)
            if section == 'numFmts':
                ident = max([163] + [int(x.get('numFmtId')) for x in dest]) + 1
                mapping[int(child.get('numFmtId'))] = ident
                child.set('numFmtId', str(ident))
            else:
                mapping[index] = len(dest)
                for attr, group in [('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders'),
                                    ('numFmtId', 'numFmts'), ('xfId', 'cellStyleXfs')]:
                    if attr in child.attrib and int(child.get(attr)) in maps.get(group, {}):
                        child.set(attr, str(maps[group][int(child.get(attr))]))
            dest.append(child)
        dest.set('count', str(len(dest)))
        maps[section] = mapping
    return xml(old, original), maps['cellXfs']


def preserve_billing(original_bytes, generated_bytes):
    with zipfile.ZipFile(io.BytesIO(original_bytes)) as archive:
        parts = {p.filename: archive.read(p.filename) for p in archive.infolist()}
    with zipfile.ZipFile(io.BytesIO(generated_bytes)) as archive:
        new = {p.filename: archive.read(p.filename) for p in archive.infolist()}
    workbook = ET.fromstring(parts['xl/workbook.xml'])
    relations = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    by_id = {r.get('Id'): r for r in relations}
    sheets = workbook.find(f'{{{S}}}sheets')
    factura = next(s for s in sheets if s.get('name') == 'Factura Detallada')
    fact_id = factura.get(f'{{{R}}}id')
    fact_path = target('xl/workbook.xml', by_id[fact_id].get('Target'))
    original_sheet = parts[fact_path]
    keep = dependencies(parts, [fact_path])
    remove_roots = []
    retained_cache_ids = set()
    for item in list(relations):
        kind = item.get('Type').rsplit('/', 1)[-1]
        path = target('xl/workbook.xml', item.get('Target'))
        if ((kind == 'worksheet' and item.get('Id') != fact_id) or kind == 'calcChain'
                or (kind == 'pivotCacheDefinition' and path not in keep)):
            remove_roots.append(path)
            relations.remove(item)
        elif kind == 'pivotCacheDefinition':
            retained_cache_ids.add(item.get('Id'))
        else:
            keep |= dependencies(parts, [path])
    removed = dependencies(parts, remove_roots) - keep
    old_index = list(sheets).index(factura)
    max_sheet_id = max(int(s.get('sheetId')) for s in sheets)
    for item in list(sheets):
        if item is not factura:
            sheets.remove(item)
    names = workbook.find(f'{{{S}}}definedNames')
    if names is not None:
        for item in list(names):
            local = item.get('localSheetId')
            if local is not None:
                if int(local) != old_index:
                    names.remove(item)
                else:
                    item.set('localSheetId', '0')
            elif re.search(r"(?:'?(?:Hoja1|Hoja2)'?!)", item.text or ''):
                names.remove(item)
    caches = workbook.find(f'{{{S}}}pivotCaches')
    if caches is None:
        caches = ET.Element(f'{{{S}}}pivotCaches')
        ext = workbook.find(f'{{{S}}}extLst')
        workbook.insert(list(workbook).index(ext) if ext is not None else len(workbook), caches)
    for item in list(caches):
        if item.get(f'{{{R}}}id') not in retained_cache_ids:
            caches.remove(item)
    mapping = {}
    for part in new:
        if part.startswith(('xl/worksheets/', 'xl/tables/', 'xl/pivotTables/', 'xl/pivotCache/')):
            mapping[part] = pp.join(pp.dirname(part), 'claro_' + pp.basename(part))
    # Los nombres de los .rels deben corresponder a los nombres nuevos.
    for part in list(mapping):
        if '/_rels/' in part:
            owner = pp.join(pp.dirname(pp.dirname(part)), pp.basename(part)[:-5])
            mapping[part] = relpath(mapping[owner])
    style_bytes, style_map = merge_styles(parts['xl/styles.xml'], new['xl/styles.xml'])
    styles = {'xl/styles.xml': style_bytes}
    new_book = ET.fromstring(new['xl/workbook.xml'])
    new_rels = {x.get('Id'): x for x in ET.fromstring(new['xl/_rels/workbook.xml.rels'])}
    next_rel = max([0] + [int(r.get('Id')[3:]) for r in relations if re.fullmatch(r'rId\d+', r.get('Id'))]) + 1
    cache_map = {}
    max_cache = max([0] + [int(x.get('cacheId')) for x in caches])
    for index, item in enumerate(new_book.find(f'{{{S}}}pivotCaches'), 1):
        cache_map[item.get('cacheId')] = str(max_cache + index)
        new_id = f'rId{next_rel}'; next_rel += 1
        part = target('xl/workbook.xml', new_rels[item.get(f'{{{R}}}id')].get('Target'))
        ET.SubElement(relations, f'{{{P}}}Relationship', Id=new_id, Type=R + '/pivotCacheDefinition', Target=pp.relpath(mapping[part], 'xl'))
        ET.SubElement(caches, f'{{{S}}}pivotCache', cacheId=str(max_cache + index), **{f'{{{R}}}id': new_id})
    for index, item in enumerate(new_book.find(f'{{{S}}}sheets'), 1):
        new_id = f'rId{next_rel}'; next_rel += 1
        part = target('xl/workbook.xml', new_rels[item.get(f'{{{R}}}id')].get('Target'))
        ET.SubElement(relations, f'{{{P}}}Relationship', Id=new_id, Type=R + '/worksheet', Target=pp.relpath(mapping[part], 'xl'))
        ET.SubElement(sheets, f'{{{S}}}sheet', name=item.get('name'), sheetId=str(max_sheet_id + index), **{f'{{{R}}}id': new_id})
    for view in workbook.findall(f'.//{{{S}}}workbookView'):
        view.set('activeTab', '0'); view.set('firstSheet', '0')
    calc = workbook.find(f'{{{S}}}calcPr')
    if calc is not None:
        calc.set('fullCalcOnLoad', '1'); calc.set('calcMode', 'auto')
    replacements = {**styles, 'xl/workbook.xml': xml(workbook, parts['xl/workbook.xml']),
                    'xl/_rels/workbook.xml.rels': xml(relations, parts['xl/_rels/workbook.xml.rels'])}
    for part, destination in mapping.items():
        root = ET.fromstring(new[part])
        if part.endswith('.rels'):
            owner = pp.join(pp.dirname(pp.dirname(part)), pp.basename(part)[:-5])
            for item in root:
                old_target = target(owner, item.get('Target'))
                if old_target in mapping:
                    item.set('Target', pp.relpath(mapping[old_target], pp.dirname(mapping[owner])))
        elif part.startswith('xl/worksheets/'):
            for cell in root.findall(f'.//{{{S}}}c'):
                cell.set('s', str(style_map[int(cell.get('s', 0))]))
        elif part.startswith('xl/pivotTables/'):
            root.set('cacheId', cache_map[root.get('cacheId')])
        replacements[destination] = xml(root, new[part])
    types = ET.fromstring(parts['[Content_Types].xml'])
    for item in list(types):
        if item.get('PartName', '').lstrip('/') in removed:
            types.remove(item)
    for item in ET.fromstring(new['[Content_Types].xml']):
        part = item.get('PartName', '').lstrip('/')
        if part in mapping:
            ET.SubElement(types, f'{{{C}}}Override', PartName='/' + mapping[part], ContentType=item.get('ContentType'))
    replacements['[Content_Types].xml'] = xml(types, parts['[Content_Types].xml'])
    if 'docProps/app.xml' in parts:
        app = ET.fromstring(parts['docProps/app.xml'])
        for item in app.iter():
            if item.tag.endswith('}TitlesOfParts'):
                vector = item[0]
                child_tag = vector[0].tag if len(vector) else '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpstr'
                vector.clear(); vector.set('size', '3'); vector.set('baseType', 'lpstr')
                for title in ['Factura Detallada', 'Hoja2', 'Hoja1']:
                    ET.SubElement(vector, child_tag).text = title
            if item.tag.endswith('}HeadingPairs'):
                for value in item.iter():
                    if value.tag.endswith('}i4'):
                        value.text = '3'
        replacements['docProps/app.xml'] = xml(app, parts['docProps/app.xml'])
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as archive:
        for part, data in parts.items():
            if part not in removed:
                archive.writestr(part, replacements.pop(part, data))
        for part, data in replacements.items():
            archive.writestr(part, data)
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as archive:
        if archive.read(fact_path) != original_sheet:
            raise ValueError('La hoja Factura Detallada debe permanecer intacta.')
    return out.getvalue()
