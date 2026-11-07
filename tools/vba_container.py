"""Build a source-only MS-OVBA container; Excel must compile it on first open.

No compiled cache or third-party/operational VBA project is copied. Implements
MS-OVBA compression and a small CFB v3 writer for this demo's streams.
"""
import struct
from pathlib import Path


def compressed(data):
    """MS-OVBA chunks represent 4096 source bytes except the final chunk."""
    out = bytearray(b'\x01')
    for offset in range(0, len(data), 4096):
        chunk = data[offset:offset + 4096]
        payload = bytearray(); position = 0
        while position < len(chunk):
            flag_at = len(payload); payload.append(0); flags = 0
            for bit in range(8):
                if position >= len(chunk):
                    break
                offset_bits = max(4, (position - 1).bit_length())
                length_bits = 16 - offset_bits
                max_length = min((1 << length_bits) + 2, len(chunk) - position)
                best_length = 0; best_offset = 0
                if max_length >= 3:
                    candidate = chunk.rfind(chunk[position:position + 3], max(0, position - (1 << offset_bits)), position)
                    attempts = 0
                    while candidate >= 0 and attempts < 128:
                        length = 3
                        while length < max_length and chunk[candidate + length] == chunk[position + length]:
                            length += 1
                        if length > best_length:
                            best_length = length; best_offset = position - candidate
                        if length == max_length:
                            break
                        candidate = chunk.rfind(chunk[position:position + 3], max(0, position - (1 << offset_bits)), candidate)
                        attempts += 1
                if best_length >= 3:
                    flags |= 1 << bit
                    payload += struct.pack('<H', ((best_offset - 1) << length_bits) | (best_length - 3))
                    position += best_length
                else:
                    payload.append(chunk[position]); position += 1
            payload[flag_at] = flags
        if len(payload) <= 4096:
            out += struct.pack('<H', 0xB000 | (len(payload) - 1)) + payload
        else:
            if len(chunk) != 4096:
                raise ValueError('Incompressible short chunk is unsupported by this small writer')
            out += b'\xff\x3f' + chunk
    return bytes(out)


def rec(kind, value):
    return struct.pack('<HI', kind, len(value)) + value


def u16(value):
    return struct.pack('<H', value)


def u32(value):
    return struct.pack('<I', value)


def protection(data):
    """MS-OVBA project-property obfuscation, with no password/protection."""
    project_id = '{F2814620-A1EE-49CC-97B1-ABCDEF123456}'
    seed = 0x22
    key = sum(project_id.encode('ascii')) & 255
    version_enc = seed ^ 2; key_enc = seed ^ key
    result = bytearray([seed, version_enc, key_enc])
    previous_plain, previous_enc, second_enc = key, key_enc, version_enc
    plain = bytes((seed & 6) >> 1) + struct.pack('<I', len(data)) + data
    for value in plain:
        encoded = value ^ ((second_enc + previous_plain) & 255)
        result.append(encoded)
        previous_plain, second_enc, previous_enc = value, previous_enc, encoded
    return result.hex().upper()


def make_project(modules):
    # Required project records, followed by registered Excel reference.
    directory = rec(1, u32(1)) + rec(2, u32(0x409)) + rec(0x14, u32(0x409))
    directory += rec(3, u16(1252)) + rec(4, b'HotlistDemo')
    directory += rec(5, b'') + rec(0x40, b'') + rec(6, b'') + rec(0x3D, b'')
    directory += rec(7, u32(0)) + rec(8, u32(0))
    directory += struct.pack('<HIIH', 9, 4, 1, 0) + rec(0xC, b'') + rec(0x3C, b'')
    reference = b'*\\G{00020813-0000-0000-C000-000000000046}#1.9#0##Microsoft Excel 16.0 Object Library'
    directory += rec(0x16, b'Excel') + rec(0x3E, 'Excel'.encode('utf-16le'))
    directory += rec(0xD, u32(len(reference)) + reference + u32(0) + u16(0))
    directory += rec(0xF, u16(len(modules))) + rec(0x13, u16(0))
    streams = {}
    project = ['ID="{F2814620-A1EE-49CC-97B1-ABCDEF123456}"']
    wm = b''
    for name, (source, document) in modules.items():
        encoded = name.encode('ascii'); wide = name.encode('utf-16le')
        directory += rec(0x19, encoded) + rec(0x47, wide)
        directory += rec(0x1A, encoded) + rec(0x32, wide)
        directory += rec(0x1C, b'') + rec(0x48, b'') + rec(0x31, u32(0))
        directory += rec(0x1E, u32(0)) + rec(0x2C, u16(0))
        directory += rec(0x22 if document else 0x21, b'') + rec(0x2B, b'')
        source = source.replace('\r\n', '\n').replace('\n', '\r\n').encode('cp1252')
        streams['VBA/' + name] = compressed(source)
        project.append(('Document=' + name + '/&H00000000') if document else 'Module=' + name)
        wm += encoded + b'\0' + wide + b'\0\0'
    directory += rec(0x10, b'')
    project += ['Name="HotlistDemo"', 'HelpContextID="0"', 'VersionCompatible32="393222000"', 'CMG="' + protection(bytes(4)) + '"', 'DPB="' + protection(bytes(1)) + '"', 'GC="' + protection(bytes([255])) + '"', '', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '']
    streams['PROJECT'] = '\r\n'.join(project).encode('ascii')
    streams['PROJECTwm'] = wm + b'\0\0'
    streams['VBA/dir'] = compressed(directory)
    # MS-OVBA header with invalidated performance cache version; source recompiles.
    streams['VBA/_VBA_PROJECT'] = bytes.fromhex('cc61ffff000000')
    return compound(streams)


def compound(streams):
    free, end, fatmark = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
    nodes = [{'name': 'Root Entry', 'type': 5, 'left': free, 'right': free, 'child': free, 'start': end, 'size': 0, 'color': 1}]
    nodes.append({'name': 'VBA', 'type': 1, 'left': free, 'right': free, 'child': free, 'start': end, 'size': 0, 'color': 1})
    for path, data in streams.items():
        nodes.append({'name': path.split('/')[-1], 'path': path, 'type': 2, 'left': free, 'right': free, 'child': free, 'start': end, 'size': len(data), 'color': 1})
    for parent, ids in [(0, [1] + [i for i in range(2, len(nodes)) if '/' not in nodes[i]['path']]), (1, [i for i in range(2, len(nodes)) if '/' in nodes[i]['path']])]:
        ids.sort(key=lambda i: (len(nodes[i]['name']), nodes[i]['name'].upper()))
        deepest = []
        def tree(items, depth=0):
            if not items:
                return free
            mid = len(items) // 2; index = items[mid]
            deepest.append((index, depth))
            nodes[index]['left'] = tree(items[:mid], depth + 1)
            nodes[index]['right'] = tree(items[mid + 1:], depth + 1)
            return index
        nodes[parent]['child'] = tree(ids)
        maxdepth = max(depth for _, depth in deepest)
        for index, depth in deepest:
            nodes[index]['color'] = 0 if depth == maxdepth and depth > 0 else 1
    sectors, fat = [], []
    def allocate(data):
        start = len(sectors)
        count = (len(data) + 511) // 512
        for i in range(count):
            sectors.append(data[i * 512:(i + 1) * 512].ljust(512, b'\0'))
            fat.append(start + i + 1 if i < count - 1 else end)
        return start if count else end
    mini = bytearray(); minifat = []
    for node in nodes[2:]:
        data = streams[node['path']]
        if len(data) < 4096:
            start = len(minifat); count = (len(data) + 63) // 64
            node['start'] = start if count else end
            for i in range(count):
                minifat.append(start + i + 1 if i < count - 1 else end)
                mini += data[i * 64:(i + 1) * 64].ljust(64, b'\0')
        else:
            node['start'] = allocate(data)
    nodes[0]['start'] = allocate(bytes(mini)); nodes[0]['size'] = len(mini)
    mini_start = allocate(b''.join(u32(n) for n in minifat).ljust(((len(minifat) + 127) // 128) * 512, b'\xff'))
    mini_sectors = (len(minifat) + 127) // 128
    entries = b''
    for n in nodes:
        name = (n['name'] + '\0').encode('utf-16le')
        entry = name.ljust(64, b'\0') + u16(len(name)) + bytes([n['type'], n['color']])
        entry += u32(n['left']) + u32(n['right']) + u32(n['child']) + b'\0' * 16 + u32(0) + b'\0' * 16
        entry += u32(n['start']) + struct.pack('<Q', n['size'])
        assert len(entry) == 128
        entries += entry
    dir_start = allocate(entries)
    nf = 1
    while nf * 128 < len(sectors) + nf:
        nf += 1
    fat_ids = list(range(len(sectors), len(sectors) + nf))
    fat.extend([fatmark] * nf)
    fat.extend([free] * (nf * 128 - len(fat)))
    sectors.extend(b''.join(u32(n) for n in fat[i * 128:(i + 1) * 128]) for i in range(nf))
    header = bytes.fromhex('d0cf11e0a1b11ae1') + b'\0' * 16
    header += struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\0' * 6
    header += struct.pack('<IIIIIIIII', 0, nf, dir_start, 0, 4096, mini_start, mini_sectors, end, 0)
    header += b''.join(u32(n) for n in (fat_ids + [free] * (109 - nf)))
    assert len(header) == 512
    return header + b''.join(sectors)
