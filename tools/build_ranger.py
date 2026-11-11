#!/usr/bin/env python3
"""Builds ranger.html from tools/ranger.src.html + the shelf/zone/photo code of
Stock1981.html (ZONES, LOCS, zone photos and shelf boxes), so both apps draw
the same map. Run after changing zones or built-in shelf boxes in Stock1981."""
import re, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
st = (root / 'Stock1981.html').read_text()
def line(prefix):
    m = re.search(r'^' + re.escape(prefix) + r'.*$', st, re.M); assert m, prefix; return m.group(0)
def block(start, end_marker):
    i = st.index(start); j = st.index(end_marker, i); return st[i:j]
parts = [
    line('const ZONES = '),
    block('const LOCS = [', '];') + '];',
    line('const bkey = '),
    block('function parseShape(str){', '// Zones with no shelves'),
    block('function zoneSpots(z){', 'function locName('),
    block('const TRACES_BUILTIN = ', 'function spotState('),
    line('const AZ_MAP = '),
]
shared = '\n'.join(parts)
src = (root / 'tools' / 'ranger.src.html').read_text()
out = src.replace('/*@@SHARED@@*/', shared)
(root / 'ranger.html').write_text(out)
print('ranger.html built,', len(out), 'bytes')
