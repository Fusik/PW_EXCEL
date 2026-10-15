#!/usr/bin/env python3
"""Szybka kontrola src/*.bas|cls|frm przed importem: dlugosc linii (<1000) i niedomkniete literaly psFile.WriteLine."""
import glob, sys
bad = 0
for f in sorted(glob.glob(sys.argv[1] + '/*')):
    for i, l in enumerate(open(f, encoding='utf-8').read().split('\n'), 1):
        if len(l) > 1000: print('DLUGA LINIA', f, i, len(l)); bad = 1
        s = l.strip()
        if s.startswith('psFile.WriteLine "') and "' " not in s.split('"')[-1] and s.count('"') % 2 == 1:
            print('NIEDOMKNIETY LITERAL', f, i, s[:80]); bad = 1
sys.exit(bad)
