#!/bin/bash
# Regenerate photo-manifest.js from OneDrive folder listings
# Works with cloud-only stubs — only reads folder/file names, not image data
# Run from the repo root: ./update-manifest.sh

PHOTO_DIR="$HOME/Library/CloudStorage/OneDrive-ThePennsylvaniaStateUniversity/students/undergrads/tornadoREUs/2025stlouis/Reference_Overall_Photo_Links/StLouis_2025_Tornado_Photos"

if [ ! -d "$PHOTO_DIR" ]; then
  echo "Error: Photo directory not found at $PHOTO_DIR"
  exit 1
fi

python3 << 'PYEOF'
import os, json, re

base = os.path.expanduser("~/Library/CloudStorage/OneDrive-ThePennsylvaniaStateUniversity/students/undergrads/tornadoREUs/2025stlouis/Reference_Overall_Photo_Links/StLouis_2025_Tornado_Photos/")

SUFFIXES = {'drive':'dr','dr':'dr','place':'pl','pl':'pl','lane':'ln','ln':'ln',
            'terrace':'ter','ter':'ter','avenue':'ave','ave':'ave','court':'ct','ct':'ct',
            'road':'rd','rd':'rd','street':'st','st':'st','boulevard':'blvd','blvd':'blvd',
            'way':'way','circle':'cir','cir':'cir'}

def normalize(addr):
    if not addr: return ''
    parts = addr.lower().strip().split(',')[0].strip().split()
    nums, rest = [], []
    for w in parts:
        if re.match(r'^\d+$', w) and not rest: nums.append(w)
        else: rest.append(w)
    if rest and rest[-1].rstrip('.') in SUFFIXES:
        rest[-1] = SUFFIXES[rest[-1].rstrip('.')]
    return ' '.join(nums) + ' ' + ' '.join(rest)

manifest = {}
for folder in sorted(os.listdir(base)):
    full = os.path.join(base, folder)
    if not os.path.isdir(full): continue
    files = sorted([f for f in os.listdir(full) if f.lower().endswith(('.png','.jpg','.jpeg'))])
    if not files: continue
    photos = []
    for f in files:
        name_lower = f.rsplit('.', 1)[0].lower()
        view = 'Other'
        for v in ['aerial', 'front', 'back', 'left', 'right']:
            if v in name_lower:
                view = v.capitalize()
                break
        time = 'after' if 'after' in name_lower else 'before'
        photos.append({'file': f, 'view': view, 'time': time})
    manifest[normalize(folder)] = {'folder': folder, 'photos': photos}

out = os.path.join(os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else '.', 'photo-manifest.js')
out = 'photo-manifest.js'
with open(out, 'w') as fh:
    fh.write('var PHOTO_MANIFEST = ')
    json.dump(manifest, fh, separators=(',', ':'))
    fh.write(';\n')
print(f"Updated photo-manifest.js: {len(manifest)} buildings")
PYEOF
