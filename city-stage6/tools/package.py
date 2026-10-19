from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
root=Path(__file__).resolve().parents[1]
out=root/'release/daon-city-stage6-rc1.zip'
files=[root/'README.md',root/'index.html',root/'package.json',root/'release/daon-city-stage6-rc1.html']
for folder in ['src','docs','tests','tools','artifacts']:
    files.extend(p for p in (root/folder).rglob('*') if p.is_file())
with ZipFile(out,'w',ZIP_DEFLATED) as z:
    for p in sorted(files): z.write(p,Path('daon-city-stage6-rc1')/p.relative_to(root))
with ZipFile(out) as z: assert z.testzip() is None
sha=hashlib.sha256(out.read_bytes()).hexdigest()
(root/'release/SHA256SUMS.txt').write_text(f'{sha}  {out.name}\n')
print(out, out.stat().st_size, sha)
