from pathlib import Path
import zipfile,hashlib,json
root=Path(__file__).resolve().parent.parent
output=root/'release/daon-city-stage5-source.zip'
files=[]
for name in ['src','tests','tools','docs','artifacts']:
 files.extend(p for p in (root/name).rglob('*') if p.is_file())
files.extend(root/n for n in ['README.md','index.html','package.json','.gitignore','release/daon-city-stage5.html'])
with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
 for p in sorted(files): z.write(p,Path('city-stage5')/p.relative_to(root))
manifest={p.name:{'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in [output,root/'release/daon-city-stage5.html']}
(root/'release/checksums.json').write_text(json.dumps(manifest,indent=2)+'\n')
print(json.dumps(manifest,indent=2))
