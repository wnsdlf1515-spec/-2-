from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
root=Path(__file__).resolve().parents[2]
out=root/'classroom/release/daon-classroom.zip'
out.parent.mkdir(exist_ok=True)
with ZipFile(out,'w',ZIP_DEFLATED) as z:
    for base in [root/'classroom',root/'city-stage6/src']:
        for f in sorted(base.rglob('*')):
            if f.is_file() and not any(p in {'data','release','node_modules','.env'} for p in f.relative_to(base).parts):
                z.write(f,Path('daon-classroom')/f.relative_to(root))
    for f in [root/'city-stage6/index.html',root/'.dockerignore']:
        z.write(f,Path('daon-classroom')/f.relative_to(root))
with ZipFile(out) as z: assert z.testzip() is None
print(out,out.stat().st_size)
