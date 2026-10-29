from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import hashlib,json,os,shutil

work=Path('/private/tmp/hinton-place-properties')
audit=json.loads((work/'audit.json').read_text())
report=json.loads((work/'validation.json').read_text())
assert report['result']=='passed'
target=Path(audit['source']);staged=work/'campus_academic_ontology.owl'
digest=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
assert digest(target)==audit['source_sha256'],'The source file changed after inspection; do not overwrite.'
assert digest(staged)==audit['output_sha256'],'The staged file changed after validation.'
stamp=datetime.now(ZoneInfo('Asia/Shanghai')).strftime('%Y%m%d-%H%M%S')
backup=target.with_name(target.stem+'.before-place-properties-'+stamp+target.suffix)
assert not backup.exists()
shutil.copy2(target,backup)
assert digest(backup)==audit['source_sha256']
pending=target.with_name('.'+target.name+'.properties-'+stamp+'.tmp')
try:
    shutil.copyfile(staged,pending);shutil.copymode(target,pending)
    assert digest(pending)==audit['output_sha256']
    assert digest(target)==audit['source_sha256'],'Source changed during installation.'
    os.replace(pending,target)
finally:
    if pending.exists():pending.unlink()
assert digest(target)==audit['output_sha256']
result={'updated':str(target),'backup':str(backup),'sha256':digest(target)}
(work/'installed.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
print(json.dumps(result,ensure_ascii=False,indent=2))
