"""Recompute the SHA-256 of every file listed in export/EXPORT-MANIFEST.json and report."""
import hashlib, json, sys
from pathlib import Path

root = Path(__file__).resolve().parent / "export"
man = json.loads((root / "EXPORT-MANIFEST.json").read_text(encoding="utf-8"))
bad = []
for rel, want in man["files"].items():
    p = root / rel
    got = hashlib.sha256(p.read_bytes()).hexdigest() if p.exists() else "missing"
    if got != want:
        bad.append((rel, want[:12], got[:12]))
print(f"{len(man['files']) - len(bad)} of {len(man['files'])} files match the manifest (export built {man['built_at']})")
for rel, want, got in bad:
    print("MISMATCH", rel, "manifest", want, "file", got)
sys.exit(1 if bad else 0)
