#!/usr/bin/env bash
# Builds dist/MoveTA.exe on Linux (cloud sessions) with the Mono C# compiler.
# The output targets .NET Framework 4.x and runs on Windows like the csc.exe build.
# Needs: mono-mcs, libmono-system-windows-forms4.0-cil (apt) and python3 + pefile (pip)
#        only when build/MoveTA.ico is missing (the icon is taken from the previous dist/MoveTA.exe).
set -euo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"
cd "$root"
if [ ! -f build/MoveTA.ico ]; then
  python3 - <<'PY'
import pefile, struct
pe = pefile.PE('dist/MoveTA.exe')
icons, group = {}, None
for t in pe.DIRECTORY_ENTRY_RESOURCE.entries:
    for e in t.directory.entries:
        for l in e.directory.entries:
            d = pe.get_data(l.data.struct.OffsetToData, l.data.struct.Size)
            if t.id == 3: icons[e.id] = d
            elif t.id == 14: group = d
n = struct.unpack('<HHH', group[:6])[2]
ents = [struct.unpack('<BBBBHHIH', group[6 + 14 * i:20 + 14 * i]) for i in range(n)]
head, body, off = struct.pack('<HHH', 0, 1, n), b'', 6 + 16 * n
for w, h, c, r, planes, bits, _, iid in ents:
    data = icons[iid]
    head += struct.pack('<BBBBHHII', w, h, c, r, planes, bits, len(data), off); off += len(data); body += data
open('build/MoveTA.ico', 'wb').write(head + body)
PY
fi
mcs -nologo -target:winexe -platform:anycpu -optimize+ -out:dist/MoveTA.exe -win32icon:build/MoveTA.ico \
    -r:System.Windows.Forms.dll \
    -resource:Move-TA.ps1,Move-TA.ps1 -resource:usp_IPS_MoveTA.sql,usp_IPS_MoveTA.sql \
    build/MoveTA.Launcher.cs
ls -la dist/MoveTA.exe
