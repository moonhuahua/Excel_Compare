from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import urllib.request, urllib.parse, zipfile, io, xml.etree.ElementTree as ET
import os, shutil, subprocess, tempfile, hashlib

root = Path(__file__).resolve().parent / 'src'
packages = set()
for file in [root/'ExcelMerge.GUI/packages.config', root/'ExcelMerge/packages.config']:
    for item in ET.parse(file).getroot():
        packages.add((item.attrib['id'], item.attrib['version']))
packages.add(('Microsoft.NETFramework.ReferenceAssemblies.net452', '1.0.3'))
packages = {item for item in packages if item[0] not in ('SKCore', 'SKCore.Wpf')}
def restore(item):
    name, version = item
    dest = root/'packages'/f'{name}.{version}'
    if dest.is_dir():
        return name + ' cached'
    normalized = version if version.count('.') >= 2 else version + '.0'
    url = f'https://api.nuget.org/v3-flatcontainer/{name.lower()}/{normalized}/{name.lower()}.{normalized}.nupkg'
    data = urllib.request.urlopen(url, timeout=90).read()
    dest.mkdir(parents=True)
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        z.extractall(dest)
    for path in sorted(dest.rglob('*'), key=lambda p: len(p.parts), reverse=True):
        decoded = urllib.parse.unquote(path.name)
        if decoded != path.name and not path.with_name(decoded).exists():
            path.rename(path.with_name(decoded))
    return name + ' restored'
with ThreadPoolExecutor(max_workers=6) as pool:
    for result in pool.map(restore, sorted(packages)):
        print(result, flush=True)

# SKCore packages are no longer available from NuGet. Recover the exact DLLs
# from the author's pinned official release without running its installer.
legacy = {
    'EE19B099646060E7966D7AF23E1DC685C4CC802E2762A746AFA0D6E968262331':
        root/'packages/SKCore.1.0.1/lib/net452/SKCore.dll',
    '892647CFDA1BCF864C5B5CF1540FEF230578F310B4D7FDDB6C13CA9EDECE16F4':
        root/'packages/SKCore.Wpf.1.0.2/lib/net452/SKCore.Wpf.dll'
}
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest().upper()

# Restore the two pinned legacy packages from the source tree first. Their DLL
# hashes must match the original official release; no installer is executed.
for sha, path in legacy.items():
    if path.is_file() and digest(path) == sha:
        continue
    package = root/'.nuget/source'/(path.parents[2].name + '.nupkg')
    if package.is_file():
        with zipfile.ZipFile(package) as archive:
            data = archive.read(path.relative_to(path.parents[2]).as_posix())
        if hashlib.sha256(data).hexdigest().upper() != sha:
            raise SystemExit('Bundled legacy package hash mismatch: ' + package.name)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
if not all(path.is_file() and digest(path) == sha for sha,path in legacy.items()):
    sevenzip = shutil.which('7z') or str(Path(os.environ.get('LOCALAPPDATA', str(Path.home())))/'Excel_Compare/BuildTools/7zip-2601/app/7z.exe')
    if not Path(sevenzip).is_file():
        raise SystemExit('7z.exe required to recover SKCore from the official MSI.')
    with tempfile.TemporaryDirectory(prefix='excelmerge_restore_') as temp:
        folder = Path(temp)
        msi = folder/'ExcelMergeSetup.msi'
        url = 'https://github.com/skanmera/ExcelMerge/releases/download/ver1.3.4/ExcelMergeSetup_v1.3.4.msi'
        msi.write_bytes(urllib.request.urlopen(url, timeout=120).read())
        assert digest(msi) == '48D658E578DEDCF41D3B4FB4D45815792CC04C4DDDECFDBC7ACD04401DEDB297', 'Upstream MSI hash changed'
        flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        subprocess.run([sevenzip,'x',str(msi),'-o'+str(folder/'msi'),'-y'],check=True,capture_output=True,timeout=90,creationflags=flags)
        cabinets=[p for p in (folder/'msi').iterdir() if p.is_file() and p.read_bytes()[:4]==b'MSCF']
        assert len(cabinets)==1
        subprocess.run([sevenzip,'x',str(cabinets[0]),'-o'+str(folder/'payload'),'-y'],check=True,capture_output=True,timeout=90,creationflags=flags)
        for file in (folder/'payload').iterdir():
            if file.is_file():
                sha=digest(file)
                if sha in legacy:
                    legacy[sha].parent.mkdir(parents=True,exist_ok=True)
                    shutil.copyfile(file,legacy[sha])
    assert all(path.is_file() and digest(path)==sha for sha,path in legacy.items())
print('SKCore and SKCore.Wpf verified against official release hashes.')
