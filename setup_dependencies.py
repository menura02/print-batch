"""Install the pinned PDF engine and optional legacy Office test fixtures."""
import argparse
import hashlib
import io
from pathlib import Path
from urllib.request import urlopen
from zipfile import ZipFile

BASE = Path(__file__).resolve().parent
SUMATRA = 'SumatraPDF-3.6.1-64'
SUMATRA_URL = f'https://www.sumatrapdfreader.org/dl/rel/3.6.1/{SUMATRA}.zip'
POI = 'https://raw.githubusercontent.com/apache/poi/trunk/test-data/'
FIXTURES = {
    'doc': ('document/simple-list.doc', '666f851727aeb9013d34586fe3ab5df687a36108ca3eb2ba193954cbd084be70'),
    'xls': ('spreadsheet/Simple.xls', '2b5ec8298bacffdca61c50f6f7607da0a4e63cf76fa011c62f203e6a66bc072a'),
    'ppt': ('slideshow/basic_test_ppt_file.ppt', 'e8fd4c38ad72ce3735acbca0b60dd5870fb60f819c169e3947d166f4f1bcc79d'),
}


def checked_download(url, sha256):
    with urlopen(url, timeout=60) as response:
        data = response.read()
    if hashlib.sha256(data).hexdigest() != sha256:
        raise RuntimeError(f'SHA-256 mismatch: {url}')
    return data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--fixtures', action='store_true', help='Also download legacy Office test fixtures')
    args = parser.parse_args()
    engine = BASE / 'tools' / f'{SUMATRA}.exe'
    if not engine.is_file():
        data = checked_download(SUMATRA_URL, '98b33a518d42986856d225064b0cd2d3643ecf78cbf84ab873d26cc51877a544')
        with ZipFile(io.BytesIO(data)) as archive:
            engine.write_bytes(archive.read(engine.name))
    if args.fixtures:
        for ext, (remote, sha256) in FIXTURES.items():
            target = BASE / 'tests' / 'fixtures' / f'sample.{ext}'
            if not target.is_file() or hashlib.sha256(target.read_bytes()).hexdigest() != sha256:
                target.write_bytes(checked_download(POI + remote, sha256))


if __name__ == '__main__':
    main()
