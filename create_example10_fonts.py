#!/usr/bin/env python3
"""Generate the Example10_Fonts template document used to verify custom font propagation."""

from pathlib import Path
import shutil
import tempfile
import zipfile

SCRIPT_DIR = Path(__file__).resolve().parent
EXAMPLES_DIR = SCRIPT_DIR / 'DocumentAssemblerSdk.Examples'


def sdt(tag_text: str, placeholder_id: int, indent: int = 4, rich_text: bool = False) -> str:
    pad = ' ' * indent
    pad2 = ' ' * (indent + 2)
    pad3 = ' ' * (indent + 4)
    pad4 = ' ' * (indent + 6)
    control_tag = 'w:richText' if rich_text else 'w:text'
    return ''.join(
        [
            f'{pad}<w:sdt>\n',
            f'{pad2}<w:sdtPr>\n',
            f'{pad3}<w:id w:val="{placeholder_id}"/>\n',
            f'{pad3}<w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>\n',
            f'{pad3}<{control_tag}/>\n',
            f'{pad2}</w:sdtPr>\n',
            f'{pad2}<w:sdtContent>\n',
            f'{pad3}<w:p>\n',
            f'{pad4}<w:r>\n',
            f'{pad4}  <w:t>{tag_text}</w:t>\n',
            f'{pad4}</w:r>\n',
            f'{pad3}</w:p>\n',
            f'{pad2}</w:sdtContent>\n',
            f'{pad}</w:sdt>\n',
        ]
    )


def sdt_with_rpr(tag_text: str, placeholder_id: int, rpr: str, indent: int = 4) -> str:
    pad = ' ' * indent
    pad2 = ' ' * (indent + 2)
    pad3 = ' ' * (indent + 4)
    pad4 = ' ' * (indent + 6)
    return ''.join(
        [
            f'{pad}<w:sdt>\n',
            f'{pad2}<w:sdtPr>\n',
            f'{pad3}<w:id w:val="{placeholder_id}"/>\n',
            f'{pad3}<w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>\n',
            f'{pad3}<w:richText/>\n',
            f'{pad2}</w:sdtPr>\n',
            f'{pad2}<w:sdtContent>\n',
            f'{pad3}<w:p>\n',
            f'{pad4}<w:r>\n',
            f'{pad4}  {rpr}\n',
            f'{pad4}  <w:t>{tag_text}</w:t>\n',
            f'{pad4}</w:r>\n',
            f'{pad3}</w:p>\n',
            f'{pad2}</w:sdtContent>\n',
            f'{pad}</w:sdt>\n',
        ]
    )


def create_fonts_template() -> None:
    source = EXAMPLES_DIR / 'Example01_Basic' / 'TemplateDocument.docx'
    target_dir = EXAMPLES_DIR / 'Example10_Fonts'
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / 'TemplateFontsDocument.docx'

    if not source.exists():
        raise FileNotFoundError(f'Base template not found: {source}')

    shutil.copy(source, target)

    parts: list[str] = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    parts.append('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n')
    parts.append('            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n')
    parts.append('  <w:body>\n')
    parts.append('    <w:p>\n')
    parts.append('      <w:pPr><w:jc w:val="center"/></w:pPr>\n')
    parts.append('      <w:r>\n')
    parts.append('        <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>\n')
    parts.append('        <w:t>Font verification template</w:t>\n')
    parts.append('      </w:r>\n')
    parts.append('    </w:p>\n')

    parts.append('    <w:p>\n')
    parts.append('      <w:r><w:t xml:space="preserve">Customer: </w:t></w:r>\n')
    parts.append('    </w:p>\n')
    parts.append(sdt('&lt;Content Select="FontSample/CustomerName"/&gt;', 1, indent=4))

    parts.append('    <w:p>\n')
    parts.append('      <w:r><w:t>Barcode preview (uses Libre Barcode 128 Text):</w:t></w:r>\n')
    parts.append('    </w:p>\n')
    barcode_rpr = (
        '<w:rPr>'
        '<w:rFonts w:ascii="Libre Barcode 128 Text" w:hAnsi="Libre Barcode 128 Text" w:cs="Libre Barcode 128 Text"/>'
        '<w:sz w:val="72"/>'
        '</w:rPr>'
    )
    parts.append(sdt_with_rpr('&lt;Content Select="FontSample/Barcode"/&gt;', 2, barcode_rpr, indent=4))

    parts.append('    <w:p>\n')
    parts.append('      <w:r><w:t xml:space="preserve">Nota: questo paragrafo usa il font di default per evidenziare la differenza visiva.</w:t></w:r>\n')
    parts.append('    </w:p>\n')

    parts.append('    <w:sectPr>\n')
    parts.append('      <w:pgSz w:w="12240" w:h="15840"/>\n')
    parts.append('      <w:pgMar w:top="1000" w:right="1200" w:bottom="1000" w:left="1200"/>\n')
    parts.append('    </w:sectPr>\n')
    parts.append('  </w:body>\n')
    parts.append('</w:document>\n')

    document_xml = ''.join(parts)

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(target, 'r') as archive:
            archive.extractall(tmpdir)

        document_path = Path(tmpdir) / 'word' / 'document.xml'
        document_path.write_text(document_xml, encoding='utf-8')

        target.unlink()
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as output:
            for path in Path(tmpdir).rglob('*'):
                if path.is_file():
                    output.write(path, path.relative_to(tmpdir))

    print(f'✓ Created {target}')


def main() -> None:
    print('Creating Example10_Fonts template document...')
    create_fonts_template()
    print('✓ Done.')


if __name__ == '__main__':
    main()
