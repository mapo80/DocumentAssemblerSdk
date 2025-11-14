#!/usr/bin/env python3
"""
Create template documents for Example08_Signature
"""

from pathlib import Path
import zipfile
import shutil
import tempfile

SCRIPT_DIR = Path(__file__).resolve().parent
EXAMPLES_DIR = SCRIPT_DIR / "DocumentAssemblerSdk.Examples"

def create_signature_template():
    source = EXAMPLES_DIR / "Example01_Basic" / "TemplateDocument.docx"
    target_dir = EXAMPLES_DIR / "Example08_Signature"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / "TemplateSignatureDocument.docx"

    shutil.copy(source, target)

    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
        <w:t>Signature Demonstration</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:r>
        <w:t xml:space="preserve">Questo esempio mostra come utilizzare il nuovo tag </w:t>
      </w:r>
      <w:r>
        <w:t>Signature</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> per inserire punti firma nel template.</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:r><w:t>Richiedente:</w:t></w:r>
    </w:p>
    <w:sdt>
      <w:sdtPr>
        <w:id w:val="101"/>
        <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
        <w:text/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t xml:space="preserve">&lt;# &lt;Signature Id="Richiedente" Label="Firma Richiedente" Width="220px" Height="60px" /&gt; #&gt;</w:t>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>

    <w:p>
      <w:r><w:t>Responsabile:</w:t></w:r>
    </w:p>
    <w:sdt>
      <w:sdtPr>
        <w:id w:val="102"/>
        <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
        <w:text/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t xml:space="preserve">&lt;# &lt;Signature Id="Responsabile" Label="Firma Responsabile" Width="220px" Height="60px" /&gt; #&gt;</w:t>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>

    <w:p>
      <w:r>
        <w:t xml:space="preserve">Durante la conversione in PDF, questi placeholder verranno sostituiti con campi firma AcroForm.</w:t>
      </w:r>
    </w:p>

    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(target, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        document_path = Path(tmpdir) / "word" / "document.xml"
        document_path.write_text(document_xml, encoding="utf-8")

        target.unlink()
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as docx:
            for path in Path(tmpdir).rglob("*"):
                if path.is_file():
                    docx.write(path, path.relative_to(tmpdir))

    print(f"✓ Created {target}")

def main():
    print("Creating Example08_Signature template document...")
    create_signature_template()
    print("✓ Done.")

if __name__ == "__main__":
    main()
