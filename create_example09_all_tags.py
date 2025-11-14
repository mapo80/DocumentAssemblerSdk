#!/usr/bin/env python3
"""Generate the Example09_AllTags template document."""

from pathlib import Path
import shutil
import tempfile
import zipfile

SCRIPT_DIR = Path(__file__).resolve().parent
EXAMPLES_DIR = SCRIPT_DIR / 'DocumentAssemblerSdk.Examples'


def sdt(tag_text: str, placeholder_id: int, indent: int = 4) -> str:
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
            f'{pad3}<w:text/>\n',
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


def table_cell(tag_text: str, placeholder_id: int) -> str:
    return '      <w:tc>\n' + sdt(tag_text, placeholder_id, indent=8) + '      </w:tc>\n'


def table_cell_xpath(x_path: str) -> str:
    return ''.join(
        [
            '      <w:tc>\n',
            '        <w:p>\n',
            '          <w:r>\n',
            f'            <w:t>{x_path}</w:t>\n',
            '          </w:r>\n',
            '        </w:p>\n',
            '      </w:tc>\n',
        ]
    )


def create_all_tags_template() -> None:
    source = EXAMPLES_DIR / 'Example01_Basic' / 'TemplateDocument.docx'
    target_dir = EXAMPLES_DIR / 'Example09_AllTags'
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / 'TemplateAllTagsDocument.docx'

    shutil.copy(source, target)

    parts: list[str] = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    parts.append('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\n')
    parts.append('            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n')
    parts.append('  <w:body>\n')
    parts.append('    <w:p>\n')
    parts.append('      <w:pPr><w:jc w:val="center"/><w:shd w:val="clear" w:color="auto" w:fill="1F4E78"/></w:pPr>\n')
    parts.append('      <w:r><w:rPr><w:color w:val="FFFFFF"/><w:b/><w:sz w:val="48"/></w:rPr><w:t>Strategic Delivery Report</w:t></w:r>\n')
    parts.append('    </w:p>\n')
    parts.append('    <w:p>\n')
    parts.append('      <w:pPr><w:jc w:val="center"/></w:pPr>\n')
    parts.append('      <w:r><w:rPr><w:color w:val="666666"/><w:sz w:val="28"/></w:rPr><w:t>Quarterly executive overview powered by DocumentAssembler</w:t></w:r>\n')
    parts.append('    </w:p>\n')
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Questo documento dimostra l’utilizzo congiunto di Content, Repeat, Table, Image, Conditional, Else, EndConditional e Signature su un layout multi-pagina.</w:t></w:r></w:p>\n')

    parts.append('    <w:tbl>\n')
    parts.append('      <w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="D6DCE5"/><w:left w:val="single" w:sz="12" w:space="0" w:color="D6DCE5"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="D6DCE5"/><w:right w:val="single" w:sz="12" w:space="0" w:color="D6DCE5"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="FFFFFF"/><w:insideV w:val="single" w:sz="6" w:space="0" w:color="FFFFFF"/></w:tblBorders></w:tblPr>\n')
    parts.append('      <w:tblGrid><w:gridCol w:w="6000"/><w:gridCol w:w="6000"/></w:tblGrid>\n')
    parts.append('      <w:tr>\n')
    parts.append('        <w:tc>\n')
    parts.append('          <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F7FBFF"/></w:tcPr>\n')
    parts.append('          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Cliente</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Customer/FullName"/&gt;', 101, indent=10))
    parts.append('          <w:p><w:r><w:t xml:space="preserve">Ruolo: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Customer/Title"/&gt;', 102, indent=10))
    parts.append('          <w:p><w:r><w:t xml:space="preserve">Sede: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Customer/Location/City"/&gt;', 103, indent=10))
    parts.append('          <w:p><w:r><w:t xml:space="preserve">Middle name (opzionale): </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Customer/MiddleName" Optional="true"/&gt;', 104, indent=10))
    parts.append('        </w:tc>\n')
    parts.append('        <w:tc>\n')
    parts.append('          <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="EEF3FB"/></w:tcPr>\n')
    parts.append('          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Membership</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Customer/MembershipType"/&gt;', 105, indent=10))
    parts.append('          <w:p><w:r><w:t xml:space="preserve">Customer loyalty: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Customer/LoyaltyScore"/&gt;', 106, indent=10))
    parts.append('          <w:p><w:r><w:t xml:space="preserve">Photo: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Image Select="Report/Customer/Photo" MaxWidth="140px" MaxHeight="140px"/&gt;', 107, indent=10))
    parts.append('        </w:tc>\n')
    parts.append('      </w:tr>\n')
    parts.append('    </w:tbl>\n')

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Esperienza personalizzata</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Conditional Select="Report/Customer/MembershipType" Match="Platinum"/&gt;', 201))
    parts.append('    <w:p><w:r><w:rPr><w:color w:val="0F6FC6"/><w:b/></w:rPr><w:t>Accesso prioritario a laboratori, consulenze dedicate e roadmap congiunte.</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Else/&gt;', 202))
    parts.append('    <w:p><w:r><w:rPr><w:color w:val="8A2D10"/><w:b/></w:rPr><w:t>Piano Essentials: onboarding accelerato e monitoraggio su base mensile.</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;EndConditional/&gt;', 203))

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Highlights strategici</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Repeat Select="Report/Highlights/Highlight"/&gt;', 301))
    parts.append('    <w:tbl>\n')
    parts.append('      <w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders></w:tblPr>\n')
    parts.append('      <w:tr>\n')
    parts.append('        <w:tc><w:p><w:r><w:t xml:space="preserve">Icona: </w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc>\n')
    parts.append(sdt('&lt;Content Select="./Icon"/&gt;', 302, indent=12))
    parts.append('        </w:tc>\n')
    parts.append('        <w:tc>\n')
    parts.append('          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Titolo</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Title"/&gt;', 303, indent=12))
    parts.append('          <w:p><w:r><w:t xml:space="preserve">Impatto: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Impact"/&gt;', 304, indent=12))
    parts.append('        </w:tc>\n')
    parts.append('      </w:tr>\n')
    parts.append('    </w:tbl>\n')
    parts.append(sdt('&lt;EndRepeat/&gt;', 305))

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Indicatori chiave</w:t></w:r></w:p>\n')
    parts.append('    <w:tbl>\n')
    parts.append('      <w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="D0D7E8"/><w:left w:val="single" w:sz="8" w:color="D0D7E8"/><w:bottom w:val="single" w:sz="8" w:color="D0D7E8"/><w:right w:val="single" w:sz="8" w:color="D0D7E8"/><w:insideH w:val="single" w:sz="4" w:color="D0D7E8"/><w:insideV w:val="single" w:sz="4" w:color="D0D7E8"/></w:tblBorders></w:tblPr>\n')
    parts.append('      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>\n')
    parts.append('      <w:tr>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Revenue YTD</w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Crescita</w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Soddisfazione</w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Retention</w:t></w:r></w:p></w:tc>\n')
    parts.append('      </w:tr>\n')
    parts.append('      <w:tr>\n')
    parts.append(table_cell('&lt;Content Select="Report/KPIs/RevenueYTD"/&gt;', 401))
    parts.append(table_cell('&lt;Content Select="Report/KPIs/Growth"/&gt;', 402))
    parts.append(table_cell('&lt;Content Select="Report/KPIs/Satisfaction"/&gt;', 403))
    parts.append(table_cell('&lt;Content Select="Report/KPIs/Retention"/&gt;', 404))
    parts.append('      </w:tr>\n')
    parts.append('    </w:tbl>\n')

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Dashboard rapida</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Image Select="Report/Charts/Performance" MaxWidth="480px" MaxHeight="220px"/&gt;', 405))

    parts.append('    <w:p><w:r><w:br w:type="page"/></w:r></w:p>\n')

    parts.append('    <w:p><w:r><w:rPr><w:color w:val="1F4E78"/><w:b/><w:sz w:val="36"/></w:rPr><w:t>Operational deep dive</w:t></w:r></w:p>\n')
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Analisi dei reparti e della pipeline ordini.</w:t></w:r></w:p>\n')

    parts.append(sdt('&lt;Repeat Select="Report/Departments/Department"/&gt;', 501))
    parts.append('    <w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="12" w:color="C6D7F7"/></w:pBdr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Dipartimento</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Name"/&gt;', 502))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Focus: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Focus"/&gt;', 503))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Headcount: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./HeadCount"/&gt;', 504))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Budget allocato / speso: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Budget/Allocated"/&gt;', 505))
    parts.append(sdt('&lt;Content Select="./Budget/Spent"/&gt;', 506))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Rischio: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./RiskLevel"/&gt;', 507))
    parts.append(sdt('&lt;Repeat Select="./Achievements/Achievement"/&gt;', 508))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">✔ </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="."/&gt;', 509))
    parts.append(sdt('&lt;EndRepeat/&gt;', 510))
    parts.append(sdt('&lt;EndRepeat/&gt;', 511))

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Ordini chiave</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Table Select="Report/Orders/Order"/&gt;', 601))
    parts.append('    <w:tbl>\n')
    parts.append('      <w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="BCC8DD"/><w:left w:val="single" w:sz="8" w:color="BCC8DD"/><w:bottom w:val="single" w:sz="8" w:color="BCC8DD"/><w:right w:val="single" w:sz="8" w:color="BCC8DD"/><w:insideH w:val="single" w:sz="4" w:color="BCC8DD"/><w:insideV w:val="single" w:sz="4" w:color="BCC8DD"/></w:tblBorders></w:tblPr>\n')
    parts.append('      <w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="3400"/><w:gridCol w:w="1200"/><w:gridCol w:w="1800"/><w:gridCol w:w="2400"/></w:tblGrid>\n')
    parts.append('      <w:tr>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Codice</w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Prodotto</w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Q.tà</w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Prezzo</w:t></w:r></w:p></w:tc>\n')
    parts.append('        <w:tc><w:p><w:r><w:b/><w:t>Note</w:t></w:r></w:p></w:tc>\n')
    parts.append('      </w:tr>\n')
    parts.append('      <w:tr>\n')
    parts.append(table_cell_xpath("./@code"))
    parts.append(table_cell_xpath("./Product"))
    parts.append(table_cell_xpath("./Quantity"))
    parts.append(table_cell_xpath("./Price"))
    parts.append(table_cell_xpath("./Notes"))
    parts.append('      </w:tr>\n')
    parts.append('    </w:tbl>\n')
    parts.append(sdt('&lt;Repeat Select="Report/Orders/Order"/&gt;', 607))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Codice ordine dettagliato: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./@code"/&gt;', 608))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Finestra consegna: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./DeliveryWindow/Start"/&gt;', 609))
    parts.append(sdt('&lt;Content Select="./DeliveryWindow/End"/&gt;', 610))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Stato: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./@status"/&gt;', 611))
    parts.append(sdt('&lt;EndRepeat/&gt;', 612))

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Insight sintetico</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Insights/Text"/&gt;', 613))
    parts.append('    <w:p><w:r><w:rPr><w:color w:val="107C41"/><w:i/></w:rPr><w:t>La combinazione di condizioni, ripetizioni e immagini consente dashboard dinamiche.</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Image Select="Report/Charts/Heatmap" MaxWidth="460px" MaxHeight="200px"/&gt;', 614))

    parts.append('    <w:p><w:r><w:br w:type="page"/></w:r></w:p>\n')

    parts.append('    <w:p><w:r><w:rPr><w:sz w:val="36"/><w:b/><w:color w:val="2F5496"/></w:rPr><w:t>Roadmap &amp; governance</w:t></w:r></w:p>\n')
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Milestone principali con attributi e nested repeat.</w:t></w:r></w:p>\n')

    parts.append(sdt('&lt;Repeat Select="Report/Milestones/Milestone"/&gt;', 701))
    parts.append('    <w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="FDF2D0"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Milestone</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Title"/&gt;', 702))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Owner: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Owner"/&gt;', 703))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Due date: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./DueDate"/&gt;', 704))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Status: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Status"/&gt;', 705))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Codice: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./@code"/&gt;', 706))
    parts.append(sdt('&lt;EndRepeat/&gt;', 707))

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Allegati e materiali</w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Repeat Select="Report/Attachments/Attachment"/&gt;', 801))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">🔗 </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="./Label"/&gt;', 802))
    parts.append(sdt('&lt;Content Select="./Description"/&gt;', 803))
    parts.append(sdt('&lt;Content Select="./Url"/&gt;', 804))
    parts.append(sdt('&lt;EndRepeat/&gt;', 805))

    parts.append('    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Approvazioni</w:t></w:r></w:p>\n')
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Firma primaria:</w:t></w:r></w:p>\n')
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Responsabile approvazione: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Approvals/PrimarySigner"/&gt;', 903))
    parts.append(sdt('&lt;Signature Id="PrimarySigner" Label="Firma primaria" Width="230px" Height="70px"/&gt;', 901))
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Firma backup:</w:t></w:r></w:p>\n')
    parts.append('    <w:p><w:r><w:t xml:space="preserve">Delegato sostitutivo: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Approvals/BackupSigner"/&gt;', 904))
    parts.append(sdt('&lt;Signature Id="BackupSigner" Label="Firma sostitutiva" Width="230px" Height="70px"/&gt;', 902))

    parts.append('    <w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Nota finale: </w:t></w:r></w:p>\n')
    parts.append(sdt('&lt;Content Select="Report/Customer/PremiumMessage"/&gt;', 905))

    parts.append('    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1000" w:right="1200" w:bottom="1000" w:left="1200"/></w:sectPr>\n')
    parts.append('  </w:body>\n</w:document>\n')

    document_xml = ''.join(parts)

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(target, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        document_path = Path(tmpdir) / 'word' / 'document.xml'
        document_path.write_text(document_xml, encoding='utf-8')

        target.unlink()
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as docx:
            for path in Path(tmpdir).rglob('*'):
                if path.is_file():
                    docx.write(path, path.relative_to(tmpdir))

    print(f'✓ Created {target}')


def main() -> None:
    print('Creating Example09_AllTags template document...')
    create_all_tags_template()
    print('✓ Done.')


if __name__ == '__main__':
    main()
