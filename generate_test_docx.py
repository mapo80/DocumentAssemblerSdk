#!/usr/bin/env python3
"""
Word Document Generator for DocumentAssembler Test Files
=========================================================

This script creates .docx test documents with DocumentAssembler content controls.

USAGE:
------
    # Generate all test documents (default)
    python3 generate_test_docx.py

    # Generate to a specific directory
    python3 generate_test_docx.py /path/to/output/dir

    # View this help
    python3 generate_test_docx.py --help

WHAT IT DOES:
-------------
The script generates three test documents:
1. DA270-ConditionalWithElse.docx - Basic Conditional with Else using Match
2. DA271-ConditionalWithElseNotMatch.docx - Conditional with Else using NotMatch
3. DA272-NestedConditionalWithElse.docx - Nested Conditionals with Else

Each document tests different aspects of the Conditional/Else functionality.

HOW IT WORKS:
-------------
1. Copies structure from Example01_Basic/TemplateDocument.docx
2. Creates custom document.xml with content controls
3. Packages everything into valid .docx files (ZIP format)

REQUIREMENTS:
-------------
- Python 3.x (no external dependencies)
- Template: DocumentAssemblerSdk.Examples/Example01_Basic/TemplateDocument.docx

OUTPUT:
-------
Generated files are placed in: DocumentAssemblerSdk.Tests/TestFiles/
(or custom directory if specified)
"""

import zipfile
import os
import sys
from xml.dom import minidom

def create_document_xml_with_else():
    """
    Create document.xml content with Conditional/Else structure.
    Uses raw XML string to avoid namespace issues.
    """
    xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    mc:Ignorable="w14 w15 wp14">
    <w:body>
        <!-- Title -->
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                </w:rPr>
                <w:t>Customer Membership Status</w:t>
            </w:r>
        </w:p>

        <!-- Empty line -->
        <w:p/>

        <!-- Customer name label -->
        <w:p>
            <w:r>
                <w:t xml:space="preserve">Customer: </w:t>
            </w:r>
        </w:p>

        <!-- Content control for Customer Name -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="100001"/>
                <w:placeholder>
                    <w:docPart w:val="DefaultPlaceholder_1081868574"/>
                </w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>&lt;Content Select="./Name"/&gt;</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <!-- Empty line -->
        <w:p/>

        <!-- Conditional start -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="100002"/>
                <w:placeholder>
                    <w:docPart w:val="DefaultPlaceholder_1081868574"/>
                </w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>&lt;Conditional Select="./MembershipType" Match="Premium"/&gt;</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <!-- If content (Premium) -->
        <w:p>
            <w:r>
                <w:t>You are a PREMIUM member! You get:</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>- 20% discount on all purchases</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>- Free shipping worldwide</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>- Priority customer support</w:t>
            </w:r>
        </w:p>

        <!-- Else tag -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="100003"/>
                <w:placeholder>
                    <w:docPart w:val="DefaultPlaceholder_1081868574"/>
                </w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>&lt;Else/&gt;</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <!-- Else content (Standard) -->
        <w:p>
            <w:r>
                <w:t>You are a Standard member. You get:</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>- 5% discount on purchases</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>- Standard shipping rates apply</w:t>
            </w:r>
        </w:p>

        <!-- EndConditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="100004"/>
                <w:placeholder>
                    <w:docPart w:val="DefaultPlaceholder_1081868574"/>
                </w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>&lt;EndConditional/&gt;</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <!-- Empty line -->
        <w:p/>

        <!-- Closing -->
        <w:p>
            <w:r>
                <w:t>Thank you for your business!</w:t>
            </w:r>
        </w:p>

        <!-- Section properties -->
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
            <w:docGrid w:linePitch="360"/>
        </w:sectPr>
    </w:body>
</w:document>'''

    return xml_content

def create_document_xml_with_else_notmatch():
    """Create document.xml with Conditional/Else using NotMatch"""
    xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>Shipping Information</w:t>
            </w:r>
        </w:p>
        <w:p/>

        <!-- Conditional with NotMatch -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="200001"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>&lt;Conditional Select="./Country" NotMatch="USA"/&gt;</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p>
            <w:r><w:t>International shipping rates apply.</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Estimated delivery: 10-15 business days.</w:t></w:r>
        </w:p>

        <!-- Else -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="200002"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;Else/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p>
            <w:r><w:t>Domestic USA shipping - FREE!</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Estimated delivery: 2-3 business days.</w:t></w:r>
        </w:p>

        <!-- EndConditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="200003"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;EndConditional/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
    </w:body>
</w:document>'''

    return xml_content

def create_document_xml_nested_else():
    """Create document.xml with nested Conditional/Else"""
    xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>Customer Benefits</w:t>
            </w:r>
        </w:p>
        <w:p/>

        <!-- Outer Conditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="300001"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;Conditional Select="./MembershipType" Match="Premium"/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p>
            <w:r><w:t>Premium Member Benefits:</w:t></w:r>
        </w:p>

        <!-- Inner Conditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="300002"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;Conditional Select="./Points" Match="5000"/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p>
            <w:r><w:t>- PLATINUM status with 5000 points!</w:t></w:r>
        </w:p>

        <!-- Inner Else -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="300003"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;Else/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p>
            <w:r><w:t>- You have Premium status</w:t></w:r>
        </w:p>

        <!-- Inner EndConditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="300004"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;EndConditional/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <!-- Outer Else -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="300005"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;Else/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p>
            <w:r><w:t>Standard Member Benefits:</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>- Basic discounts apply</w:t></w:r>
        </w:p>

        <!-- Outer EndConditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="300006"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r><w:t>&lt;EndConditional/&gt;</w:t></w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
    </w:body>
</w:document>'''

    return xml_content

def create_simple_rels():
    """Create a simple document.xml.rels without glossary"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>'''

def create_docx_from_template(output_path, template_path, document_xml_content):
    """
    Create a .docx file by copying template and replacing document.xml

    Args:
        output_path: Path where to save the new .docx
        template_path: Path to the template .docx to copy from
        document_xml_content: XML string content for document.xml
    """
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        # Copy all files from template except document.xml, glossary, and rels
        with zipfile.ZipFile(template_path, 'r') as template_zip:
            for filename in template_zip.namelist():
                # Skip files we'll replace or don't need
                if (filename == 'word/document.xml' or
                    filename.startswith('word/glossary/') or
                    filename == 'word/_rels/document.xml.rels'):
                    continue

                content = template_zip.read(filename)
                docx.writestr(filename, content)

        # Write custom document.xml.rels
        docx.writestr('word/_rels/document.xml.rels', create_simple_rels())

        # Write custom document.xml
        docx.writestr('word/document.xml', document_xml_content)

    print(f"✓ Created {output_path}")

def main():
    """Main function to generate all test documents"""

    # Handle help flag
    if len(sys.argv) > 1 and sys.argv[1] in ('--help', '-h', 'help'):
        print(__doc__)
        sys.exit(0)

    # Configuration
    template_path = "DocumentAssemblerSdk.Examples/Example01_Basic/TemplateDocument.docx"
    output_dir = "DocumentAssemblerSdk.Tests/TestFiles"

    # Allow custom output directory from command line
    if len(sys.argv) > 1:
        output_dir = sys.argv[1]

    # Check template exists
    if not os.path.exists(template_path):
        print(f"❌ Error: Template not found at {template_path}")
        print("Please ensure you're running from the project root directory")
        sys.exit(1)

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    print(f"Creating test documents in: {output_dir}\n")

    # Generate test documents
    test_files = [
        ("DA270-ConditionalWithElse.docx", create_document_xml_with_else()),
        ("DA271-ConditionalWithElseNotMatch.docx", create_document_xml_with_else_notmatch()),
        ("DA272-NestedConditionalWithElse.docx", create_document_xml_nested_else()),
    ]

    for filename, xml_content in test_files:
        output_path = os.path.join(output_dir, filename)
        create_docx_from_template(output_path, template_path, xml_content)

    print(f"\n✅ All test documents created successfully!")
    print(f"\nTo run tests:")
    print(f"  dotnet test DocumentAssemblerSdk.Tests/DocumentAssemblerSdk.Tests.csproj --logger 'console;verbosity=normal'")

if __name__ == "__main__":
    main()
