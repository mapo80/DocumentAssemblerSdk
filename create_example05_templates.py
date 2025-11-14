#!/usr/bin/env python3
"""
Create template documents for Example05_ConditionalElse
"""

import zipfile
import os
import shutil

def create_template_document():
    """Create TemplateDocument.docx for basic Conditional with Else"""

    template_path = "DocumentAssemblerSdk.Examples/Example01_Basic/TemplateDocument.docx"
    output_path = "DocumentAssemblerSdk.Examples/Example05_ConditionalElse/TemplateDocument.docx"

    # Copy template structure
    shutil.copy(template_path, output_path)

    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <!-- Title -->
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="32"/>
                </w:rPr>
                <w:t>Membership Benefits Letter</w:t>
            </w:r>
        </w:p>

        <w:p/>

        <!-- Customer Name -->
        <w:p>
            <w:r>
                <w:t>Dear </w:t>
            </w:r>
        </w:p>
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="1"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
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

        <w:p/>

        <!-- Conditional: Premium Member -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="2"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
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

        <!-- Premium content -->
        <w:p>
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>Thank you for being a PREMIUM member!</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>As a Premium member, you enjoy:</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Free shipping on all orders</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• 24/7 Priority customer support</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Exclusive access to premium products</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Double reward points on every purchase</w:t>
            </w:r>
        </w:p>

        <!-- Else tag -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="3"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
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

        <!-- Standard content -->
        <w:p>
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>Thank you for being a valued member!</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>Your current benefits include:</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Standard shipping on orders over $50</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Email support (24-hour response time)</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Monthly newsletter with special offers</w:t>
            </w:r>
        </w:p>

        <w:p/>

        <w:p>
            <w:r>
                <w:rPr><w:i/></w:rPr>
                <w:t>Consider upgrading to Premium for enhanced benefits!</w:t>
            </w:r>
        </w:p>

        <!-- EndConditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="4"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
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

        <w:p/>

        <!-- Current points -->
        <w:p>
            <w:r>
                <w:t>Your current reward points: </w:t>
            </w:r>
        </w:p>
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="5"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>&lt;Content Select="./Points"/&gt;</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p/>

        <!-- Closing -->
        <w:p>
            <w:r>
                <w:t>We appreciate your business!</w:t>
            </w:r>
        </w:p>

        <w:p/>

        <w:p>
            <w:r>
                <w:t>Sincerely,</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>Customer Relations Team</w:t>
            </w:r>
        </w:p>

        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>'''

    # Update document.xml in the docx
    # Read the existing docx, replace document.xml, and write to new file
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract all files
        with zipfile.ZipFile(output_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        # Write new document.xml
        with open(os.path.join(tmpdir, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
            f.write(document_xml)

        # Create new docx
        os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    docx.write(file_path, arcname)

    print(f"✓ Created {output_path}")

def create_notmatch_document():
    """Create TemplateNotMatchDocument.docx for NotMatch with Else"""

    template_path = "DocumentAssemblerSdk.Examples/Example01_Basic/TemplateDocument.docx"
    output_path = "DocumentAssemblerSdk.Examples/Example05_ConditionalElse/TemplateNotMatchDocument.docx"

    # Copy template structure
    shutil.copy(template_path, output_path)

    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <!-- Title -->
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="32"/>
                </w:rPr>
                <w:t>Shipping Information</w:t>
            </w:r>
        </w:p>

        <w:p/>

        <!-- Customer Name -->
        <w:p>
            <w:r>
                <w:t>Dear </w:t>
            </w:r>
        </w:p>
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="1"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
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

        <w:p/>

        <!-- Conditional: International (NotMatch USA) -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="2"/>
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

        <!-- International content -->
        <w:p>
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>International Shipping</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>Your order will be shipped internationally to: </w:t>
            </w:r>
        </w:p>
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="10"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
                <w:text/>
            </w:sdtPr>
            <w:sdtContent>
                <w:p>
                    <w:r>
                        <w:t>&lt;Content Select="./Country"/&gt;</w:t>
                    </w:r>
                </w:p>
            </w:sdtContent>
        </w:sdt>

        <w:p/>

        <w:p>
            <w:r>
                <w:t>Please note:</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• International shipping may take 7-14 business days</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Additional customs fees may apply</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Tracking information will be provided</w:t>
            </w:r>
        </w:p>

        <!-- Else tag -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="3"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
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

        <!-- Domestic (USA) content -->
        <w:p>
            <w:r>
                <w:rPr><w:b/></w:rPr>
                <w:t>Domestic Shipping (USA)</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>Your order will be shipped within the United States.</w:t>
            </w:r>
        </w:p>

        <w:p/>

        <w:p>
            <w:r>
                <w:t>Shipping benefits:</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Fast delivery: 2-5 business days</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• Free shipping on orders over $50</w:t>
            </w:r>
        </w:p>

        <w:p>
            <w:r>
                <w:t>• No customs or international fees</w:t>
            </w:r>
        </w:p>

        <!-- EndConditional -->
        <w:sdt>
            <w:sdtPr>
                <w:id w:val="4"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>
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

        <w:p/>

        <!-- Closing -->
        <w:p>
            <w:r>
                <w:t>Thank you for your order!</w:t>
            </w:r>
        </w:p>

        <w:p/>

        <w:p>
            <w:r>
                <w:t>Shipping Department</w:t>
            </w:r>
        </w:p>

        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>'''

    # Update document.xml in the docx
    # Read the existing docx, replace document.xml, and write to new file
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract all files
        with zipfile.ZipFile(output_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        # Write new document.xml
        with open(os.path.join(tmpdir, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
            f.write(document_xml)

        # Create new docx
        os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    docx.write(file_path, arcname)

    print(f"✓ Created {output_path}")


def main():
    print("Creating Example05_ConditionalElse template documents...")
    print()

    create_template_document()
    create_notmatch_document()

    print()
    print("✓ All template documents created successfully!")
    print()
    print("Templates created:")
    print("  1. TemplateDocument.docx - Basic Conditional with Else (Match)")
    print("  2. TemplateNotMatchDocument.docx - Conditional with Else (NotMatch)")

if __name__ == "__main__":
    main()
