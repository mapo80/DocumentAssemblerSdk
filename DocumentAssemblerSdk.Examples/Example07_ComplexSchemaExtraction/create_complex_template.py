#!/usr/bin/env python3
"""
Creates a complex DocumentAssembler template for testing schema extraction.
This template includes: Content, Image, Table, Repeat, Conditional with Else, nested structures.
"""

import zipfile
import os
from datetime import datetime

def create_complex_template(output_path="ComplexTemplate.docx"):
    """
    Creates a comprehensive template with all DocumentAssembler tag types.

    Template represents an E-Commerce Order Report with:
    - Customer information (Content tags)
    - Customer photo (Image tag)
    - Order items table (Table tag)
    - Shipping addresses (Repeat tag)
    - Membership benefits (Conditional with Else)
    - Nested conditionals for payment methods
    - Product categories with nested items
    """

    # Document XML with comprehensive template structure
    document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <!-- Header -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>📦 E-Commerce Order Report</w:t></w:r>
    </w:p>

    <!-- Order Information -->
    <w:p>
      <w:r><w:t>Order ID: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/OrderID" Optional="false"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Order Date: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/OrderDate"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Total Amount: $</w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/TotalAmount"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Customer Section -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>👤 Customer Information</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:t>Name: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/Customer/FirstName"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
      <w:r><w:t> </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/Customer/LastName"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Email: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/Customer/Email"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Phone: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/Customer/Phone"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Customer Photo -->
    <w:p>
      <w:r><w:t>Customer Photo:</w:t></w:r>
    </w:p>
    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Image Select="Order/Customer/ProfilePhoto" MaxWidth="200px" MaxHeight="200px" Align="center"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Membership Status with Conditional Else -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>⭐ Membership Benefits</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Conditional Select="Order/Customer/MembershipType" Match="Premium"&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>✓ PREMIUM Member Benefits:</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>  • 20% discount on all purchases</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>  • Free express shipping worldwide</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>  • Priority customer support 24/7</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>  • Exclusive access to new products</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Else&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Standard Member Benefits:</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>  • 5% discount on purchases</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>  • Standard shipping rates</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>  • Email support (24-48h response)</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;EndConditional&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Order Items Table -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>📋 Order Items</w:t></w:r>
    </w:p>

    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="pct"/>
      </w:tblPr>

      <!-- Header Row -->
      <w:tr>
        <w:tc><w:p><w:r><w:t>Product Name</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>SKU</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Quantity</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Unit Price</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Subtotal</w:t></w:r></w:p></w:tc>
      </w:tr>

      <!-- Data Row (prototype) -->
      <w:tr>
        <w:tc>
          <w:p>
            <w:sdt>
              <w:sdtContent>
                <w:r><w:t>&lt;Table Select="Order/Items/Item"/&gt;</w:t></w:r>
              </w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
        <w:tc><w:p></w:p></w:tc>
        <w:tc><w:p></w:p></w:tc>
        <w:tc><w:p></w:p></w:tc>
        <w:tc><w:p></w:p></w:tc>
      </w:tr>

      <w:tr>
        <w:tc>
          <w:p>
            <w:sdt>
              <w:sdtContent>
                <w:r><w:t>&lt;Content Select="ProductName"/&gt;</w:t></w:r>
              </w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:sdt>
              <w:sdtContent>
                <w:r><w:t>&lt;Content Select="SKU"/&gt;</w:t></w:r>
              </w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:sdt>
              <w:sdtContent>
                <w:r><w:t>&lt;Content Select="Quantity"/&gt;</w:t></w:r>
              </w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>$</w:t></w:r>
            <w:sdt>
              <w:sdtContent>
                <w:r><w:t>&lt;Content Select="UnitPrice"/&gt;</w:t></w:r>
              </w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:r><w:t>$</w:t></w:r>
            <w:sdt>
              <w:sdtContent>
                <w:r><w:t>&lt;Content Select="Subtotal"/&gt;</w:t></w:r>
              </w:sdtContent>
            </w:sdt>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>

    <!-- Product Categories with Repeat -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>🏷️ Product Categories</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Repeat Select="Order/Categories/Category"&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Category: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="CategoryName"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>  Items in this category: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="ItemCount"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>  Description: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Description"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>---</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;EndRepeat&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Shipping Addresses with Repeat -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>🚚 Shipping Addresses</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Repeat Select="Order/ShippingAddresses/Address"&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Type: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="AddressType"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Street"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="City"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
      <w:r><w:t>, </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="State"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
      <w:r><w:t> </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="ZipCode"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Country"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>---</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;EndRepeat&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Payment Information with Nested Conditionals -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>💳 Payment Information</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:t>Payment Method: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/Payment/Method"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Nested Conditional for Credit Card -->
    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Conditional Select="Order/Payment/Method" Match="CreditCard"&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Card Type: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/Payment/CardType"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Last 4 Digits: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/Payment/Last4Digits"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Nested Conditional for Card Type -->
    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Conditional Select="Order/Payment/CardType" Match="Visa"&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>✓ Visa benefits: No foreign transaction fees</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Else&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Standard card processing applies</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;EndConditional&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Else&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>Alternative payment method used</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;EndConditional&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Special Notes (Optional Field) -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>📝 Special Notes</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/SpecialNotes"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Delivery Instructions (Conditional) -->
    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Conditional Select="Order/ExpressDelivery" Match="true"&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:p>
      <w:r><w:t>⚡ EXPRESS DELIVERY REQUESTED</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Estimated delivery: Next business day</w:t></w:r>
    </w:p>

    <w:p>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;EndConditional&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <!-- Footer -->
    <w:p>
      <w:r><w:t>────────────────────────────</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Report generated: </w:t></w:r>
      <w:sdt>
        <w:sdtContent>
          <w:r><w:t>&lt;Content Select="Order/ReportGeneratedDate"/&gt;</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>

    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    # Create minimal DOCX structure
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        # [Content_Types].xml
        docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''')

        # _rels/.rels
        docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')

        # word/_rels/document.xml.rels
        docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''')

        # word/document.xml
        docx.writestr('word/document.xml', document_xml)

    print(f"✓ Created complex template: {output_path}")
    print(f"  Size: {os.path.getsize(output_path)} bytes")
    print(f"\nTemplate includes:")
    print("  • Order information (Content tags)")
    print("  • Customer details with profile photo (Image tag)")
    print("  • Membership benefits (Conditional with Else)")
    print("  • Order items table (Table tag)")
    print("  • Product categories (Repeat tag)")
    print("  • Multiple shipping addresses (Repeat tag)")
    print("  • Payment info with nested conditionals")
    print("  • Optional special notes")
    print("  • Express delivery conditional")

if __name__ == "__main__":
    import sys
    output_file = sys.argv[1] if len(sys.argv) > 1 else "ComplexTemplate.docx"
    create_complex_template(output_file)
