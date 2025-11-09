# Example 07: Complex Schema Extraction

This example demonstrates **comprehensive schema extraction** from a real-world, production-grade DocumentAssembler template that includes **all tag types** and **complex nested structures**.

## Template Overview

The example uses an **E-Commerce Order Report** template that represents a realistic business scenario with:

### 📋 Template Structure

1. **Order Information** (Content tags)
   - Order ID (Required field)
   - Order Date
   - Total Amount

2. **Customer Information** (Content + Image)
   - First Name, Last Name
   - Email, Phone
   - Profile Photo (Image with sizing metadata)

3. **Membership Benefits** (Conditional with Else)
   - Premium vs Standard membership tiers
   - Different benefits for each tier

4. **Order Items** (Table)
   - Dynamic table with product details
   - Multiple fields per row (Name, SKU, Quantity, Price, Subtotal)

5. **Product Categories** (Repeat)
   - Repeating block for each category
   - Category name, item count, description

6. **Shipping Addresses** (Repeat)
   - Multiple delivery addresses
   - Full address structure (Street, City, State, Zip, Country)

7. **Payment Information** (Nested Conditionals)
   - Payment method selection
   - Credit card details (nested conditional)
   - Card type benefits (double-nested conditional with Else)

8. **Special Features**
   - Optional special notes
   - Express delivery conditional
   - Report generation date

## What This Example Tests

✅ **All Tag Types**
- Content (simple and required fields)
- Image (with metadata attributes)
- Table (dynamic row generation)
- Repeat (multiple repeating blocks)
- Conditional (with Else tags)

✅ **Complex Scenarios**
- Nested conditionals (3 levels deep)
- Multiple repeating structures
- Optional vs required fields
- Attribute extraction (MaxWidth, MaxHeight, Align, etc.)

✅ **Performance**
- Extraction time measurement
- Large template handling
- Efficient parsing validation

✅ **Output Quality**
- XML well-formedness validation
- Hierarchical structure verification
- Comment annotations for optional/repeating fields

## Running the Example

```bash
cd DocumentAssemblerSdk.Examples/Example07_ComplexSchemaExtraction
dotnet run
```

The example will:
1. **Generate the template** (if not present) using the Python script
2. **Extract the schema** using TemplateSchemaExtractor
3. **Display detailed analysis** of discovered fields
4. **Save the XML schema** to `Schema_ComplexTemplate.xml`
5. **Validate** the output

## Expected Output

```
=== Example 07: Complex Schema Extraction ===

Loading template: ComplexTemplate.docx
Template size: X bytes

Extracting XML schema...
✓ Extraction completed in XXms

═══════════════════════════════════════
           EXTRACTION SUMMARY
═══════════════════════════════════════
Total fields discovered: 35+
Root element name: Order

Fields by Type:
───────────────────────────────────────
  Conditional     :   4 field(s)
  Content         :  28 field(s)
  Image           :   1 field(s)
  Repeat          :   2 field(s)
  Table           :   1 field(s)

Optional fields: 33
Required fields: 1

Repeating structures: 3
```

## Generated XML Schema

The extractor produces a well-structured XML template like:

```xml
<Data>
  <Order>
    <OrderID>[value]</OrderID>  <!-- REQUIRED -->
    <OrderDate>[value] <!-- Optional --></OrderDate>
    <TotalAmount>[value] <!-- Optional --></TotalAmount>
    <Customer>
      <FirstName>[value] <!-- Optional --></FirstName>
      <LastName>[value] <!-- Optional --></LastName>
      <Email>[value] <!-- Optional --></Email>
      <Phone>[value] <!-- Optional --></Phone>
      <ProfilePhoto>[value] <!-- Optional, Base64 encoded image --></ProfilePhoto>
      <MembershipType>[value] <!-- Optional --></MembershipType>
    </Customer>
    <Items>
      <Item> <!-- Repeating -->
        <ProductName>[value] <!-- Optional --></ProductName>
        <SKU>[value] <!-- Optional --></SKU>
        <Quantity>[value] <!-- Optional --></Quantity>
        <UnitPrice>[value] <!-- Optional --></UnitPrice>
        <Subtotal>[value] <!-- Optional --></Subtotal>
      </Item>
      <Item> <!-- Repeating -->
        <ProductName>[value] <!-- Optional --></ProductName>
        <SKU>[value] <!-- Optional --></SKU>
        <Quantity>[value] <!-- Optional --></Quantity>
        <UnitPrice>[value] <!-- Optional --></UnitPrice>
        <Subtotal>[value] <!-- Optional --></Subtotal>
      </Item>
    </Items>
    <Categories>
      <Category> <!-- Repeating -->
        <CategoryName>[value] <!-- Optional --></CategoryName>
        <ItemCount>[value] <!-- Optional --></ItemCount>
        <Description>[value] <!-- Optional --></Description>
      </Category>
      <Category> <!-- Repeating -->
        <CategoryName>[value] <!-- Optional --></CategoryName>
        <ItemCount>[value] <!-- Optional --></ItemCount>
        <Description>[value] <!-- Optional --></Description>
      </Category>
    </Categories>
    <ShippingAddresses>
      <Address> <!-- Repeating -->
        <AddressType>[value] <!-- Optional --></AddressType>
        <Street>[value] <!-- Optional --></Street>
        <City>[value] <!-- Optional --></City>
        <State>[value] <!-- Optional --></State>
        <ZipCode>[value] <!-- Optional --></ZipCode>
        <Country>[value] <!-- Optional --></Country>
      </Address>
      <Address> <!-- Repeating -->
        <AddressType>[value] <!-- Optional --></AddressType>
        <Street>[value] <!-- Optional --></Street>
        <City>[value] <!-- Optional --></City>
        <State>[value] <!-- Optional --></State>
        <ZipCode>[value] <!-- Optional --></ZipCode>
        <Country>[value] <!-- Optional --></Country>
      </Address>
    </ShippingAddresses>
    <Payment>
      <Method>[value] <!-- Optional --></Method>
      <CardType>[value] <!-- Optional --></CardType>
      <Last4Digits>[value] <!-- Optional --></Last4Digits>
    </Payment>
    <SpecialNotes>[value] <!-- Optional --></SpecialNotes>
    <ExpressDelivery>[value] <!-- Optional --></ExpressDelivery>
    <ReportGeneratedDate>[value] <!-- Optional --></ReportGeneratedDate>
  </Order>
</Data>
```

## Key Validations

The example performs comprehensive validations:

1. ✅ **All tag types detected** - Ensures Content, Image, Table, Repeat, and Conditional tags are found
2. ✅ **XML well-formedness** - Validates generated XML structure
3. ✅ **Performance check** - Measures extraction time
4. ✅ **Attribute preservation** - Verifies metadata attributes are captured
5. ✅ **Hierarchical structure** - Confirms proper nesting of elements
6. ✅ **Optional/Required distinction** - Validates field classification

## Technical Details

### Template Complexity Metrics

- **Total tags**: 35+ DocumentAssembler tags
- **Nesting depth**: 3 levels (triple-nested conditionals)
- **Repeating blocks**: 3 (Items, Categories, Addresses)
- **Table rows**: Dynamic with 5 columns
- **Conditional branches**: 4 with Else tags
- **Image fields**: 1 with sizing metadata

### Performance Characteristics

- **Extraction time**: < 50ms typical
- **Template size**: ~15-20 KB
- **Memory usage**: Minimal (single-pass algorithm)
- **XML output size**: ~2-3 KB

## Use Cases

This example demonstrates how to:

1. **Understand complex templates** - Reverse-engineer production templates
2. **Generate API contracts** - Create data specifications for integrations
3. **Validate template coverage** - Ensure all fields are documented
4. **Migrate templates** - Extract structure for template versioning
5. **Train developers** - Onboard team members with clear field mapping

## Prerequisites

- **.NET 8.0** SDK
- **Python 3.x** (for template generation)
- **DocumentAssemblerSdk** project reference

## Files Generated

After running:
- `ComplexTemplate.docx` - The test template (auto-generated if missing)
- `Schema_ComplexTemplate.xml` - Extracted XML schema

## Notes

- The template is **auto-generated** using `create_complex_template.py` if not present
- Template uses **both formats**: `<#Tag#>` custom format and `<Tag />` XML format
- All **edge cases** are covered: nested structures, optional fields, metadata attributes
- The example serves as a **comprehensive validation suite** for the schema extractor

## Real-World Application

This example represents a **realistic e-commerce scenario** that could be used for:
- Order confirmation emails/documents
- Invoice generation
- Shipping manifests
- Customer reports
- Warehouse picking lists
- Financial reconciliation reports

The complexity level matches production systems where templates combine multiple data sources and conditional logic.
