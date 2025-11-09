# Example 05: Conditional with Else

This example demonstrates the `<#Conditional#>` tag with `<#Else#>` functionality, allowing you to create if-else logic in your Word templates.

## What This Example Shows

1. **Basic Conditional with Else (Match)**: Display different content based on whether a field matches a specific value
2. **Conditional with Else (NotMatch)**: Display different content based on whether a field does NOT match a specific value
3. **Real-world scenarios**: Premium vs Standard membership, International vs Domestic shipping

## Template Tags Used

### Conditional with Match and Else
```
<#Conditional Select="./MembershipType" Match="Premium"#>
    Content shown when MembershipType equals "Premium"
<#Else/#>
    Content shown when MembershipType does NOT equal "Premium"
<#EndConditional/#>
```

### Conditional with NotMatch and Else
```
<#Conditional Select="./Country" NotMatch="USA"#>
    Content shown when Country is NOT "USA" (international)
<#Else/#>
    Content shown when Country IS "USA" (domestic)
<#EndConditional/#>
```

## Examples Included

### Example 1: Premium Member
- Template: `TemplateDocument.docx`
- Data: MembershipType = "Premium"
- Result: Shows premium benefits (free shipping, 24/7 support, exclusive products, double points)

### Example 2: Standard Member
- Template: `TemplateDocument.docx`
- Data: MembershipType = "Standard"
- Result: Shows standard benefits (conditional shipping, email support, newsletter)

### Example 3: International Customer
- Template: `TemplateNotMatchDocument.docx`
- Data: Country = "Canada" (not USA)
- Result: Shows international shipping info (7-14 days, customs fees)

### Example 4: US Customer
- Template: `TemplateNotMatchDocument.docx`
- Data: Country = "USA"
- Result: Shows domestic shipping info (2-5 days, free over $50, no customs)

## Running the Example

```bash
dotnet run
```

This will generate 4 output documents:
- `Output_PremiumMember.docx` - Premium membership letter
- `Output_StandardMember.docx` - Standard membership letter
- `Output_InternationalCustomer.docx` - International shipping info
- `Output_USCustomer.docx` - Domestic shipping info

## Key Concepts

### If-Else Logic
The `<#Else#>` tag allows you to create if-else logic:
- If the condition is TRUE → content before `<#Else#>` is shown
- If the condition is FALSE → content after `<#Else#>` is shown

### Match vs NotMatch
- **Match**: Condition is true when the field value EQUALS the specified value
- **NotMatch**: Condition is true when the field value DOES NOT EQUAL the specified value

### Use Cases
Perfect for:
- Membership tiers (Premium, Standard, Basic)
- Geographic variations (International vs Domestic)
- Status-based content (Active, Inactive, Pending)
- Any binary choice in your documents

## See Also
- Example04: Conditional blocks (without Else)
- Main README: Full documentation of all template tags
