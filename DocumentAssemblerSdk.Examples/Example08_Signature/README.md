# Example 08 – Signature Placeholders

This example demonstrates how to place the new `<Signature>` tag inside a DOCX template so
that, after assembly and PDF conversion, the placeholder text is replaced by an AcroForm
signature field ready for electronic signing.

## Files

| File | Description |
|------|-------------|
| `TemplateSignatureDocument.docx` | Template that contains two signature placeholders (`Richiedente`, `Responsabile`). |
| `Example08_Signature.csproj`, `Program.cs` | Console sample that assembles the template into `Output_SignatureDemo.docx`. |
| `DocumentAssembler/create_example08_signature.py` | Helper script to regenerate the template DOCX. |

## Regenerating the template

```bash
python DocumentAssembler/create_example08_signature.py
```

The script copies the base template from Example01 and injects the signature placeholders.

## Running the sample

```bash
cd DocumentAssembler/DocumentAssemblerSdk.Examples/Example08_Signature
dotnet run
```

The program loads `TemplateSignatureDocument.docx`, assembles it (no data fields are required)
and produces `Output_SignatureDemo.docx`. You can then run the usual DOCX → PDF pipeline
(Collabora, UnoServer, etc.) and the signature placeholders will become AcroForm `/Sig`
fields automatically.

## Syntax recap

```
<# <Signature Id="Responsabile"
              Label="Firma Responsabile"
              Width="220px"
              Height="60px" /> #>
```

- `Id` must be unique per document and becomes the form field name.
- `Label` controls the inline text (and the PDF tooltip).
- `Width` / `Height` accept `px`, `cm`, `mm`, `in`, or `emu`.

Place the markup wherever you need a signature line (inside a paragraph, table cell, etc.);
the conversion pipeline will measure the placeholder and swap it with an AcroForm widget.
