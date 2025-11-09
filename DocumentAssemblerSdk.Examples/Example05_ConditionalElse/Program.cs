using DocumentAssembler.Core;
using System.Xml.Linq;

namespace Example05_ConditionalElse
{
    /// <summary>
    /// Example 05: Conditional with Else
    ///
    /// This example demonstrates the use of Conditional tags with Else blocks
    /// for if-else logic in document templates.
    ///
    /// Features demonstrated:
    /// - Basic Conditional with Else (Match)
    /// - Conditional with Else using NotMatch
    /// - Nested Conditionals with Else
    /// - Multiple conditions with different membership types
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("===========================================");
            Console.WriteLine("Example 05: Conditional with Else");
            Console.WriteLine("===========================================\n");

            // Example 1: Premium Member
            Console.WriteLine("Example 1: Premium Member");
            Console.WriteLine("------------------------------------------");
            GenerateDocument(
                "TemplateDocument.docx",
                CreatePremiumMemberData(),
                "Output_PremiumMember.docx"
            );

            // Example 2: Standard Member
            Console.WriteLine("\nExample 2: Standard Member");
            Console.WriteLine("------------------------------------------");
            GenerateDocument(
                "TemplateDocument.docx",
                CreateStandardMemberData(),
                "Output_StandardMember.docx"
            );

            // Example 3: International Customer
            Console.WriteLine("\nExample 3: International Customer");
            Console.WriteLine("------------------------------------------");
            GenerateDocument(
                "TemplateNotMatchDocument.docx",
                CreateInternationalCustomerData(),
                "Output_InternationalCustomer.docx"
            );

            // Example 4: US Customer
            Console.WriteLine("\nExample 4: US Customer");
            Console.WriteLine("------------------------------------------");
            GenerateDocument(
                "TemplateNotMatchDocument.docx",
                CreateUSCustomerData(),
                "Output_USCustomer.docx"
            );

            Console.WriteLine("\n===========================================");
            Console.WriteLine("All examples completed successfully!");
            Console.WriteLine("===========================================");
        }

        static void GenerateDocument(string templateName, XElement data, string outputName)
        {
            try
            {
                // Load template
                var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, templateName);
                var wmlTemplate = new WmlDocument(templatePath);

                // Assemble document
                var wmlAssembled = DocumentAssembler.Core.DocumentAssembler.AssembleDocument(
                    wmlTemplate,
                    data,
                    out bool templateError
                );

                // Save output
                var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, outputName);
                wmlAssembled.SaveAs(outputPath);

                if (templateError)
                {
                    Console.WriteLine($"⚠️  Template errors detected!");
                    Console.WriteLine($"   Output: {outputPath}");
                }
                else
                {
                    Console.WriteLine($"✓ Generated: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: {ex.Message}");
            }
        }

        static XElement CreatePremiumMemberData()
        {
            return new XElement("Customer",
                new XElement("Name", "Alice Johnson"),
                new XElement("MembershipType", "Premium"),
                new XElement("Country", "USA"),
                new XElement("Points", "5000"),
                new XElement("HighValueCustomer", "True")
            );
        }

        static XElement CreateStandardMemberData()
        {
            return new XElement("Customer",
                new XElement("Name", "Bob Smith"),
                new XElement("MembershipType", "Standard"),
                new XElement("Country", "USA"),
                new XElement("Points", "500"),
                new XElement("HighValueCustomer", "False")
            );
        }

        static XElement CreateInternationalCustomerData()
        {
            return new XElement("Customer",
                new XElement("Name", "Carlos Rodriguez"),
                new XElement("MembershipType", "Premium"),
                new XElement("Country", "Mexico"),
                new XElement("Points", "3000"),
                new XElement("HighValueCustomer", "True")
            );
        }

        static XElement CreateUSCustomerData()
        {
            return new XElement("Customer",
                new XElement("Name", "Diana Wilson"),
                new XElement("MembershipType", "Standard"),
                new XElement("Country", "USA"),
                new XElement("Points", "1200"),
                new XElement("HighValueCustomer", "False")
            );
        }
    }
}
