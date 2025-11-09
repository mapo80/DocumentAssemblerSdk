using DocumentAssembler.Core;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace PerfMeasurementTool
{
    internal static class Program
    {
        private static readonly int MeasurementRuns = 10;

        private static readonly string TemplateDirectory = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".."));

        private const string TinyPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";

        private static readonly Scenario SimpleScenario = new Scenario(
            "Simple",
            ResolveTemplatePath("SimpleTemplate.docx"),
            new XElement("Customer",
                new XElement("CustomerId", "C123"),
                new XElement("Name", "Jane Doe"),
                new XElement("Email", "jane.doe@example.com"),
                new XElement("Phone", "555-0100"),
                new XElement("Photo", TinyPngBase64)));

        private static readonly Scenario ComplexScenario = new Scenario(
            "Complex",
            ResolveTemplatePath("ComplexTemplate.docx"),
            new XElement("Company",
                new XElement("Name", "Contoso Ltd."),
                new XElement("Address",
                    new XElement("Street", "1 Infinite Loop"),
                    new XElement("City", "Cupertino"),
                    new XElement("PostalCode", "95014")),
                new XElement("HasPremium", "true"),
                new XElement("PremiumCode", "PRM-001"),
                new XElement("Orders",
                    new XElement("Order",
                        new XElement("OrderId", "1001"),
                        new XElement("Product", "UltraWidget"),
                        new XElement("Quantity", "3"),
                        new XElement("HasDiscount", "true"),
                        new XElement("Discount", "5%"),
                        new XElement("Signature", TinyPngBase64),
                        new XElement("LineItems",
                            new XElement("Item",
                                new XElement("Description", "Main widget"),
                                new XElement("Amount", "199")),
                            new XElement("Item",
                                new XElement("Description", "Add-on"),
                                new XElement("Amount", "49")))),
                    new XElement("Order",
                        new XElement("OrderId", "1002"),
                        new XElement("Product", "MegaWidget"),
                        new XElement("Quantity", "2"),
                        new XElement("HasDiscount", "false"),
                        new XElement("Signature", TinyPngBase64),
                        new XElement("LineItems",
                            new XElement("Item",
                                new XElement("Description", "Core module"),
                                new XElement("Amount", "249")),
                            new XElement("Item",
                                new XElement("Description", "Support"),
                                new XElement("Amount", "79"))))),
                new XElement("Partners",
                    new XElement("Partner",
                        new XElement("Name", "Partner A"),
                        new XElement("Role", "Reseller")),
                    new XElement("Partner",
                        new XElement("Name", "Partner B"),
                        new XElement("Role", "Distributor"))),
                new XElement("Inventory",
                    new XElement("Item",
                        new XElement("Name", "WidgetA"),
                        new XElement("Stock", "257")),
                    new XElement("Item",
                        new XElement("Name", "WidgetB"),
                        new XElement("Stock", "143")))));


        private static int Main()
        {
            Console.WriteLine("PerfMeasurementTool - DocumentAssembler baseline profiling");
            Console.WriteLine($"{MeasurementRuns} runs per scenario (first run discarded). Timings in milliseconds.\n");

            RunScenario(SimpleScenario);
            RunScenario(ComplexScenario);

            return 0;
        }

        private static void RunScenario(Scenario scenario)
        {
            Console.WriteLine($"Scenario: {scenario.Name}");
            var measurements = MeasureScenario(scenario);
            Console.WriteLine($"Raw timings: {string.Join(", ", measurements.Select(m => m.ToString("F1", CultureInfo.InvariantCulture)))}");

            var trimmed = measurements.Skip(1).ToArray();
            var average = trimmed.Average();
            Console.WriteLine($"Average (drop first): {average.ToString("F1", CultureInfo.InvariantCulture)} ms\n");
        }

        private static double[] MeasureScenario(Scenario scenario)
        {
            var results = new double[MeasurementRuns];
            for (var i = 0; i < MeasurementRuns; i++)
            {
                var template = scenario.CreateTemplate();
                var dataCopy = new XElement(scenario.Data);

                var sw = Stopwatch.StartNew();
                DocumentAssembler.Core.DocumentAssembler.AssembleDocument(template, dataCopy, out var templateError);
                sw.Stop();

                if (templateError)
                {
                    Console.WriteLine($"Warning: template reported errors during {scenario.Name} run {i + 1}");
                }

                results[i] = sw.Elapsed.TotalMilliseconds;
            }

            return results;
        }

        private static string ResolveTemplatePath(string fileName) =>
            Path.Combine(TemplateDirectory, fileName);

        private sealed record Scenario(string Name, string TemplatePath, XElement Data)
        {
            public WmlDocument CreateTemplate() => new WmlDocument(TemplatePath);
        }

    }
}
