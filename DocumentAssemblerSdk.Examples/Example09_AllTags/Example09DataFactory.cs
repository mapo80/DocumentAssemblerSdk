using System.Xml.Linq;

namespace Example09_AllTags;

public static class Example09DataFactory
{
    public static XElement CreateSampleData()
    {
        var report = new XElement("Report",
            new XElement("Customer",
                new XElement("FullName", "Mario Rossi"),
                new XElement("MiddleName", "Luigi"),
                new XElement("Title", "Chief Operations Officer"),
                new XElement("MembershipType", "Platinum"),
                new XElement("LoyaltyScore", "96 / 100"),
                new XElement("PremiumMessage", "Salvataggio automatico, revisioni istantanee e governance centralizzata eliminano gli errori manuali."),
                new XElement("Photo", BuildPlaceholderImage()),
                new XElement("Location",
                    new XElement("City", "Milano"),
                    new XElement("Country", "Italia"))
            ),
            new XElement("KPIs",
                new XElement("RevenueYTD", "€ 12,4M"),
                new XElement("Growth", "+18% QoQ"),
                new XElement("Satisfaction", "4,7 / 5"),
                new XElement("Retention", "96%")
            ),
            new XElement("Highlights",
                new XElement("Highlight",
                    new XElement("Title", "Innovation Sprint"),
                    new XElement("Impact", "+18% efficienza end-to-end"),
                    new XElement("Icon", "⚡")),
                new XElement("Highlight",
                    new XElement("Title", "AI Service Desk"),
                    new XElement("Impact", "Riduzione ticket -32%"),
                    new XElement("Icon", "🤖")),
                new XElement("Highlight",
                    new XElement("Title", "Sostenibilità"),
                    new XElement("Impact", "Filiera carbon neutral"),
                    new XElement("Icon", "🌿"))
            ),
            new XElement("Departments",
                new XElement("Department",
                    new XElement("Name", "Digital Factory"),
                    new XElement("Focus", "Automazione processi core"),
                    new XElement("HeadCount", 48),
                    new XElement("RiskLevel", "Low"),
                    new XElement("Budget",
                        new XElement("Allocated", "€ 4,2M"),
                        new XElement("Spent", "€ 3,1M")),
                    new XElement("Achievements",
                        new XElement("Achievement", "Lancio di 3 prodotti smart"),
                        new XElement("Achievement", "Riduzione errori QA del 25%")
                    )
                ),
                new XElement("Department",
                    new XElement("Name", "Customer Experience"),
                    new XElement("Focus", "Journey omnicanale"),
                    new XElement("HeadCount", 35),
                    new XElement("RiskLevel", "Moderate"),
                    new XElement("Budget",
                        new XElement("Allocated", "€ 3,0M"),
                        new XElement("Spent", "€ 2,5M")),
                    new XElement("Achievements",
                        new XElement("Achievement", "Nuovo portale mobile"),
                        new XElement("Achievement", "NPS +12 p.p.")
                    )
                )
            ),
            new XElement("Orders",
                new XElement("Order",
                    new XAttribute("code", "ORD-2042"),
                    new XAttribute("status", "Ready"),
                    new XElement("Product", "Modulo IoT Edge"),
                    new XElement("Quantity", 240),
                    new XElement("Price", "€ 1.320,00"),
                    new XElement("Notes", "Configurazione completa con firmware v3"),
                    new XElement("DeliveryWindow",
                        new XElement("Start", "05 Set 2024"),
                        new XElement("End", "30 Set 2024"))
                ),
                new XElement("Order",
                    new XAttribute("code", "ORD-2051"),
                    new XAttribute("status", "In produzione"),
                    new XElement("Product", "Console Collabora"),
                    new XElement("Quantity", 60),
                    new XElement("Price", "€ 890,00"),
                    new XElement("Notes", "Aggiornare layout per filiali"),
                    new XElement("DeliveryWindow",
                        new XElement("Start", "10 Ott 2024"),
                        new XElement("End", "15 Nov 2024"))
                )
            ),
            new XElement("Insights",
                new XElement("Text", "La domanda enterprise rimane robusta; prioritizzare bundle AI + assistenza 24/7."))
            ,
            new XElement("Charts",
                new XElement("Performance", BuildPlaceholderImage()),
                new XElement("Heatmap", BuildPlaceholderImage())),
            new XElement("Milestones",
                new XElement("Milestone",
                    new XAttribute("code", "M1"),
                    new XElement("Title", "Deploy piattaforma wave 2"),
                    new XElement("Owner", "PMO"),
                    new XElement("DueDate", "31 Ott 2024"),
                    new XElement("Status", "On Track")
                ),
                new XElement("Milestone",
                    new XAttribute("code", "M2"),
                    new XElement("Title", "UAT personalizzazioni banking"),
                    new XElement("Owner", "CX Lab"),
                    new XElement("DueDate", "15 Nov 2024"),
                    new XElement("Status", "At Risk")
                )
            ),
            new XElement("Attachments",
                new XElement("Attachment",
                    new XElement("Label", "Executive memo"),
                    new XElement("Description", "Sintesi portafoglio e metriche di salute"),
                    new XElement("Url", "https://example.com/memo.pdf")),
                new XElement("Attachment",
                    new XElement("Label", "Roadmap digitale"),
                    new XElement("Description", "Versione aggiornata Q4"),
                    new XElement("Url", "https://example.com/roadmap"))
            ),
            new XElement("Approvals",
                new XElement("PrimarySigner", "Luca Bianchi"),
                new XElement("BackupSigner", "Sara Greco"))
        );
        return new XElement("Data", report);
    }

    public static string BuildPlaceholderImage()
    {
        var bytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAI0lEQVR4nGNgGAWjYBSMglEwCkbDqADEpGAYRUM1gWgAAFOlAh7Ate0oAAAAAElFTkSuQmCC");
        return Convert.ToBase64String(bytes);
    }
}
