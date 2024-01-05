using System.Reflection;
using QuestPDF.Infrastructure;

namespace PdfProject;

internal static class Program
{
    public static async Task Main(string[] args)
    {
        QuestPDF.Settings.License = LicenseType.Community;
        await CreatePdfDocument.Create();
    }
}