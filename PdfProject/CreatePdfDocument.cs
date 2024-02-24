using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Previewer;

namespace PdfProject;

public static class CreatePdfDocument
{
    public static async Task Create()
    {
        var document = Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.MarginLeft(3, Unit.Centimetre);
                page.MarginRight(1.5f, Unit.Centimetre);
                
                page.DefaultTextStyle(TextStyle.Default.FontFamily("Times New Roman").FontSize(10));
                
                page.Content().Column(column =>
                {
                    column.Item()
                        .AlignRight()
                        .Text("Додаток 3");

                    column.Item()
                        .Text(text =>
                        {
                            text.AlignCenter();
                            text.Line("НАЦІОНАЛЬНИЙ УНІВЕРСИТЕТ «ОДЕСЬКА ПОЛІТЕХНІКА»");
                            text.Line("НАВЧАЛЬНО-НАУКОВИЙ ІНСТИТУТ КОМП'ЮТЕРНИХ СИСТЕМ"); // From Model
                        });
                    
                    column.Item()
                        .Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.ConstantColumn(70);
                                columns.ConstantColumn(280);
                                columns.RelativeColumn();
                                columns.ConstantColumn(40);
                                columns.ConstantColumn(60);
                            });

                            table.Cell().Text("Спеціальність");
                            table.Cell().Border(1).AlignCenter().Text("Фармація, промислова фармація");
                            table.Cell();
                            table.Cell().Text("Семестр");
                            table.Cell().Border(1).AlignCenter().AlignMiddle().Text("1");
                            
                            table.Cell().Text("Освіт.програма");
                            table.Cell().Border(1).AlignCenter().Text("Технології фармацевтичних препаратів");
                            table.Cell();
                            table.Cell().Text("Курс");
                            table.Cell().Border(1).AlignCenter().AlignMiddle().Text("1");
                            
                            table.Cell().Text("Навчальний рік");
                            table.Cell().Border(1).AlignCenter().Text("2023");
                            table.Cell();
                            table.Cell().Text("Група");
                            table.Cell().Border(1).AlignCenter().AlignMiddle().Text("ЗЗЩХФ-181");
                        });

                    column.Item()
                        .Text(text =>
                        {
                            text.EmptyLine();
                            text.AlignCenter();
                            text.Line("Відомість обліку успішності № ІКС-3/2023");
                            text.Line("від: 25 жовтня 2023");
                        });
                    
                    column.Item()
                        .PaddingTop(-10)
                        .Row(row =>
                        {
                            row.RelativeItem().Text("з дисципліни");
                            row.ConstantItem(280)
                                .Border(1).AlignCenter().Text("Бази данних Бази данних Бази данних Бази данних Бази данних Бази данних Бази данних Бази данних");
                        });
                    
                    column.Item()
                        .Row(row =>
                        {
                            row.RelativeItem().Text("Форма семестрового контролю");
                            row.ConstantItem(280)
                                .Border(1).AlignCenter().Text("Залік");
                        });
                            
                    column.Item()
                        .Row(row =>
                        {
                            row.RelativeItem().Text("Підсумкову оцінку виставив викладач");
                            row.ConstantItem(280)
                                .Border(1).AlignCenter().Text("Григоренко Світлана Миколаївна");
                        });
                            
                    column.Item()
                        .Row(row =>
                        {
                            row.RelativeItem().Text("Поточний контроль здійснив викладач");
                            row.ConstantItem(280)
                                .Border(1).AlignCenter().Text("Євсеєв Олександр Володимирович");
                        });
                    
                    column.Item()
                        .PaddingTop(10)
                        .Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.ConstantColumn(20);
                                columns.ConstantColumn(200);
                                columns.ConstantColumn(30);
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                            });
                            
                            table.Header(header =>
                            {
                                header.Cell().RowSpan(3).Border(1).AlignCenter().AlignMiddle().Text("№");
                                header.Cell().RowSpan(3).Border(1).AlignCenter().AlignMiddle().Text("Прізвище, ініціали студента");
                                header.Cell().RowSpan(3).Border(1).AlignCenter().AlignMiddle().Text("№ інд. навч. плану");
                                header.Cell().ColumnSpan(3).Border(1).AlignCenter().Text("Підсумкова оцінка за шкалою");
                                header.Cell().RowSpan(3).Border(1).AlignCenter().AlignMiddle().Text("Підпис викладача");

                                header.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("Стобальна");
                                header.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("Національна");
                                header.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("ECTS");
                            });
                            
                            for (int i = 0; i < 30; i++)
                            {
                                table.Cell().Border(1).AlignCenter().Text($"{i+1}");
                                table.Cell().Border(1).PaddingLeft(5).Text("Костін Олексій Леонідович");
                                table.Cell().Border(1).AlignCenter().Text("51");
                                table.Cell().Border(1).AlignCenter().Text("60");
                                table.Cell().Border(1).AlignCenter().Text("3");
                                table.Cell().Border(1).AlignCenter().Text("E");
                                table.Cell().Border(1).Text("");
                            }
                        });

                    column.Item()
                        .Text(text =>
                        {
                            text.EmptyLine();
                            text.AlignCenter();
                            text.Line("Підсумки складання екзамену (заліку)");
                        });
                    
                    column.Item()
                        .PaddingTop(-10)
                        .Row(row =>
                        {
                            row.Spacing(50);
                            
                            row.ConstantItem(250)
                                .Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(35);
                                        columns.ConstantColumn(35);
                                        columns.ConstantColumn(35);
                                        columns.RelativeColumn();
                                        columns.RelativeColumn();
                                    });
                                    
                                    table.Header(header =>
                                    {
                                        header.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("Всього оцінок");
                                        header.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("Бали");
                                        header.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("Оцінка ECTS");
                                        header.Cell().ColumnSpan(2).Border(1).AlignCenter().Text("Оцінка за національною шкалою");

                                        header.Cell().Border(1).AlignCenter().Text("екзамен");
                                        header.Cell().Border(1).AlignCenter().Text("залік");
                                    });

                                    table.Cell().Border(1).AlignCenter().Text("0");
                                    table.Cell().Border(1).AlignCenter().Text("90-100");
                                    table.Cell().Border(1).AlignCenter().Text("A");
                                    table.Cell().Border(1).AlignCenter().Text("відмінно");
                                    table.Cell().RowSpan(5).Border(1).AlignCenter().AlignMiddle().Text("Зараховано");

                                    table.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("0");
                                    table.Cell().Border(1).AlignCenter().Text("82-89");
                                    table.Cell().Border(1).AlignCenter().Text("B");
                                    table.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("добре");
                                    
                                    table.Cell().Border(1).AlignCenter().Text("75-81");
                                    table.Cell().Border(1).AlignCenter().Text("C");
                                    
                                    table.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("1");
                                    table.Cell().Border(1).AlignCenter().Text("67-74");
                                    table.Cell().Border(1).AlignCenter().Text("D");
                                    table.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("задовільно");
                                    
                                    table.Cell().Border(1).AlignCenter().Text("60-66");
                                    table.Cell().Border(1).AlignCenter().Text("E");
                                    
                                    table.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("0");
                                    table.Cell().Border(1).AlignCenter().Text("35-59");
                                    table.Cell().Border(1).AlignCenter().Text("FX");
                                    table.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("незадовільно");
                                    table.Cell().RowSpan(2).Border(1).AlignCenter().AlignMiddle().Text("Не зараховано");
                                    
                                    table.Cell().Border(1).AlignCenter().Text("0-34");
                                    table.Cell().Border(1).AlignCenter().Text("F");
                                });

                            row.RelativeItem()
                                .Text(text =>
                                {
                                    text.Line("Директор інституту:");
                                    text.Line("Антощук Світлана Григорівна");
                                    text.EmptyLine();
                                    text.Line("_________________________________");
                                    text.EmptyLine();
                                    text.Line("Екзаменатор (викладач):");
                                    text.Line("Григоренко Світлана Миколаївна");
                                    text.EmptyLine();
                                    text.Line("_________________________________");
                                });
                        });
                });
            });
        });
        
        document.GeneratePdf("test.pdf");
        
        // Install Previewer for QuestPdf
        //await document.ShowInPreviewerAsync();
    }
}