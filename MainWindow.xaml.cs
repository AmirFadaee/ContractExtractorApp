using ClosedXML.Excel;
using Microsoft.Win32;
using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using UglyToad.PdfPig;

namespace ContractExtractorApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string selectedPath;
        private List<ContractInfo> extractedContracts = new List<ContractInfo>();
        public MainWindow()
        {
            InitializeComponent();
        }
        private void btnSelectPdf_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf"
            };
            if (ofd.ShowDialog() == true)
            {
                selectedPath = ofd.FileName;
                txtStatus.Text = "در حال پردازش فایل PDF...";
                extractedContracts = ExtractDataFromPdf(selectedPath);
                if (extractedContracts.Any())
                {
                    txtStatus.Text = "اطلاعات قرارداد با موفقیت استخراج شد.";
                }
                else
                {
                    txtStatus.Text = "هیچ اطلاعاتی از قرارداد پیدا نشد.";
                }
            }
        }
        private void btnSelectWord_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc"
            };
            if (ofd.ShowDialog() == true)
            {
                selectedPath = ofd.FileName;
                txtStatus.Text = "در حال پردازش فایل Word...";

                extractedContracts = ExtractDataFromWord(selectedPath);

                if (extractedContracts.Any())
                {
                    txtStatus.Text = "اطلاعات قرارداد Word با موفقیت استخراج شد.";
                }
                else
                {
                    txtStatus.Text = "هیچ اطلاعاتی از قرارداد Word پیدا نشد.";
                }
            }

        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (extractedContracts == null || extractedContracts.Count == 0)
            {
                MessageBox.Show("ابتدا یک قرارداد را استخراج کنید.");
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "Contracts.xlsx");


            // بررسی اینکه فایل وجود دارد یا نه
            XLWorkbook workbook;
            if (File.Exists(filePath))
            {
                workbook = new XLWorkbook(filePath);
            }
            else
            {
                workbook = new XLWorkbook(); // فایل جدید بساز
            }

            var contract = extractedContracts.First();

            // ساخت نام شیت با شماره واحد
            string sheetName = $"واحد {contract.UnitNumber}";
            if (workbook.Worksheets.Any(x => x.Name == sheetName))
            {

                int counter = 1;
                while (workbook.Worksheets.Any(x => x.Name == $"{sheetName}_{counter}"))
                {
                    counter++;
                }
                sheetName = $"{sheetName}_{counter}";
            }

            var ws = workbook.Worksheets.Add(sheetName);

            // تیتر بالای جدول
            ws.Cell(1, 1).Value = $"اطلاعات مربوط به قرارداد شماره واحد {contract.UnitNumber}";
            ws.Range("A1:H1").Merge().Style
                .Font.SetBold().Font.SetFontSize(14).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            ws.Cell(2, 1).Value = "نوع واحد";
            ws.Cell(2, 2).Value = "شماره واحد";
            ws.Cell(2, 3).Value = "طبقه";
            ws.Cell(2, 4).Value = "متراژ";
            ws.Cell(2, 5).Value = "قیمت نهایی";
            ws.Cell(2, 6).Value = "مراحل پرداخت";
            ws.Cell(2, 7).Value = "تاریخ خرید";
            ws.Cell(2, 8).Value = "تاریخ تحویل";

            ws.Cell(3, 1).Value = contract.ContractType;
            ws.Cell(3, 2).Value = contract.UnitNumber;
            ws.Cell(3, 3).Value = contract.Floor;
            ws.Cell(3, 4).Value = contract.Area;
            ws.Cell(3, 5).Value = contract.TotalPrice;
            ws.Cell(3, 6).Value = contract.InstallmentDescription;
            ws.Cell(3, 7).Value = contract.PurchaseDate;
            ws.Cell(3, 8).Value = contract.DeliveryDate;

            ws.Columns().AdjustToContents();
            workbook.SaveAs(filePath);

            MessageBox.Show("اطلاعات قرارداد با موفقیت به فایل Contracts.xlsx اضافه شد.");
        }


        private List<ContractInfo> ExtractDataFromPdf(string pdfPath)
        {
            var list = new List<ContractInfo>();
            using (var document = PdfDocument.Open(pdfPath))
            {
                string fullText = "";
                foreach (var page in document.GetPages())
                {
                    fullText += page.Text + "\n";
                }

                var contract = new ContractInfo
                {
                    ContractType = fullText.Contains("واحد تجاری") ? "تجاری" :
                                   fullText.Contains("واحد اداری") ? "اداری" : "نامشخص",

                    UnitNumber = ExtractField(fullText, @"واحد\s+([A-Z]?\d{1,3})"),

                    Floor = ExtractField(fullText, @"طبقه\s+([^\s،\n]+)"),

                    Area = ExtractDoubleField(fullText, @"مساحت\s+تقریبی\s+(\d{1,3}[.,]\d{1,3})"),

                    TotalPrice = ExtractTotalPrice(fullText),

                    PurchaseDate = ExtractPurchaseDate(fullText),

                    InstallmentDescription = ExtractInstallmentSummary(fullText),

                    DeliveryDate = CalculateDeliveryDate(fullText)
                };

                list.Add(contract);
            }

            return list;
        }

        private string ExtractPurchaseDate(string text)
        {
            return ExtractField(text, @"تاريخ\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})");
        }



        private List<ContractInfo> ExtractDataFromWord(string wordPath)
        {
            var list = new List<ContractInfo>();
            var doc = new Spire.Doc.Document(wordPath);
            doc.LoadFromFile(wordPath);
            string text = doc.GetText();
            var contract = new ContractInfo
            {
                ContractType = text.Contains("واحد تجاری") ? "تجاری" :
                   text.Contains("واحد اداری") ? "اداری" : "نامشخص",

                UnitNumber = ExtractField(text, @"واحد\s+([A-Z]?\d{1,3})"),

                Floor = ExtractField(text, @"طبقه\s+([^\s،]+)"),

                Area = ExtractDoubleField(text, @"مساحت\s+تقريبي\s*(\d+)\s*متر"),

                TotalPrice = ExtractTotalPriceFromTable(doc),

                PurchaseDate = ExtractPurchaseDate(text),

                InstallmentDescription = ExtractInstallmentSummary(text),

                DeliveryDate = CalculateDeliveryDate(text)
            };
            list.Add(contract);
            return list;
        }

        private double ExtractDoubleField(string text, string pattern)
        {
            var str = ExtractField(text, pattern);
            return double.TryParse(str.Replace(",", ""), out var val) ? val : 0;
        }

        private string ExtractInstallmentSummary(string text)
        {
            var match = Regex.Match(text,
                @"ماده\s*5[\-–:]?\s*(.*?)\s*(?=ماده\s*\d|تاریخ تحویل|پیوست|ضمیمه|$)",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            return match.Success
                ? match.Groups[1].Value.Trim().Replace("\r", "").Replace("\n", " ")
                : "اطلاعات یافت نشد";
        }


        private string CalculateDeliveryDate(string text)
        {
            string baseDateStr = ExtractPurchaseDate(text);
            string monthsText = ExtractField(text, @"7-1.*?(\d{1,3})\s*ماه", RegexOptions.Singleline);

            if (!string.IsNullOrEmpty(baseDateStr) && int.TryParse(monthsText, out int monthsToAdd))
            {
                try
                {
                    var parts = baseDateStr.Split('/');
                    int year = int.Parse(parts[2]);
                    int month = int.Parse(parts[1]);
                    int day = int.Parse(parts[0]);

                    var pc = new System.Globalization.PersianCalendar();
                    DateTime baseDate = pc.ToDateTime(year, month, day, 0, 0, 0, 0);

                    DateTime deliveryDate = baseDate.AddMonths(monthsToAdd);
                    return $"{pc.GetDayOfMonth(deliveryDate):00}/{pc.GetMonth(deliveryDate):00}/{pc.GetYear(deliveryDate)}";
                }
                catch
                {
                    return "خطا در محاسبه";
                }
            }

            return "نامشخص";
        }



        private string ExtractField(string text, string pattern, RegexOptions options = RegexOptions.IgnoreCase)
        {
            var match = Regex.Match(text, pattern, options);
            return match.Success ? match.Groups[1].Value.Trim() : "";
        }

        private decimal ExtractTotalPriceFromTable(Spire.Doc.Document doc)
        {
            foreach (Section section in doc.Sections)
            {
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    if (obj.DocumentObjectType == DocumentObjectType.Table)
                    {
                        var table = obj as Table;

                        if (table == null) continue;

                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            var row = table.Rows[r];

                            for (int c = 0; c < row.Cells.Count; c++)
                            {
                                string label = row.Cells[c].Paragraphs[0].Text.Trim();

                                if (label.Contains("قیمت پس از کسر تخفیف"))
                                {
                                    // حالا مقدار سلول پایین یعنی در ردیف r+1 و ستون c
                                    int nextRowIndex = r + 1;

                                    if (nextRowIndex < table.Rows.Count)
                                    {
                                        var nextRow = table.Rows[nextRowIndex];

                                        // اگر ستون c در ردیف بعدی وجود دارد
                                        if (c < nextRow.Cells.Count)
                                        {
                                            string valueText = nextRow.Cells[c].Paragraphs[0].Text.Trim();
                                            string cleaned = valueText.Replace(",", "").Replace("،", "").Replace(" ", "");

                                            if (decimal.TryParse(cleaned, out decimal result))
                                                return result;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return 0;
        }

        private decimal ExtractTotalPrice(string text)
        {
            var match = Regex.Match(text,
                @"(?:مبلغ\s+کل\s+ثمن\s+معامله\s*(?:پس\s+از\s+اعمال\s+تخفیف)?\s*[:\-]?\s*)(\d{1,3}(?:,\d{3})*)\s*ریال",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            if (match.Success)
                return decimal.TryParse(match.Groups[1].Value.Replace(",", ""), out var val) ? val : 0;

            return 0;
        }
    }
}
