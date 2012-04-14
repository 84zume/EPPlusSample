using System;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

namespace EPPlusSample
{
    class Program
    {
        static void Main()
        {
            var outputDir = new DirectoryInfo(@"C:\temp\EPPlusSample");

            if (!outputDir.Exists)
            {
                throw new Exception(@"サンプルを出力するフォルダー(C:\temp\SampleApp)を作成してください。");
            }
            var outputPath = RunSample1(outputDir);
            Console.WriteLine("{0}を作成しました。", outputPath);
            Console.ReadKey();
        }

        private static string RunSample1(FileSystemInfo outputDir)
        {
            var newFile = new FileInfo(outputDir.FullName + @"\kakeibo.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\kakeibo.xlsx");
            }

            using (var package = new ExcelPackage(newFile))
            {
                var ws = package.Workbook.Worksheets.Add("4月の家計簿");

                //ヘッダー部分
                ws.Cells[1, 1].Value = "カテゴリー";
                ws.Cells[1, 2].Value = "商品名";
                ws.Cells[1, 3].Value = "値段";
                ws.Cells[1, 4].Value = "利用日";

                //データの部分
                ws.Cells["A2"].Value = "水道光熱費";
                ws.Cells["B2"].Value = "電気";
                ws.Cells["C2"].Value = 3000;
                ws.Cells["D2"].Value = new DateTime(2012, 4, 1);

                ws.Cells["A3"].Value = "水道光熱費";
                ws.Cells["B3"].Value = "ガス";
                ws.Cells["C3"].Value = 4000;
                ws.Cells["D3"].Value = new DateTime(2012, 4, 1);

                ws.Cells["A4"].Value = "水道光熱費";
                ws.Cells["B4"].Value = "水道";
                ws.Cells["C4"].Value = 1000;
                ws.Cells["D4"].Value = new DateTime(2012, 4, 1);

                ws.Cells["A5"].Value = "食費";
                ws.Cells["B5"].Value = "天ぷら";
                ws.Cells["C5"].Value = 800;
                ws.Cells["D5"].Value = new DateTime(2012, 4, 2);

                ws.Cells["A6"].Value = "食費";
                ws.Cells["B6"].Value = "かつ丼";
                ws.Cells["C6"].Value = 600;
                ws.Cells["D6"].Value = new DateTime(2012, 4, 3);

                ws.Cells["A7"].Value = "食費";
                ws.Cells["B7"].Value = "うどん";
                ws.Cells["C7"].Value = 300;
                ws.Cells["D7"].Value = new DateTime(2012, 4, 4);

                //合計金額を表示
                ws.Cells["A9"].Value = "合計";
                ws.Cells["B9"].Formula = "SUM(C2:C7)";

                //フォーマット・スタイル定義
                ws.Cells["D2:D7"].Style.Numberformat.Format = "yyyy年MM月dd日";
                ws.Cells["A1:D10"].Style.Font.Name = "ＭＳ Ｐゴシック";
                ws.Cells["A1:D7"].AutoFilter = true;
                ws.Cells.AutoFitColumns(0);

                //デザインを付ける
                ws.Cells["A1:D7"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D7"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D7"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D7"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(Color.Bisque);

                //図を描く
                var chart = (ws.Drawings.AddChart("4月のお金", eChartType.BarStacked) as ExcelBarChart);
                if (chart == null)
                {
                    throw new Exception(@"chartオブジェクトが作成されていません。");
                }
                chart.Title.Text = "4月のお金";
                chart.SetPosition(0, 400);
                chart.SetSize(300, 300);
                chart.Series.Add("C2:C7", "D2:D7");

                //ファイルのプロパティを設定
                package.Workbook.Properties.Title = "家計簿";
                package.Workbook.Properties.Author = "84zume";
                package.Save();
            }

            return newFile.FullName;
        }
    }
}
