using System;
using Excel = Microsoft.Office.Interop.Excel;
using MongoDB.Bson;
using MongoDB.Driver;
using Topshelf;
using System.Runtime.InteropServices;

namespace mula_generator
{
    class Program
    {
        static void Main()
        {
            HostFactory.Run(hostConfigurator =>
            {
                hostConfigurator.Service<Mula>(s =>
                {
                    s.ConstructUsing(name => new Mula());
                    s.WhenStarted(m => m.Start());
                    s.WhenStopped(m => m.Stop());
                });
                hostConfigurator.RunAsLocalSystem();

                hostConfigurator.SetDescription("Mula spreadsheet generator");
                hostConfigurator.SetDisplayName("Mula Generator");
                hostConfigurator.SetServiceName("mula-gen");
            });
        }
    }
    
    class Mula
    {
        private bool _doWork;
        public Mula() { }
        public async void Start()
        {
            var conn = new MongoClient("mongodb://mula.rlsolutions.com:27017").GetDatabase("quixpense");
            var exports = conn.GetCollection<BsonDocument>("exports");
            var expenses = conn.GetCollection<Expense>("expenses");
            var options = new FindOptions<BsonDocument> { CursorType = CursorType.TailableAwait };
            var builder = Builders<BsonDocument>.Filter;
            var filter = builder.Eq("action", "generate") & builder.Gt("submitted", new BsonDateTime(DateTime.Now));
            var projection = Builders<Expense>.Projection.Exclude("sheet");

            _doWork = true;

            var asdf = new BsonDocument();
            asdf["_id"] = new ObjectId();
            asdf["action"] = "init";
            asdf["submitted"] = DateTime.Now;
            Console.WriteLine("done");
            await exports.InsertOneAsync(asdf);

            while (_doWork)
            {
                using (var cursor = await exports.FindAsync(filter, options))
                {
                    await cursor.ForEachAsync(async export =>
                    {
                        Console.WriteLine(export.ToString());
                        var expense = await expenses.Find(new BsonDocument("_id", ObjectId.Parse(export["expenseId"].ToString()))).Project<Expense>(projection).FirstAsync();

                        if (generate(expense))
                        {
                            var send = export;
                            send["_id"] = new ObjectId();
                            send["action"] = "mail";
                            send["submitted"] = DateTime.Now;
                            Console.WriteLine("done");
                            await exports.InsertOneAsync(send);
                        }
                    });
                }
            }
        }
        public void Stop()
        {
            _doWork = false;
        }

        private bool generate(Expense expense)
        {
            Excel.Application excelApp = new Excel.Application();
            string workbookPath = @"C:\Mula\expenseclaim_v3.05.xlsm";
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets[1];
            Excel.Worksheet uploadWorksheet = (Excel.Worksheet)excelSheets[6];
            Excel.Worksheet wseetemplateWorksheet = (Excel.Worksheet)excelSheets[7];

            try
            {

                DateTime oldest = Convert.ToDateTime(expense.oldestBillDate).Date;
                DateTime penaltyDate = oldest.AddDays(56);
                bool fillDate = false;

                if(oldest < DateTime.Now.AddDays(-57)) { fillDate = true; }

                excelWorksheet.Cells[3, "G"] = expense.fullname;
                excelWorksheet.Cells[4, "G"] = DateTime.Now.Date.ToString(@"MM/dd/yyyy"); ;
                excelWorksheet.Cells[5, "G"] = expense.expCurrency;
                excelWorksheet.Cells[6, "G"] = expense.reimbCurrency;
                excelWorksheet.Cells[7, "G"] = oldest.ToString(@"MM/dd/yyyy");

                foreach (Project project in expense.projects)
                {
                    foreach (Row row in project.row)
                    {
                        if (row.sheetNumber == 1) //TODO: FIX!
                        {
                            string an = row.number.ToString();
                            excelWorkbook.Names.Item("num" + an + "Assignment").RefersToRange.Value = project.assignment;
                            excelWorkbook.Names.Item("num" + an + "Project").RefersToRange.Value = project.name;
                            excelWorksheet.Cells[Int32.Parse(an) + 3, "X"] = project.description;
                        }
                    }
                }

                foreach (Receipt receipt in expense.receipts)
                {
                    int rowReceiptCount = 17 + receipt.receiptNumber;
                    if (receipt.sheetNumber == 1) //TODO: FIX!
                    {
                        excelWorksheet.Cells[rowReceiptCount, "E"] = receipt.receiptNumber;
                        excelWorksheet.Cells[rowReceiptCount, "F"] = receipt.projectNumber;
                        excelWorksheet.Cells[rowReceiptCount, "G"] = receipt.where;
                        excelWorksheet.Cells[rowReceiptCount, GetCategoryIndex(receipt.type)] = receipt.amount.ToString("0.00");
                        excelWorksheet.Cells[rowReceiptCount, "V"] = receipt.description;

                        if (fillDate) { excelWorksheet.Cells[rowReceiptCount, "U"] = Convert.ToDateTime(receipt.date).ToString(@"MM/dd/yyyy"); }
                    }
                }

                Excel.Range sourceRange = uploadWorksheet.Cells;
                Excel.Range destinationRange = wseetemplateWorksheet.Cells;
                sourceRange.Copy(Type.Missing);
                destinationRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);

                excelApp.DisplayAlerts = false;              
                excelWorkbook.SaveAs(@"C:\Mula\export\" + expense._id.ToString() + ".xlsm");
            }
            catch(Exception e)
            {
                return false;
            }
            finally {
                excelWorkbook.Close(false);
                excelApp.Application.Quit();
                excelApp.Quit();

                Marshal.ReleaseComObject(excelSheets);
                Marshal.ReleaseComObject(excelWorksheet);
                Marshal.ReleaseComObject(uploadWorksheet);
                Marshal.ReleaseComObject(wseetemplateWorksheet);
                Marshal.ReleaseComObject(excelWorkbook);
                Marshal.ReleaseComObject(excelApp);
            }
            return true;
        }
        
        private string GetCategoryIndex(string rawCategory)
        {
            switch (rawCategory)
            {
                case "Airfare": return "H";
                case "Car Rentl, Parking, Toll": return "I";
                case "Hotel": return "I";
                case "Meal & Entertainmt": return "K";
                case "Taxi": return "L";
                case "Mileage (KM)": return "I";
                default: return "ERROR";
            }
        }
    }
}
