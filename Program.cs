using OfficeOpenXml;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace ConvertProAxess
{
    static class Program
    {
        // List of Columns in the work sheet
        // Col 1 - Access Card Reader
        // Col 2 - Name - Under strange circumstances this could be area
        // Col 3 - SSN - Personnummer 000000-1111 / 000000-utl
        // Col 4 - Company
        // Col 5 - Division
        // Col 6 - Tillhörighet - Under strance circumstances this could be attestant
        // Col 7 - Time schedule of accesses
        // Col 8 - From which date access is approved
        // Col 9 - To which date access is approved
        // Col 10 - Extra field 1
        // Col 11 - Attestant
        // Col 12 - Area
        const int noOfColumns = 12;
        static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] != "/?" )
            {
                ConvertExcel(args);
            }
            else if(args.Length >= 1 && args[0] == "/?")
            {
                PrintHelpText();
            } else
            {
                PrintErrorTest();
            }
        }

        private static void ConvertExcel(string[] args)
        {
            var inputFile = new FileInfo(args[0]);
            var outputFile = new FileInfo(args[1]);
            var unmatchFile = new FileInfo(args[2]);

            var inpStartRowNo = 1;   // First row in the worksheet
            var outpRowNo = 1;
            var unmatchRowNo = 1;
            const int inpFirstWorkSheet = 0;

            Console.WriteLine("Starting to fix badly formatted excel file...");
            // Open input file
            using (var inp = new ExcelPackage(inputFile))
            {
                Console.WriteLine("Opening input file...");
                var inpWorkSheet = inp.Workbook.Worksheets[inpFirstWorkSheet]; // Load the work sheet from the input file
                var inpTotRows = inpWorkSheet.Dimension.Rows;

                // Open output file
                using (var outp = new ExcelPackage(outputFile))
                {
                    var outpWorkSheet = outp.Workbook.Worksheets.Add("Sheet1"); // Add the work sheet to the output
                    Console.WriteLine("Opening output file...");

                    // Open a recovery file in case of emergency ococures during the translation
                    using (var unmatch = new ExcelPackage(unmatchFile))
                    {
                        var unmatchWorkSheet = unmatch.Workbook.Worksheets.Add("Unmatched");
                        Console.WriteLine("Opening unmatched file...");

                        Console.WriteLine("Rows: " + inpTotRows.ToString());
                        Console.Write("Row 0\t of " + inpTotRows.ToString());

                        for (int rowNo = inpStartRowNo; rowNo < inpTotRows; rowNo++)
                        {
                            // Save output file and unmatch file every 250th row for safety reasons
                            if (rowNo % 250 == 0)
                            {
                                outp.Save();
                                unmatch.Save();
                            }

                            (rowNo, outpRowNo, unmatchRowNo) = FindPersonalNumber(inpWorkSheet,
                                                                                    outpWorkSheet,
                                                                                    unmatchWorkSheet,
                                                                                    rowNo,
                                                                                    outpRowNo,
                                                                                    unmatchRowNo);
                        }
                        unmatch.Save();
                    }
                    outp.Save();
                }
            }
            PrintWorkDoneText(outpRowNo, unmatchRowNo);
        }

        private static void TryWriteExcel( int inpColNo, 
                                            int outpColNo, 
                                            int inpRowNo, 
                                            int outpRowNo, 
                                            ExcelWorksheet inpWorkSheet, 
                                            ExcelWorksheet outpWorkSheet)
        {
            try
            {
                outpWorkSheet.Cells[outpRowNo, outpColNo].Value = inpWorkSheet.Cells[inpRowNo, inpColNo].Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                if (ex is NullReferenceException)
                {
                    outpWorkSheet.Cells[outpRowNo, outpColNo].Value = inpWorkSheet.Cells[inpRowNo, inpColNo].Value.ToString().Trim();
                }
                else if (ex is ArgumentException)
                {
                    // Just catch this exception
                }
            }
        }

        private static int ChangeColumns(int colNo)
        {
            var newColNo = 0;
            if (colNo == 2)
            {
                // Move col 2 to col 12
                newColNo = 12;
            }
            else if (colNo == 6)
            {
                // Move col 6 to col 11
                newColNo = 11;
            }
            return newColNo;
        }

        private static bool CheckPersonalNumber(    ExcelWorksheet inpWorkSheet, 
                                                    int inpColNo, 
                                                    int inpRowNo)
        {
            try
            {
                var personalNo = inpWorkSheet.Cells[inpRowNo, inpColNo].Value.ToString().Trim();
                // is the column containing a PSN?
                return MatchPersonalNumber(personalNo);
            }
            catch (NullReferenceException)
            {
                return false;
            }
        }

        private static bool MatchPersonalNumber(string personalNo)
        {
            var matchSSN = new Regex(@"^\d{6}[\-]?(\d{4}|\butl)$", RegexOptions.IgnoreCase);
            return matchSSN.IsMatch(personalNo);
        }

        private static (int, int, int) FindPersonalNumber(  ExcelWorksheet inpWorkSheet, 
                                                            ExcelWorksheet outpWorkSheet, 
                                                            ExcelWorksheet unmatchWorkSheet,
                                                            int rowNo,
                                                            int outpRowNo,
                                                            int unmatchRowNo)
        {
            var inpColPSN = 3;
            var inpStartColNo = 1;
            var rowGoToUnmatch = false;
            
            var inpTotRows = inpWorkSheet.Dimension.Rows;
            
            // Check if default column is containing a personal number
            var isPersonalNumber = CheckPersonalNumber(inpWorkSheet, inpColPSN, rowNo);

            if (!isPersonalNumber)
            {
                // If default col not containing a PSN, search if there is a personal number at all.
                for (int x = inpStartColNo; x < noOfColumns; x++)
                {
                    try
                    {
                        var tryPSN = inpWorkSheet.Cells[rowNo, x].Value.ToString().Trim();
                        if (MatchPersonalNumber(tryPSN))
                        {
                            // If we find personal number in another column, do this
                            if (x == 4)
                            {
                                // If we find PSN in column 4 do sorting
                                for(int colNo = inpStartColNo; colNo < noOfColumns; colNo++)
                                {
                                    int newColNo = colNo - 1;
                                    newColNo = ChangeColumns(colNo);
                                    TryWriteExcel(  colNo, 
                                                    newColNo, 
                                                    rowNo, 
                                                    outpRowNo, 
                                                    inpWorkSheet, 
                                                    outpWorkSheet);
                                }
                                outpRowNo++;
                                rowGoToUnmatch = false;
                                break;
                            }
                        }
                        else
                        {
                            // If do not find PSN in any col set rowGoToUmatched to true
                            rowGoToUnmatch = true;
                        }
                    }
                    catch (NullReferenceException)
                    {
                        rowGoToUnmatch = true;
                    }
                }
                if (rowGoToUnmatch)
                {
                    // If we can't find a personal number output to unmatched file for manual processing
                    for (int colNo = inpStartColNo; colNo < noOfColumns; colNo++)
                    {
                        TryWriteExcel(  colNo, 
                                        colNo, 
                                        rowNo, 
                                        unmatchRowNo, 
                                        inpWorkSheet, 
                                        unmatchWorkSheet);
                    }
                    unmatchRowNo++;
                }
            }
            else
            {
                // If it is containing a personal no, just write to output file.
                for (int colNo = inpStartColNo; colNo < noOfColumns; colNo++)
                {
                    TryWriteExcel(  colNo, 
                                    colNo, 
                                    rowNo, 
                                    outpRowNo, 
                                    inpWorkSheet, 
                                    outpWorkSheet);
                }
                outpRowNo++;
            }
            Console.Write("\rRow: " + rowNo + " of " + inpTotRows.ToString());
            return (rowNo, outpRowNo, unmatchRowNo);
        }

        private static void PrintWorkDoneText(int outpRowNo, int unmatchRowNo)
        {
            Console.WriteLine("");
            Console.WriteLine("Lines written to output:   " + outpRowNo);
            Console.WriteLine("Lines written to recovery: " + unmatchRowNo);
            Console.WriteLine("Formatting is done!");
        }

        private static void PrintErrorTest()
        {
            Console.WriteLine("ERROR: Commands not given.");
            Console.WriteLine("");
            PrintHelpText();
        }

        private static void PrintHelpText()
        {
            Console.WriteLine("DESCRIPTION:");
            Console.WriteLine("\tThis command fixes badly formatted output files from PACS");
            Console.WriteLine("\t(Physical Access Control System) ProAxess. This commands ");
            Console.WriteLine("\twas specially made for a project to a customer.");
            Console.WriteLine("");
            Console.WriteLine("\tVersion 1.0 © Copyright Marcus Fredlund");
            Console.WriteLine("");
            Console.WriteLine("USAGE:");
            Console.WriteLine("\tConvertProAxess.exe [input file] [output file] [recovery file]");
            Console.WriteLine("");
            Console.WriteLine("\t[input file] = badly formatted excel file");
            Console.WriteLine("\t[output file] = better formatted excel file");
            Console.WriteLine("\t[recovery file] = this contains rows which is unmatched.");
        }
    }
}
