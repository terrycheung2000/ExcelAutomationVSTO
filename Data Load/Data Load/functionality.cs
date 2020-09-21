/**************************************************************
 * Author: Terry Cheung
 * Date: July 28, 2020
 *                                          
 * File: functionality.cs
 * 
 * Description: Contains the functionality of the DataLoad ribbon
 * 
 * Input: template sheet, input sheet, output sheet
 * 
 * Output: function dependent 
 * 
 **************************************************************/

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Data_Load {
    /// <summary>
    /// Functionality class. Contains all functionality used in the Data Load add-in
    /// </summary>
    class functionality {
        /// <summary>
        /// Loading bar used throughout background worker
        /// </summary>
        public static loadInput loadbar;
        /// <summary>
        /// Background worker using separate thread. 
        /// Updates loadbar.
        /// </summary>
        public static BackgroundWorker backgroundWorker1;
        /// <summary>
        /// Opens the file specified in OpenFileDialog and sets a reference to it.
        /// </summary>
        /// <param name="app">A reference to the excel application</param>
        /// <param name="wb">A reference to the workbook to be opened</param>
        /// <returns>
        /// Returns true if file opened successfully, false if the file returned by Workbooks.Open is null
        /// </returns>
        // opens the file using openFileDialog and finds excel extensions
        public static bool OpenFile(ref Excel.Application app, ref Workbook wb) {
            using (OpenFileDialog openFileDialog = new OpenFileDialog()) {
                openFileDialog.Filter = "Excel files (*.xlsm,*.xls,*.xlsx)|*.xlsm;*.xls;*.xlsx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                // opens the excel file chosen 
                if (openFileDialog.ShowDialog() == DialogResult.OK) {
                    app = new Excel.Application();
                    wb = app.Workbooks.Open(openFileDialog.FileName);
                    app.Visible = true;
                    return true;
                } else {
                    return false;
                }
            }
        }
        /// <summary>
        /// Creates and runs a background worker to copy input specified by user.
        /// </summary>
        /// <param name="input">Input workbook specified by user</param>
        /// <param name="output">Output workbook as a copy of the template</param>
        // gets the input from the assigned test file
        public static void getInput(Workbook input, Workbook output) {
            // initializing worksheets and variables
            Worksheet inputData = input.Worksheets[1];
            Worksheet inputCopy = output.Worksheets["Input"];
            Worksheet outsheet = output.Worksheets["Output"];
            outsheet.Range["A1", outsheet.Cells[outsheet.Rows.EntireRow.Count, outsheet.Columns.EntireColumn.Count]].Locked = false;
            // setting list objects to populate input workpage
            List<List<object>> copyData = new List<List<object>>(Enumerable.Range(1, visibleRows(inputData)).Select(x => new List<object>(Enumerable.Range(1, 10).Select(y => ""))));
            List<List<object>> resultTable = new List<List<object>>(Enumerable.Range(1, visibleRows(inputData)).Select(x => new List<object>(Enumerable.Range(1, 10).Select(y => new object()))));
            // initaliaze backgroundworker arguments and apply override to methods
            backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            List<object> args = new List<object>();
            args.Add(inputData);
            args.Add(inputCopy);
            args.Add(output.Worksheets["Output"]);
            args.Add(copyData);
            args.Add(resultTable);
            inputCopy.Select();
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            // run the background worker 
            backgroundWorker1.RunWorkerAsync(args);
            // start the loading bar
            loadbar = new loadInput();
            loadbar.Activate();
            // check for cancelation and wait until loading is complete
            while (backgroundWorker1.IsBusy) {
                DialogResult result = loadbar.ShowDialog();
                if (result == DialogResult.Cancel) {
                    backgroundWorker1.CancelAsync();
                }
                System.Threading.Thread.Sleep(50);
            }
            // end the background worker
            backgroundWorker1.Dispose();
        }
        /// <summary>
        /// Counts all visible rows of the given worksheet
        /// </summary>
        /// <param name="sheet">Worksheet to be counted</param>
        /// <returns>Returns number of rows visible</returns>
        // count visible rows of given worksheet
        private static int visibleRows(Worksheet sheet) {
            Range visible = sheet.UsedRange.SpecialCells(XlCellType.xlCellTypeVisible);
            int count = 0;
            foreach (Range area in visible.Areas) {
                foreach (Range row in area.Rows) {
                    count++;
                }
            }
            return count;
        }
        /// <summary>
        /// Main function to write to copy/paste template table to the output sheet.
        /// </summary>
        /// <param name="output">Output workbook</param>
        // performs the calculations required from data
        public static void postCalc(Workbook output) {
            // initalizing variables
            string even = "A", odd = "E", prevAR = "";
            string[] prevAdd = { "", "", "", "" };
            int row = 1;
            int last = -1;
            Worksheet template = output.Worksheets["Template"];
            Worksheet input = output.Worksheets["Input"];
            Worksheet outsheet = output.Worksheets["Output"];
            outsheet.Select();
            // writes to the template sheet and copies to two columns of the output sheet
            for (int i = 2; i < input.UsedRange.Rows.Count + 1; i++) {
                writeTemplate(template, input, i);
                if ((i - 2) % 2 == 0) {
                    copyTemplate(template, outsheet, even + row.ToString(), ref prevAdd, ref prevAR);
                    last = 1;
                } else {
                    copyTemplate(template, outsheet, odd + row.ToString(), ref prevAdd, ref prevAR);
                    row += 17;
                    last = 2;
                }
                // paste signature box at the end of every page
                if ((i - 1) % 4 == 0) {
                    template.Range["E1", "H7"].Copy();
                    outsheet.Range["A" + (row + 1).ToString()].PasteSpecial(XlPasteType.xlPasteAll);
                    row += 9;
                    last = 0;
                }
            }
            // pastes signature box at the end of the document
            if (last == 2) {
                template.Range["E1", "H7"].Copy();
                outsheet.Range["A" + (row + 1).ToString()].PasteSpecial(XlPasteType.xlPasteAll);
            } else if (last == 1) {
                row += 17;
                template.Range["E1", "H7"].Copy();
                outsheet.Range["A" + (row + 1).ToString()].PasteSpecial(XlPasteType.xlPasteAll);
            }
            outsheet.Columns.AutoFit();
            // protecting the sheet
            outsheet.Protect("password", AllowFormattingCells: true, AllowFormattingColumns: true);
        }
        /// <summary>
        /// Copies and creates a new sheet containing the template
        /// </summary>
        /// <param name="tempbook">Template workbook to be copied</param>
        /// <returns>Returns output workbook with template</returns>
        // copies and creates a new sheet containing the template
        public static Excel.Workbook createOutputBook(Workbook tempbook) {
            Excel.Application outputExcel = new Excel.Application();
            Workbook output = outputExcel.Application.Workbooks.Add(tempbook.FullName);
            outputExcel.Visible = true;
            return output;
        }
        /// <summary>
        /// Writes data given from input worksheet to the template table
        /// </summary>
        /// <param name="template">Template worksheet containing the worksheet table</param>
        /// <param name="input">Input worksheet containing specifications and product data</param>
        /// <param name="row">Current row of input worksheet</param>
        // writing data from input sheet to template sheet
        public static void writeTemplate(Worksheet template, Worksheet input, int row) {
            // writes attachment number every 4 tables
            string spec = "";
            if ((row - 2) % 4 == 0) {
                template.Range["AttachmentNum"].Value2 = "Attachment-02-" + (((row - 2) / 4) + 1).ToString().PadLeft(2, '0');
            } else {
                template.Range["AttachmentNum"].ClearContents();
            }
            //read the spec string of the row and find the unit used
            if (input.Range["J" + row].Value != null) {
                spec = input.Range["J" + row].Value.ToString().ToLower();
            }
            string baseUnit = baseUnits(spec);
            // setting template cells
            template.Range["AR_No"].Value2 = input.Range["C" + row.ToString()];
            template.Range["Compound"].Value2 = input.Range["E" + row.ToString()];
            template.Range["Spec"].Value2 = input.Range["J" + row.ToString()];
            template.Range["Average_Weight"].FormulaR1C1 = checkSpec(spec, template);
            template.Range["Average_Weight_unit"].Value2 = "Average Weight " + baseUnit;
            template.Range["Weight_Taken_unit"].Value2 = "Weight Taken " + baseUnit;
            // 20 units for any tablet, gummy, capsule, or softgel 
            if (spec.Contains("tablet") || spec.Contains("gummy") || spec.Contains("capsule") || spec.Contains("softgel") || spec.Contains("caplet")) {
                template.Range["Number_Units"].Value2 = "20 UNITS";
            } else {
                template.Range["Number_Units"].Value2 = "";
            }
            specRange(template, spec);
            setConversion(template, spec);
        }
        /// <summary>
        /// Gets the proper value of average weight depending on given specification.
        /// Writes the formula for entered weight if the specification contains capsule or softgel
        /// </summary>
        /// <param name="spec">Specification string</param>
        /// <param name="template">Template Worksheet to write to</param>
        /// <returns>Returns the average weight</returns>
        // checks the proper value of Average_Weight depending on spec
        private static dynamic checkSpec(string spec, Worksheet template) {
            dynamic avgWeight = 1;
            double tryparse;
            string numUnit, unit;
            template.Range["Average_Weight"].NumberFormat = "0.0000";
            template.Range["C14"].ClearContents();

            // if the spec is a percentage, average weight is 1
            if (spec.Contains("%") && (!spec.Contains("(") && !spec.Contains(")") && !spec.Contains("/"))) {
                avgWeight = 1;
            }
            // if the spec is a tablet or gummy, take the first weight value and divide by 20 
            else if (spec.Contains("tablet") || spec.Contains("gummy") || spec.Contains("lozenge") || spec.Contains("caplet")) {
                avgWeight = "=R[9]C[0]/20";
                template.Range["C14"].FormulaR1C1 = "=(R[0]C[-1])/20";
            }
            // if the spec is a capsule or softgel, take first and second weight value, find difference, and divide by 20
            else if (spec.Contains("capsule") || spec.Contains("softgel")) {
                avgWeight = "=(R[9]C[0]-R[10]C[0])/20";
                template.Range["C14"].FormulaR1C1 = "=(R[0]C[-1]-R[1]C[-1])/20";
            }
            // if the spec has average specified, use given average
            else if (spec.Contains("/")) {
                numUnit = spec.Split('/')[1].Trim();
                // splits the string to interpret scoop average weight
                if (numUnit.Contains("scoop")) {
                    if (numUnit.Contains("-")) {
                        numUnit = numUnit.Split('-')[1].Trim();
                    } else if (numUnit.Contains("=")) {
                        numUnit = numUnit.Split('=')[1].Trim();
                    }
                }
                numUnit = unitsFormat(numUnit);
                // gets the average weight from numbers in front of the unit
                if (numUnit.Split(' ').Count() > 1) {
                    unit = numUnit.Split(' ')[1];
                    numUnit = numUnit.Split(' ')[0];
                    if (double.TryParse(numUnit, out tryparse)) {
                        avgWeight = tryparse;
                        if (avgWeight.ToString().Contains(".")) {
                            template.Range["Average_Weight"].NumberFormat = 0.ToString("N" + avgWeight.ToString().Split('.')[1].Trim().Length);
                        } else {
                            template.Range["Average_Weight"].NumberFormat = "0";
                        }
                    }
                    // divides average weight depending on unit and requirement
                    if (unit.Contains("mg")) {
                        template.Range["Average_Weight"].NumberFormat = 0.ToString("N" + avgWeight.ToString().Split('.')[0].Trim().Length);
                        avgWeight = avgWeight / 1000;
                    } else if (unit.Contains("mcg") || unit.Contains("µg")) {
                        template.Range["Average_Weight"].NumberFormat = 0.ToString("N" + avgWeight.ToString().Split('.')[0].Trim().Length);
                        avgWeight = avgWeight / 1000000;
                    }
                }
            }
            // sets number format for default average weight
            if (avgWeight.ToString() == "1") {
                template.Range["Average_Weight"].NumberFormat = "0";
            }
            return avgWeight;
        }
        /// <summary>
        /// Applies a space in front of the unit in a given specification string
        /// </summary>
        /// <param name="toformat">Specification string to be formatted</param>
        /// <returns>Returns formatted string</returns>
        // inserts space in front of unit to properly interpret 
        private static string unitsFormat(string toformat) {
            if (!toformat.Contains(" ")) {
                if (toformat.Contains("µ")) {
                    toformat = toformat.Insert(toformat.IndexOf("µ"), " ");
                } else if (toformat.Contains("m")) {
                    toformat = toformat.Insert(toformat.IndexOf("m"), " ");
                } else if (toformat.Contains("g")) {
                    toformat = toformat.Insert(toformat.IndexOf("g"), " ");
                } else if (toformat.Contains("%")) {
                    toformat = toformat.Insert(toformat.IndexOf("%"), " ");
                } else if (toformat.Contains("pp")) {
                    toformat = toformat.Insert(toformat.IndexOf("pp"), " ");
                } else if (toformat.Contains("iu")) {
                    toformat = toformat.Insert(toformat.IndexOf("iu"), " ");
                }
            }
            return toformat;
        }
        /// <summary>
        /// Gets the units of the given specification
        /// </summary>
        /// <param name="spec">Specification to get units</param>
        /// <returns>Returns the units of the specification</returns>
        // check units within spec string
        private static string specUnits(string spec) {
            // defaults to g when no units detected
            string numUnit, unit = "g";
            numUnit = spec.Split('/')[0];
            if (numUnit.Contains("%")) {
                unit = "%";
            } else if (numUnit.Contains("mg")) {
                unit = "mg";
            } else if (numUnit.Contains("µg") || numUnit.Contains("mcg") || numUnit.Contains("ppm")) {
                unit = "µg";
            } else if (numUnit.Contains("ppb")) {
                unit = "ppb";
            }
            return unit;
        }
        /// <summary>
        /// Gets the range of values for valid results.
        /// Formats the result field according to the range.
        /// See <see cref="formatResult(Worksheet, string)"/> and <see cref="formatResult(Worksheet, string, string)"/> for formatting.
        /// </summary>
        /// <param name="template">Template worksheet to format</param>
        /// <param name="spec">Specification string</param>
        // getting min max values for conditional formating
        private static void specRange(Worksheet template, string spec) {
            // initializing variables
            string min = "0", max = "0", minMax, range;
            template.Range["result"].Interior.Color = XlRgbColor.rgbPink;
            template.Range["result"].Font.Color = XlRgbColor.rgbRed;
            if ((spec.Contains("/") && spec.Contains("-")) || (spec.Contains("%") && spec.Contains("-"))) {
                // splits the spec into components only if they exist and gets min max
                if (spec.Contains("/")) {
                    range = spec.Split('/')[0];
                } else {
                    range = spec.Split('%')[0];
                }
                min = range.Split('-')[0];
                minMax = range.Split('-')[1];
                minMax = minMax.Trim();
                max = minMax.Split(' ')[0];
                formatResult(template, min, max);
            } else if (spec.Contains("%")) {
                formatResult(template, spec.Split('%')[0]);
            } else if (spec.Contains("/") || spec.Contains("ppm") || spec.Contains("ppb")) {
                formatResult(template, spec.Split('/')[0]);
            } else {
                // if range cannot be gotten, make the cell yellow
                template.Range["result"].FormatConditions.Delete();
                template.Range["result"].Interior.Color = XlRgbColor.rgbYellow;
                template.Range["result"].Font.Color = XlRgbColor.rgbBlack;
            }
        }
        /// <summary>
        /// Sets the conversion formula for the result cell
        /// </summary>
        /// <param name="template">Template worksheet to write the formula</param>
        /// <param name="spec">Specification string</param>
        // sets the formula for the result to use conversion table depending on the unit
        private static void setConversion(Worksheet template, string spec) {
            Range resultRange = template.Range["result"];
            Range resultTitle = template.Range["ResultTitle"];
            Range units = template.Range["units"];
            string compound = template.Range["Compound"].Value.ToString().ToLower();
            // if spec has µgrae, then set the formula to use the RAELookup table
            if (spec.Contains("µgrae") || spec.Contains("mcgrae") || spec.Contains("µg rae")) {
                units.Value2 = "µgRAE";
                if (spec.Contains("%") && spec.Contains("(") && spec.Contains(")")) {
                    resultRange.Formula = "= ((B9*B10)/B11)/(VLOOKUP(\"" + compound + "\",RAELookup,2,FALSE)*" + getdivision(spec, units).ToString() + ")*100";
                } else {
                    resultRange.Formula = "= ((B9*B10)/B11)/VLOOKUP(B3,RAELookup,2,FALSE)";
                }
            // if spec has IU set the formula to use the IULookup table and set the compound as the lookup
            } else if (spec.Contains("iu")) {
                units.Value2 = "IU";
                if (spec.Contains("%") && spec.Contains("(") && spec.Contains(")")) {
                    resultRange.Formula = "= ((B9*B10)/B11)/(VLOOKUP(\"" + compound + "\",IULookup,2,FALSE)*" + getdivision(spec, units).ToString() + ")*100";
                } else if (spec.Contains("miu")) {
                    units.Value2 = "MIU";
                    resultRange.Formula = "= ((B9*B10)/B11)/(VLOOKUP(\"" + compound + "\",IULookup,2,FALSE)*1000000)";
                } else {
                    resultRange.Formula = "= ((B9*B10)/B11)/VLOOKUP(B3,IULookup,2,FALSE)";
                }
            // if spec has any vitamin E conversion then set the conversion
            } else if (spec.Contains("mgat") || spec.Contains("mg at") || spec.Contains("mgte") || spec.Contains("mg te")) {
                if (compound.Contains("alpha tocopherol")) {
                    units.Value2 = "mg";
                    if (spec.Contains("%") && spec.Contains("(") && spec.Contains(")")) {
                        resultRange.Formula = "= ((B9*B10)/B11)/(VLOOKUP(C12,tblUnitsPost,2,FALSE)*" + getdivision(spec, units).ToString() + ")*100";
                    } else {
                        resultRange.Formula = "= ((B9*B10)/B11)/VLOOKUP(C12,tblUnitsPost,2,FALSE)";
                    }
                } else if (compound.Contains("alpha tocopheryl acetate")) {
                    units.Value2 = "mg AT";
                    if (spec.Contains("%") && spec.Contains("(") && spec.Contains(")")) {
                        resultRange.Formula = "= (((B9*B10)/B11)/(1000000*" + getdivision(spec, units).ToString() + "))*91";
                    } else {
                        resultRange.Formula = "= (((B9*B10)/B11)/1000000)*0.91";
                    }
                    resultTitle.Value2 = "Final Result (CF = 0.91)";
                } else if (compound.Contains("alpha tocopheryl succinate")) {
                    units.Value2 = "mg TE";
                    if (spec.Contains("%") && spec.Contains("(") && spec.Contains(")")) {
                        resultRange.Formula = "= (((B9*B10)/B11)/(1000000*" + getdivision(spec, units).ToString() + "))*81";
                    } else {
                        resultRange.Formula = "= (((B9*B10)/B11)/1000000)*0.81";
                    }
                    resultTitle.Value2 = "Final Result (CF = 0.81)";
                }
                // set the percentage conversion
            } else if (spec.Contains("%") && spec.Contains("(") && spec.Contains(")")) {
                units.Value2 = specUnits(spec);
                resultRange.Formula = "= ((B9*B10)/B11)/(VLOOKUP(C12,tblUnitsPost,2,FALSE)*" + getdivision(spec, units).ToString() + ")";
            // default use the units lookup table
            } else {
                units.Value2 = specUnits(spec);
                resultRange.Formula = "= ((B9*B10)/B11)/VLOOKUP(C12,tblUnitsPost,2,FALSE)";
            }
            // conversion for specific compounds
            if (compound == "pantothenic acid") {
                resultRange.Formula += "*0.92";
                resultTitle.Value2 = "Final Result (CF = 0.92)";
            } else if (compound == "choline") {
                resultRange.Formula += "*0.41";
                resultTitle.Value2 = "Final Result (CF = 0.41)";
            } else if (compound == "vitamin b1 (thiamine mononitrate)") {
                resultRange.Formula += "*0.97";
                resultTitle.Value2 = "Final Result (CF = 0.97)";
            } else if (compound == "vitamin b6 (pyridoxine)") {
                resultRange.Formula += "*0.82";
                resultTitle.Value2 = "Final Result (CF = 0.82)";
            } else if (compound == "citruilline malate") {
                resultRange.Formula += "*1.77";
                resultTitle.Value2 = "Final Result (CF = 1.77)";
            } else {
                resultTitle.Value2 = "Final Result";
            }
        }
        /// <summary>
        /// Gets the division factor of the specification
        /// </summary>
        /// <param name="spec">Specification string</param>
        /// <param name="units">Worksheet range with units of the specification</param>
        /// <returns>Returns the division factor</returns>
        // get the amount to divide for labor claim
        private static double getdivision(string spec, Range units) {
            string unit = "";
            double divide;
            // splits the string to get the percentage division
            spec = spec.Substring(spec.IndexOf('(') + 1, spec.IndexOf(')') - (spec.IndexOf('(') + 1));
            if (spec.Contains("/") && spec.Contains("=")) {
                unit = unitsFormat(spec.Substring(spec.IndexOf('=') + 1, spec.IndexOf('/') - (spec.IndexOf('=') + 1)).Trim()).Trim().Split(' ')[1];
                spec = spec.Substring(spec.IndexOf('=') + 1, spec.IndexOf('/') - (spec.IndexOf('=') + 1)).Trim();
            } else if (spec.Contains("/") && spec.Contains(":")) {
                unit = unitsFormat(spec.Substring(spec.IndexOf(':') + 1, spec.IndexOf('/') - (spec.IndexOf(':') + 1)).Trim()).Trim().Split(' ')[1];
                spec = spec.Substring(spec.IndexOf(':') + 1, spec.IndexOf('/') - (spec.IndexOf(':') + 1)).Trim();
            } else if (spec.Contains(":")) {
                unit = unitsFormat(spec.Substring(spec.IndexOf(':') + 1, spec.Length - (spec.IndexOf(':') + 1)).Trim()).Trim().Split(' ')[1];
                spec = spec.Substring(spec.IndexOf(':') + 1, spec.Length - (spec.IndexOf(':') + 1)).Trim();
            } else if (spec.Contains("=")) {
                unit = unitsFormat(spec.Substring(spec.IndexOf('=') + 1, spec.Length - (spec.IndexOf('=') + 1)).Trim()).Trim().Split(' ')[1];
                spec = spec.Substring(spec.IndexOf('=') + 1, spec.Length - (spec.IndexOf('=') + 1)).Trim();
            }
            spec = unitsFormat(spec).Trim().Split(' ')[0];
            units.Value2 = "%";
            // divide by the units only if a double is read
            if (double.TryParse(spec, out divide)) {
                if (!unit.Contains("µgrae") && !unit.Contains("mcgrae") && !unit.Contains("µg rae") && !unit.Contains("mgat") && !unit.Contains("mg at") && !unit.Contains("mgte") && !unit.Contains("mg te")) {
                    if (unit.Contains("mg")) {
                        divide = divide / 1000;
                    } else if (unit.Contains("mcg") || unit.Contains("µg")) {
                        divide = divide / 1000000;
                    }
                }
            } else {
                divide = 1;
            }
            return divide;
        }
        /// <summary>
        /// Sets the format of the result cell according to the max and min values.
        /// </summary>
        /// <param name="template">Template worksheet to be written</param>
        /// <param name="min">Minimum value accepted</param>
        /// <param name="max">Maximum value accepted</param>
        // sets the conditional formating depending on the range
        private static void formatResult(Worksheet template, string min, string max) {
            Range resultRange = template.Range["result"];
            int minDecimals, maxDecimals;
            // sets the number formatting depending on the decimals given by min or max whichever has less
            if (min.Contains(".") && max.Contains(".")) {
                minDecimals = min.Split('.')[1].Trim().Length;
                maxDecimals = max.Split('.')[1].Trim().Length;
                if (minDecimals == maxDecimals || minDecimals < maxDecimals) {
                    resultRange.NumberFormat = 0.ToString("N" + min.Split('.')[1].Trim().Length);
                } else if (minDecimals > maxDecimals) {
                    resultRange.NumberFormat = 0.ToString("N" + max.Split('.')[1].Trim().Length);
                }
            } else {
                resultRange.NumberFormat = "0";
            }
            // setting the format condition depending on the range
            resultRange.FormatConditions.Delete();
            FormatCondition format = (FormatCondition)resultRange.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlBetween, min, max);
            format.Interior.Color = XlRgbColor.rgbPaleGreen;
            format.Font.Color = XlRgbColor.rgbDarkGreen;
        }
        /// <summary>
        /// Sets the format of the result cell according to the full specification string
        /// </summary>
        /// <param name="template">Template worsheet to be written</param>
        /// <param name="spec">Specification string</param>
        // override of first formatResult method to find range using full specification string
        private static void formatResult(Worksheet template, string spec) {
            FormatCondition format;
            Range resultRange = template.Range["result"];
            resultRange.FormatConditions.Delete();
            // checks for unit
            unitsFormat(spec);
            spec = spec.Split(' ')[0].Trim();
            // sets condition depending on given condition
            if (spec.Contains(">=")) {
                format = (FormatCondition)resultRange.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlGreaterEqual, spec.Trim().Split('=')[1]);
            } else if (spec.Contains(">")) {
                format = (FormatCondition)resultRange.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlGreater, spec.Trim().Split('>')[1]);
            } else if (spec.Contains("<=")) {
                format = (FormatCondition)resultRange.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLessEqual, spec.Trim().Split('=')[1]);
            } else if (spec.Contains("<")) {
                format = (FormatCondition)resultRange.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlLess, spec.Trim().Split('<')[1]);
            } else {
                format = (FormatCondition)resultRange.FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, spec);
            }
            if (spec.Contains(".")) {
                resultRange.NumberFormat = 0.ToString("N" + spec.Split('.')[1].Trim().Length);
            } else {
                resultRange.NumberFormat = "0";
            }
            format.Interior.Color = XlRgbColor.rgbPaleGreen;
            format.Font.Color = XlRgbColor.rgbDarkGreen;
        }
        /// <summary>
        /// Gets the units of average weight and weight taken
        /// </summary>
        /// <param name="spec">Specification string</param>
        /// <returns>Returns the base unit</returns>
        // sets base units depending on spec 
        private static string baseUnits(string spec) {
            string unit = "G";
            if (spec.ToLower().Contains("ml") || spec.ToLower().Contains("drop")) {
                unit = "mL";
            }
            return unit;
        }
        /// <summary>
        /// Copies the template table and pastes into the given range
        /// </summary>
        /// <param name="template">Template worksheet to be copied</param>
        /// <param name="outsheet">Output worksheet to be pasted</param>
        /// <param name="pasteRange">Range to be pasted</param>
        /// <param name="prevAdd">Array of previous addresses</param>
        /// <param name="prevAR">Previous AR number</param>
        // copies template and pastes on the output sheet
        public static void copyTemplate(Worksheet template, Worksheet outsheet, String pasteRange, ref String[] prevAdd, ref String prevAR) {
            Range tempRange = template.Range["A1", "C16"];
            tempRange.Copy();
            outsheet.Range[pasteRange].PasteSpecial(XlPasteType.xlPasteAll);
            // tracks the previous AR number provided and links the weight taken to the first instance of the AR number
            string currAR = outsheet.Range[pasteRange].Cells[2, 2].Value.ToString();
            // if the AR number matches then link
            if (prevAR == currAR) {
                outsheet.Range[pasteRange].Cells[11, 2].Formula = "=" + prevAdd[0];
                outsheet.Range[pasteRange].Cells[5, 2].Formula = "=" + prevAdd[1];
                outsheet.Range[pasteRange].Cells[11, 2].Locked = true;
                outsheet.Range[pasteRange].Cells[5, 2].Locked = true;
                // set the given weights of the same AR to be the first given weight
                if (template.Range["Average_Weight"].HasFormula) {
                    outsheet.Range[pasteRange].Cells[14, 2].Formula = "=" + prevAdd[2];
                    outsheet.Range[pasteRange].Cells[14, 2].Locked = true;
                    // shell weight are copied only for softgel and capsule
                    if (template.Range["Spec"].Value.ToString().ToLower().Contains("softgel") || template.Range["Spec"].Value.ToString().ToLower().Contains("capsule")) {
                        outsheet.Range[pasteRange].Cells[15, 2].Formula = "=" + prevAdd[3];
                        outsheet.Range[pasteRange].Cells[15, 2].Locked = true;
                    }
                }
            }
            // updating address and AR
            prevAdd[0] = outsheet.Range[pasteRange].Cells[11, 2].Address.ToString();
            prevAdd[1] = outsheet.Range[pasteRange].Cells[5, 2].Address.ToString();
            prevAdd[2] = outsheet.Range[pasteRange].Cells[14, 2].Address.ToString();
            prevAdd[3] = outsheet.Range[pasteRange].Cells[15, 2].Address.ToString();
            prevAR = currAR;
        }
        /// <summary>
        /// Background worker purposed to write the given input workbook to the output worksheet
        /// </summary>
        /// <param name="sender">Method sender</param>
        /// <param name="e">Method event args</param>
        // background worker for loading bar and multi-threading
        private static void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) {
            // initializing variables
            int i = 1, k = 0;
            double percent = 0.0;
            string[] resultHead = { "AR", "Compound", "Result", "Spec" };
            string[] dilutionHead = { "AR", "Compound", "Dilution" };
            // unpacking arguments
            List<object> args = e.Argument as List<object>;
            Worksheet inputData = (Worksheet)args[0];
            Worksheet inputCopy = (Worksheet)args[1];
            Worksheet outsheet = (Worksheet)args[2];
            List<List<object>> copyData = (List<List<object>>)args[3];
            var resultTable = (List<List<object>>)args[4];
            // initializing empty lists to clear lists read
            var empty4 = Enumerable.Range(1, 4).Select(x => new object());
            var empty10 = Enumerable.Range(1, 10).Select(x => new object());

            Excel.Range visible = inputData.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            // looping through all visible rows
            foreach (Excel.Range area in visible.Areas) {
                foreach (Excel.Range row in area.Rows) {
                    List<object> arow = new List<object>(empty10);
                    // gets the data from each used cell and puts it into an array
                    for (int j = 1; j < inputData.UsedRange.Columns.Count + 1; j++) {
                        arow[j - 1] = inputData.Cells[row.Row, j].Value2;
                        copyData[i - 1][j - 1] = inputData.Cells[row.Row, j].Value2;
                        percent += 100.0 / (inputData.Columns.Count * area.Rows.Count * visible.Areas.Count);
                        // if at any point backgroundworker is cancled then stop the method
                        if (backgroundWorker1.CancellationPending) {
                            e.Cancel = true;
                            return;
                        }
                    }
                    //copyData[i-1] = arow;
                    List<object> resultRow = new List<object>(empty4);
                    resultRow[0] = inputData.Cells[row.Row, 3].Value;
                    resultRow[1] = inputData.Cells[row.Row, 5].Value;
                    resultRow[2] = "";
                    resultRow[3] = inputData.Cells[row.Row, 10].Value;
                    resultTable[k] = resultRow;
                    // data for aflatoxin and terpene lactones need to be split into its individual compounds
                    if (inputData.Cells[row.Row, 9].Value == "Aflatoxin (B1+B2+G1+G2)" || inputData.Cells[row.Row, 9].Value == "(B1+B2+G1+G2)" || inputData.Cells[row.Row, 9].Value == "Aflatoxins (B1+B2+G1+G2)") {
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Aflatoxin B2";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData[i - 1] = new List<object>(arow);
                        i++;
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Aflatoxin G1";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData.Insert(i - 1, new List<object>(arow));
                        i++;
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Aflatoxin G2";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData.Insert(i - 1, new List<object>(arow));
                    } else if (inputData.Cells[row.Row, 9].Value == "Aflatoxin B1" || inputData.Cells[row.Row, 9].Value == "B1" || inputData.Cells[row.Row, 9].Value == "Aflatoxin (B1)" || inputData.Cells[row.Row, 9].Value == "Aflatoxins B1") {
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Aflatoxin B1";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData[i - 1] = new List<object>(arow);
                    } else if (inputData.Cells[row.Row, 9].Value == "Terpene Lactones" || inputData.Cells[row.Row, 9].Value == "Total Terpene Lactones") {
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Ginkgolide-A";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData[i - 1] = new List<object>(arow);
                        i++;
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Ginkgolide-J";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData.Insert(i - 1, new List<object>(arow));
                        i++;
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Ginkgolide-C";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData.Insert(i - 1, new List<object>(arow));
                        i++;
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "Ginkgolide-B";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData.Insert(i - 1, new List<object>(arow));
                        i++;
                        arow[2] = inputData.Cells[row.Row, 3];
                        arow[4] = "BiloBide";
                        arow[9] = inputData.Cells[row.Row, 10];
                        copyData.Insert(i - 1, new List<object>(arow));
                    }
                    i++;
                    k++;
                    backgroundWorker1.ReportProgress((int)percent);
                }
            }
            // pastes the array
            inputCopy.Range["A1", "J" + (i - 1).ToString()].Value2 = to2dArray(copyData);
            inputCopy.Columns.AutoFit();
            // protecting the sheet 
            inputCopy.Protect("password", AllowFormattingColumns: true, AllowFormattingCells: true);
            // writing result table and dilution table
            outsheet.Range["J1", "M" + (resultTable.Count).ToString()].Value2 = to2dArray(resultTable);
            outsheet.Range["J1", "M1"].Value2 = resultHead;
            AllBorders(outsheet.Range["J1", "M" + (resultTable.Count).ToString()].Borders);
            outsheet.Range["J" + (resultTable.Count + 3).ToString(), "L" + (resultTable.Count * 2 + 2).ToString()].Value2 = to2dArray(resultTable);
            outsheet.Range["J" + (resultTable.Count + 3).ToString() , "L" + (resultTable.Count + 3).ToString()].Value2 = dilutionHead;
            AllBorders(outsheet.Range["J" + (resultTable.Count + 3).ToString(), "L" + (resultTable.Count*2+2).ToString()].Borders);
            outsheet.Range["J1", "M" + (resultTable.Count).ToString()].Locked = true;
            outsheet.Range["L2", "L" + (resultTable.Count).ToString()].Locked = false;
            outsheet.Range["J" + (resultTable.Count + 3).ToString(), "K" + (resultTable.Count * 2 + 2).ToString()].Locked = true;
            backgroundWorker1.ReportProgress(100);
        }
        /// <summary>
        /// Converts the given nested list to a 2d array
        /// </summary>
        /// <param name="list">Nested list to be converted</param>
        /// <returns>Converted list as a 2d array</returns>
        // converts a given list to a 2d Array
        private static object[,] to2dArray(List<List<object>> list) {
            if (list.Count > 0) {
                object[,] array = new object[list.Count(), list[0].Count()];
                for (int i = 0; i < list.Count(); i++) {
                    for (int j = 0; j < list[i].Count(); j++) {
                        array[i, j] = list[i][j];
                    }
                }
                return array;
            } else {
                return null;
            }
        }
        /// <summary>
        /// Sets all border around a given range
        /// </summary>
        /// <param name="_borders">Borders to be set</param>
        // sets all borders around a range
        private static void AllBorders(Excel.Borders _borders) {
            _borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders.Color = Color.Black;
        }
        /// <summary>
        /// Triggers when background worker reports progress.
        /// Sends the progress percentage given.
        /// </summary>
        /// <param name="sender">Method sender</param>
        /// <param name="e">Method event args</param>
        // called when progress changed
        private static void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            // update the loading bar
            loadbar.progressBar1.Value = e.ProgressPercentage;
        }
        /// <summary>
        /// Triggers when background worker is completed.
        /// Closes loading bar and prompts user input is loaded.
        /// </summary>
        /// <param name="sender">Method sender</param>
        /// <param name="e">Method event args</param>
        // when the background worker is completed or canceled
        private static void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            // close the loading bar
            loadbar.Close();
            // if the backgroundworker ends normally, notify the user
            if (!e.Cancelled) {
                MessageBox.Show("Input Loaded");
            }
        }
    }
}

