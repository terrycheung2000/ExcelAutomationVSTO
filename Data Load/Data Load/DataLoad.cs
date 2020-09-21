/**************************************************************
 * Author: Terry Cheung
 * Date: July 28, 2020
 *                                          
 * File: DataLoad.cs
 * 
 * Description: Excel Addin to generate post 
 * calculation templates and validate input data from CAL tests 
 * 
 * Input: assigned tests specification, AR number, compound name
 * 
 * Output: Copied templates each with data gathered from
 * specification. Sets specific units and converts results to 
 * proper units.
 * 
 **************************************************************/

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Data_Load {
    /// <summary>
    /// Data Load contains all functionality used for the Add-in.
    /// </summary>
    class NamespaceDoc {
    }
    /// <summary>
    /// Main class containing excel ribbon functionality.
    /// Contains the button used to run the Add-in
    /// </summary>
    // Main class containing excel ribbon and button
    public partial class DataLoad {
        Excel.Application inputExcel, tempExcel;
        Excel.Workbook input, output, tempbook;
        /// <summary>
        /// Unused. Runs when the ribbon is loaded.
        /// </summary>
        /// <param name="sender">Method sender</param>
        /// <param name="e">Method event args</param>
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {

        }
        /// <summary>
        /// Functionality when the postCalc button is clicked.
        /// Prompts user to run on the template sheet or select it as input reference.
        /// </summary>
        /// <param name="sender">Method sender</param>
        /// <param name="e">Method event args</param>
        // actions taken on clicking the post calc button
        private void postCalc_click(object sender, RibbonControlEventArgs e) {
            bool quit = false;
            // checks if the current workbook is the template
            if (Globals.ThisAddIn.Application.ActiveWorkbook.Name != "sample_prep_LCMS_WS_Template.xlsm") {
                // quits if user cancels the add in 
                DialogResult result = System.Windows.Forms.MessageBox.Show("Please run Data Load in sample_prep_LCMS_WS_Template.xlsm or select template workbook", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.Cancel) {
                    Globals.ThisAddIn.Application.ActiveWindow.Close(false);
                    Globals.ThisAddIn.Application.Quit();
                    quit = true;
                } else {
                    // does not continue if the file isnt opened 
                    if (!functionality.OpenFile(ref tempExcel, ref tempbook)) {
                        quit = true;
                    }
                }
            }
            // continue if the current worksheet is the template
            else {
                // tracks the main workbook
                tempbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            }
            if (!quit) {
                // opens the input data workbook
                System.Windows.Forms.MessageBox.Show(new Form { TopMost = true }, "Please select the assigned test workbook", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (functionality.OpenFile(ref inputExcel, ref input)) {
                    // copy template workbook
                    output = functionality.createOutputBook(tempbook);
                    // do post calculation
                    try {
                        functionality.getInput(input, output);
                        functionality.postCalc(output);
                    } catch (Exception error) {
                        System.Windows.MessageBox.Show(error.Message, error.GetType().ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
    }
}
