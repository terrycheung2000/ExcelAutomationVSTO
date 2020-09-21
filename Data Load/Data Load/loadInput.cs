/**************************************************************
 * Author: Terry Cheung
 * Date: July 28, 2020
 *                                          
 * File: loadinput.cs
 * 
 * Description: Windows forms loading bar
 * 
 * Input: N/A
 * 
 * Output: N/A
 * 
 **************************************************************/

using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.Windows.Forms;

namespace Data_Load
{
    /// <summary>
    /// loadInput class for loading bar
    /// </summary>
    public partial class loadInput : Form
    {
        /// <summary>
        /// Default contstructor for the loadInput. 
        /// Sets the default minimum and maximum of the loading bar.
        /// </summary>
        public loadInput()
        {
            InitializeComponent();
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
        }
    }
}
