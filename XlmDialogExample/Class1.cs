using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

using ExcelDna.Integration;
using ExcelDna.XlDialogBox;
namespace XlmDialogExample
{
    public class Class1
    {

        /// <summary>
        /// This is a dummy validation routine
        /// Validation routines only matter if you use a trigger on a control within an XlDialogBox
        /// </summary>
        /// <param name="index">the row index of the control that caused a trigger</param>
        /// <param name="dialogResult">the object array, that the Dialog worked with</param>
        /// <param name="Controls">the collection of controls, that can be edited in the callback function</param>
        /// <returns>
        /// return true, to show the dialog (again) with the updated control settings
        /// return false, if no more changes need to be made
        /// return false will have the same effect as pressing the OK button
        /// </returns>
        static bool validate(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            // just some code to set a break point
            int i = index;

            return true; // return to dialog
        }

    [ExcelCommand(
    Name = "Dialog1",
    Description = "Starts an example Dialog",
    HelpTopic = "XlmDialogExample-AddIn.xll.chm!1001",
    ShortCut = "^R")]
        public static void Dialog1Macro()
        {
            var dialog = new XlDialogBox()                      { Width = 494, Height = 210, Text = "Generic Sample Dialog" };

            // To do: use reflection to get the HelpTopic as defined in the [ExcelCommand(HelpTopic)] attribute.  
            dialog.HelpTopic = "1001";

            var okBtn = new XlDialogBox.OkButton()              { X = 209, Y = 174, Width = 075, Height = 023 };
            var cancelBtn = new XlDialogBox.CancelButton()      { X = 296, Y = 174, Width = 075, Height = 023 };
            var helpBtn = new XlDialogBox.HelpButton2()         { X = 384, Y = 174, Width = 075, Height = 023 };

            var nameLabel = new XlDialogBox.Label               { X = 019, Y = 011, Text = "&Name:" };
            var nameEdit = new XlDialogBox.TextEdit             { X = 019, Y = 029, IO_string = "<Name>" };

            var refLabel = new XlDialogBox.Label                { X = 019, Y = 050, Text = "&Reference" };
            var refEdit = new XlDialogBox.RefEdit               { X = 019, Y = 067, Width = 253, };

            var listEdit = new XlDialogBox.ListBox()            { X = 019, Y = 099, Width = 160, Height = 96, IO_index = 2, Text = "GENERIC_List1" };
            listEdit.Items.AddRange(new string[]                { "Bake", "Broil", "Sizzle", "Fry", "Saute" });

            var educateBox = new XlDialogBox.GroupBox           { X = 305, Y = 015, Width = 154, Height = 073, Text = "College" };
            var RadioGroup = new XlDialogBox.RadioButtonGroup   { IO_index = 1 };
            var RadioHarvr = new XlDialogBox.RadioButton        { Text = "&Harvard" };
            var RadioOther = new XlDialogBox.RadioButton        { Text = "&Other" };

            var qualiGroup = new XlDialogBox.GroupBox           { X = 209, Y = 093, Width = 250, Height = 063, Text = "&Qualifications" };
            var BaBsCheck = new XlDialogBox.CheckBox            { Text = "&BA / BS", IO_checked = true };
            var MaMsCheck = new XlDialogBox.CheckBox            { Text = "&MA / MS", IO_checked = true };

            // note: setting Trigger = true for PhD_Check (or any other triggerable control) will initiate the DDV callback function
            var PhD_Check = new XlDialogBox.CheckBox            { Text = "&PhD / other Grad", Trigger = true };

            // The sequence of adding controls is important in view of the tab order.
            // Note: always put the 'labels' *in front* of their (edit/list) controls.
            dialog.Controls.Add(nameLabel);
            dialog.Controls.Add(nameEdit);

            dialog.Controls.Add(refLabel);
            dialog.Controls.Add(refEdit);

            dialog.Controls.Add(listEdit);

            dialog.Controls.Add(educateBox);
            dialog.Controls.Add(RadioGroup);
            dialog.Controls.Add(RadioHarvr);
            dialog.Controls.Add(RadioOther);

            dialog.Controls.Add(qualiGroup);
            dialog.Controls.Add(BaBsCheck);
            dialog.Controls.Add(MaMsCheck);
            dialog.Controls.Add(PhD_Check);

            dialog.Controls.Add(okBtn);
            dialog.Controls.Add(cancelBtn);
            dialog.Controls.Add(helpBtn);

            bool bOK = dialog.ShowDialog(validate);
            if (bOK == false) return;

            // now it is time to play around with the command to get things done
            var xlApp = (Application)ExcelDnaUtil.Application;
            var ws = xlApp.Sheets[1] as Worksheet;
            var range = ws.Cells[1, 1] as Range;
            range.Value2 = nameEdit.IO_string;
        }
    }
}
