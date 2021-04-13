using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

using ExcelDna.Integration;

// The two original source files are coming from: https://github.com/zwq00000/ExcelDna-XlDialog.
// Text in this file has been translated from Japanese ("文本编辑控件") to English ("Text editing controls") using:
// https://translate.yandex.com/?lang=zh-en&text=%E6%96%87%E6%9C%AC%E7%BC%96%E8%BE%91%E6%8E%A7%E4%BB%B6
// Several bugs have been fixed and various enhancements have been added.

// Useful information on Excel 4.0 macro functions can be found here :
// https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf

#region XlDialogBox Introduction

/*    
    The dialog box definition table must be at least two rows high, and shall be seven columns wide. 
    The definitions of each column in a dialog box definition table are listed in the following table.

    | Column type               |Col.|Index|
    | --------------------------|----|-----|
    | Item number               |  1 |  0  |
    | Horizontal position       |  2 |  1  |
    | Vertical position         |  3 |  2  |
    | Item width                |  4 |  3  |
    | Item height               |  5 |  4  |
    | Item text                 |  6 |  5  |
    | Initial value / result    |  7 |  6  |

    The first row of the dialog box definition table defines the position, size, and name of the dialog box itself. 
    It can also specify the default selected item and the reference for the Help button. 
    The dialog position is specified in columns 2 and 3, the size in columns 4 and 5, and the name in column 6. 
    To specify a default start item, place the item's position number in column 7. 

    You can place the reference for the Help button in row 1, column 1 of the table, 
    but the preferred location is column 7 in the row where the Help button is defined. 
    Row 1, column 1 is usually left blank.

    The following table lists the numbers for the items you can display in a dialog box.

    | Dialog-box item                       |Item No |
    |---------------------------------------|--------|
    | Default OK button                     |       1|
    | Cancel button                         |       2|
    | OK button                             |       3|
    | Default Cancel button                 |       4|
    | Static text                           |       5|
    | Text edit box                         |       6|
    | Integer edit box                      |       7|
    | Number edit box                       |       8|
    | Formula edit box                      |       9|
    | Reference edit box                    |      10|
    | Option button group                   |      11|
    | Option button                         |      12|
    | Check box                             |      13|
    | Group box                             |      14|
    | List box                              |      15|
    | Linked list box                       |      16|
    | Icons                                 |      17|
    | Linked file list box                  |      18|  (Microsoft Excel for Windows only) 
    | Linked drive and directory box        |      19|  (Microsoft Excel for Windows only)
    | Directory text box                    |      20|
    | Drop-down list box                    |      21|
    | Drop-down combination edit/list box   |      22|
    | Picture button                        |      23|
    | Help button                           |      24|

    ### Remarks

    Adding 100 to certain item numbers causes the function to return control to the DLL code when the item is clicked on with the dialog still displayed. 
    This "trigger feature" enables the xlfDialogBox command to alter the dialog, validate input and so on, and returning for more user interaction. 
    The position of the item number chosen in this way is returned in the 1st row, 7th column of the returned array. (Accessible through XlDialogBox.IO.)
    This feature does not work with static text (item 5) edit boxes (6, 7, 8, 9 and 10), group boxes (14), pictures (23) or the help button (24). 
    Those controls just ignore the trigger request if 100 would be added to their item numbers.
    
    Adding 200 to any item number, disables (greys-out) the item. A disabled item cannot be chosen or selected. For example, 203 is a disabled OK button. 
    You could for instance use item 223 to include a picture in your dialog box that does not behave like a button.
    
    Most of the dialog items are simple and no further explanation is required. For some a little more explanation is helpful.

    ### Text and edit boxes

    Vertical alignment of a text label to the text that appears in an edit box is important aesthetically. 
    For edit boxes with the default height (set by leaving the height field blank)
    This is achieved by setting the vertical position of the text to be that of the edit box + 3.

    ### Buttons

    Selecting a cancel button (item 2 or 4) causes the dialog to terminate returning FALSE. 
    Pressing any other button causes the function to return the offset of that button in the definition table in the 7th column, 1st row of the returned array.

    ### Radio buttons

    A group of radio buttons (12) must be preceded immediately by a radio group item (11) and must be uninterrupted by other item types. 
    If the radio group item has no text label, the group is not contained within a border. 
    If the height and/or width of the radio group are omitted but text is provided, a border is drawn that surrounds the radio buttons and their labels.

    ### List-boxes
    
    The text supplied in a list box item row should either be a name (DLL-internal or on a worksheet) that resolves to a literal array or range of cells. 
    It can also be a string that looks like a literal array, e.g. "{1,2,3,4,5,\"A\",\"B\",\"C\"}" (where coded in a C source file). 
    List-boxes return the position (counting from 1) of the selected item in the list in the 7th column of the list-box item line. 
    Drop-down list-boxes (21) behave exactly as list boxes (15) except that the list is only displayed when the item is selected.

    ### Linked list-boxes

    Linked list-boxes (16), linked file-boxes (18) and drop-down combo-boxes (22) should be preceded immediately by an edit box that can support the data types in the list. 
    The lists themselves are drawn from the text field of the definition row which should be a range name or a string that represents a static array. 
    A linked path box (19) must be preceded immediately by a linked file-box (18).
    Drop down combo-boxes return the value selected in the 7th column of the associated edit box and the position (counting from 1) of the selected item
    in the list in the 7th column of the combo-box item line.


*** This example dialog comes from GENERIC.C from the Microsoft "2013 Office System Developer Resources" XLL-toolkit ***

#define g_rgDialogRows 16
#define g_rgDialogCols 7

static LPWSTR g_rgDialog[g_rgDialogRows][g_rgDialogCols] =
{
    {L"\000",   L"\000",    L"\000",    L"\003494", L"\003210", L"\025Generic Sample Dialog", L"\000"},
    {L"\0011",  L"\003330", L"\003174", L"\00288",  L"\000",    L"\002OK",                    L"\000"},
    {L"\0012",  L"\003225", L"\003174", L"\00288",  L"\000",    L"\006Cancel",                L"\000"},
    {L"\0015",  L"\00219",  L"\00211",  L"\000",    L"\000",    L"\006&Name:",                L"\000"},
    {L"\0016",  L"\00219",  L"\00229",  L"\003251", L"\000",    L"\000",                      L"\000"},
    {L"\00214", L"\003305", L"\00215",  L"\003154", L"\00273",  L"\010&College",              L"\000"},
    {L"\00211", L"\000",    L"\000",    L"\000",    L"\000",    L"\000",                      L"\0011"},
    {L"\00212", L"\000",    L"\000",    L"\000",    L"\000",    L"\010&Harvard",              L"\0011"},
    {L"\00212", L"\000",    L"\000",    L"\000",    L"\000",    L"\006&Other",                L"\000"},
    {L"\0015",  L"\00219",  L"\00250",  L"\000",    L"\000",    L"\013&Reference:",           L"\000"},
    {L"\00210", L"\00219",  L"\00267",  L"\003253", L"\000",    L"\000",                      L"\000"},
    {L"\00214", L"\003209", L"\00293",  L"\003250", L"\00263",  L"\017&Qualifications",       L"\000"},
    {L"\00213", L"\000",    L"\000",    L"\000",    L"\000",    L"\010&BA / BS",              L"\0011"},
    {L"\00213", L"\000",    L"\000",    L"\000",    L"\000",    L"\010&MA / MS",              L"\0011"},
    {L"\00213", L"\000",    L"\000",    L"\000",    L"\000",    L"\021&PhD / Other Grad",     L"\0010"},
    {L"\00215", L"\00219",  L"\00299",  L"\003160", L"\00296",  L"\015GENERIC_List1",         L"\0011"},
};

This table above is hard to read, it is best to strip off the L'00x bits for readability. This results in 

    { 00, 000, 000, 494, 210, "Generic Sample Dialog", 0},
    { 01, 330, 174, 088, 000, "OK",                    0},
    { 02, 225, 174, 088, 000, "Cancel",                0},
    { 05, 019, 011, 000, 000, "&Name:",                0},
    { 06, 019, 029, 251, 000, 0000,                    0},
    { 14, 305, 015, 154, 073, "&College",              0},
    { 11, 000, 000, 000, 000, 0000,                    1},
    { 12, 000, 000, 000, 000, "&Harvard",              1},
    { 12, 000, 000, 000, 000, "&Other",                0},
    { 05, 019, 050, 000, 000, "&Reference:",           0},
    { 10, 019, 067, 253, 000, 000,                     0},
    { 14, 209, 093, 250, 063, "&Qualifications",       0},
    { 13, 000, 000, 000, 000, "&BA / BS",              1},
    { 13, 000, 000, 000, 000, "&MA / MS",              1},
    { 13, 000, 000, 000, 000, "&PhD / Other Grad",     0},
    { 15, 019, 099, 160, 096, "GENERIC_List1",         1},
           
This translates to the following using XlDialogBox

    var dialog = new XlDialogBox()                         { Width = 494, Height = 210, Text = "Generic Sample Dialog" };

    var okBtn = new XlDialogBox.OkButton()                 { X = 330, Y = 174, Width = 088, Height = 023, Text = "&OK" };
    var cancelBtn = new XlDialogBox.CancelButton()         { X = 225, Y = 174, Width = 088, Height = 023, Text = "&Cancel" };

    var nameLabel = new XlDialogBox.Label                  { X = 019, Y = 011, Text = "&Name:" };
    var nameEdit = new XlDialogBox.TextEdit                { X = 019, Y = 029};

    var educatBox = new XlDialogBox.GroupBox               { X = 305, Y = 015, Width = 154, Height = 073, Text = "College" };
    var RadioGroup = new XlDialogBox.RadioButtonGroup      { };
    var RadioHarvr = new XlDialogBox.RadioButton           { Text = "&Harvard" };
    var RadioOther = new XlDialogBox.RadioButton           { Text = "&Other" };

    var refLabel = new XlDialogBox.Label                   { X = 019, Y = 050, Text = "&Reference" };
    var refEdit = new XlDialogBox.RefEdit                  { X = 019, Y = 067, Width = 253, };

    var qualiGroup = new XlDialogBox.GroupBox              { X = 209, Y = 093, Width = 250, Height = 063, Text = "&Qualifications" };
    var BaBsCheck = new XlDialogBox.CheckBox               { Text = "&BA / BS", Value = true };
    var MaMsCheck = new XlDialogBox.CheckBox               { Text = "&MA / MS", Value = true };
    var PhD_Check = new XlDialogBox.CheckBox               { Text = "&PhD / other Grad" };

    var listEdit = new XlDialogBox.ListBox()               { X = 019, Y = 099, Width = 160, Height = 96, SelectedIndex = 2 };
    listEdit.Items.AddRange(new string[]                { "Bake", "Broil", "Sizzle", "Fry", "Saute" });

    // The sequence of adding controls is important in view of the tab order.
    // Note: always put the 'labels' *in front* of their (edit/list) controls.

    dialog.Controls.Add(nameLabel);
    dialog.Controls.Add(nameEdit);

    dialog.Controls.Add(educatBox);
    dialog.Controls.Add(RadioGroup);
    dialog.Controls.Add(RadioHarvr);
    dialog.Controls.Add(RadioOther);

    dialog.Controls.Add(refLabel);
    dialog.Controls.Add(refEdit);

    dialog.Controls.Add(qualiGroup);
    dialog.Controls.Add(BaBsCheck);
    dialog.Controls.Add(MaMsCheck);
    dialog.Controls.Add(PhD_Check);

    dialog.Controls.Add(listEdit);

    dialog.Controls.Add(okBtn);
    dialog.Controls.Add(cancelBtn);

    dialog.ShowDialog();

*** The next example is from the book : Excel Add-in Development in C/C++ - Applications in Finance - By Steve Dalton (page 159) ***

#define NUM_DIALOG_COLUMNS 7
#define NUM_DIALOG_ROWS 10
cpp_xloper UsernameDlg[NUM_DIALOG_ROWS * NUM_DIALOG_COLUMNS] =
{
    //1,   2,   3,   4,   5,                                          6,           7, // Column
     "",  "",  "", 372, 200,                                    "Logon",          "", // Dialog box size & title
    001, 100, 170, 090,  "",                                       "OK",          "", // Default OK button
    002, 200, 170, 090,  "",                                   "Cancel",          "", // Cancel button
    005, 040, 010,  "",  "", "Please enter your username and password.",          "", // Static text
    014, 040, 035, 290, 100,                                         "",          "", // Group box
    005, 050, 053,  "",  "",                                 "Username",          "", // Static text
    006, 150, 050,  "",  "",                                         "",    "MyName", // Text edit box (pre-filled)
    005, 050, 073,  "",  "",                                 "Password",          "", // Static text
    006, 150, 070,  "",  "",                                         "", "*********", // Text edit box (pre-filled)
    013, 050, 110,  "",  "",           "Remember username and password",        true, // Checkbox (initialised as "true")
};

*** The last example comes from : https://github.com/zwq00000/ExcelDna-XlDialog ***

    var dialog = new XlDialogBox()                         { Width = 337, Height = 255, Text = "TestDialog" };

    // here are the dialog controls from top left to bottom right
    var groupBox = new XlDialogBox.GroupBox()              { X = 012, Y = 012, Width = 312, Height = 183, Text = "Define cells" };

    var labelForNames = new XlDialogBox.Label("&Name:")    { X = 029, Y = 040, Width = 066 };
    var nameEdit = new XlDialogBox.DropdownList()          { X = 101, Y = 037, Width = 202, SelectedIndex = 2 };
    nameEdit.Items.AddRange(new string[] { "Item 1", "Item 2", "Item 3" });

    var labelForCaption = new XlDialogBox.Label("&Title:") { X = 029, Y = 078, Width = 068 };
    var captionEdit = new XlDialogBox.TextEdit()           { X = 101, Y = 075, Width = 202, Text = "<enter title>" };

    var labelForAddress = new XlDialogBox.Label("&Cell:")  { X = 027, Y = 117, Width = 068 };
    var addressEdit = new XlDialogBox.RefEdit()            { X = 101, Y = 114, Width = 202 };

    var labelForValue = new XlDialogBox.Label("&Value:")   { X = 027, Y = 156, Width = 068 };
    var valueEdit = new XlDialogBox.TextEdit()              { X = 101, Y = 153, Width = 202, Value = "7" };

    var okBtn = new XlDialogBox.OkButton()                 { X = 169, Y = 220, Width = 075, Height = 023, Text = "&OK" };
    var cancelBtn = new XlDialogBox.CancelButton()         { X = 250, Y = 220, Width = 075, Height = 023, Text = "&Cancel" };

    // The sequence of adding controls is important in view of the tab order.
    // Note: always put the 'labels' *in front* of their (edit/list) controls.
    dialog.Controls.Add(groupBox);

    dialog.Controls.Add(labelForNames);
    dialog.Controls.Add(nameEdit);

    dialog.Controls.Add(labelForCaption);
    dialog.Controls.Add(captionEdit);

    dialog.Controls.Add(labelForAddress);
    dialog.Controls.Add(addressEdit);

    dialog.Controls.Add(labelForValue);
    dialog.Controls.Add(valueEdit);

    dialog.Controls.Add(okBtn);
    dialog.Controls.Add(cancelBtn);

    
    dialog.ShowDialog();


*** This code snippet helped me to set up a ShowDialog function using data validation ***
    Data validation is (only) applied when you use 'triggers' in the dialog controls

    Example of a typical dialog validation loop
    int __stdcall get_username(void)
    {
        xloper ret_val;
        int xl4;
        cpp_xloper DialogDef((WORD) NUM_DIALOG_ROWS,
        (WORD) NUM_DIALOG_COLUMNS, UsernameDlg);
        do
        {
            xl4 = Excel4(xlfDialogBox, &ret_val, 1, &DialogDef);
            if (xl4 || (ret_val.xltype == xltypeBool
            && ret_val.val._bool == 0))
                break;
            // Process the input from the dialog by reading
            // the 7th column of the returned array.
            // ... code omitted
            Excel4(xlFree, 0, 1, &ret_val);
            ret_val.xltype = xltypeNil;
        }
        while (1);
        Excel4(xlFree, 0, 1, &ret_val);
        return 1;
    }
*/
#endregion XlDialogBox Introduction

#region RefEdit in WPF Forms
/*
 * Please not that the use of XlfDialogBox has been a "step back in time" compared to using WPF forms.
 * One reason NOT to use WPF forms is that there is no good functioning RefEdit control available
 * Several projects on CodeProject created RefEdit alternatives, but these projects were hard to compile.
 * For that reason I reverted back to using the XLM macro "XlfDialogBox" to create a modal dialog
 * When I have time to develop a WPF Dialog Wizard instead, I'll need the code snippet below.

// Code from : https://www.breezetree.com/blog/excel-refedit-in-c-sharp/
#region ---- Select range ----------------------------------------
private void btnSelectRange_Click(object sender, EventArgs e)
{
    string prompt = "Select the range";
    string title = "Select Range";
    try
    {
        string address = Utilities.PromptForRangeAddress(this, title, prompt);
        if (!String.IsNullOrEmpty(address))
        {
            txtBaseShapeCell.Text = address;
        }
    }
    catch
    {
        MessageBox.Show("An error occurred when selecting the range.", "Range Error");
    }
}
#endregion

#region ---- static utility methods ----------------------------------------
// Requires:
// using System.Runtime.InteropServices;
// using System.Windows.Forms;
DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
[StructLayout(LayoutKind.Sequential)]
private struct RECT
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}

public static string PromptForRangeAddress(Form form, string title, string prompt)
{
    Size windowSize = form.Size;
    form.Size = SystemInformation.MinimumWindowSize;
    Point location = form.Location;
    SetFormPositionForInputBox(form);
    string rangeAddress = string.Empty;
    Excel.Range range = null;
    try
    {
        range = XL.App.InputBox(prompt, title, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, 8) as Excel.Range;
        if (range != null)
            rangeAddress = range.get_AddressLocal(Office.MsoTriState.msoFalse,
                                                  Office.MsoTriState.msoFalse,
                                                  Excel.XlReferenceStyle.xlA1,
                                                  Office.MsoTriState.msoFalse,
                                                  Type.Missing);
    }
    catch
    {
        throw new Exception("An error occured when selecting the range.");
    }
    finally
    {
        form.Location = location;
        form.Size = windowSize;
        MRCO(range);
    }
    return rangeAddress;
}

public static Excel.Range PromptForRange(Form form, string title, string prompt)
{
    Size windowSize = form.Size;
    Point location = form.Location;
    form.Size = SystemInformation.MinimumWindowSize;
    SetFormPositionForInputBox(form);
    Excel.Range range = null;
    try
    {
        range = XL.App.InputBox(prompt, title, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, 8) as Excel.Range;
    }
    catch
    {
        throw new Exception("An error occured when selecting the range.");
    }
    finally
    {
        form.Location = location;
        form.Size = windowSize;
    }
    return range;
}
public static void SetFormPositionForInputBox(Form form)
{
    int x = form.Location.X;
    bool isSet = false;
    try
    {
        System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("Excel");
        if (processes != null && processes.Length > 0 && processes[0] != null)
        {
            RECT rect;
            IntPtr ptrXL = processes[0].MainWindowHandle;
            if (!ptrXL.Equals(IntPtr.Zero) && GetWindowRect(ptrXL, out rect))
            {
                form.Location = new Point(x, rect.Bottom - SystemInformation.MinimumWindowSize.Height);
                isSet = true;
            }
        }
    }
    finally
    {
        if (!isSet)
        {
            form.Location = new Point(x, 0);
        }
    }
}
public static void MRCO(object obj)
{
    if (obj == null) { return; }
    try
    {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
    }
    catch
    {
        // ignore, cf: http://support.microsoft.com/default.aspx/kb/317109
    }
    finally
    {
        obj = null;
    }
}
#endregion

*/


/*
 * this code snippet is not "at home" here. It is intended for use with WPF dialogs, where you cannot use RefEdit. 
 * It may be useful, when stepping away from the (old) XML macro language that underpins the xlfDialogBox calls
    try
    {
        range = _excel.Application.InputBox("InputRange", "Title", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8) as Excel.Range;

        if (range != null)
        {
            *//* rangeAddress = range.get_AddressLocal(Office.MsoTriState.msoFalse,
                                                    Office.MsoTriState.msoFalse,
                                                    Excel.XlReferenceStyle.xlA1,
                                                    Office.MsoTriState.msoFalse,
                                                    Type.Missing);

            rangeAddress = range.get_Address(0, 0, Excel.XlReferenceStyle.xlA1, 0, Type.Missing);

            // See: https://stackoverflow.com/questions/27069471/excel-buttons-to-call-wpf-windows for WPF & Excel
            // See: https://stackoverflow.com/questions/51018904/get-workbook-name-and-worksheet-name-from-a-range-in-excel-vba                    
            string wbookName = string.Empty;
            string wbookCode = string.Empty;
            string sheetName = string.Empty;
            wbookName = range.Parent.Parent.Name;     // Name of the workbook
            wbookCode = range.Parent.CodeName;        // Code Name of the worksheet
            sheetName = range.Parent.Name;            // Name of the worksheet
            // example range: '[Hours_worked.xlsx]Hours March'!$E$11:$E$16

            rangeAddress = "'[" + wbookName + "]" + sheetName + "'!" + rangeAddress;
        }
    }
    catch
    {

    }
    finally
    {
        MessageBox.Show("Hello from range: " + rangeAddress);
        //            DataWriter.WriteData();
    }
*/

#endregion RefEdit in WPF Forms

/// <summary>
/// The namespace XlDialogBox is named after the XLM Macro function DIALOG.BOX
/// The XLM macro's date back from Excel 4.0, prior to the use of VBA macro's
/// See: https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf
/// See: https://exceloffthegrid.com/using-excel-4-macro-functions/ for some guidance & examples
/// The 'old' XLM DIALOG.BOX macro is still supported today, by calling Excel with xlfDialogBox:
/// 
///     Excel4(xlfDialogBox, &ret_val, 1, &DialogDef);
/// 
/// The value of DIALOG.BOX lays in the fact that neither WinForms nor WPF have a good RefEdit control :(
/// </summary>
namespace ExcelDna.XlDialogBox 
{
    #region Data validation

    /// <summary>
    ///     Define a delegate function for use as callback during dialog data validation
    ///     The purpose of the DDV routine is to expose some of the 'guts' of the dialog
    ///     A validation routine can change some parameters which are fed back to the dialog
    /// </summary>
    /// <param name="index">
    ///     indicates the row of the Control in the dialogResult table that caused the dialog to return
    /// </param>
    /// <param name="dialogResult">
    ///     By accessing to the dialogResult array, one can check what has happended with the dialog data
    ///     Any changes to the controls that **exist** in the current dialog need to be made here     
    /// </param>
    /// <param name="Controls">
    ///     Access to the Dialog Controls is given for the purpose of making some controls (in-)visible.
    ///     The invisible controls are not part of the controls that **exist** in the current dialog box. 
    ///     Before the dialog is shown (again) the dialog_ref table is built up from the Dialog Controls.
    ///     So controls that were earlier invisible, now become visible, and vice versa.
    /// </param>
    /// <returns>
    /// 'true' if we need to give control back to the dialog; 'false' if we are done with the dialog
    /// when 'true' is returned, the ShowDialog function updates the Controls with info from dialogResult
    /// </returns>
    public delegate bool DDV(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls);
    
    #endregion Data validation

    /// <summary>
    ///     DIALOG.BOX(dialog_ref)
    ///     Dialog_ref is a reference to a dialog box definition table on sheet, or an array containing the definition table.
    /// </summary>
    public class XlDialogBox 
    {
        #region Enumerations
        /// <summary>
        ///     XlDialogBox Control type
        /// </summary>
        public enum XlControl
        {
            /// <summary>
            ///     Empty (undefined) type
            /// </summary>
            XlEmpty = -1,

            /// <summary>
            ///     Default OK button   1
            /// </summary>
            XlDefaultOkButton = 1,

            /// <summary>
            ///     Cancel button   2
            /// </summary>
            XlCancelButton = 2,

            /// <summary>
            ///     OK button   3
            /// </summary>
            XlOkButton = 3,

            /// <summary>
            ///     Default Cancel button   4
            /// </summary>
            XlDefaultCancelButton = 4,

            /// <summary>
            ///     Static text   5
            /// </summary>
            XlStaticText = 5,

            /// <summary>
            ///     Text edit box   6
            /// </summary>
            XlTextBox = 6,

            /// <summary>
            ///     Integer edit box   7
            /// </summary>
            XlIntegerEedit = 7,

            /// <summary>
            ///     Number edit box   8
            /// </summary>
            XlNumberEdit = 8,

            /// <summary>
            ///     Formula edit box   9
            /// </summary>
            XlFormulaEdit = 9,

            /// <summary>
            ///     Reference edit box   10
            /// </summary>
            XlReferenceEdit = 10,

            /// <summary>
            ///     Radio button group   11
            /// </summary>
            XlRadioButtonGroup = 11,

            /// <summary>
            ///     Radio button   12
            /// </summary>
            XlRadioButton = 12,

            /// <summary>
            ///     Check box   13
            /// </summary>
            XlCheckBox = 13,

            /// <summary>
            ///     Group box   14
            /// </summary>
            XlGroupBox = 14,

            /// <summary>
            ///     List box   15
            /// </summary>
            XlListBox = 15,

            /// <summary>
            ///     Linked list box   16
            /// </summary>
            XlLinkedListBox = 16,

            /// <summary>
            ///     Icons   17
            /// </summary>
            XlIcons = 17,

            /// <summary>
            ///     Linked file list box   18 
            ///     (Microsoft Excel for Windows only)
            /// </summary>
            XlLinkedFileListBox = 18,

            /// <summary>
            ///     Linked drive and directory box   19 
            ///     (Microsoft Excel for Windows only)
            /// </summary>
            XlLinkedDriveDirBox = 19,

            /// <summary>
            ///     Directory text box   20
            /// </summary>
            XlDirectoryTextbox = 20,

            /// <summary>
            ///     Drop-down list box   21
            /// </summary>
            XlDropdownList = 21,

            /// <summary>
            ///     Drop-down combination edit/list box   22
            /// </summary>
            XlCombobox = 22,

            /// <summary>
            ///     Picture button   23
            /// </summary>
            XlPictureButton = 23,

            /// <summary>
            ///     Help button   24
            /// </summary>
            XlHelpButton = 24,

            /// <summary>
            ///     trigger + ItemNum
            /// </summary>
            XlTrigger = 100,

            /// <summary>
            ///     disable + ItemNum
            /// </summary>
            XlDisable = 200,

            // dummy to test code
            XlInvisibleTextBox = 999,

            /// <summary>
            ///     invisible + ItemNum
            /// </summary>
            /// <remarks>
            ///     UNDOCUMENTED 'hack' to make items invisible by giving a dialog item an ItemNumber > 24
            ///     This is a POSITIVE offset shifting the XlControl number out of range of the available control numbers
            ///     A negative offset (e.g. by flipping the polarity) would lead to exceptions in Excel and must be avoided
            ///     The concept is in line with "Trigger" which adds 100 and "Disable" which adds 200
            /// </remarks>
            XlInvisible = 50
        }

        /// <summary>
        ///     XlDialogBox Control type
        /// </summary>
        public enum XlColumn
        {
            /// <summary>
            ///     Item number in first column
            /// </summary>
            XlNumberColumn = 0,

            /// <summary>
            ///     Horizontal position in second column
            /// </summary>
            XlHoriPosColumn = 1,

            /// <summary>
            ///     Vertical position in third column
            /// </summary>
            XlVertPosColumn = 2,

            /// <summary>
            ///     Item width in fourth column
            /// </summary>
            XlWidthColumn = 3,

            /// <summary>
            ///     Item height in fifth column
            /// </summary>
            XlHeightColumn = 4,

            /// <summary>
            ///     Item text in sixt column
            /// </summary>
            XlTextColumn = 5,

            /// <summary>
            ///     Item value in seventh column
            /// </summary>
            XlIOColumn = 6
        }
        #endregion Enumerations

        #region Class members
        /// <summary>
        ///     The collection of controls containing a record for the dialog and the buttons it 'owns'
        /// </summary>
        public readonly XlDialogControlCollection Controls = new XlDialogControlCollection();

        /// <summary>
        ///     Result object from XlDialogBox call
        /// </summary>
        private object _resultObject;

        /// <summary>
        ///     Result array from XlDialogBox call
        /// </summary>
        private object[,] _resultArray;

        /// <summary>
        ///     Dialog form definition
        /// </summary>
        /// <remarks>
        ///     The first row of dialog_ref defines the position, size, and name of the dialog box.
        ///     It can also specify the default selected item and the reference for the Help button.
        ///     The position is specified in columns 2 and 3, the size in columns 4 and 5, and the name in column 6.
        ///     To specify a default item, place the item's position number in column 7.
        ///     You can place the reference for the Help button in row 1, column 1 of the table,
        ///     but the preferred location is column 7 in the row where the Help button is defined. 
        ///     Row 1, column 1 is usually left blank.
        /// </remarks>
        private readonly ControlItem _formControl;

        public System.Reflection.MethodBase CallingMethod = null;

        #endregion Class members

        #region Dialog Constructor
        public XlDialogBox() 
        {
            _formControl = new ControlItem(XlControl.XlEmpty);
            Controls.Add(_formControl); // _formControl becomes the first row of the Controls collection
            W = 300;
            H = 200;
            Text = "XlDialogBox";
        }
        #endregion Dialog Constructor

        #region Dialog Get Set routines
        /// <remarks>
        ///     _formControl is created in the Constructor, so we don't have to check for IsNull() here....
        /// </remarks>

        /// <summary>
        ///     Dialog position X
        /// </summary>
        public int X 
        {
            get { return _formControl.X; }
            set { _formControl.X = value; }
        }

        /// <summary>
        ///     Dialog position Y
        /// </summary>
        public int Y 
        {
            get { return _formControl.Y; }
            set { _formControl.Y = value; }
        }

        /// <summary>
        ///    Dialog width
        /// </summary>
        public int W 
        {
            get { return _formControl.W; }
            set { _formControl.W = value; }
        }

        /// <summary>
        ///     Dialog height
        /// </summary>
        public int H 
        {
            get { return _formControl.H; }
            set { _formControl.H = value; }
        }

        /// <summary>
        ///     Dialog title
        /// </summary>
        public string Text 
        {
            get { return _formControl.Text; }
            set { _formControl.Text = value; }
        }

        /// <summary>
        ///     Access to 'raw' Dialog IO object
        /// </summary>
         public virtual object IO
        {
            get { return (_formControl.IO); }
            set { _formControl.IO = value; }
        }


        /// <summary>
        ///     access to index contained in Dialog IO
        /// </summary>
        public int IO_index
        {
            get { return Convert.ToInt32(_formControl.IO); }
            set { _formControl.IO = value; }
        }

        #endregion Dialog Get Set routines

        #region ShowDialog() implementations
        /// <summary>
        /// Show the dialog; and allows for data validation to take place
        /// </summary>
        /// <param name="x">x position of Dialog Box</param>
        /// <param name="y">y position of Dialog Box</param>
        /// <param name="dataValidation">callback function to check dialog controls when a control in the dialog issues a trigger</param>
        /// <returns>true if OK is selected, false if CANCEL is selected</returns>
        public virtual bool ShowDialog(int x, int y, DDV dataValidation = null)
        {
            X = x;
            Y = y;

            return ShowDialog(dataValidation);
        }

        /// <summary>
        /// Show the dialog; without any data validation to take place
        /// </summary>
        /// <param name="x">x position of Dialog Box</param>
        /// <param name="y">y position of Dialog Box</param>
        /// <returns>true if OK is selected, false if CANCEL is selected</returns>
        public virtual bool ShowDialog(int x, int y)
        {
            X = x;
            Y = y;

            return ShowDialog(null);
        }

        /// <summary>
        /// Show the dialog; and allows for data validation to take place
        /// </summary>
        /// <param name="dataValidation">callback function to check dialog controls when a control in the dialog issues a trigger</param>
        /// <returns>true if OK is selected, false if CANCEL is selected</returns>
        public virtual bool ShowDialog(DDV dataValidation = null)
        {
            try
            {
                bool loop = true;
                do
                {
                    // Build the array used as input for the XlCall.xlfDialogBox call
                    object[,] dialogDef = Controls.Build();

                    // Now start a modal dialog
                    _resultObject = XlCall.Excel(XlCall.xlfDialogBox, dialogDef);

                    // Convert the return value into an object array.
                    _resultArray = _resultObject as object[,];

                    // If cancel has been selected, _resultArray will be null
                    // In that case we can skip the parameter verification and quit
                    if (_resultArray == null)
                        return false;

                    // Need to update the IO-results; the (= IO) column of the Controls collection
                    //  (a) before returning back to dialog; because of pressing the "help" button
                    //  (b) before using the outcome of the dialog when OK has been pressed
                    Controls.UpdateResult(_resultArray);

                    // The code directly below is very ugly.
                    // Need to define ControlParameters as a class for easy parameter access !

                    // Get the type of control from the first column out of the IO_index row
                    XlControl triggerItem = (XlControl) Convert.ToInt32(_resultArray[this.IO_index, 0]);

                    // If this is the (default) OK button, we are done with the dialog, unless the 
                    if (triggerItem == XlControl.XlOkButton || triggerItem == XlControl.XlDefaultOkButton)
                    {
                        // We are creating a hack here to overcome the non-functioning help button
                        // if anybody knows how the properly use the help button please advise... 
                        // Right now, the help button is an OK button with a value > 0 in the IO column

                        // the IO column may contain a null pointer, so guard yourself against this
                        object triggerIO = _resultArray[this.IO_index, 6];

                        // check for OK and Helpo conditions
                        bool bOk = (triggerIO.IsNull() || Convert.ToInt32(triggerIO) == 0);
                        bool bHelp = (!triggerIO.IsNull() && Convert.ToInt32(triggerIO) < 0);

                        if (bOk) return true; // perform the normal OK exit

                        // if the 'help' button brought us here; launch help and revert back to the unaltered dialog
                        // do so without any data validation, as no trigger has occurred from a dialog item
                        if (bHelp)
                        {
                            // get the Path of xll file;
                            string xllPath = ExcelDnaUtil.XllPath;
                            string xllDir  = System.IO.Path.GetDirectoryName(xllPath);

                            if (CallingMethod != null)
                            {   // is there an ExcelCommandAttribute attribute decorating the method where ShowDialog has been called from ?
                                ExcelCommandAttribute attr = (ExcelCommandAttribute)CallingMethod.GetCustomAttributes(typeof(ExcelCommandAttribute), true)[0];
                                if (attr != null)
                                {   
                                    // get the HelpTopic string and split it in two parts ([a] file name and [b] helptopic)
                                    string[] parts = attr.HelpTopic.Split('!');

                                    // the complete helpfile path consists of the xll directory + first part of HelpTopic attribute string 
                                    string chmPath = System.IO.Path.Combine(xllDir, parts[0]);

                                    // See : http://www.help-info.de/en/Help_Info_HTMLHelp/hh_command.htm
                                    // Example of opening a help topic using help ID = 12030
                                    // ID is a number that you've defined in the [MAP] section of your project (*.hhp) file
                                    // and mapped to the required topic in the [ALIAS] section.
                                    // Note: The "-map ID chm" command line became available in HH 1.1b.
                                    // C:\> HH.EXE -mapid 12030 ms-its:C:/xTemp/XMLconvert.chm

                                    // get some help WITHOUT specifying HelpTopic
                                    // System.Diagnostics.Process.Start(chmPath);

                                    // get some help WITH specifying HelpTopic 
                                    // string helpArguments = "-mapid " + HelpTopic + " ms-its:" + "\"" + chmPath + "\"";
                                    System.Diagnostics.Process hh = new System.Diagnostics.Process();
                                    string helpArguments = "-mapid " + parts[1] + " ms-its:" + chmPath;
                                    hh.StartInfo.FileName = "HH.exe";
                                    hh.StartInfo.Arguments = helpArguments;
                                    hh.Start();
                                }
                            }
                            else
                            {
                                MessageBox.Show("Can't show context sensitive Help; XlDialogBox.CallingMethod is undefined");

                                // to do: show dialogbox : no ExcelCommandAttribute found with help information
                            }
                            continue; // continue with do/while loop and skip data validation when calling for help
                        }

                        // If we arrive here it is because the OK button wants us to keep looping through data validation
                        // Therefore do nothing, as 'loop' is already true
                    }
                                            
                    // If we arrive here it is because neither a CANCEL button nor an OK button was pressed.
                    // So now it is time to do the data validation check and see if we can leave the dialog.
                    if (dataValidation != null)
                        loop = dataValidation(this.IO_index, this._resultArray, this.Controls);
                    else
                        loop = true; // return control to the dialog in case no data validation is done
                }
                while (loop == true);

                return true;
            }
            finally
            {
                Controls.Dispose();

                // to do; call xlfree on _resultObject to prevent memory leaks
            }
        }

        /// <summary>
        ///     Show dialog box; this is the original code to show the Dialog Box
        ///     For reference; in future it should be merged by simply implementing:
        ///     public virtual bool dialog() { return ShowDialog(null); }
        ///     see other 'variations' of ShowDialog(...) above
        /// </summary>
        /// <returns></returns>
        public virtual bool ShowDialog()
        {
            try
            {
                var dialogDef = Controls.Build();
                var result = XlCall.Excel(XlCall.xlfDialogBox, dialogDef);
                _resultArray = result as object[,];
                if (_resultArray != null)
                {
                    Controls.UpdateResult(_resultArray);
                    return true;
                }
                return false;
            }
            finally
            {
                Controls.Dispose();
            }
        }
        #endregion ShowDialog() implementations

        #region ControlItem definition
        /// <summary>
        ///     XlDialogBox Control interface;
        ///     interface at the basis of ControlItem
        /// </summary>
        private interface IXlDialogControl : IDisposable 
        {
            /// <summary>
            ///     Item number; basically the Control Type defined by the XlControl enumeration
            ///     First column of a multi-row array with 7 columns
            /// </summary>
            XlControl ItemNumber { get; }

            /// <summary>
            ///     X Coordinate, if the value is less than 0, the default value is used
            ///     Second column of a multi-row array with 7 columns
            /// </summary>
            int X { get; set; }

            /// <summary>
            ///     Y Coordinate, if the value is less than 0, the default value is used
            ///     Third column of a multi-row array with 7 columns
            /// </summary>
            int Y { get; set; }

            /// <summary>
            ///     Width, if the value is less than 0, the default value is used
            ///     Fourth column of a multi-row array with 7 columns
            /// </summary>
            int W { get; set; }

            /// <summary>
            ///     Height, if the value is less than 0, the default value is used
            ///     Fift column of a multi-row array with 7 columns
            /// </summary>
            int H { get; set; }

            /// <summary>
            ///     Text content
            ///     Sixt column of a multi-row array with 7 columns
            /// </summary>
            string Text { get; set; }

            // Note: the seventh column contains an Initial Value or the Result upon return from dialog.  
            // Referred to as the "IO" member of a dialog control. Can be overloaded to be int or string.
            object IO { get; set; }
        }

        /// <summary>
        ///     base class for all dialog controls
        /// </summary>
        public class ControlItem : IXlDialogControl 
        {
            public ControlItem() 
            {
            }

            protected internal ControlItem(XlControl itemNumber)
            {
                ItemNumber = itemNumber;

                Trigger = false;
                Visible = true;
                Enable = true;
            }

            /// <summary>
            ///     Control definition array
            /// </summary>
            protected readonly object[] ControlParameters = new object[7];


            /// <summary>
            ///     Control index
            /// </summary>
            internal int Index { get; set; }

            internal object this[int index] 
            {
                get { return ControlParameters[index]; }
                set { ControlParameters[index] = value; }
            }

            /// <summary>
            ///     Control type
            /// </summary>
            public XlControl ItemNumber 
            {
                get 
                {
                    if (ControlParameters[(int)XlColumn.XlNumberColumn].IsNull())
                        return XlControl.XlEmpty;
                    else 
                        return (XlControl)ControlParameters[(int)XlColumn.XlNumberColumn];
                }

                protected set
                {
                    if (value < 0)
                        ControlParameters[(int)XlColumn.XlNumberColumn] = null;
                    else
                        ControlParameters[(int)XlColumn.XlNumberColumn] = (int)value;
                }
            }

            void ResetItemNumber()
            {
                if (ItemNumber > XlControl.XlDisable)
                    ItemNumber -= XlControl.XlDisable;

                // subtract 100 if this can be done 
                if (ItemNumber > XlControl.XlTrigger)
                    ItemNumber -= XlControl.XlTrigger;

                // subtract 50 if this can be done 
                if (ItemNumber > XlControl.XlInvisible)
                    ItemNumber -= (int)XlControl.XlInvisible;
            }

            /// <summary>
            ///     X Coordinate, if the value is less than 0, the default value is used
            /// </summary>
            public virtual int X 
            {
                get 
                {
                    if (ControlParameters[(int)XlColumn.XlHoriPosColumn].IsNull()) 
                        return -1;
                    else
                        return (int)ControlParameters[(int)XlColumn.XlHoriPosColumn];
                }
                
                set 
                {
                    if (value < 0) 
                        ControlParameters[(int)XlColumn.XlHoriPosColumn] = null;
                    else 
                        ControlParameters[(int)XlColumn.XlHoriPosColumn] = value;
                }
            }

            /// <summary>
            ///     Y Coordinate, if the value is less than 0, the default value is used
            /// </summary>
            public virtual int Y 
            {
                get 
                {
                    if (ControlParameters[(int)XlColumn.XlVertPosColumn].IsNull()) 
                        return -1;
                    else 
                        return (int)ControlParameters[(int)XlColumn.XlVertPosColumn];
                }
    
                set 
                {
                    if (value < 0) 
                        ControlParameters[(int)XlColumn.XlVertPosColumn] = null;
                    else 
                        ControlParameters[(int)XlColumn.XlVertPosColumn] = value;
                }
            }

            /// <summary>
            ///     Width, if the value is less than 0, the default value is used
            /// </summary>
            public virtual int W 
            {
                get 
                {
                    if (ControlParameters[(int)XlColumn.XlWidthColumn].IsNull()) 
                        return -1;
                    else
                        return (int)ControlParameters[(int)XlColumn.XlWidthColumn];
                }
                set {
                    if (value < 0)
                        ControlParameters[(int)XlColumn.XlWidthColumn] = null;
                    else 
                        ControlParameters[(int)XlColumn.XlWidthColumn] = value;
                }
            }

            /// <summary>
            ///     Height, if the value is less than 0, the default value is used
            /// </summary>
            [DefaultValue(20)]
            public virtual int H 
            {
                get 
                {
                    if (ControlParameters[(int)XlColumn.XlHeightColumn].IsNull()) 
                        return -1;
                    else
                        return (int)ControlParameters[(int)XlColumn.XlHeightColumn];
                }

                set 
                {
                    if (value < 0) 
                        ControlParameters[(int)XlColumn.XlHeightColumn] = -1;
                    else 
                        ControlParameters[(int)XlColumn.XlHeightColumn] = value;
                }
            }

            /// <summary>
            ///     Text content for column 6
            /// </summary>
            public virtual string Text 
            {
                get
                {
                    if (ControlParameters[(int)XlColumn.XlTextColumn].IsNull())
                        return "NULL";
                    else
                        return (string)ControlParameters[(int)XlColumn.XlTextColumn]; 
                }

                set 
                { 
                    ControlParameters[(int)XlColumn.XlTextColumn] = value; 
                }
            }

            /// <summary>
            ///     Information exchange in column 7
            /// </summary>
            public virtual object IO
            {
                get
                {
                    if (ControlParameters[(int)XlColumn.XlIOColumn].IsNull())
                        return null;
                    else
                        return ControlParameters[(int)XlColumn.XlIOColumn]; 
                }

                set
                {
                    ControlParameters[(int)XlColumn.XlIOColumn] = value;
                }
            }

            /// <summary>
            ///     Defines whether the control is visible; unlike 'Enable' this is not a built-in dialog control property.
            ///     It is added to the XlDialogBox class, by adding 50 to a dialog item to allow for *not* drawing it.
            /// </summary>
            [DefaultValue(true)]
            public bool Visible
            {
                get
                {
                    if (ItemNumber == XlControl.XlEmpty)
                        //Dialog form definition
                        return true;
                    else
                    {
                        // First make a copy of ItemNumber
                        XlControl tmp = this.ItemNumber;

                        // subtract 200 if this can be done 
                        if (tmp > XlControl.XlDisable)
                            tmp -= XlControl.XlDisable;

                        // subtract 100 if this can be done 
                        if (tmp > XlControl.XlTrigger)
                            tmp -= XlControl.XlTrigger;

                        // are we now inside the allow range ?
                        return ((tmp > 0) && (tmp <= XlControl.XlHelpButton));
                    }
                }

                set
                {
                    if (ItemNumber != XlControl.XlEmpty)
                    //Dialog form definition
                    {
                        if (value != Visible) // only take action if we need to make changes
                        {
                            if (value) // make visible; also remove trigger and disabled condition
                            {
                                ResetItemNumber();
                            }
                            else  // make invisible
                            {
                                ResetItemNumber();

                                // add 50 
                                ItemNumber += (int)XlControl.XlInvisible;
                            }
                        }
                    }
                }
            }

            /// <summary>
            ///     Is the control enabled ?
            /// </summary>
            [DefaultValue(true)]
            public bool Enable
            {
                get
                {
                    if (ItemNumber == XlControl.XlEmpty)
                        //Dialog form definition
                        return true;
                    else
                    {
                        // First make a copy of ItemNumber
                        XlControl tmp = this.ItemNumber;

                        // subtract 200 if this can be done 
                        if (tmp > XlControl.XlDisable)
                            tmp -= XlControl.XlDisable;

                        // subtract 100 if this can be done 
                        if (tmp > XlControl.XlTrigger)
                            tmp -= XlControl.XlTrigger;

                        // are we now inside the allow range ?
                        return ((tmp > 0) && (tmp <= XlControl.XlHelpButton) && ItemNumber < XlControl.XlDisable);
                    }

                }

                set
                {
                    if (ItemNumber != XlControl.XlEmpty)
                    //Dialog form definition
                    {
                        if (value != Enable) // only take action if we need to make changes
                        {
                            if (value) // enable; also remove trigger and invisible conditions
                            {
                                ResetItemNumber();
                            }
                            else  // disable
                            {
                                ResetItemNumber();

                                // finally add 200 
                                ItemNumber += (int)XlControl.XlDisable;
                            }
                        }
                    }
                }
            }

            /// <summary>
            ///     Is the control acting as a trigger ?
            /// </summary>
            [DefaultValue(false)]
            public bool Trigger
            {
                get
                {
                    if (ItemNumber == XlControl.XlEmpty)
                        //Dialog form definition
                        return false;
                    else
                    {
                        // First make a copy of ItemNumber
                        XlControl tmp = this.ItemNumber;

                        // Can't be a trigger when disabled
                        if (tmp > XlControl.XlDisable) 
                            return false;

                        return ((tmp > XlControl.XlTrigger) && (tmp - (int)XlControl.XlTrigger <= XlControl.XlHelpButton));
                    }
                }

                set
                {
                    if (ItemNumber != XlControl.XlEmpty)
                    // Don't do this on the Dialog form definition
                    {
                        if (value != Trigger) // only take action if we need to make changes
                        {
                            if (value) // enable trigger
                            {
                                ResetItemNumber();

                                // add 100 
                                ItemNumber += (int)XlControl.XlTrigger;
                            }
                            else  // disable trigger; also remove invisible and disabled condition
                            {
                                ResetItemNumber();
                            }
                        }
                    }
                }
            }

            /// <summary>
            ///     Get N x 7-parameter DialogDefinition table
            /// </summary>
            public virtual IEnumerable<object[]> GetControlParameters()
            {
                return new[] { ControlParameters };
            }

            /// <summary>
            ///     Perform application-defined tasks related to releasing or resetting unmanaged resources.
            /// </summary>
            public virtual void Dispose() { }

            /// <summary>
            ///     Call before building the dialog box
            /// </summary>
            protected internal virtual void OnBeforeBuild() {
            }
        }
        #endregion ControlItem  definition

        #region Ok-Cancel-Help Buttons

        /// <summary>
        ///     OK button; No longer sealed as other buttons are derived from it
        /// </summary>
        public class OkButton : ControlItem
        {
            public OkButton() : base(XlControl.XlOkButton)
            {
                Text = "&OK";
            }

            /// <summary>
            /// This constructor takes the button text as input
            /// </summary>
            /// <param name="text">Button text</param>
            public OkButton(string text) : base(XlControl.XlOkButton)
            {
                Text = text;
            }

            /// <summary>
            ///     Is it the default button ?
            /// </summary>
            public bool Default
            {
                get { return ItemNumber == XlControl.XlDefaultOkButton; }
                set { ItemNumber = value ? XlControl.XlDefaultOkButton : XlControl.XlOkButton; }
            }

            // experimental; see if we can use the IO Column for some fancy stuff
            public int IO_int
            {
                get { return Convert.ToInt32(base.IO); }
                set { base.IO = Convert.ToInt32(value); }
            }
        }

        /// <summary>
        ///     Cancel button
        /// </summary>
        public sealed class CancelButton : ControlItem
        {
            public CancelButton() : base(XlControl.XlCancelButton)
            {
                Text = "&Cancel";
            }

            /// <summary>
            /// This constructor takes the button text as input
            /// </summary>
            /// <param name="text">Button text</param>
            public CancelButton(string text) : base(XlControl.XlOkButton)
            {
                Text = text;
            }

            /// <summary>
            ///     Is it the default button
            /// </summary>
            public bool Default
            {
                get { return ItemNumber == XlControl.XlDefaultCancelButton; }
                set { ItemNumber = value ? XlControl.XlDefaultCancelButton : XlControl.XlCancelButton; }
            }
        }

        /// <summary>
        ///     Help button; it does NOT work !
        /// </summary>
        /// <remarks>
        ///     It really does NOT work !
        ///     As a Workaround, use an OkButton with "IO_int" set at '-1', and with "&Help" as button text.
        ///     When the dialog box returns, "IO_int" will be evaluated and the help file will be called.
        ///     
        ///     In the mean time I ordered "Greg Harvey's Excel 4.0 for the MAC" from the US: 
        ///     https://www.amazon.nl/gp/product/0679790446/ref=ppx_od_dt_b_asin_title_s00?ie=UTF8&psc=1
        ///     Let's see if this book provides some more information how to deal with DIALOG.BOX
        /// </remarks>
        public sealed class HelpButton : ControlItem
        {
            public HelpButton() : base(XlControl.XlHelpButton)
            {
                Text = "&Help";
            }

            // Alas, Help does not work as intended ...
            // Therefore use HelpButton2 as a workaround
            public string IO_string
            {
                get { return Convert.ToString(base.IO); }
                set { base.IO = Convert.ToString(value); }
            }
        }

        /// <summary>
        ///     Help button; replacing the one that does not work...
        /// </summary>
        /// <remarks>
        ///     As a Workaround, use OkButton with "IO_int" set at '-1', and with "&Help" as button text.
        ///     ShowDialog will launch hh.exe with its _HelpTopic; and return to the dialog without data valdation.
        /// </remarks>
        public sealed class HelpButton2 : OkButton
        {
            public HelpButton2()
            {
                Text = "&Help";
                IO_int = -1;
            }
        }

        /// <summary>
        ///     Next button; for use in a dialog wizard
        /// </summary>
        /// <remarks>
        ///     An OkButton is used with "IO_int" set at '1', and with "&Next >" as button text.
        ///     When the dialog box returns, "IO_int" will be evaluated and data validation will be done.
        /// </remarks>
        public sealed class NextButton : OkButton
        {
            public NextButton()
            {
                Text = "&Next >";
                IO_int = 1;
            }
        }

        /// <summary>
        ///     Back button; for use in a dialog wizard
        /// </summary>
        /// <remarks>
        ///     An OkButton is used with "IO_int" set at '2', and with "< &Back" as button text.
        ///     When the dialog box returns, "IO_int" will be evaluated and data validation will be done.
        /// </remarks>
        public sealed class BackButton : OkButton
        {
            public BackButton()
            {
                Text = "< &Back";
                IO_int = 2;
            }
        }

        /// <summary>
        ///     Apply button; for use in a dialog wizard
        /// </summary>
        /// <remarks>
        ///     An OkButton is used with "IO_int" set at '3', and with "&Apply" as button text.
        ///     When the dialog box returns, "IO_int" will be evaluated and data validation will be done.
        /// </remarks>
        public sealed class ApplyButton : OkButton
        {
            public ApplyButton()
            {
                Text = "&Apply";
                IO_int = 3;
            }
        }

        #endregion Ok-Cancel-Help Buttons

        #region Edits-CheckBoxes-RadioButtons
        /// <summary>
        ///     Static text label
        /// </summary>
        public sealed class Label : ControlItem
        {
            public Label() : base(XlControl.XlStaticText)
            {
            }

            public Label(string text) : this()
            {
                Text = text;
            }
        }

        /// <summary>
        ///     Text box
        /// </summary>
        public class TextEdit : ControlItem
        {
            public TextEdit() : base(XlControl.XlTextBox)
            {
            }

            public TextEdit(string text) : this()
            {
                IO_string = text;
            }

            /// <summary>
            ///    Text box editing content 
            ///    The override forces string entry into the IO object
            /// </summary>
            public string IO_string
            {
                get { return Convert.ToString(base.IO); }
                set { base.IO = Convert.ToString(value); }
            }
        }


        /// <summary>
        ///     Integer edit box
        /// </summary>
        /// <remarks>
        /// this control does internal datavalidation, before allowing "OK" to exit the dialog
        /// </remarks>
        public class IntegerEdit : ControlItem
        {

            public IntegerEdit() : base(XlControl.XlIntegerEedit)
            {
            }

            public IntegerEdit(int input) : this()
            {
                IO_int = input;
            }

            public int IO_int 
            {
                get { return Convert.ToInt32(base.IO); }
                set { base.IO = Convert.ToInt32(value); }
            }
        }

        /// <summary>
        ///     Double edit box
        /// </summary>
        /// <remarks>
        /// this control does internal datavalidation, before allowing "OK" to exit the dialog
        /// </remarks>
        public class DoubleEdit : ControlItem
        {
            public DoubleEdit() : base(XlControl.XlNumberEdit)
            {
            }

            public DoubleEdit(double input) : this()
            {
                IO_double = input;
            }

            public double IO_double
            {
                get { return Convert.ToDouble(base.IO); }
                set { base.IO = Convert.ToDouble(value); }
            }
        }

        /// <summary>
        ///     Formula editor control
        /// </summary>
        /// <remarks>
        /// this control does internal datavalidation, before allowing "OK" to exit the dialog
        /// </remarks>
        public class FormulaEdit : ControlItem
        {
            public FormulaEdit() : base(XlControl.XlFormulaEdit)
            {
            }

            public FormulaEdit(string text) : this()
            {
                IO_string = text;
            }

            /// <summary>
            ///    Formula content
            ///    The override forces string entry into the IO object
            /// </summary>
            public string IO_string
            {
                get { return Convert.ToString(base.IO); }
                set { base.IO = Convert.ToString(value); }
            }
        }

        /// <summary>
        ///     Cell reference edit control
        /// </summary>
        /// <remarks>
        /// this control does internal datavalidation, before allowing "OK" to exit the dialog
        /// </remarks>
        public class RefEdit : ControlItem 
        {
            public RefEdit() : base(XlControl.XlReferenceEdit)
            {
            }
             
            public RefEdit(string text) : this()
            {
                IO_string = text;
            }

            /// <summary>
            ///     Reference Address (R1C1)
            ///    The override forces string entry into the IO object
            /// </summary>
            public string IO_string
            {
                get { return Convert.ToString(base.IO); }
                set { base.IO = Convert.ToString(value); }
            }

        }

        /// <summary>
        ///     Radio button group
        /// </summary>
        public class RadioButtonGroup : ControlItem
        {
            public RadioButtonGroup() : base(XlControl.XlRadioButtonGroup)
            {
            }

            public RadioButtonGroup(string text) : this()
            {
                this.Text = text;
            }

            /// <summary>
            ///     Select the index of the list starting from 0
            ///     Not selected as -1
            /// </summary>
            /// <remarks>
            ///     The built-in index starts at 1, externally it is exposed a 0-starting index as per .NET General Convention 
            /// </remarks>
            [DefaultValue(-1)]
            public int IO_index
            {
                get
                {
                    if (IO.IsNull())
                        return -1;
                    else
                        return Convert.ToInt32(base.IO) - 1;
                }

                set
                {
                    if (value < 0)
                        base.IO = null;
                    else
                        base.IO = Convert.ToInt32(value + 1);
                }
            }
        }

        /// <summary>
        ///     Radio button
        /// </summary>
        public class RadioButton : ControlItem
        {
            public RadioButton() : base(XlControl.XlRadioButton)
            {
            }

            public RadioButton(string text) : this()
            {
                this.Text = text;
            }

            /// <summary>
            ///    Mainly used to *get* the status of the radio control being interrogated
            ///    To *set* the active radio control, please use the SelectedIndex of the preceeding RadioGroupButton
            /// </summary>
            [DefaultValue(false)]
            public bool IO_selected
            {
                get
                {
                    if (IO.IsNull())
                        return false;
                    else
                        return Convert.ToBoolean(IO);
                }

                set 
                {
                    IO = Convert.ToBoolean(value);
                }
            }
        }

        /// <summary>
        ///     Check box
        /// </summary>
        public sealed class CheckBox : ControlItem
        {
            public CheckBox() : base(XlControl.XlCheckBox)
            {
            }

            public CheckBox(string text) : base(XlControl.XlCheckBox)
            {
                this.Text = text;
            }

            /// <summary>
            ///    (Re-)sets the checked status of the control
            /// </summary>
            public bool IO_checked
            {
                get
                {
                    if (IO.IsNull())
                        return false;
                    else
                        return Convert.ToBoolean(IO);
                }

                set
                {
                    IO = Convert.ToBoolean(value);
                }
            }
        }

        /// <summary>
        ///     Group Box
        /// </summary>
        public sealed class GroupBox : ControlItem
        {
            public GroupBox() : base(XlControl.XlGroupBox)
            {
                X = Y = 10;
            }

            public GroupBox(string text) : this()
            {
                Text = text;
            }
        }
        #endregion Edits-CheckBoxes-RadioButtons

        #region ListBox Controls
        public abstract class AbstractListControl : ControlItem
        {

            protected AbstractListControl(XlControl itemNumber) : base(itemNumber)
            {
                this.Items = new StringCollection();
            }

            /// <summary>
            ///     Select the index of the list starting from 0
            ///     Not selected as -1
            /// </summary>
            /// <remarks>
            ///     The built-in index starts at 1, externally behaves as a 0-starting index of the composite .NET General Convention 
            /// </remarks>
            [DefaultValue(-1)]
            public int IO_index
            {
                get
                {
                    if (IO == null)
                        return -1;
                    else
                        return Convert.ToInt32(base.IO) - 1;
                }

                set
                {
                    if (value < 0)
                        IO = null;
                    else
                        base.IO = Convert.ToInt32(value + 1);
                }
            }

            /// <summary>
            ///     Gets or sets an object that represents a collection of items contained in this <see cref="T: ComboBox" />
            /// </summary>
            public StringCollection Items { get; }

            /// <summary>
            ///     List name
            /// </summary>
            private string ListName
            {
                get { return base.Text; }
                set { base.Text = value; }
            }

            /// <summary>
            ///     Called before building the dialog box
            /// </summary>
            protected internal override void OnBeforeBuild()
            {
                string[] listArray;
                if (Items != null && Items.Any())
                {
                    listArray = Items.ToArray();
                }
                else
                {
                    // There must be a list or an error will occur
                    listArray = new[] { string.Empty };
                }

                if (string.IsNullOrEmpty(ListName))
                {
                    ListName = $"Gen_{GetType().Name}_{Index}";
                }

                if (IO_index > Items.Count)
                {
                    IO_index = -1;
                }
                
                // by calling XlCall.xlfSetName with the ListName WHILE passing the listArray, we make the ListName Known in Excel.
                XlCall.Excel(XlCall.xlfSetName, ListName, listArray);
            }

            /// <summary>
            ///     Perform application-defined tasks related to releasing or resetting unmanaged resources.
            /// </summary>
            public override void Dispose()
            {
                // by calling XlCall.xlfSetName with the ListName WITHOUT passing the listArray, we relinquish the ListName in Excel.
                if ((bool)XlCall.Excel(XlCall.xlfSetName, ListName))
                {
                    // Bart: the next line causes Excel to crash, we can't ask Excel to free memory it hasn't allocated !
                    // XlCall.Excel(XlCall.xlFree, ListName);
                };
                base.Dispose();
            }

            public class StringCollection : Collection<String>
            {

                public void AddRange(IEnumerable<string> items)
                {
                    foreach (var item in items)
                    {
                        this.Add(item);
                    }
                }
            }
        }

        /// <summary>
        ///     List Box
        /// </summary>
        public class ListBox : AbstractListControl
        {
            public ListBox() : base(XlControl.XlListBox) {}
        }

        /// <summary>
        ///     Combo box control
        /// </summary>
        /// <remarks>
        ///     A Text edit control is required directly before the combo box
        /// </remarks>
        public class ComboBox : AbstractListControl 
        {
            private readonly TextEdit _innerTextBox = new TextEdit();

            /// <summary>
            /// </summary>
            public ComboBox() : base(XlControl.XlCombobox) {}

            /// <summary>
            ///     X Coordinate, if the value is less than 0, the default value is used
            /// </summary>
            public override int X 
            {
                get { return base.X; }
                set { base.X = value;  _innerTextBox.X = value; }
            }

            /// <summary>
            ///     Y Coordinate, if the value is less than 0, the default value is used
            /// </summary>
            public override int Y 
            {
                get { return base.Y; }
                set 
                {
                    base.Y = value;
                    _innerTextBox.Y = value;
                }
            }

            /// <summary>
            ///     Width, if the value is less than 0, the default value is used
            /// </summary>
            public override int W 
            {
                get { return base.W; }
                set 
                {
                    base.W = value;
                    _innerTextBox.W = value;
                }
            }

            /// <summary>
            ///     Height, if the value is less than 0, the default value is used
            /// </summary>
            public override int H
            {
                get { return base.H; }
                set 
                {
                    base.H = value;
                    _innerTextBox.H = value;
                }
            }

            /// <summary>
            ///     Text content 
            /// </summary>
            public override string Text {
                get { return (string) _innerTextBox.IO; }
                set {
                    _innerTextBox.IO= value;
                    if (this.Items != null) {
                        IO_index = Items.IndexOf(value);
                    }
                }
            }

            /// <summary>
            ///     Active control definition array
            /// </summary>
            /// <returns></returns>
            public override IEnumerable<object[]> GetControlParameters() 
            {
                return new[] { _innerTextBox.GetControlParameters().FirstOrDefault(), base.ControlParameters };
            }
        }

        /// <summary>
        ///     Drop-down list controls
        /// </summary>
        public class DropdownList : AbstractListControl {
            public DropdownList() : base(XlControl.XlDropdownList) {
            }

            /// <summary>
            ///     Selected values
            /// </summary>
            public string ValueAtIndex {
                get {
                    int index = IO_index;
                    if (index < 0) {
                        return string.Empty;
                    }
                    return Items.ElementAt(index);
                }
            }
        }

        #endregion ListBox Controls

        #region Dialog Control Collection
        /// <summary>
        ///     Collection of dialog controls
        /// </summary>
        /// <remarks>
        /// I'm not sure if a collection is the best way to encapsulate the ControlParameters array.
        /// A list seems a more logical choice; need to review later...
        /// </remarks>
        public class XlDialogControlCollection : Collection<ControlItem>, IDisposable 
        {
            internal XlDialogControlCollection()
            {
            }

            /// <summary>
            ///     Performs application-defined tasks related to releasing or resetting unmanaged resources.
            /// </summary>
            public void Dispose() {
                foreach (var item in Items) {
                    item.Dispose();
                }
            }

            /// <summary>
            ///     Insert an element at the specified index <see cref="T:System.Collections.ObjectModel.Collection`1" /> 
            /// </summary>
            /// <param name="index">A zero-based index that should be inserted at that location <paramref name="item" />。</param>
            /// <param name="item">The object to insert. For reference types, the value can be null.</param>
            /// <exception cref="T:System.ArgumentOutOfRangeException">
            ///     <paramref name="index" /> Less than zero. Or Greater than -<paramref name="index" />
            ///     <see cref="P:System.Collections.ObjectModel.Collection`1.Count" />。
            /// </exception>
            protected override void InsertItem(int index, ControlItem item) 
            {
                base.InsertItem(index, item);
                item.Index = index;
            }

            /// <summary>
            ///     Remove the element at the specified index. <see cref="T:System.Collections.ObjectModel.Collection`1" /> 
            /// </summary>
            /// <param name="index">The zero-based index of the element to be removed.</param>
            /// <exception cref="T:System.ArgumentOutOfRangeException">
            ///     <paramref name="index" /> Less than zero. Or equal to or greater than <paramref name="index" />
            ///     <see cref="P:System.Collections.ObjectModel.Collection`1.Count" />。
            /// </exception>
            protected override void RemoveItem(int index)
            {
                base.RemoveItem(index);
                UpdateItemsIndex(index);
            }

            /// <summary>
            ///     Update element index; needed when removing an item from the collection
            /// </summary>
            private void UpdateItemsIndex(int startIndex)
            {
                for (var i = startIndex; i < Count; i++)
                {
                    Items[i].Index = i;
                }
            }

            /// <summary>
            ///     Building an array of control definitions
            /// </summary>
            /// <returns>object[,] array </returns>
            internal object[,] Build()
            {
                int rows = Items.Count();
                object[,] result = new object[rows, 7];
                int rowIndex = 0;
                foreach (var item in Items)
                {
                    item.OnBeforeBuild();
                    var defArray = item.GetControlParameters();
                    foreach (var array in defArray)
                    {
                        for (int i = 0; i < 7; i++)
                        {
                            result[rowIndex, i] = array[i];
                        }
                        rowIndex++;
                    }
                }
                return result;

                /*                ControlItem [] visibleControls = Items.Where(i => i.Visible).ToArray();
                                int rows = visibleControls.Sum(i => i.GetControlParameters().Count());
                                object [,] result = new object[rows, 7];
                                int rowIndex = 0;
                                foreach (var item in visibleControls)
                                {
                                    item.OnBeforeBuild();
                                    var defArray = item.GetControlParameters();
                                    foreach (var array in defArray)
                                    {
                                        for (int i = 0; i < 7; i++) 
                                        {
                                            result[rowIndex, i] = array[i];
                                        }
                                        rowIndex++;
                                    }
                                }
                                return result;
                */
            }

            /// <summary>
            ///     Parse the return value and write it to the control collection
            /// </summary>
            internal void UpdateResult(object[,] result)
            {
                try
                {
                    int index = 0;
                    foreach (var item in Items)
                    {
                        var controlDefs = item.GetControlParameters();
                        foreach (var defItem in controlDefs)
                        {
                            defItem[(int)XlColumn.XlIOColumn] = result[index, (int)XlColumn.XlIOColumn];
                            index++;
                        }
                    }

                    /*                    int index = 0;
                                        var visibleControls = Items.Where(i => i.Visible).ToArray();
                                        foreach (var item in visibleControls)
                                        {
                                            var controlDefs = item.GetControlParameters();
                                            foreach (var defItem in controlDefs)
                                            {
                                                defItem[(int)XlColumn.XlIOColumn] = result[index, (int)XlColumn.XlIOColumn];
                                                index++;
                                            }
                                        }
                    */
                }
                catch
                {
                    ;
                }
            }

        }
        #endregion Dialog Control Collection
    }
}