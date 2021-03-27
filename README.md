# ExcelDna-XlDialogBox
### C# wrapper for XLM macro DIALOG.BOX

*Note; this code builds further on the ExcelDna-XlDialog class, available [here](https://github.com/zwq00000/ExcelDna-XlDialog) on GitHub.*

The DIALOG.BOX macro is part of the XLM macro's that predate *Visual Basic for Applications* (VBA) and were introduced in Excel 4.0.  Documentation of these macro functions is hard to find, but a comprehensive [function reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf) can be found on cloudfront.net.

You may wonder, why spend time and effort on working on such ancient technology, but the point is that there are no good RefEdit controls available for Winforms or WPF-forms.  This is discussed in-depth [here](https://www.breezetree.com/blog/excel-refedit-in-c-sharp/).

The XLM functions and commands are still available in Excel-365 and are exposed to developers  through the `2013 Office System Developer Resources` that provide support in writing  Excel 2013 XLL's'. The following function call creates a dialog box: 

```c
Excel4(xlfDialogBox, &returValue, 1, &DialogDefinition)
```

To use this  functionality, requires working with LPXLOPER12 structures that are complex to deal with, because of the  various overloads that exist, and the interaction with Excel's internal memory. A dialog box call, running in a verification loop could look like:

```C#
xloper ret_val;
int xl12;
cpp_xloper DialogDef((WORD) NUM_DIALOG_ROWS, (WORD) NUM_DIALOG_COLUMNS, UsernameDlg);
// now set up the N x 7 dialog definition array
do
{
    xl12 = Excel4(xlfDialogBox, &ret_val, 1, &DialogDef);
    if (xl4 || (ret_val.xltype == xltypeBool && ret_val.val._bool == 0))
        break;
    // Process the input from the dialog by reading
    // the 7th column of the returned array.
    // ... code omitted
    Excel2(xlFree, 0, 1, &ret_val);
    ret_val.xltype = xltypeNil;
}
while (1);
Excel2(xlFree, 0, 1, &ret_val);
return 1;
```

Developing Excel extensions is a daunting task, which is easy to understand, when you start reading through the [online documentation](https://docs.microsoft.com/en-us/office/client-developer/excel/developing-excel-xlls). Here is where `Excel-DNA` comes to the rescue. Thanks to [Excel-DNA,](https://excel-dna.net/) writing Excel extensions  has become much easier, the code base is easier to maintain and context sensitive help (_compiled into an *.chm file_) is now straightforward to implement. A call to set up a modal dialog box now becomes:

```C#
var result = XlCall.Excel(XlCall.xlfDialogBox, dialogDef);
```

Where :

* `XlCall.xlfDialogBox` is an enumerated value, telling Excel to create a dialog box. 

* `dialogDef` is a N x 7 two-dimensional array that defines the contents of the dialog box。

The following figure closely resembles the dialog example from `GENERIC.C` from the  `2013 Office System Developer Resources` .

In `GENERIC.C` the following parameter table is defined:

```c
static LPWSTR g_rgDialog[g_rgDialogRows][g_rgDialogCols] =
{
	{L"\000",   L"\000",    L"\000",    L"\003494", L"\003210", L"\025Generic Sample Dialog", L"\000"},
	{L"\0011",  L"\003330", L"\003174", L"\00288",  L"\000",    L"\002OK",                    L"\000"},
	{L"\0012",  L"\003225", L"\003174", L"\00288",  L"\000",    L"\006Cancel",                L"\000"},
	{L"\0015",  L"\00219",  L"\00211",  L"\000",    L"\000",    L"\006&Name:",                L"\000"},
	{L"\0016",  L"\00219",  L"\00229",  L"\003251", L"\000",    L"\000",                      L"\000"},
	{L"\00214", L"\003305", L"\00215",  L"\003154", L"\00273",  L"\010&College",              L"\000"},
	{L"\00211", L"\000",    L"\000",    L"\000",    L"\000",    L"\000",                      L"\0012"},
	{L"\00212", L"\000",    L"\000",    L"\000",    L"\000",    L"\010&Harvard",              L"\000"},
	{L"\00212", L"\000",    L"\000",    L"\000",    L"\000",    L"\006&Other",                L"\0011"},
	{L"\0015",  L"\00219",  L"\00250",  L"\000",    L"\000",    L"\013&Reference:",           L"\000"},
	{L"\00210", L"\00219",  L"\00267",  L"\003253", L"\000",    L"\000",                      L"\000"},
	{L"\00214", L"\003209", L"\00293",  L"\003250", L"\00263",  L"\017&Qualifications",       L"\000"},
	{L"\00213", L"\000",    L"\000",    L"\000",    L"\000",    L"\010&BA / BS",              L"\0011"},
	{L"\00213", L"\000",    L"\000",    L"\000",    L"\000",    L"\010&MA / MS",              L"\0011"},
	{L"\00213", L"\000",    L"\000",    L"\000",    L"\000",    L"\021&PhD / Other Grad",     L"\0010"},
	{L"\00215", L"\00219",  L"\00299",  L"\003160", L"\00296",  L"\015GENERIC_List1",         L"\0011"},
};
```



The same information is used for the sample dialog shown below, be it that a help button has been added. 

![image](./images/Dialog1.png) 

Figure 1. Example dialog



The C# code to create the above dialog box is shown below:

```c#
var dialog = new XlDialogBox() 						{ Width = 494, Height = 210, Text = "Generic Sample Dialog" };

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
// Note: always put the  'labels' in front of their (edit/list) controls.
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

```

#### Some guidance 

The dialog box definition table must be at least two rows high, and shall be seven columns wide.  The definition of the dialog itself is in the first row of the table. This row also specifies the default selected item and may contain the reference for the Help button in the Item number column.  

The definitions of each column in a dialog definition table are listed in the following table.

| Column type            | Column (1-referenced) | Index (0-referenced) |
| ---------------------- | :-------------------: | :------------------: |
| Item number            |           1           |          0           |
| Horizontal position    |           2           |          1           |
| Vertical position      |           3           |          2           |
| Item width             |           4           |          3           |
| Item height            |           5           |          4           |
| Item text              |           6           |          5           |
| Initial value / result |           7           |          6           |

#### Item number

The first column in each row of the dialog definition table contains the item number. It is an enumeration with one out of 24 values shown in the following table that defines the type of dialog control being displayed. See the table below for different dialog control types.

#### Horizontal, vertical position, width and height

These are integer numbers. They can be left undefined when items are part of a group box. To position an item at least horizontal and vertical position need to be defined. In that case the control gets a default size

#### Item text

Describes the static text shown by an item. 

#### Initial value / result

The last column in each row is used for data exchange. Whereas the data types of column 1 - 6 are given (string or int), column 7 can contain a number or a string.  To work with this column an IO object has been defined in the  `ControlItem` class.



***

#### XlControl enumeration table

| Dialog-box item                                              | Item number |
| :----------------------------------------------------------- | ----------: |
| Default OK button                                            |           1 |
| Cancel button                                                |           2 |
| OK button                                                    |           3 |
| Default Cancel button                                        |           4 |
| Static text                                                  |           5 |
| Text edit box                                                |           6 |
| Integer edit box                                             |           7 |
| Number edit box                                              |           8 |
| Formula edit box                                             |           9 |
| Reference edit box                                           |          10 |
| Option button group                                          |          11 |
| Option button                                                |          12 |
| Check box                                                    |          13 |
| Group box                                                    |          14 |
| List box                                                     |          15 |
| Linked list box                                              |          16 |
| Icons                                                        |          17 |
| Linked file list box     (Microsoft Excel for Windows only)  |          18 |
| Linked drive and directory box     (Microsoft Excel for Windows only) |          19 |
| Directory text box                                           |          20 |
| Drop-down list box                                           |          21 |
| Drop-down combination edit/list box                          |          22 |
| Picture button                                               |          23 |
| Help button                                                  |          24 |

#### Remarks

A number of controls (*Integer edit box, Number edit box, Formula edit box and Reference edit box*) do internal data validation and may therefore prevent the OK button from exiting the dialog.

#### Help (!!), Next, Back and Apply buttons

I have **not** been able to make the `Help` button work. As a workaround a `help2` button has been derived from an OK button with `Initial value = -1`.
Likewise, `Next`, `Back` and `Apply` buttons have been defined to help creating Wizard functionality.

#### Triggers

Adding 100 to certain item numbers causes the function to return control to the XLL when the item is clicked on with the dialog still displayed. 
This "trigger feature" enables the `xlfDialogBox` command to alter the dialog, validate input and so on, and return for more user interaction. 
The position of the item number chosen in this way is returned in the 1st row, 7th column of the returned array. 

This feature does not work with static text (item 5) edit boxes (6, 7, 8, 9 and 10), group boxes (14), pictures (23) or the help button (24). 
Those controls just ignore the trigger request if 100 would be added to their item numbers.

In the code for the dialog controls this is accomplished by setting `Trigger = true / false`. 

#### Disabling

Adding 200 to any item number, disables (greys-out) the item. A disabled item cannot be chosen or selected. For example, 203 is a disabled OK button. 
You could for instance use item 223 to include a picture in your dialog box that does not behave like a button.

In the code for the dialog controls this is accomplished by setting `Enable = true / false`. 

#### (In-) visible

The `Visible` property  is **not part of** the built-in properties of the parameter table, but defined in the ControlItem class. When this property is false, the ControlItem is not added to the parameter table (***i. e. not shown at all***) in the dialog. Therefore the number of controls passed to the `ShowDialog(...)` function, need not be the same as the number of controls added to the dialog during initialization, using the `dialog.Controls.Add(...)` function. You need to be aware of this, when doing data validation.

In the code for the dialog controls this is accomplished by setting `Visible = true / false`. 

Most of the dialog items are simple and no further explanation is required. For some a little more explanation is helpful.

#### Text and edit boxes

Vertical alignment of a text label to the text that appears in an edit box is important aesthetically. For edit boxes with the default height (set by leaving the height field blank) This is achieved by setting the vertical position of the text to be that of the edit box + 3.

#### Buttons

Selecting a cancel button (item 2 or 4) causes the dialog to terminate returning FALSE.  Pressing any other button causes the function to return the offset of that button in the definition table in the 7th column, 1st row of the returned array that describes the dialog itself.

#### Radio buttons

A group of radio buttons (12) must be preceded immediately by a radio group item (11) and must be uninterrupted by other item types. 
If the radio group item has no text label, the group is not contained within a border. 
If the height and/or width of the radio group are omitted but text is provided, a border is drawn that surrounds the radio buttons and their labels.

#### List-boxes

The text supplied in a list box item row should either be a name (DLL-internal or on a worksheet) that resolves to a literal array or range of cells. 
It can also be a string that looks like a literal array, e. g. "{1, 2, 3, 4, 5, \"A\", \"B\", \"C\"}" (where coded in a C source file). 
List-boxes return the position (counting from 1) of the selected item in the list in the 7th column of the list-box item line. 
Drop-down list-boxes (21) behave exactly as list boxes (15) except that the list is only displayed when the item is selected.

#### Linked list-boxes

Linked list-boxes (16), linked file-boxes (18) and drop-down combo-boxes (22) should be preceded immediately by an edit box that can support the data types in the list.  The lists themselves are drawn from the text field of the definition row which should be a range name or a string that represents a static array.  A linked path box (19) must be preceded immediately by a linked file-box (18).
Drop down combo-boxes return the value selected in the 7th column of the associated edit box and the position (counting from 1) of the selected item
in the list in the 7th column of the combo-box item line.

