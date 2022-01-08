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
    public class Examples
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
        static bool Validate(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            // just some code to set a break point
            // int i = index;

            return true; // return to dialog
        }

        [ExcelCommand(
        Name = "Generic_Example",
        Description = "Starts the example dialog 'Generic.c' from the Excel2013 XLL SDK",
        HelpTopic = "XlmDialogExample-AddIn.chm!1001",
        ShortCut = "^R")]
        public static void Cmd_Generic()
        {

/*          This is what you get using the tables from DialogBox.xlsb
 * 
 *          var dialog  = new XlDialogBox()                  {                    W = 494, H = 210, Text = "Generic Sample Dialog",  };
            var ctrl_01 = new XlDialogBox.Label()            {  X = 020, Y = 010,                   Text = "&Name:",  };
            var ctrl_02 = new XlDialogBox.TextEdit()         {  X = 020, Y = 026, W = 250,          };
            var ctrl_03 = new XlDialogBox.Label()            {  X = 020, Y = 050,                   Text = "&Reference:",  };
            var ctrl_04 = new XlDialogBox.RefEdit()          {  X = 020, Y = 066, W = 250,          };
            var ctrl_05 = new XlDialogBox.ListBox()          {  X = 020, Y = 099, W = 160, H = 096, Text = "List_05", IO = 2, };
            var ctrl_06 = new XlDialogBox.GroupBox()         {  X = 305, Y = 015, W = 154, H = 073, Text = "&College",  };
            var ctrl_07 = new XlDialogBox.RadioButtonGroup() {                                      IO = 1, };
            var ctrl_08 = new XlDialogBox.RadioButton()      {                                      Text = "&Harvard",  };
            var ctrl_09 = new XlDialogBox.RadioButton()      {                                      Text = "&Other",  IO = 1, };
            var ctrl_10 = new XlDialogBox.GroupBox()         {  X = 209, Y = 093, W = 250, H = 073, Text = "&Qualifications",  };
            var ctrl_11 = new XlDialogBox.CheckBox()         {                                      Text = "&BA / BS",  IO = false, Trigger = true, };
            var ctrl_12 = new XlDialogBox.CheckBox()         {                                      Text = "&MA / MS",  IO = false, };
            var ctrl_13 = new XlDialogBox.CheckBox()         {                                      Text = "&PhD / Other Grad",  IO = true, Enable = false, };
            var ctrl_14 = new XlDialogBox.OkButton()         {  X = 209, Y = 174, W = 075,          Text = "&OK", Default = true, };
            var ctrl_15 = new XlDialogBox.CancelButton()     {  X = 296, Y = 174, W = 075,          Text = "&Cancel",  };
            var ctrl_16 = new XlDialogBox.HelpButton2()      {  X = 383, Y = 174, W = 075,          Text = "&Help",  };
            ctrl_05.Items.AddRange(new string[]              { "Bake", "Broil", "Sizzle", "Fry", "Saute", "Deep fry",  });

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);
            dialog.Controls.Add(ctrl_11);
            dialog.Controls.Add(ctrl_12);
            dialog.Controls.Add(ctrl_13);
            dialog.Controls.Add(ctrl_14);
            dialog.Controls.Add(ctrl_15);
            dialog.Controls.Add(ctrl_16);
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;
*/

            // the following approach uses control-names more inline with Generic.c
            // By setting IO = 4 for the dialog; the RefEdit is activated first

            var dialog = new XlDialogBox()                      { W = 494, H = 210, Text = "Generic Sample Dialog" , IO = 4};

            var okBtn = new XlDialogBox.OkButton()              { X = 209, Y = 174, W = 075, H = 023 };
            var cancelBtn = new XlDialogBox.CancelButton()      { X = 296, Y = 174, W = 075, H = 023 };
            var helpBtn = new XlDialogBox.HelpButton2()         { X = 384, Y = 174, W = 075, H = 023 };
//          var helpBtn = new XlDialogBox.HelpButton()          { X = 384, Y = 174, W = 075, H = 023, IO_string = "D:\\Source\\VS19\\XlmDialogExample\\XlmDialogExample\\bin\\Debug\\XlmDialogExample-AddIn.chm!1001" };

            var nameLabel = new XlDialogBox.Label               { X = 019, Y = 011, Text = "&Name:" };
            var nameEdit = new XlDialogBox.TextEdit             { X = 019, Y = 029, IO_string = "<Name>" };

            var refLabel = new XlDialogBox.Label                { X = 019, Y = 050, Text = "&Reference" };
            var refEdit = new XlDialogBox.RefEdit               { X = 019, Y = 067, W = 253 };

            var listEdit = new XlDialogBox.ListBox()            { X = 019, Y = 099, W = 160, H = 96, IO_index = 2, Text = "GENERIC_List1" };
            listEdit.Items.AddRange(new string[]                { "Bake", "Broil", "Sizzle", "Fry", "Saute" });

            var CollegeBox = new XlDialogBox.GroupBox           { X = 305, Y = 015, W = 154, H = 073, Text = "College" };
            var RadioGroup = new XlDialogBox.RadioButtonGroup   { IO_index = 1, Enable = false };
            var RadioHarvr = new XlDialogBox.RadioButton        { Text = "&Harvard", Enable = false };
            var RadioOther = new XlDialogBox.RadioButton        { Text = "&Other", Enable = false };

            var qualiGroup = new XlDialogBox.GroupBox           { X = 209, Y = 093, W = 250, H = 063, Text = "&Qualifications" };
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

            dialog.Controls.Add(CollegeBox);
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

            // define the method that is calling the dialog box so we can select the correct HelpTopic from ExcelCommand attribute 
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod();

            // only change scaling (default = 100) if the dialog has been designed on a display with a higher DPI setting than the standard 96 DPI.
            // dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

            // now it is time to play around with the parameters chosen in the dialog box to get things done
            var xlApp = (Application)ExcelDnaUtil.Application;
            var ws = xlApp.Sheets[1] as Worksheet;
            var range = ws.Cells[1, 1] as Range;
            range.Value2 = nameEdit.IO_string;
        } // Cmd_Generic

#pragma warning disable CS0649 // Dialogs are only initialized, the results aren't being used in these examples
#pragma warning disable IDE0044 // Dialogs are only initialized, the results aren't being used in these examples
        static string DI_Range;
        static string Vp_Range;
        static string Un_Range;
        static double MaxAngle = 60;
        static int    NrOfRays = 51;
        static double VertSamp = 25;
        static bool   Rays_upw = true;
        static bool   MakePlot = true;
#pragma warning restore CS0649 // Dialogs are only initialized, the results aren't being used in these examples
#pragma warning restore IDE0044 // Dialogs are only initialized, the results aren't being used in these examples

        [ExcelCommand(
            Name = "Ray_Tracer",
            Description = "Creates a 1D-ray tracer on a new sheet, using a wizard-type dialog, that is available from the GeoLib ribbon.",
        HelpTopic = "XlmDialogExample-AddIn.chm!1002",
            ShortCut = "^R")]
        public static void Cmd_RayTracer()
        {
            var dialog = new XlDialogBox() { W = 515, H = 330, Text = "Ray Tracer Wizard", IO = 3, };
            var ctrl_01 = new XlDialogBox.GroupBox() { X = 015, Y = 010, W = 480, H = 160, Text = "Over&burden parameters  ➔   Vp && depth ranges need same nr of rows", };
            var ctrl_02 = new XlDialogBox.Label() { X = 030, Y = 031, Text = "&Depth Information in 1 column", };
            var ctrl_03 = new XlDialogBox.RefEdit() { X = 030, Y = 046, W = 200, };
            var ctrl_04 = new XlDialogBox.Label() { X = 240, Y = 042, W = 235, H = 040, Text = "Exclude Z = 0 value. Use same nr. of data points as used for Vp (see below).", };
            var ctrl_05 = new XlDialogBox.Label() { X = 030, Y = 073, Text = "V&p Interval velocities in 1 (or 2) columns", };
            var ctrl_06 = new XlDialogBox.RefEdit() { X = 030, Y = 088, W = 200, };
            var ctrl_07 = new XlDialogBox.Label() { X = 240, Y = 084, W = 235, H = 040, Text = "Vp = Vp0 + K * z. With Vp0 in 1st column, optionally provide K values in 2nd column.", };
            var ctrl_08 = new XlDialogBox.Label() { X = 030, Y = 115, Text = "Cell containing Vp in &underburden", };
            var ctrl_09 = new XlDialogBox.RefEdit() { X = 030, Y = 130, W = 200, };
            var ctrl_10 = new XlDialogBox.Label() { X = 240, Y = 126, W = 235, H = 040, Text = "Leave empty for unconstraint Angle of Incidence at target depth.", };
            var ctrl_11 = new XlDialogBox.GroupBox() { X = 015, Y = 190, W = 180, H = 120, Text = "A&ngles at target", };
            var ctrl_12 = new XlDialogBox.Label() { X = 030, Y = 211, Text = "&Maximum AoI  [deg]", };
            var ctrl_13 = new XlDialogBox.DoubleEdit() { X = 030, Y = 226, W = 140, IO = 60, };
            var ctrl_14 = new XlDialogBox.Label() { X = 030, Y = 253, Text = "Nr of rays within +/- AoI", };
            var ctrl_15 = new XlDialogBox.IntegerEdit() { X = 030, Y = 268, W = 140, IO = 25, };
            var ctrl_16 = new XlDialogBox.GroupBox() { X = 215, Y = 190, W = 180, H = 120, Text = "Depth && M&isc parameters", };
            var ctrl_17 = new XlDialogBox.Label() { X = 230, Y = 211, Text = "&Vertical sampling [m]", };
            var ctrl_18 = new XlDialogBox.DoubleEdit() { X = 230, Y = 226, W = 140, IO = 25, };
            var ctrl_19 = new XlDialogBox.CheckBox() { X = 230, Y = 261, Text = "Shoot rays up&wards", IO = true, };
            var ctrl_20 = new XlDialogBox.CheckBox() { X = 230, Y = 281, W = 140, Text = "Create Rays Char&t", IO = true, };
            var ctrl_21 = new XlDialogBox.OkButton() { X = 415, Y = 224, W = 080, Text = "&OK", Default = true, };
            var ctrl_22 = new XlDialogBox.CancelButton() { X = 415, Y = 254, W = 080, Text = "&Cancel", };
            var ctrl_23 = new XlDialogBox.HelpButton2() { X = 415, Y = 289, W = 080, Text = "&Help", };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);
            dialog.Controls.Add(ctrl_11);
            dialog.Controls.Add(ctrl_12);
            dialog.Controls.Add(ctrl_13);
            dialog.Controls.Add(ctrl_14);
            dialog.Controls.Add(ctrl_15);
            dialog.Controls.Add(ctrl_16);
            dialog.Controls.Add(ctrl_17);
            dialog.Controls.Add(ctrl_18);
            dialog.Controls.Add(ctrl_19);
            dialog.Controls.Add(ctrl_20);
            dialog.Controls.Add(ctrl_21);
            dialog.Controls.Add(ctrl_22);
            dialog.Controls.Add(ctrl_23);

            ctrl_03.IO_string = DI_Range;
            ctrl_06.IO_string = Vp_Range;
            ctrl_09.IO_string = Un_Range;
            ctrl_13.IO_double = MaxAngle;
            ctrl_15.IO_int = NrOfRays;
            ctrl_18.IO_double = VertSamp;
            ctrl_19.IO_checked = Rays_upw;
            ctrl_20.IO_checked = MakePlot;

            // GetCurrentMethod reflection is used to use the ExcelCommand attribute
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod();
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI instead of 96 dpi
            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;
        } // Cmd_RayTracer

        // used in the depth chart dialog
#pragma warning disable CS0649 // Dialogs are only initialized, the results aren't being used in these examples
#pragma warning disable IDE0044 // Dialogs are only initialized, the results aren't being used in these examples
        static string ChartTitle = "Common-depth chart";
        static string TitleRange;
        static string DepthRange;
        static string Data_Range;
        static int    RangesVert = 0;
        static int    Chart_type = 2;
        static bool   PlotColors = true;
#pragma warning restore CS0649 // Dialogs are only initialized, the results aren't being used in these examples
#pragma warning restore IDE0044 // Dialogs are only initialized, the results aren't being used in these examples

        [ExcelCommand(
            Name = "Chart_Dialog",
            Description = "Dialog to create a scatter-chart, where the y-axis is common between all series instead of the x-axis.",
            HelpTopic = "XlmDialogExample-AddIn.chm!1003",
            ShortCut = "^T")]
        public static void Cmd_ChartDialog()
        {

            var dialog = new XlDialogBox() { W = 410, H = 220, Text = "Create common-depth Scatter Chart", IO = 2, };
            var ctrl_01 = new XlDialogBox.Label() { X = 020, Y = 010, Text = "Chart &Title", };
            var ctrl_02 = new XlDialogBox.TextEdit() { X = 020, Y = 026, W = 250, };
            var ctrl_03 = new XlDialogBox.Label() { X = 020, Y = 050, Text = "Range of Series &Names  (leave blank for default names)", };
            var ctrl_04 = new XlDialogBox.RefEdit() { X = 020, Y = 066, W = 250, };
            var ctrl_05 = new XlDialogBox.Label() { X = 020, Y = 090, Text = "Range of &Depth Points", };
            var ctrl_06 = new XlDialogBox.RefEdit() { X = 020, Y = 106, W = 250, };
            var ctrl_07 = new XlDialogBox.Label() { X = 020, Y = 130, Text = "Range of Data &Points", };
            var ctrl_08 = new XlDialogBox.RefEdit() { X = 020, Y = 146, W = 250, };
            var ctrl_09 = new XlDialogBox.GroupBox() { X = 290, Y = 020, W = 100, H = 050, Text = "Data &as:", };
            var ctrl_10 = new XlDialogBox.RadioButtonGroup() { IO = 2, };
            var ctrl_11 = new XlDialogBox.RadioButton() { Y = 035, Text = "Co&lumns", };
            var ctrl_12 = new XlDialogBox.RadioButton() { Y = 050, Text = "Ro&ws", IO = 1, };
            var ctrl_13 = new XlDialogBox.GroupBox() { X = 290, Y = 080, W = 100, H = 085, Text = "Plotti&ng", };
            var ctrl_14 = new XlDialogBox.RadioButtonGroup() { IO = 3, };
            var ctrl_15 = new XlDialogBox.RadioButton() { Y = 095, Text = "L&ines", };
            var ctrl_16 = new XlDialogBox.RadioButton() { Y = 110, Text = "&Markers", };
            var ctrl_17 = new XlDialogBox.RadioButton() { Y = 125, Text = "&Both", };
            var ctrl_18 = new XlDialogBox.CheckBox() { Y = 145, Text = "Colo&rs", };
            var ctrl_19 = new XlDialogBox.OkButton() { X = 020, Y = 180, W = 075, Text = "&OK", Default = true, };
            var ctrl_20 = new XlDialogBox.CancelButton() { X = 107, Y = 180, W = 075, Text = "&Cancel", };
            var ctrl_21 = new XlDialogBox.HelpButton2() { X = 194, Y = 180, W = 075, Text = "&Help", };
            // To maintain the whitespace between above {} brackets in VS19, go to Tools->Options->Text Editor->C#->Code Style->Formatting. Uncheck "Automatically format on paste" 

            // initilise ranges with static variables, saved earlier
            ctrl_02.IO_string = ChartTitle;
            ctrl_04.IO_string = TitleRange;
            ctrl_06.IO_string = DepthRange;
            ctrl_08.IO_string = Data_Range;
            ctrl_10.IO_index = RangesVert;
            ctrl_14.IO_index = Chart_type;
            ctrl_18.IO_checked = PlotColors;

            // The sequence of adding controls is important in view of the tab order.
            // Note: always put the 'labels' *in front* of their (edit/list) controls.
            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);
            dialog.Controls.Add(ctrl_11);
            dialog.Controls.Add(ctrl_12);
            dialog.Controls.Add(ctrl_13);
            dialog.Controls.Add(ctrl_14);
            dialog.Controls.Add(ctrl_15);
            dialog.Controls.Add(ctrl_16);
            dialog.Controls.Add(ctrl_17);
            dialog.Controls.Add(ctrl_18);
            dialog.Controls.Add(ctrl_19);
            dialog.Controls.Add(ctrl_20);
            dialog.Controls.Add(ctrl_21);
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod();
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

        } // Cmd_ChartDialog


        [ExcelCommand(
            Name = "Version_Info",
            Description = "Shows a dialog with information on library version and compilation date & time",
            HelpTopic = "XlmDialogExample-AddIn.chm!1004")]
        public static void Cmd_ShowVersion()
        {
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string version = v.ToString();
            System.DateTime date_time = System.IO.File.GetLastWriteTime(ExcelDnaUtil.XllPath);
            string date = date_time.ToString();


            var dialog  = new XlDialogBox()             {                    W = 313, H = 200, Text = "Version Info"};
            var ctrl_01 = new XlDialogBox.GroupBox()    {  X = 013, Y = 013, W = 287, H = 130, Text = "Geophysical and Geomatics function library",  };
            var ctrl_02 = new XlDialogBox.Label()       {  X = 031, Y = 039,                   Text = "Library version",  };
            var ctrl_03 = new XlDialogBox.TextEdit()    {  X = 031, Y = 058, W = 250,          };
            var ctrl_04 = new XlDialogBox.Label()       {  X = 031, Y = 091,                   Text = "Library compile date",  };
            var ctrl_05 = new XlDialogBox.TextEdit()    {  X = 031, Y = 110, W = 250,          };
            var ctrl_06 = new XlDialogBox.OkButton()    {  X = 031, Y = 160, W = 100,          Text = "&OK", Default = true, };
            var ctrl_07 = new XlDialogBox.HelpButton2() {  X = 181, Y = 160, W = 100,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);

            ctrl_03.IO_string = version;
            ctrl_05.IO_string = date;
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;
        } // Cmd_ShowVersion

        static bool ValidateAbout(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            System.Diagnostics.Process.Start("https://www.github.com/duijndam-dev/");
            return true; // return to dialog
        }

        [ExcelCommand(
            Name = "About_GeoLib",
            Description = "Shows a dialog with a copy right statement and a list of referenced NuGet packages",
            HelpTopic = "XlmDialogExample-AddIn.chm!1005"
            )]
        public static void Cmd_About()
        {
            var dialog  = new XlDialogBox()                  {                   W = 333, H = 240, Text = "About GeoLib",  };
            var ctrl_01 = new XlDialogBox.GroupBox()         { X = 013, Y = 013, W = 307, H = 130, Text = "This library uses the following NuGet packages",  };
            var ctrl_02 = new XlDialogBox.ListBox()          { X = 031, Y = 038, W = 270,          Text = "List_02" };
            var ctrl_03 = new XlDialogBox.OkButton()         { X = 031, Y = 160, W = 270,          Text = "DuijndamDev   |   Copyright © 2020 - 2021", IO = 1, };
            var ctrl_04 = new XlDialogBox.OkButton()         { X = 031, Y = 200, W = 100,          Text = "&OK", Default = true, };
            var ctrl_05 = new XlDialogBox.HelpButton2()      { X = 201, Y = 200, W = 100,          Text = "&Help",  };

            ctrl_02.Items.AddRange(new string[]              {
                "DotSpatial.Positioning   version=2.0.0-rc1", 
                "DotSpatial.Projections   version=2.0.0-rc1", 
                "ExcelDna.AddIn   version=1.1.1", 
                "ExcelDna.Integration   version=1.1.0", 
                "ExcelDna.IntelliSense   version=1.4.2", 
                "ExcelDna.Interop   version=14.0.1", 
                "ExcelDna.Registration   version=1.1.0", 
                "ExcelDna.Utilities   version=0.1.6", 
                "ExcelDna.XmlSchemas   version=1.0.0", 
                "ExcelDnaDoc   version=1.1.0-beta2", 
                "MathNet.Numerics   version=4.15.0", 
                "MathNet.Numerics.MKL.Win   version=2.5.0", });

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            bool bOK = dialog.ShowDialog(ValidateAbout);
            if (bOK == false) return;

        } // Cmd_about

        [ExcelCommand(
        Name = "File_Selector",
        Description = "Starts the File Selector Dialog",
        HelpTopic = "XlmDialogExample-AddIn.chm!1006",
        ShortCut = "^R")]
        public static void Cmd_FileSelector()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 420, H = 240, Text = "File finder",  IO =  7, };
            var ctrl_01 = new XlDialogBox.GroupBox()         {	 X = 013, Y = 010, W = 394, H = 040, Text = "Current directory at launch of dialog. Use ⭮ button to refresh ",  };
            var ctrl_02 = new XlDialogBox.OkButton()         {	 X = 018, Y = 026, W = 025,          Text = "⭮",  IO = 3, };
            var ctrl_03 = new XlDialogBox.DirectoryLabel()   {	 X = 048, Y = 030, W = 357,          };
            var ctrl_04 = new XlDialogBox.GroupBox()         {	 X = 013, Y = 055, W = 394, H = 140, Text = "File selector. Use *.* to search for all files in a folder ",  };
            var ctrl_05 = new XlDialogBox.TextEdit()         {	 X = 031, Y = 073, W = 170,          IO = "*.*", };
            var ctrl_06 = new XlDialogBox.LinkedFilesList()  {	 X = 220, Y = 073, W = 170, H = 110, IO = 2, };
            var ctrl_07 = new XlDialogBox.LinkedDriveList()  {	 X = 031, Y = 096,          H = 090, };
            var ctrl_08 = new XlDialogBox.OkButton()         {	 X = 151, Y = 205, W = 075,          Text = "&OK",  };
            var ctrl_09 = new XlDialogBox.CancelButton()     {	 X = 238, Y = 205, W = 075,          Text = "&Cancel",  };
            var ctrl_10 = new XlDialogBox.HelpButton2()      {	 X = 330, Y = 205, W = 075,          Text = "&Help",  IO = -1, };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI
            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

            // now it is time to play around with the parameters chosen in the dialog box to get things done

            string directory = System.IO.Directory.GetCurrentDirectory();
            directory = directory.TrimEnd('\\');    // network drives keep trailing backslash
            string file = ctrl_05.IO_string;
            string path = directory + "\\" + file;

            var xlApp = (Application)ExcelDnaUtil.Application;
            var ws = xlApp.Sheets[1] as Worksheet;
            var range = ws.Cells[1, 1] as Range;
            range.Value2 = path;
        } // Cmd_FileSelector

        [ExcelCommand(
            Name = "Show_Help",
            Description = "Shows the Compiled Help file",
            HelpTopic = "XlmDialogExample-AddIn.chm!1007"
            )]
        public static void Cmd_ShowHelp()
        {
            // get the Path of xll file;
            string xllPath = ExcelDnaUtil.XllPath;
            string xllDir  = System.IO.Path.GetDirectoryName(xllPath);

            var CallingMethod = System.Reflection.MethodBase.GetCurrentMethod();
            if (CallingMethod != null)
            {   // is there an ExcelCommandAttribute attribute decorating the method where ShowDialog has been called from ?
                ExcelCommandAttribute attr = (ExcelCommandAttribute)CallingMethod.GetCustomAttributes(typeof(ExcelCommandAttribute), true)[0];
                if (attr != null)
                {
                    // get the HelpTopic string and split it in two parts ([a] file name and [b] helptopic)
                    string[] parts = attr.HelpTopic.Split('!');

                    // the complete helpfile path consists of the xll directory + first part of HelpTopic attribute string 
                    string chmPath = System.IO.Path.Combine(xllDir, parts[0]);

                    // don't bother to start at a particular help topic
                    System.Diagnostics.Process.Start(chmPath);
                }
            }
        } // Cmd_ShowHelp

    }

}


