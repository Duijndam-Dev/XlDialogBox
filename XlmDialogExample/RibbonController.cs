using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

using ExcelDna.Integration.CustomUI;

namespace Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private Excel.Application _excel;

        public override string GetCustomUI(string RibbonID)
        {
            _excel = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

            return @"
              <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
              <ribbon>
                <tabs>
                  <tab id='tab1' label='XlDialogBox'>
                    <group id='group1' label='Dialog Examples'>
                      <button id='button0' label='Generic Example' imageMso='DataFormExcel' size='large' onAction='OnButton0Pressed'/>
                      <button id='button1' label='Ray Tracer'      imageMso='CreateQueryFromWizard' size='large' onAction='OnButton1Pressed'/>
                      <button id='button2' label='Chart Dialog'    imageMso='DataFormExcel' size='large' onAction='OnButton2Pressed'/>
                      <button id='button3' label='Version Info'    imageMso='ReviewCompareMajorVersion' size='large' onAction='OnButton3Pressed'/>
                      <button id='button4' label='About GeoLib'    imageMso='FontDialog'    size='large' onAction='OnButton4Pressed' />
                      <button id='button5' label='File Selector'   imageMso='FileOpen' size='large' onAction='OnButton5Pressed'/>
                    </group >
                    <group id='group2' label='Help'>
                      <button id='button6' label='Show Help' imageMso='TentativeAcceptInvitation' size='large' onAction='OnButton6Pressed' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnButton0Pressed(IRibbonControl control)
        {
            _excel.Application.Run("Generic_Example");
        }

        public void OnButton1Pressed(IRibbonControl control)
        {
            _excel.Application.Run("Ray_Tracer");
        }

        public void OnButton2Pressed(IRibbonControl control)
        {
            _excel.Application.Run("Chart_Dialog");
        }

        public void OnButton3Pressed(IRibbonControl control)
        {
            _excel.Application.Run("Version_Info");
        }

        public void OnButton4Pressed(IRibbonControl control)
        {
            _excel.Application.Run("About_GeoLib");
        }

        public void OnButton5Pressed(IRibbonControl control)
        {
            _excel.Application.Run("File_Selector");
        }

        public void OnButton6Pressed(IRibbonControl control)
        {
            _excel.Application.Run("Show_Help");
        }

    }
}
