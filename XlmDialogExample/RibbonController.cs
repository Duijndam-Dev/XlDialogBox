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
                      <button id='button1' label='Dialog-1' imageMso='DataFormExcel' size='large' onAction='OnButton1Pressed'/>
                    </group >
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnButton1Pressed(IRibbonControl control)
        {
            _excel.Application.Run("Dialog1");
        }
    }
}
