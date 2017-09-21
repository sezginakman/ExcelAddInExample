using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void showButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                await Task.Run(() => {

                    var worksheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                    foreach (var worksheetComment in worksheet.Comments)
                    {
                        var comment = worksheetComment as Excel.Comment;
                        if (comment != null)
                            comment.Visible = true;
                    }
                });
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                MessageBox.Show(e.ToString());
            }
        }

        private void hideButton_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private async void trimButton_Click(object sender, RibbonControlEventArgs e)
        {
try
            {
                await Task.Run(() =>
                {
                    var worksheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

                    foreach (var worksheetComment in worksheet.Comments)
                    {
                        var comment = worksheetComment as Excel.Comment;
                        if (comment != null)
                            comment.Shape.TextFrame.AutoSize = true;

                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
