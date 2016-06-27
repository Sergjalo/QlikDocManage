using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel; 

namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        QlikInstance QV;
    
        public Form1()
        {
            InitializeComponent();
        }

        public class QlikInstance
        {
            public dynamic doc;
            public dynamic v;

            public dynamic Doc
            {
                get { return doc; }
                set { doc = value; }
            }

            public dynamic V
            {
                get { return v; }
                set { v = value;  }
            }

            public HashSet<String> varS;

            public QlikInstance (string sFileName)
            {
                String[] a = new String[] {"CD","QvPath",	"QvRoot",	"QvWorkPath",	"QvWorkRoot",	"WinPath",	"WinRoot",	"ErrorMode",	"StripComments",
                "OpenUrlTimeout",	"ScriptErrorCount",	"ScriptErrorList",	
                "ScriptError",	"ThousandSep",	"DecimalSep",	"MoneyThousandSep",	"MoneyDecimalSep",	"MoneyFormat",	"TimeFormat",	"DateFormat",	
                "TimestampFormat",	"MonthNames",	"DayNames",	"HidePrefix",	"ScriptErrorDetails", "FLOPPY"};
                varS = new HashSet<String>(a);
                System.Type objType = System.Type.GetTypeFromProgID("QlikTech.QlikView");
                dynamic comObject = System.Activator.CreateInstance(objType);
                Doc = comObject.OpenDoc(sFileName); 
                V = Doc.GetVariableDescriptions();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //txVarNames.Text = "";
            dgVars.Rows.Clear();


           dOF.FileName = Path.GetFileName(txFlName.Text);
           dOF.InitialDirectory = Path.GetDirectoryName(txFlName.Text);
           dOF.FilterIndex = 1;

            if (dOF.ShowDialog() == DialogResult.OK)
            {
                try {
                   QV.doc.CloseDoc();
                }
                catch { }
                txFlName.Text = dOF.FileName;
                QV = new QlikInstance(dOF.FileName);
                lCnt.Text = "Non-system variables count - " + QV.v.count.ToString();

                string[] row0 = {"","","",""};
                string varname;
                string VarVal;
                string comment;
                for (int i = 0; i < QV.v.count; i++)
                {
                    varname = QV.v.Item(i).Name;
                    if (!QV.varS.Contains(varname))
                    {
                        //txVarNames.AppendText(varname);
                        //txVarNames.AppendText(Environment.NewLine);

                        VarVal = QV.doc.Variables(varname).GetRawContent();
                        comment = QV.doc.Variables(varname).GetComment();
                        row0[0] = varname;
                        row0[1] = VarVal;
                        row0[2] = comment;
                        //row0[3] = i.ToString();
                        dgVars.Rows.Add(row0);
                    }
                }
                //QV.doc.CloseDoc();
                bSaveChg.Enabled = false;
            }
        }

        private void dgVars_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgVars_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

            txValCorrect.Text =  (string)dgVars[1, e.RowIndex].Value;
            txCommCorrect.Text = (string)dgVars[2, e.RowIndex].Value;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dgVars_DoubleClick(object sender, EventArgs e)
        {

        }

        private void txCommCorrect_Leave(object sender, EventArgs e)
        {
            //string t = (string)dgVars[3, dgVars.CurrentRow.Index].Value;
            dgVars[2, dgVars.CurrentRow.Index].Value = txCommCorrect.Text;
            //int t = dgVars.CurrentRow.Index;
            QV.doc.Variables(dgVars[0, dgVars.CurrentRow.Index].Value).SetComment(txCommCorrect.Text);
            //string t= QV.doc.Variables(dgVars[0, dgVars.CurrentRow.Index].Value).GetComment();
            //QV.doc.Save();
            bSaveChg.Enabled = true;
        }

        private void bSaveChg_Click(object sender, EventArgs e)
        {
            try
            {
                bSaveChg.Enabled = false;
                for (int i = 0; i < dgVars.Rows.Count-1; i++)
                {
                    QV.doc.Variables(dgVars[0, i].Value).SetComment(dgVars[2, i].Value);    
                }
                QV.doc.Save();
            }
            catch
            {
                bSaveChg.Enabled = true;
                bSaveChg.Text = "Saving error";
            }
        }

        private void dgVars_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            /*
            if (txCommCorrect.Text != (string)dgVars[2, e.RowIndex].Value)
            {
                bSaveChg.Enabled = true;
                txCommCorrect.Text = (string)dgVars[2, e.RowIndex].Value;
                QV.doc.Variables(dgVars[0, dgVars.CurrentRow.Index].Value).SetComment(txCommCorrect.Text);
            }
             */ 
        }

        private void bExport_Click(object sender, EventArgs e)
        {
           dOF.FileName = Path.GetFileName(txFlNameComm.Text);
           dOF.InitialDirectory = Path.GetDirectoryName(txFlNameComm.Text);
           dOF.FilterIndex = 2;


            if (dOF.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            txFlNameComm.Text = dOF.FileName;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;
            
            for (i = 0; i <= dgVars.RowCount - 1; i++)
            {
                for (j = 0; j <= dgVars.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dgVars[j, i];
                    xlWorkSheet.Cells[i + 1, j + 1].Value2 = "'"+(string)cell.Value;
                }
            }

            xlWorkBook.SaveAs(dOF.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void bLoad_Click(object sender, EventArgs e)
        {
            dOF.FileName = Path.GetFileName(txFlNameComm.Text);
            dOF.InitialDirectory = Path.GetDirectoryName(txFlNameComm.Text);
            dOF.FilterIndex = 2;

            if (dOF.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            txFlNameComm.Text = dOF.FileName;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(
                               dOF.FileName, 0, true, 5,
                               "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                               0, true, 0, 0); 
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;
            /*
            for (i = 0; i <= dgVars.RowCount - 1; i++)
            {
                for (j = 0; j <= dgVars.ColumnCount - 1; j++)
                {
                    dgVars[j, i].Value = (string)xlWorkSheet.Cells[i + 1, j + 1].value2;
                }
            }
             */ 
            
            range = xlWorkSheet.UsedRange;
            /*
            range = (Excel.Range) xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[3, 3]];
            System.Array myvalues = (System.Array)range.Cells.Value;
            string[] strArray = ConvertToStringArray(myvalues);
            */
            string s;
            for (i = 1; i <= range.Rows.Count; i++)
            {
                s = (string)(range.Cells[i, 1] as Excel.Range).Value;
                for (j = 0; j <= dgVars.Rows.Count - 1; j++)
                {   // compare names of vars - from 0 col in grid and 1 in xls
                    if (dgVars[0, j].Value != null && dgVars[0, j].Value.Equals(s))
                    {
                        s = (string)(range.Cells[i, 3] as Excel.Range).Value;
                        if (dgVars[2, j].Value != null && !dgVars[2, j].Value.Equals(s))
                        {
                            dgVars[2,j].Value = s;
                        }
                        break;
                    }
                }
            }
            /*
            for (i = 1; i <= range.Rows.Count; i++)
            {
                for (j = 1; j <= range.Columns.Count; j++)
                {
                    try
                    {
                        dgVars[j-1, i-1].Value = (string)(range.Cells[i, j] as Excel.Range).Value;
                    }
                    catch
                    {
                        try
                        {
                            dgVars[j - 1, i - 1].Value = ((double)(range.Cells[i, j] as Excel.Range).Value).ToString();
                        }
                        catch
                        {
                            dgVars[j - 1, i - 1].Value = ((DateTime)(range.Cells[i, j] as Excel.Range).Value).ToString("yyyy-MM-dd HH:mm");
                        }
                    }
                    //string s= (string)(range.Cells[i, j] as Excel.Range).Value;
                }
            }
            */

            //xlWorkBook.SaveAs(Path.GetDirectoryName(dOF.FileName) + "\\test.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            bSaveChg.Enabled = true;
        }
        /*
        string[] ConvertToStringArray(System.Array values)
        {
            string[] theArray = new string[values.Length];
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }
            return theArray;
        }
        */
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
