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
                addRows();
                bSaveChg.Enabled = false;
            }
        }

        private void addRows()
        {

            string[] row0 = { "", "", "", "" };
            string varname;
            string VarVal;
            string comment;
            for (int i = 0; i < QV.v.count; i++)
            {
                varname = QV.v.Item(i).Name;
                if (!QV.varS.Contains(varname))
                {
                    VarVal = QV.doc.Variables(varname).GetRawContent();
                    comment = QV.doc.Variables(varname).GetComment();
                    row0[0] = varname;
                    row0[1] = VarVal;
                    row0[2] = comment;
                    dgVars.Rows.Add(row0);
                }
            }
        }

        private void dgVars_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

            txValCorrect.Text =  (string)dgVars[1, e.RowIndex].Value;
            txCommCorrect.Text = (string)dgVars[2, e.RowIndex].Value;
        }

        private void txCommCorrect_Leave(object sender, EventArgs e)
        {
            //string t = (string)dgVars[3, dgVars.CurrentRow.Index].Value;
            dgVars[2, dgVars.CurrentRow.Index].Value = txCommCorrect.Text;
            //int t = dgVars.CurrentRow.Index;
            if (dgVars[0, dgVars.CurrentRow.Index].Value != null) 
                QV.doc.Variables(dgVars[0, dgVars.CurrentRow.Index].Value).SetComment(txCommCorrect.Text);
            //string t= QV.doc.Variables(dgVars[0, dgVars.CurrentRow.Index].Value).GetComment();
            //QV.doc.Save();
            bSaveChg.Enabled = true;
        }

        private void bSaveChg_Click(object sender, EventArgs e)
        {
            try
            {
                bSaveChg.Text = "Save changes";
                bSaveChg.Enabled = false;
                for (int i = 0; i < dgVars.Rows.Count-1; i++)
                {
                    QV.doc.Variables(dgVars[0, i].Value).SetComment(dgVars[2, i].Value);
                    QV.doc.Variables(dgVars[0, i].Value).SetContent(dgVars[1, i].Value, true);
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
           dSF.FileName = Path.GetFileName(txFlNameComm.Text);
           dSF.InitialDirectory = Path.GetDirectoryName(txFlNameComm.Text);
           dSF.FilterIndex = 2;


            if (dSF.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            txFlNameComm.Text = dSF.FileName;

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

            xlWorkBook.SaveAs(dSF.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
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
                bool bFieldExists = false;
                s = (string)(range.Cells[i, 1] as Excel.Range).Value;
                for (j = 0; j <= dgVars.Rows.Count - 1; j++)
                {   
                    // compare names of vars - from 0 col in grid and 1 in xls
                    if (dgVars[0, j].Value != null && dgVars[0, j].Value.Equals(s))
                    {
                        if (chkNewImport.Checked)
                        {
                            bFieldExists = true;
                        }
                        var ss = (string)(range.Cells[i, 3] as Excel.Range).Value;
                        if (dgVars[2, j].Value != null && !dgVars[2, j].Value.Equals(ss))
                        {
                            dgVars[2, j].Value = ss;
                        }
                        ss = (string)(range.Cells[i, 2] as Excel.Range).Value;
                        if (dgVars[1, j].Value != null && !dgVars[1, j].Value.Equals(ss))
                        {
                            dgVars[1, j].Value = ss;
                        }
                        break;
                    }
                }
                if ((chkNewImport.Checked)&&(!bFieldExists))
                {
                    string[] row0 = { "", "", "" };
                    row0[0] = (string)(range.Cells[i, 1] as Excel.Range).Value;
                    row0[1] = (string)(range.Cells[i, 2] as Excel.Range).Value;
                    row0[2] = (string)(range.Cells[i, 3] as Excel.Range).Value;
                    dgVars.Rows.Add(row0);

                    QV.doc.CreateVariable(row0[0]);
                    //QV.doc.Variables(row0[0]).SetContent(row0[1], true);
                }
            }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            bSaveChg.Enabled = true;
        }

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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
