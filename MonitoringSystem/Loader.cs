using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;

namespace MonitoringSystem
{
    class Loader
    {
        string FileName;
        MainForm f;
        private Dictionary<string, double> data = new Dictionary<string, double>();
        public Loader(MainForm f, string FileName)
        {
            this.f = f;
            this.FileName = FileName;
        }

        public Dictionary<string, double> GetData()
        {
            return data;
        }

        public Dictionary<string, double> XlsRead()
        {
            try
            {
                //добавить отлов исключения при неналичии экселя
                Excel.Application excelapp = null;
                Excel.Workbooks excelappworkbooks;
                Excel.Workbook excelappworkbook;
                Excel.Sheets excelsheets;
                Excel.Worksheet excelworksheet;
                Excel.Range excelcells;
                Excel.Range excelcells2;
                excelapp = new Excel.Application();
                excelapp.Visible = false;
                excelappworkbooks = excelapp.Workbooks;
                //Открываем книгу и получаем на нее ссылку
                excelappworkbook = excelapp.Workbooks.Open(@FileName,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
                excelsheets = excelappworkbook.Worksheets;
                //Получаем ссылку на лист 1
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                try
                {
                    int m;
                    bool eof = false;
                    for (m = 2; eof == false; m++)
                    {
                        excelcells = (Excel.Range)excelworksheet.Cells[m, 1];
                        excelcells2 = (Excel.Range)excelworksheet.Cells[m, 2];
                        if (excelcells.Value2 != null)
                        {
                            data.Add(Convert.ToString(excelcells.Value2), Convert.ToDouble(excelcells2.Value2));
                            //f.SetTextbox1(m.ToString()+" "+Convert.ToString(excelcells.Value2));
                        }
                        else eof = true;
                        if (m == 200) { }
                    }
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Ошибка чтения!");
                    excelcells = null;
                    excelappworkbook.Close();
                    excelapp.Workbooks.Close();
                    excelappworkbooks = null;
                    excelworksheet = null;
                    excelsheets = null;
                    excelappworkbook = null;
                    excelapp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
                    excelapp = null;
                    GC.Collect();
                }
                finally
                {
                    excelcells = null;
                    excelappworkbook.Close();
                    excelapp.Workbooks.Close();
                    excelappworkbooks = null;
                    excelworksheet = null;
                    excelsheets = null;
                    excelappworkbook = null;
                    excelapp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
                    excelapp = null;
                    GC.Collect();
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Файл не существует, проверьте правильность пути!");
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
            return data;
        }
        
    }
}
