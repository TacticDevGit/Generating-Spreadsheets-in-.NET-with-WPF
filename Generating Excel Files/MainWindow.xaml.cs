using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using SpreadsheetLight;
using Microsoft.Win32;


namespace Generating_Excel_Files
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        public void CreateFile(string path)
        {

            if (path != String.Empty)
            {

                using (SLDocument Doc = new SLDocument())
                {

                    for (int i = 1; i < 11; i++)
                    {
                        Doc.SetCellValue($"B{i}", $"Item {i}");
                        Doc.SetCellValue($"C{i}", i * 2);

                    }

                    Doc.SetCellValue("C13", "=SUM(C1:C10)");


                    var table = Doc.CreateTable("A1", "C10");
                    table.SetTableStyle(SLTableStyleTypeValues.Medium15);
                    Doc.InsertTable(table);

                    Doc.SaveAs(path);
                }

            }

        }


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.DefaultExt = "xlsx";

            if (saveFile.ShowDialog() == true)
            {

                CreateFile(saveFile.FileName);


            }


        }
    }
}
