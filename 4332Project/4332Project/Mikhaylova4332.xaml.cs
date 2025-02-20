using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace _4332Project
{
    /// <summary>
    /// Логика взаимодействия для Mikhaylova4332.xaml
    /// </summary>
    public partial class Mikhaylova4332 : Window
    {
        public Mikhaylova4332()
        {
            InitializeComponent();
        }

        private void ExcelImportButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }

            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (MikhailovaISRPOEntities1 usersEntities = new MikhailovaISRPOEntities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Employee.Add(new Employee()
                    {
                        EmployeeId = list[i, 0],
                        Position = list[i, 1],
                        FIO = list[i, 2],
                        Login = list[i, 3],
                        Password = list[i, 4],
                        LastEnter = list[i, 5],
                        TypeEnter = list[i, 6],
                    });
                }
                try
                {
                    

                    usersEntities.SaveChanges();
                }
                catch (DbEntityValidationException ex)
                {
                    foreach (var eve in ex.EntityValidationErrors)
                    {
                        MessageBox.Show($"Entity of type \"{eve.Entry.Entity.GetType().Name}\" in state \"{eve.Entry.State}\" has the following validation errors:");
                           
                        foreach (var ve in eve.ValidationErrors)
                        {
                            MessageBox.Show($"- Property: \"{ve.PropertyName}\", Error: \"{ve.ErrorMessage}\"");
                               
                        }
                    }
                    throw;
                }

            }
            MessageBox.Show("Данные импортированы успешно!",
                               "Внимание!",
                               MessageBoxButton.OK,
                               MessageBoxImage.Information);

        }

        private void ExcelExportButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog()
            {
                DefaultExt = "*.xlsx",
                Filter = "файл Excel (*.xlsx)|*.xlsx",
                Title = "Сохранить файл базы данных"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    Excel.Application ObjWorkExcel = new Excel.Application();
                    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Add();

                    using (MikhailovaISRPOEntities1 dbContext = new MikhailovaISRPOEntities1())
                    {

                        var clients = dbContext.Employee.ToList();

                        var groupedClients = clients
                            .GroupBy(c => c.TypeEnter);

                        foreach (var group in groupedClients)
                        {
                            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.Add();

                            ObjWorkSheet.Cells[1, 1] = "Код сотрудника";
                            ObjWorkSheet.Cells[1, 2] = "Должность";
                            ObjWorkSheet.Cells[1, 3] = "Логин";
                           

                            var sortedClients = group.OrderBy(c => c.EmployeeId).ToList();

                            int row = 2;
                            foreach (var client in sortedClients)
                            {
                                ObjWorkSheet.Cells[row, 1] = client.EmployeeId;
                                ObjWorkSheet.Cells[row, 2] = client.Position;
                                ObjWorkSheet.Cells[row, 3] = client.Login;
                              
                                row++;
                            }
                        }
                    }

                    ObjWorkBook.SaveAs(sfd.FileName);
                    ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                    ObjWorkExcel.Quit();
                    GC.Collect();

                    MessageBox.Show("Данные успешно экспортированы в Excel!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте данных: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
        }

    }


}
    


