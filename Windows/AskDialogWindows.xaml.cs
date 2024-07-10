using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

namespace ConvertorOrder.Windows
{
    /// <summary>
    /// Логика взаимодействия для AskDialogWindows.xaml
    /// </summary>
    public partial class AskDialogWindows : System.Windows.Window
    {
        public AskDialogWindows()
        {
            InitializeComponent();
            var order = GetInfoOrder();
            if (order == null) order = new InfoOrder();
            textBlockCastimer.Text = "Заказчик " + order.nameCastumer;
            textBlockOrder.Text = order.number;
            textBlockCount.Text = "Колличество товара: " + order.countProducts.ToString();
        }
        public string GetNameFile()
        {
            DirectoryInfo directory = new DirectoryInfo(
           Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads");
            var files = directory.GetFiles("order code productCode-*.xls").ToList();
            if (files.Count < 1)
            {
                MessageBox.Show("Не найден файл, возможно вы ничего не скачали");
                return null;
            }
            DateTime tempTime = files[0].CreationTime;
            int tempIndex = 0;
            for (int i = 0; i < files.Count; i++)
            {
                if (files[i].Name.Contains("Сборочный"))
                {
                    files.RemoveAt(i);
                    i--;
                }
                else
                if (tempTime < files[i].CreationTime)
                {
                    tempTime = files[i].CreationTime;
                    tempIndex = i;
                }
            }
            if (tempTime.AddMinutes(50) > DateTime.Now)
            {
                if (files.Count > 0)
                    return files[tempIndex].FullName;
                else
                    return null;
            }
            return null;
        }
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = null;
        Excel.Worksheet worksheet = null;
        private class InfoOrder
        {
            public string number = "";
            public string nameCastumer = "";
            public int countProducts = 0;
            public int countHaveProducts = 0;
        }

        private InfoOrder GetInfoOrder()
        {
            InfoOrder order = new InfoOrder();
            DirectoryInfo directory = new DirectoryInfo(
          Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads");
            var files = directory.GetFiles("order code productCode-*.xls").ToList();
            if (files.Count < 1)
            {
                MessageBox.Show("Не найден файл, возможно вы ничего не скачали");
                return null;
            }
            DateTime tempTime = files[0].CreationTime;
            int tempIndex = 0;
            for (int i = 0; i < files.Count; i++)
            {
                if (files[i].Name.Contains("Сборочный"))
                {
                    files.RemoveAt(i);
                    i--;
                }
                else
                    if (tempTime < files[i].CreationTime)
                   {
                        tempTime = files[i].CreationTime;
                        tempIndex = i;
                    }
            }
            if (tempTime.AddMinutes(50) < DateTime.Now)
                return null;
            if (files.Count < 1)
            {
                return null;
            }
            workbook = excelApp.Workbooks.Open(files[tempIndex].FullName);
            worksheet = workbook.Sheets[1];

            order.number = GetTextCell(5, 1);
            order.nameCastumer = GetTextCell(8, 3);
            int countRow;
            int b = 1;
            while (true)
            {
                if (GetTextCell(b, 1).StartsWith("Итого"))
                {
                    countRow = b;
                    break;
                }
                b++;
                if (b > 1000)
                {
                    workbook.Close();
                    excelApp.Quit();
                    return null;
                }
            }

            order.countProducts = countRow - 14;

            return order;
        }
        private string GetTextCell(int row, int column)
        {
            if ((worksheet.Cells[row, column] as Range).Value == null) return "";
            return (worksheet.Cells[row, column] as Range).Value.ToString();
        }
        public void EditFile(string path)
        {
            if (path != null)
            {
                workbook = excelApp.Workbooks.Open(path);
                worksheet = workbook.Sheets[1];
                int countRow = 0;
                worksheet.get_Range("A1:A2").EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                worksheet.Cells[3, 1].Value = "Сборка заказа" + GetTextCell(3, 1).Split('№')[1];

                worksheet.get_Range("I1").EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);

                int i = 1;
                while (true)
                {
                    if (GetTextCell(i, 1).StartsWith("Итого"))
                    {
                        countRow = i;
                        break;
                    }
                    i++;
                    if (i > 1000)
                    {
                        MessageBox.Show("В вашем файле больше 1000 строк\nИли произошла ошибка!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                        workbook.Close();
                        excelApp.Quit();
                        return;
                    }
                }


                worksheet.Cells[8, 9].Value = "Пометка";
                worksheet.get_Range("I8:I" + (countRow - 4)).Borders.LineStyle = XlLineStyle.xlContinuous;

                worksheet.Cells[8, 10].Value = "Статус";
                worksheet.get_Range("J8:J" + (countRow - 4)).Borders.LineStyle = XlLineStyle.xlContinuous;


                worksheet.get_Range("A" + countRow).EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                worksheet.get_Range("A" + (countRow - 3)).EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                worksheet.get_Range("G1").EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                worksheet.get_Range("B1").EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);

                       
            }
            else
            {
                MessageBox.Show("Файл не скачался! Или слишком долго скачивался!\nПопробуйте снова\nИли зайдите в прочее и выберите файл!");
            }
        }

        private void buttonPrint_Click(object sender, RoutedEventArgs e)
        {
            if (GetNameFile() != null)
                EditFile(GetNameFile());
            // Сохраняем изменения и закрываем книгу
            if (GetNameFile() != null)
                workbook.SaveAs(GetNameFile().Insert(GetNameFile().Length - 4, " Сборочный "));

            if (GetNameFile() != null)
                worksheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Закрываем книгу
            if (GetNameFile() != null)
                workbook.Close();

            // Закрываем приложение Excel
            excelApp.Quit();

            if (GetNameFile() != null)
                MessageBox.Show("Успешно завершено!");
            else
                MessageBox.Show("Файл не был найден в папке загрузок, он должне быть только что скаченным (до 50 минут)");
            Environment.Exit(0);

        }

        private void buttonMore_Click(object sender, RoutedEventArgs e)
        {
            if (GetNameFile() != null)
                EditFile(GetNameFile());
            excelApp.Visible = true;
        }
    }
}
