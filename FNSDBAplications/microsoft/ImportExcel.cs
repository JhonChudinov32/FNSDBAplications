using System;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using FNSDBAplications.connection;
using Appl = Microsoft.Office.Interop.Excel;
using System.Windows;

namespace FNSDBAplications.microsoft
{
    public class ImportExcel
    {
        //универсальная выгрузка
        public static void ToExcel(string filename, string query)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;
            Appl.Range xlSheetRange;
            String sql = query;

            try
            {
                wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
                ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
                ws_New1.Cells.Locked = false;

                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int collInd = 0;
                int rowInd = 0;
                string data = "";
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    data = dt.Columns[i].ColumnName.ToString();
                    ws_New1.Cells[1, i + 1] = data;

                    //выделяем первую строку
                    xlSheetRange = ws_New1.get_Range("A1:Z1", Type.Missing);

                    //делаем полужирный текст и Авторасширение по заголовкам
                    xlSheetRange.Font.Bold = true;
                    xlSheetRange.Columns.AutoFit();
                }
                //заполняем строки
                for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                {
                    for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                    {
                        data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                        ws_New1.Cells[rowInd + 2, 2].NumberFormat = "@";
                        ws_New1.Cells[rowInd + 2, 7].NumberFormat = "@";
                        ws_New1.Cells[rowInd + 2, collInd + 1] = data;
                        ws_New1.Cells[rowInd + 2, collInd + 1].WrapText = true;
                    }
                }

                //выбираем всю область данных
                xlSheetRange = ws_New1.UsedRange;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1.Visible = true;
                Marshal.ReleaseComObject(exApp_New1);
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка доходов расходов
        public static void ToExcel_Debet_Credet(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;
           
            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;
           
            String sql = @"SELECT [РИЦ],[ИНН],[Сумма доходов],[Сумма Расходов],[Год] FROM [fns].[dbo].[DOCHODRASHOD] Where Год = Year(getDate())-1 order by РИЦ";
            try
            {
                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth5 = ws_New1.Range["E1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "Сумма доходов";
                ws_New1.Cells[1, 4] = "Сумма расходов";
                ws_New1.Cells[1, 5] = "Год";
                //Из dataGridView1 в Excel
                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    ws_New1.Cells[i, 5] = row[4].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка ГИРБО
        public static void ToExcel_Girbo(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;

            String sql = @"SELECT [РИЦ],[ИНН],[БАЛАНС],[Финансовый результат ],[Год] FROM [fns].[dbo].[finotchet] Where [Год] = Year(getDate())-1 Order by РИЦ";
            try
            {
                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth5 = ws_New1.Range["E1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "Баланс";
                ws_New1.Cells[1, 4] = "Фин.результат";
                ws_New1.Cells[1, 5] = "Год";
                //Из dataGridView1 в Excel
                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    ws_New1.Cells[i, 5] = row[4].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }

        }
        //Выгрузка Минцифры
        public static void ToExcel_Mincifri(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            try
            {
                String sql = @"SELECT [RIC],[Name],[Status],[ITLgota],[State],[Accreditation],[INN],[StatusMinc] FROM [fns].[dbo].[RicAccreditation] ORDER by [RIC] ASC,[Status] DESC";

                wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
                ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
                ws_New1.Cells.Locked = false;

                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 10;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 10;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 35;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth4.EntireColumn.ColumnWidth = 20;
                Appl.Range rangeWidth5 = ws_New1.Range["E1", Type.Missing];
                rangeWidth5.EntireColumn.ColumnWidth = 30;
                Appl.Range rangeWidth6 = ws_New1.Range["F1", Type.Missing];
                rangeWidth6.EntireColumn.ColumnWidth = 25;
                Appl.Range rangeWidth7 = ws_New1.Range["G1", Type.Missing];
                rangeWidth7.EntireColumn.ColumnWidth = 25;
                Appl.Range rangeWidth8 = ws_New1.Range["H1", Type.Missing];
                rangeWidth8.EntireColumn.ColumnWidth = 20;
                Appl.Range rangeWidth9 = ws_New1.Range["I1", Type.Missing];
                rangeWidth9.EntireColumn.ColumnWidth = 35;

                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;

                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "I1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;

                ws_New1.Cells[1, 1] = "№ П/П";
                ws_New1.Cells[1, 2] = "РИЦ";
                ws_New1.Cells[1, 3] = "Наименование (краткое)";
                ws_New1.Cells[1, 4] = "Статус";
                ws_New1.Cells[1, 5] = "IT льгота";
                ws_New1.Cells[1, 6] = "Состояние";
                ws_New1.Cells[1, 7] = "Аккредитация";
                ws_New1.Cells[1, 8] = "ИНН";
                ws_New1.Cells[1, 9] = "Статус из реестра Минцифры";


                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = i - 1;
                    ws_New1.Cells[i, 2] = row[0].ToString();
                    ws_New1.Cells[i, 3] = row[1].ToString();
                    ws_New1.Cells[i, 4] = row[2].ToString();
                    ws_New1.Cells[i, 5] = row[3].ToString();
                    ws_New1.Cells[i, 6] = row[4].ToString();
                    ws_New1.Cells[i, 7] = row[5].ToString();
                    ws_New1.Cells[i, 8].NumberFormat = "@";
                    ws_New1.Cells[i, 8] = row[6].ToString();
                    ws_New1.Cells[i, 9] = row[7].ToString();

                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;

                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка Налоги
        public static void ToExcel_Nalog(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;

            try
            {
                String sql = @"SELECT [РИЦ],[ИНН],[Наименование налога],[summUplnalog],[Год] FROM [fns].[dbo].[RICNALOG] Where Год = year(getdate())-1 Order by РИЦ";

                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth4.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth5 = ws_New1.Range["E1", Type.Missing];
                rangeWidth5.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "Наименование налога";
                ws_New1.Cells[1, 4] = "Сумма уплаты налога";
                ws_New1.Cells[1, 5] = "Год";
                //Из dataGridView1 в Excel
                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4].NumberFormat = "@";
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    ws_New1.Cells[i, 5] = row[4].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка Штрафы
        public static void ToExcel_Shtraf(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;

            try
            {
                String sql = @"SELECT  [РИЦ],[ИНН],[Правонарушение],[Год] FROM [fns].[dbo].[PRAVONAR] Where Год=Year(getDate())-1 order by РИЦ";

                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "Штрафы";
                ws_New1.Cells[1, 4] = "Год";
            
                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка Пени
        public static void ToExcel_Peni(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;

            try
            {
                String sql = @"SELECT distinct [РИЦ],[ИНН] ,[Наименование налога],[СумНедНалог],[Сумма пени],[Сумма Штрафа],[ОбщаяСуммаНедоим],[Год] FROM [fns].[dbo].[NEDOIMK] Where Год = Year(getDate())-1 order by РИЦ";

                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth5 = ws_New1.Range["E1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth6 = ws_New1.Range["F1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth7 = ws_New1.Range["G1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth8 = ws_New1.Range["H1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "Наименование налога";
                ws_New1.Cells[1, 4] = "Сумма недоимок";
                ws_New1.Cells[1, 5] = "Сумма пени";
                ws_New1.Cells[1, 6] = "Сумма штрафа";
                ws_New1.Cells[1, 7] = "Общая сумма";
                ws_New1.Cells[1, 8] = "Год";

                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    ws_New1.Cells[i, 5] = row[4].ToString();
                    ws_New1.Cells[i, 6] = row[5].ToString();
                    ws_New1.Cells[i, 7] = row[6].ToString();
                    ws_New1.Cells[i, 8] = row[7].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка РМСП
        public static void ToExcel_RMSP(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;

            try
            {
                String sql = @"SELECT [РИЦ],[ИНН],[Дата включения],[Категория],[Год] FROM [fns].[dbo].[RICRMSP] Where Год = YEAR(getDate()) order by РИЦ";

                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth5 = ws_New1.Range["E1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "Дата включения";
                ws_New1.Cells[1, 4] = "Категория";
                ws_New1.Cells[1, 5] = "Год";

                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    ws_New1.Cells[i, 5] = row[4].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка СЧР
        public static void ToExcel_SCHR(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;

            try
            {
                String sql = @"SELECT  [РИЦ],[ИНН],[Количество],[Год] FROM [fns].[dbo].[RICSSCHR] where Год = Year(getDate())-1  order by РИЦ";

                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "Количество";
                ws_New1.Cells[1, 4] = "Год";

                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
        //Выгрузка СпецРежимов
        public static void ToExcel_USN(string filename)
        {
            Appl.Application exApp_New1 = new Appl.Application();
            Appl.Workbook wb_New1 = null;
            Appl.Worksheet ws_New1 = null;

            wb_New1 = exApp_New1.Workbooks.Add(System.Reflection.Missing.Value);
            ws_New1 = (Appl.Worksheet)wb_New1.Worksheets.get_Item(1);
            ws_New1.Cells.Locked = false;

            try
            {
                String sql = @"SELECT [РИЦ],[ИНН],[УСН],[Год] FROM [fns].[dbo].[RICUSN] where Год = Year(getDate())-1  order by РИЦ";

                //Ширина столбцов
                Appl.Range rangeWidth1 = ws_New1.Range["A1", Type.Missing];
                rangeWidth1.EntireColumn.ColumnWidth = 5;
                Appl.Range rangeWidth2 = ws_New1.Range["B1", Type.Missing];
                rangeWidth2.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth3 = ws_New1.Range["C1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                Appl.Range rangeWidth4 = ws_New1.Range["D1", Type.Missing];
                rangeWidth3.EntireColumn.ColumnWidth = 15;
                //Оформление листа + печать
                ws_New1.PageSetup.Orientation = Appl.XlPageOrientation.xlPortrait;
                ws_New1.PageSetup.PaperSize = Appl.XlPaperSize.xlPaperA4;
                ws_New1.PageSetup.TopMargin = 1;
                ws_New1.PageSetup.RightMargin = 0.75;
                ws_New1.PageSetup.LeftMargin = 0.75;
                ws_New1.PageSetup.BottomMargin = 1;
                ws_New1.PageSetup.CenterHorizontally = true;
                //Шапка таблицы
                Appl.Range rangeHeader1 = ws_New1.get_Range("A1", "E1").Cells;
                rangeHeader1.Font.Name = "Times New Roman";
                rangeHeader1.Font.Size = 10;
                rangeHeader1.HorizontalAlignment = Appl.XlHAlign.xlHAlignCenter;
                rangeHeader1.VerticalAlignment = Appl.XlVAlign.xlVAlignCenter;
                rangeHeader1.Borders.LineStyle = Appl.XlLineStyle.xlContinuous;
                rangeHeader1.Borders.Weight = Appl.XlBorderWeight.xlThin;
                ws_New1.Cells[1, 1] = "РИЦ";
                ws_New1.Cells[1, 2] = "ИНН";
                ws_New1.Cells[1, 3] = "УСН";
                ws_New1.Cells[1, 4] = "Год";

                DataTable dt = new DataTable();
                SqlDataAdapter ad = new SqlDataAdapter(sql, Connect.cnn);
                ad.Fill(dt);
                int i = 2;

                foreach (DataRow row in dt.Rows)
                {
                    ws_New1.Cells[i, 1] = row[0].ToString();
                    ws_New1.Cells[i, 2].NumberFormat = "@";
                    ws_New1.Cells[i, 2] = row[1].ToString();
                    ws_New1.Cells[i, 3] = row[2].ToString();
                    ws_New1.Cells[i, 4] = row[3].ToString();
                    i = i + 1;
                }
                i = i - 1;
                exApp_New1.Visible = true;
                wb_New1.SaveAs(filename);
                Marshal.ReleaseComObject(exApp_New1);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                exApp_New1 = null;
                wb_New1 = null;
                ws_New1 = null;
                GC.Collect();
            }
        }
    }
}
