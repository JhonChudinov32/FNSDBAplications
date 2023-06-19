using FNSDBAplications.connection;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Windows;
using word = Microsoft.Office.Interop.Word;

namespace FNSDBAplications.microsoft
{
    public class ImportWord
    {
        //Универсальная выгрузка в ворд
        public static void ExportWord(string filename, string query)
        {
            Connect.cnn.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter ad = new SqlDataAdapter(query, Connect.cnn);
            ad.Fill(dt);
            word.Document oDoc = new word.Document();
            oDoc.Application.Visible = false;
            try
            {
                if (dt.Rows.Count != 0)
                {
                    int RowCount = dt.Rows.Count;
                    int ColumnCount = dt.Columns.Count;
                    Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                    //Добавление строк и ячеек
                    int r = 0;
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        for (r = 0; r <= RowCount - 1; r++)
                        {
                            DataArray[r, c] = dt.Rows[r].ItemArray[c].ToString();
                        }
                    }

                    //Ориентация листа
                    oDoc.PageSetup.Orientation = word.WdOrientation.wdOrientLandscape;


                    dynamic oRange = oDoc.Content.Application.Selection.Range;
             
                    ArrayList data = new ArrayList();
                    for (r = 0; r < RowCount; r++)
                        for (int c = 0; c < ColumnCount; c++)
                            data.Add(DataArray[r, c]);
                    oRange.Text = string.Join("\t", data.ToArray());

                    //Формат таблицы
                   // oRange.Text = oTemp;
                    object oMissing = Missing.Value;
                    object Separator = word.WdTableFieldSeparator.wdSeparateByTabs;
                    object ApplyBorders = true;
                    object AutoFit = true;
                    object AutoFitBehavior = word.WdAutoFitBehavior.wdAutoFitContent;



                    oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                    oRange.Select();

                    oDoc.Application.Selection.Tables[1].Select();
                    oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                    oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                    oDoc.Application.Selection.Tables[1].Rows[1].Select();
                    oDoc.Application.Selection.InsertRowsAbove(1);
                    oDoc.Application.Selection.Tables[1].Rows[1].Select();

                    //Стиль заголовка таблицы
                    oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 2;
                    oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                    oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                    //add header row manually
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = dt.Columns[c].ToString();
                    }

                    //Стили таблицы
                    oDoc.Application.Selection.Tables[1].Rows[1].Select();
                    oDoc.Application.Selection.Cells.VerticalAlignment = word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    oDoc.Application.Selection.Tables[1].Borders.Enable = 1;

                    //Текст шапки
                    foreach (word.Section section in oDoc.Application.ActiveDocument.Sections)
                    {
                        word.Range headerRange = section.Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Fields.Add(headerRange, word.WdFieldType.wdFieldPage);
                        headerRange.Text = "Выгрузка";
                        headerRange.Font.Size = 16;
                        headerRange.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    //Сохранение файла
                    oDoc.SaveAs(filename, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                Connect.cnn.Close();
                oDoc.Application.Visible = true;
            }
        }
    }
}