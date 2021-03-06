﻿using System;
using System.IO;
using Microsoft.Office.Interop.Word;
namespace daniil
{
    class Program
    {

        /// <summary>
        /// номера раздела, ==0 - нет разделов
        /// </summary>
        static uint _sectionNumber = 0;
        /// <summary>
        /// номера рисунков, ==0 - нет картинок
        /// </summary>
        static uint _pictureNumber = 0;
        /// <summary>
        /// номера таблиц, ==0 - нет таблиц
        /// </summary>
        static uint _tableNumber = 0;
        static void Main(string[] args)
        {
            string sourcePath = @"C:\Users\vdser\Desktop\шаблон";//путь до исходного шаблона
            string distPath = @"C:\Users\vdser\Desktop\result.rtf";//путь до выходного файла
            string csvPath = @"C:\Users\vdser\Desktop\data.csv";//путь до csv файла для создания таблицы
            //список закладок
            string[] templateStringList =
                {
                "[*имя раздела*]",///0
                "[*имя рисунка*]",///1
                "[*ссылка на следующий рисунок*]",///2
                "[*ссылка на предыдущий рисунок*]",///3
                "[*ссылка на таблицу*]",///4
                "[*таблица первая*]",///5,
                "[*код*]"///6
                };
            var application = new Application();
            application.Visible = true;
            var document = application.Documents.Open(sourcePath, false);
            Paragraph prevParagraph = null;
            Object missing = System.Type.Missing;

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                for (int i = 0; i < templateStringList.Length; i++)
                {
                    if (paragraph.Range.Text.Contains(templateStringList[i]))
                    {
                        switch (i)
                        {
                            case 0:
                                {
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 15;
                                    paragraph.Format.SpaceAfter = 12;
                                    paragraph.Range.Font.Bold = 1;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    _sectionNumber++;
                                    string replaceString = _sectionNumber.ToString();
                                    paragraph.Range.Find.Execute(templateStringList[i],
                                  ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                   0, ref missing, replaceString, 2, ref missing, ref missing,
                                  ref missing, ref missing);
                                }
                                break;
                            case 1:
                                {
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 12;
                                    paragraph.Format.SpaceAfter = 12;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    if (prevParagraph != null)
                                    {
                                        prevParagraph.Format.SpaceBefore = 12;
                                        prevParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    }

                                    _pictureNumber++;
                                    string replaceString = "Рисунок " + _sectionNumber.ToString() + "." + _pictureNumber.ToString() + " -";

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 2:
                                {
                                    string replaceString = _sectionNumber.ToString() + "." + (_pictureNumber + 1).ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 3:
                                {
                                    string replaceString = _sectionNumber.ToString() + "." + _pictureNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 4:
                                {
                                    _tableNumber++;
                                    string replaceString = _sectionNumber.ToString() + "." + _tableNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 5:
                                {
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    var range = application.Selection.Range;
                                    range.HighlightColorIndex = 0;

                                    string[] listRows = System.IO.File.ReadAllText(csvPath).Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                                    string[] listTitle = listRows[0].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                                    var wordTable = document.Tables.Add(range, listRows.Length, listTitle.Length);



                                    for (var k = 0; k < listTitle.Length; k++)
                                    {
                                        wordTable.Cell(1, k + 1).Range.Text = listTitle[k].ToString();
                                    }
                                    for (var j = 1; j < listRows.Length; j++)
                                    {
                                        string[] listValues = listRows[j].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                                        for (var k = 0; k < listValues.Length; k++)
                                        {
                                            wordTable.Cell(j + 1, k + 1).Range.Text = listValues[k].ToString();
                                        }
                                    }
                                }
                                break;
                            case 6: {
                                    insertCode(application, paragraph, templateStringList[i]);
                                }
                                break;
                        }
                    }
                    prevParagraph = paragraph;
                }


                document.SaveAs2(distPath);
                System.Console.In.Read();
                // application.Quit();
            }
        }


        public static void insertCode(Application application, Paragraph paragraph, string template)
        {
            string code = @"C:\Users\vdser\Desktop\Program.cs";
            application.Selection.Find.Execute(template);
            var range = application.Selection.Range;
            range.HighlightColorIndex = 0;
            /**
             * Рисуем граница вокруг листинга
             */
            paragraph.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            paragraph.Range.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth050pt;
            range.Borders.OutsideColorIndex = WdColorIndex.wdBlack;

            var font = application.Selection.Font;
            font.Name = "Calibri";
            paragraph.Range.Font.Size = 8;

            range.Text = System.IO.File.ReadAllText(code);
        }
    }
}
