using System;
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
            string sourcePath = @"‪C:\Users\vdser\Desktop\шаблон.rtf";//путь до исходного шаблона
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
                "[*таблица первая*]"///5
                };

            Console.WriteLine("Hello World!");
        }
    }
}
