using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelLoader
{
  class Program
  {
    #region Константы
    /// <summary>
    /// Путь к excel файлу
    /// </summary>
    private static readonly string ExcelFilePath;

    /// <summary>
    /// Имя файла для сохранения записей
    /// </summary>
    private const string SaveFileName = "sorted.txt";
    #endregion

    #region Методы
    static void Main(string[] args)
    {
      using (var loader = new ExcelPriceLoader()) 
      {
        var prices = loader.Load(ExcelFilePath);

        var filtered = prices.Where(p => p.Price > 2000)
        .OrderBy(p => p.Name);
        WriteToFile(filtered);
      }
    }
    
    /// <summary>
    /// Записывет записи в файл
    /// </summary>
    /// <param name="records">Записи</param>
    static void WriteToFile(IEnumerable<PriceRecord> records)
    {
      using(var stream = new FileStream(SaveFileName, FileMode.Create))
      using (var file = new StreamWriter(stream))
      {
        var savingRecords = new StringBuilder();
        foreach(var record in records)
        {
          savingRecords.Append($"{record.Name}  {record.Price}");
          savingRecords.AppendLine();
        }

        file.Write(savingRecords);
        file.Flush();
      }
    }
    #endregion

    #region Конструкторы
    static Program()
    {
      var executedPath = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
      ExcelFilePath = $@"{executedPath.Parent.Parent.Parent.FullName}\book.xlsx";
    }
    #endregion
  }
}
