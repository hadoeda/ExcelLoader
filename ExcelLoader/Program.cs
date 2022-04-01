using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelLoader
{
  /// <summary>
  /// Основной класс приложения.
  /// </summary>
  public class Program
  {
    #region Константы
    /// <summary>
    /// Имя файла для сохранения записей.
    /// </summary>
    private const string OutFileName = "sorted.txt";

    /// <summary>
    /// Имя excel файла.
    /// </summary>
    private const string ExcelFileName = "book.xlsx";

    #endregion

    #region Методы

    /// <summary>
    /// Точка входа в приложение.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    public static void Main(string[] args)
    {
      var excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExcelFileName);
      var outFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, OutFileName);

      using (var loader = new ExcelProductLoader())
      {
        var prices = loader.Load(excelFilePath);

        var filtered = prices.Where(p => p.Price > 2000)
          .OrderBy(p => p.Name);
        WriteToFile(filtered, outFilePath);
      }
    }

    /// <summary>
    /// Записать товары в файл.
    /// </summary>
    /// <param name="records">Товары.</param>
    /// <param name="filePath">Путь файла.</param>
    private static void WriteToFile(IEnumerable<Product> records, string filePath)
    {
      using (var file = new StreamWriter(filePath, false))
      {
        foreach (var record in records)
          file.WriteLine($"{record.Name}  {record.Price}");

        file.Flush();
      }
    }
    #endregion
  }
}
