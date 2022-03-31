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
    /// Путь к excel файлу.
    /// </summary>
    private static readonly string ExcelFilePath;

    /// <summary>
    /// Имя файла для сохранения записей.
    /// </summary>
    private const string SaveFileName = "sorted.txt";

    /// <summary>
    /// Имя excel файла.
    /// </summary>
    private const string FileName = "book.xlsx";

    #endregion

    #region Методы

    /// <summary>
    /// Точка входа в приложение.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    public static void Main(string[] args)
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
    /// Записать товары в файл.
    /// </summary>
    /// <param name="records">Товары.</param>
    private static void WriteToFile(IEnumerable<PriceRecord> records)
    {
      using (var file = new StreamWriter(SaveFileName, false))
      {
        foreach (var record in records)
          file.WriteLine($"{record.Name}  {record.Price}");

        file.Flush();
      }
    }
    #endregion

    #region Конструкторы

    /// <summary>
    /// Конструктор.
    /// </summary>
    static Program()
    {
      ExcelFilePath = $@"{AppDomain.CurrentDomain.BaseDirectory}\{FileName}";
    }

    #endregion
  }
}
