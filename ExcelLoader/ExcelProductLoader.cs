using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelLoader
{
  /// <summary>
  /// Загрузчик товаров их Excel.
  /// </summary>
  internal sealed class ExcelProductLoader : IDisposable
  {
    #region Константы

    /// <summary>
    /// Номер колонки с именем товара.
    /// </summary>
    private const int NamePosition = 1;

    /// <summary>
    /// Номер колонки с ценой товара.
    /// </summary>
    private const int PricePosition = 2;

    #endregion

    #region Свойства и поля

    /// <summary>
    /// Экземпляр Excel.Application.
    /// </summary>
    private readonly Application excelApplication;

    #endregion

    #region Методы

    /// <summary>
    /// Прочитать данные о товарах из excel файла.
    /// </summary>
    /// <param name="filename">Имя файла.</param>
    /// <param name="sheet">Номер листа.</param>
    /// <returns>Коллекция записей PriceRecord.</returns>
    /// <exception cref="InvalidDataException">Количество заполненных колонок меньше номера колонки с ценой.</exception>
    public IEnumerable<Product> Load(string filename, int sheet = 1)
    {
      var workBook = this.excelApplication.Workbooks.Open(filename, 0, true);
      var workSheet = (Worksheet)workBook.Worksheets.Item[sheet];
      var result = new List<Product>();
      try
      {
        var range = workSheet.UsedRange;
        if (range.Columns.Count < PricePosition)
          throw new InvalidDataException($"Used range columns count is less than {PricePosition}");

        for (int i = 1; i <= range.Rows.Count; i++)
        {
          var name = ((Range)range.Cells[i, NamePosition]).Value2;
          var price = ((Range)range.Cells[i, PricePosition]).Value2 as double?;
          if (name == null || price == null)
            continue;

          result.Add(new Product(name.ToString(), price.Value));
        }
      }
      finally
      {
        Marshal.ReleaseComObject(workSheet);
        workBook.Close();
        Marshal.ReleaseComObject(workBook);
      }

      return result;
    }

    #endregion

    #region IDisposable

    /// <summary>
    /// Освобождение не управляемых ресурсов.
    /// </summary>
    public void Dispose()
    {
      this.excelApplication.Quit();
      Marshal.ReleaseComObject(this.excelApplication);
    }

    #endregion

    #region Конструкторы

    /// <summary>
    /// Конструктор.
    /// </summary>
    public ExcelProductLoader()
    {
      this.excelApplication = new Application();
    }

    #endregion
  }
}
