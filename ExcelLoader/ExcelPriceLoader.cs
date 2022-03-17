using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLoader
{
  internal sealed class ExcelPriceLoader : IDisposable
  {
    #region Константы
    /// <summary>
    /// Номер колонки с именем товара
    /// </summary>
    private const int NamePosition = 1;

    /// <summary>
    /// Номер колонки с ценой товара
    /// </summary>
    private const int PricePosition = 2;
    #endregion

    #region Поля
    private readonly Application excelApplication;
    #endregion

    #region Методы
    /// <summary>
    /// Читает данные из указанного excel файла
    /// </summary>
    /// <param name="filename">Имя файла</param>
    /// <param name="sheet">Номер листа</param>
    /// <returns>Коллекция записей PriceRecord</returns>
    public IEnumerable<PriceRecord> Load(string filename, int sheet = 1)
    {
      var workBook = this.excelApplication.Workbooks.Open(filename, 0, true);
      var workSheet = (Worksheet)workBook.Worksheets.Item[sheet];
      var result = new List<PriceRecord>();
      try
      {
        var range = workSheet.UsedRange;
        if (range.Columns.Count < PricePosition)
          throw new Exception("Wrong data in the file");

        for (int i = 1; i <= range.Rows.Count; i++)
        {
          var name = ((Range)range.Cells[i, NamePosition]).Value2 as string;
          var price = ((Range)range.Cells[i, PricePosition]).Value2 as double?;
          if (name == null || price == null) continue;

          result.Add(new PriceRecord(name, price.Value));
        }
      }
      catch(Exception e)
      {
        throw e;
      }
      finally
      {
        Marshal.ReleaseComObject(workSheet);
        workBook.Close();
        Marshal.ReleaseComObject(workBook);
      }

      return result;
    }

    /// <summary>
    /// Сохраняет коллекция в указанный файл
    /// </summary>
    /// <param name="filename">Имя файла</param>
    /// <param name="records">Коллекция записей PriceRecord</param>
    public void Save(string filename, IEnumerable<PriceRecord> records)
    {
      var workBook = this.excelApplication.Workbooks.Add();
      var workSheet = workBook.Worksheets.Item[1];
      try
      {
        int i = 1;
        foreach(var record in records)
        {
          workSheet.Cells[i, NamePosition] = record.Name;
          workSheet.Cells[i, PricePosition] = record.Price;
          i++;
        }

        workBook.SaveAs(filename);
      }
      catch (Exception e)
      {
        throw e;
      }
      finally
      {
        Marshal.ReleaseComObject(workSheet);
        workBook.Close();
        Marshal.ReleaseComObject(workBook);
      }
    }
    #endregion

    #region IDisposable
    public void Dispose()
    {
      this.excelApplication.Quit();
      Marshal.ReleaseComObject(this.excelApplication); 
    }
    #endregion

    #region Конструкторы
    public ExcelPriceLoader()
    {
      this.excelApplication = new Application();
    }
    #endregion
  }
}
