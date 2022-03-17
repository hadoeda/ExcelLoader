using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLoader
{
  /// <summary>
  /// Записть "Товар - Цена"
  /// </summary>
  internal sealed class PriceRecord
  {
    /// <summary>
    /// Имя товара
    /// </summary>
    public string Name { get; private set; }
    
    /// <summary>
    /// Цена товара
    /// </summary>
    public double Price { get; private set; }

    public PriceRecord(string name, double price)
    {
      this.Name = name;
      this.Price = price;
    }
  }
}
