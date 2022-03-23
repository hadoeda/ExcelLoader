using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLoader
{
  /// <summary>
  /// Запись "Товар - Цена".
  /// </summary>
  internal sealed class PriceRecord
  {
    #region Поля и свойства

    /// <summary>
    /// Имя товара.
    /// </summary>
    public string Name { get; private set; }
    
    /// <summary>
    /// Цена товара.
    /// </summary>
    public double Price { get; private set; }

    #endregion

    #region Конструкторы
    
    /// <summary>
    /// Конструктор.
    /// </summary>
    /// <param name="name">Имя товара.</param>
    /// <param name="price">Цена товара.</param>
    public PriceRecord(string name, double price)
    {
      this.Name = name;
      this.Price = price;
    }

    #endregion
  }
}
