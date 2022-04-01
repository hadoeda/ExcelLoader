namespace ExcelLoader
{
  /// <summary>
  /// Товар.
  /// </summary>
  internal sealed class Product
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
    public Product(string name, double price)
    {
      this.Name = name;
      this.Price = price;
    }

    #endregion
  }
}
