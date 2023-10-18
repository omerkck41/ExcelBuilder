using OfficeOpenXml;

namespace ExcelBuilder
{
	public class ExcelBuilder
	{
		private readonly ExcelPackage _package;
		private ExcelWorksheet? _worksheet;
		public enum InsertPosition { Left, Right }


		public ExcelBuilder(string worksheetName, string? templatePath = null)
		{
			if(string.IsNullOrEmpty(worksheetName)) 
				throw new ArgumentException($"{nameof(worksheetName)} cannot be empty or null.");
            
			ExcelPackage.LicenseContext = LicenseContext.Commercial;

            _package = templatePath == null ? new ExcelPackage() : new ExcelPackage(new FileInfo(templatePath));

			// Zaten var olan bir çalışma sayfasını kontrol edin
			_worksheet = _package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase));

			// Eğer yoksa, yeni bir çalışma sayfası oluşturun
			_worksheet ??= _package.Workbook.Worksheets.Add(worksheetName);
		}
		public ExcelBuilder(string? templatePath = null)
		{
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            _package = templatePath == null ? new ExcelPackage() : new ExcelPackage(new FileInfo(templatePath));

			// Eğer worksheetName belirtilmemişse, ilk çalışma sayfası kullanılır. Eğer hiç çalışma sayfası yoksa, yeni bir tane oluşturulur.
			_worksheet = _package.Workbook.Worksheets.FirstOrDefault() ?? _package.Workbook.Worksheets.Add("Sheet1");
		}


		#region Rows Operation
		public ExcelBuilder SetRow(int rowNumber)
		{
			if(rowNumber <= 0) 
				throw new ArgumentException($"{nameof(rowNumber)} cannot be <= 0");

			_worksheet?.InsertRow(rowNumber + 1, 1);
			return this;
		}
		public ExcelBuilder SetCopyRow(int rowNumber, bool copyIsValues = true)
		{
			if(rowNumber <= 0) 
				throw new ArgumentException($"{nameof(rowNumber)} cannot be <= 0");

			_worksheet?.InsertRow(rowNumber + 1, 1);

			if (copyIsValues)
			{
				_worksheet?.Cells[rowNumber, 1, rowNumber, _worksheet.Dimension.End.Column]
					.Copy(_worksheet.Cells[rowNumber + 1, 1, rowNumber + 1, _worksheet.Dimension.End.Column]);
			}
			else
			{
				for (var col = 1; col <= _worksheet?.Dimension.End.Column; col++)
				{
					_worksheet.Cells[rowNumber + 1, col].StyleID = _worksheet.Cells[rowNumber, col].StyleID;
				}
			}

			return this;
		}
		public ExcelBuilder SetDeleteRow(int rowNumber)
		{
			if(rowNumber <= 0)
				throw new ArgumentException($"{nameof(rowNumber)} cannot be <= 0");

			_worksheet?.DeleteRow(rowNumber);
			return this;
		}
		#endregion

		#region Columns Operation
		public ExcelBuilder SetColumn(object cellOrColumn, InsertPosition insertTo = InsertPosition.Left)
		{
			int newColumn;

			if(string.IsNullOrEmpty(cellOrColumn.ToString()))
				throw new ArgumentException($"{nameof(cellOrColumn)} cannot be empty or null.");

			if (int.TryParse(cellOrColumn.ToString(), out int row))
			{
				//_ = insertTo == InsertPosition.Left ? 1 : _worksheet?.Dimension.End.Column + 1;
				_worksheet?.InsertColumn(row, 1);
			}
			else
			{
				var address = new ExcelAddress(cellOrColumn + "1");
				newColumn = insertTo == InsertPosition.Left ? address.Start.Column : address.Start.Column + 1;
				_worksheet?.InsertColumn(newColumn, 1);
				_ = _worksheet?.Cells[address.Start.Row, newColumn];
			}

			return this;
		}
		public ExcelBuilder SetCopyColumn(object cellOrColumn, InsertPosition insertTo, bool copyIsValues = true)
		{
			int columnToCopy;
			int newColumn;

			if(string.IsNullOrEmpty(cellOrColumn.ToString()))
				throw new ArgumentException($"{nameof(cellOrColumn)} cannot be empty or null.");

			if (int.TryParse(cellOrColumn.ToString(), out int columnIndex))
			{
				columnToCopy = columnIndex;
			}
			else
			{
				var address = new ExcelAddress(cellOrColumn + "1");
				columnToCopy = address.Start.Column;
			}

			newColumn = insertTo == InsertPosition.Left ? columnToCopy : columnToCopy + 1;
			_worksheet?.InsertColumn(newColumn, 1);

			if (copyIsValues)
			{
				_worksheet?.Cells[1, columnToCopy, _worksheet.Dimension.End.Row, columnToCopy]
					.Copy(_worksheet.Cells[1, newColumn, _worksheet.Dimension.End.Row, newColumn]);
			}
			else
			{
				for (var row = 1; row <= _worksheet?.Dimension.End.Row; row++)
				{
					_worksheet.Cells[row, newColumn].StyleID = _worksheet.Cells[row, columnToCopy].StyleID;
				}
			}

			return this;
		}
		public ExcelBuilder SetDeleteColumn(object cellOrColumn)
		{
			int columnToDelete;

			if(string.IsNullOrEmpty(cellOrColumn.ToString()))
				throw new ArgumentException($"{nameof(cellOrColumn)} cannot be empty or null.");

			if (int.TryParse(cellOrColumn.ToString(), out int columnIndex))
			{
				columnToDelete = columnIndex;
			}
			else
			{
				var address = new ExcelAddress(cellOrColumn + "1");
				columnToDelete = address.Start.Column;
			}

			_worksheet?.DeleteColumn(columnToDelete);

			return this;
		}
		#endregion

		#region WorkSheet Operation
		public ExcelBuilder CopyWorksheet(string sourceWorksheetName, string? targetWorksheetName)
		{
			if (string.IsNullOrEmpty(targetWorksheetName))
				throw new ArgumentException($"{nameof(targetWorksheetName)} cannot be empty or null.");

			var sourceWorksheet = _package.Workbook.Worksheets.FirstOrDefault(ws =>
				ws.Name.Equals(sourceWorksheetName, StringComparison.OrdinalIgnoreCase)) ?? throw new ArgumentException($"No worksheet named {nameof(sourceWorksheetName)} exists in the workbook.");

			_worksheet = _package.Workbook.Worksheets.Add(targetWorksheetName, sourceWorksheet);

			return this;
		}
		public ExcelBuilder DeleteWorksheet(string? worksheetName)
		{
			var worksheetToDelete = _package.Workbook.Worksheets.FirstOrDefault(ws =>
				ws.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase)) ?? throw new ArgumentException($"No worksheet named {nameof(worksheetName)} exists in the workbook.");

			_package.Workbook.Worksheets.Delete(worksheetToDelete);

			return this;
		}
		#endregion

		public ExcelBuilder ClearContent(string? range)
		{
			if (string.IsNullOrEmpty(range))
				throw new ArgumentException($"{nameof(range)} cannot be empty or null.");

			var cells = _worksheet?.Cells[range];

			cells?.Clear();

			return this;
		}

		public ExcelBuilder SetData(string cell, object value, string? numberFormat = null, string? formula = null)
		{
			if(string.IsNullOrEmpty(cell))
				throw new ArgumentException($"{nameof(cell)} cannot be empty or null.");

			var excelCell = _worksheet?.Cells[cell] ?? throw new ArgumentException($"{nameof(_worksheet)} cannot be empty or null.");
			excelCell.Value = value;

			if (numberFormat != null) excelCell.Style.Numberformat.Format = numberFormat;
			if (formula != null) excelCell.Formula = formula;

			return this;
		}
		public ExcelBuilder SetData(int row, int column, object value, string? numberFormat = null, string? formula = null)
		{
			if(row <=0 || column <= 0)
				throw new ArgumentException($"{nameof(row)} or {nameof(column)} cannot be <= 0");


			var excelCell = _worksheet?.Cells[row, column] ?? throw new ArgumentException($"{nameof(_worksheet)} cannot be empty or null.");
			excelCell.Value = value;

			if (numberFormat != null) excelCell.Style.Numberformat.Format = numberFormat;
			if (formula != null) excelCell.Formula = formula;

			return this;
		}
		private string ColumnIndexToLetter(int colIndex)
		{
			if (colIndex < 0) return "A";

			int dividend = colIndex;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}
			return columnName;
		}

		public ExcelBuilder SetDataList<T>(string startCell, List<T> data)
		{
			if(string.IsNullOrEmpty(startCell))
				throw new ArgumentException($"{nameof(startCell)} cannot be empty or null.");

			_worksheet?.Cells[startCell].LoadFromCollection(data);
			return this;
		}
		public ExcelBuilder SetDataList<T>(int startRow, int startColumn, List<T> data)
		{
			if(startRow <= 0  || startColumn <= 0)
				throw new ArgumentException($"{nameof(startRow)} or {nameof(startColumn)} cannot be <= 0");

			_worksheet?.Cells[startRow, startColumn].LoadFromCollection(data);
			return this;
		}


		public static void ExportToExcel<T>(string filePath, List<T> data)
		{
			if(string.IsNullOrEmpty(filePath))
				throw new ArgumentException($"{nameof(filePath)} cannot be empty or null.");
			if(data == null || data.Count == 0)
				throw new ArgumentException($"Data cannot be empty or null.");

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using ExcelPackage package = new();
			var worksheet = package.Workbook.Worksheets.Add("Sheet1");

			// Sütun başlıklarını yazdırın ve biçimlendirin
			var properties = typeof(T).GetProperties();
			for (int columnIndex = 1; columnIndex <= properties.Length; columnIndex++)
			{
				var property = properties[columnIndex - 1];
				worksheet.Cells[1, columnIndex].Value = property.Name;
				worksheet.Cells[1, columnIndex].Style.Font.Bold = true;
				worksheet.Cells[1, columnIndex].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
			}

			// Verileri hızlı bir şekilde çalışma sayfasına yükleyin
			worksheet.Cells["A2"].LoadFromCollection(data, false);

			// AutoFitColumns metodunu çağırarak sütunları otomatik genişliğe ayarlayın
			worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

			// Dosyayı kaydedin
			File.WriteAllBytes(filePath, package.GetAsByteArray());
		}
		public static List<T> ImportToEntity<T>(string filePath)
		{
			if (string.IsNullOrEmpty(filePath))
				throw new ArgumentException($"{nameof(filePath)} cannot be empty or null.");

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using var package = new ExcelPackage(new FileInfo(filePath));
			var worksheet = package.Workbook.Worksheets.First();

			var properties = typeof(T).GetProperties();
			var entities = new List<T>();

			// İlk satırı sütun başlıkları olarak alın
			var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
			var headerValues = headerRow.Select(cell => cell.Value.ToString()).ToList();

			// Veri satırlarını dolaşın
			for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
			{
				var entity = Activator.CreateInstance<T>();

				// Sütunları dolaşın ve her bir sütunun değerini ilgili özelliklere ata
				for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
				{
					var propertyName = headerValues[col - 1];
					var property = properties.FirstOrDefault(p => p.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));

					if (property != null)
					{
						var cellValue = worksheet.Cells[row, col].Value;

						if (property.PropertyType == typeof(DateTime) && cellValue is double numericValue)
						{
							var dateTimeValue = DateTime.FromOADate(numericValue);
							property.SetValue(entity, dateTimeValue);
						}
						else
						{
							var convertedValue = Convert.ChangeType(cellValue, property.PropertyType);
							property.SetValue(entity, convertedValue);
						}
					}
				}

				entities.Add(entity);
			}

			return entities;
		}
        /// <summary>
        /// Transfers by matching Excel Columns and Class Columns respectively. 		
        /// The first column of the class Id is considered the default and continues from the next property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static List<T> MapEntitiesFromExcel<T>(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException($"{nameof(filePath)} cannot be empty or null.");

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets.First();
            var entities = new List<T>();

            // Excel'deki sütun başlıklarını alın
            var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
            var headerValues = headerRow.Select(cell => cell.Text).ToList();

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var entity = Activator.CreateInstance<T>();
                var properties = typeof(T).GetProperties();

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    if (col > properties.Length)
                    {
                        // Sınıfın özellik sayısından fazla sütun varsa işlem yapmayın.
                        break;
                    }

                    var headerValue = headerValues[col - 1];
                    var property = properties[col];
                    var cellValue = worksheet.Cells[row, col].Value;

                    if (property.PropertyType == typeof(DateTime) && cellValue is double numericValue)
                    {
                        var dateTimeValue = DateTime.FromOADate(numericValue);
                        property.SetValue(entity, dateTimeValue);
                    }
                    else
                    {
                        var convertedValue = Convert.ChangeType(cellValue, property.PropertyType);
                        property.SetValue(entity, convertedValue);
                    }
                }

                entities.Add(entity);
            }

            return entities;
        }

        public void Build(string directory, string fileName, string fileExtension = ".xlsx")
		{
			if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(fileName))
				throw new ArgumentException($"{nameof(directory)} or {nameof(fileName)} cannot be empty or null.");

			var path = Path.Combine(directory, fileName + fileExtension);

			switch (fileExtension)
			{
				case ".xlsx":
					_package.SaveAs(new FileInfo(path));
					break;

				case ".csv":
					// For .csv, we need to convert the data and save it manually.
					using (var sw = new StreamWriter(path))
					{
						for (var i = 1; i <= _worksheet?.Dimension.Rows; i++)
						{
							for (var j = 1; j <= _worksheet.Dimension.Columns; j++)
							{
								sw.Write(_worksheet.Cells[i, j].Value);

								if (j != _worksheet.Dimension.Columns)
								{
									sw.Write(",");
								}
							}

							sw.WriteLine();
						}
					}
					break;

				default:
					_package.SaveAs(new FileInfo(path));
					break;
			}
		}
	}
}
