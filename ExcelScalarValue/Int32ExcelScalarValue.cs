using NPOI.SS.UserModel;
using SimpleJSON;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
	public class Int32ExcelScalarValue : ExcelScalarValue
	{
		public Int32ExcelScalarValue()
		{
			Name = "int32";
		}
		public override string ToJson(ICell cell)
		{
			int result = 0;
			if (cell != null)
			{
				if (cell.CellType == CellType.Numeric)
				{
					result = (int)cell.NumericCellValue;
				}
				else if (cell.CellType == CellType.String)
				{
					if (!string.IsNullOrEmpty(cell.StringCellValue)
						&& int.TryParse(cell.StringCellValue, out int intValue))
					{
						result = intValue;
					}
				}
			}

			return result.ToString();
		}

		public override string ToJson(string[] args)
		{
			JSONArray array = new JSONArray();
			if (args != null && args.Length > 0)
			{
				foreach (string arg in args)
				{
					JSONData data = null;
					if (!string.IsNullOrEmpty(arg)
						&& int.TryParse(arg, out int value))
					{
						data = new JSONData(value);
					}
					if (data == null)
					{
						data = new JSONData(0);
					}
					array.Add(data);
				}
			}

			return array.ToString();
		}

		public override string ToJson(string arg)
		{
			int value = 0;
			int.TryParse(arg, out value);
			return value.ToString();
		}

	}

}
