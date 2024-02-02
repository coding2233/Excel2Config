using NPOI.SS.UserModel;
using SimpleJSON;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
	public class FloatExcelScalarValue : ExcelScalarValue
	{
		public FloatExcelScalarValue()
		{
			Name = "float";
		}

		public override string ToJson(ICell cell)
		{
			float result = 0.0f;
			if (cell != null)
			{
				if (cell.CellType == CellType.Numeric)
				{
					result = (float)cell.NumericCellValue;
				}
				else if (cell.CellType == CellType.String)
				{
					if (!string.IsNullOrEmpty(cell.StringCellValue)
						&& float.TryParse(cell.StringCellValue, out float floatValue))
					{
						result = floatValue;
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
					if (string.IsNullOrEmpty(arg)
						&& float.TryParse(arg,out float value))
					{
						data = new JSONData(value);
					}
					if (data == null)
					{
						data = new JSONData(0.0f);
					}
					array.Add(data);
				}
			}

			return array.ToString();
		}

		public override string ToJson(string arg)
		{
			float value = 0.0f;
			float.TryParse(arg, out value);
			return value.ToString();
		}
	}


}
