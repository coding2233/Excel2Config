using NPOI.SS.UserModel;
using SimpleJSON;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
	public class BoolExcelScalarValue : ExcelScalarValue
	{
		public BoolExcelScalarValue()
		{
			Name = "bool";
		}
		public override string ToJson(ICell cell)
		{
			bool result = false;
			if (cell != null)
			{
				if (cell.CellType == CellType.String)
				{
					if (!string.IsNullOrEmpty(cell.StringCellValue))
					{
						if (bool.TryParse(cell.StringCellValue, out bool boolValue))
						{
							result = boolValue;
						}
					}
				}
			}

			return result.ToString().ToLower();
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
						&& bool.TryParse(arg, out bool value))
					{
						data = new JSONData(value);
					}
					if (data == null)
					{
						data = new JSONData(false);
					}
					array.Add(data);
				}
			}

			return array.ToString().ToLower();
		}

		public override string ToJson(string arg)
		{
			bool value = false;
			bool.TryParse(arg, out value);
			return value.ToString().ToLower();
		}
	}

}
