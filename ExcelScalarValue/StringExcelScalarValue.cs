using NPOI.SS.UserModel;
using SimpleJSON;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
	public class StringExcelScalarValue : ExcelScalarValue
	{
		public StringExcelScalarValue()
		{
			Name = "string";
		}

		public override string ToJson(string[] args)
		{
			JSONArray array = new JSONArray();

			if (args != null && args.Length > 0)
			{
				foreach (string arg in args)
				{
					if (string.IsNullOrEmpty(arg))
					{
						array.Add("");
					}
					else
					{
						array.Add(arg);
					}
				}
			}

			return array.ToString();
		}

		public override string ToJson(string arg)
		{
			if (!string.IsNullOrEmpty(arg))
			{
				return $"\"{arg.Trim()}\"";
			}

			return "\"\"";
		}

		public override string ToJson(ICell cell)
		{
			if (cell != null)
			{
				if (cell.CellType == CellType.String)
				{
					if (!string.IsNullOrEmpty(cell.StringCellValue))
					{
						return $"\"{cell.StringCellValue}\"";
					}
				}
			}

			return "\"\"";
		}
	}

}
