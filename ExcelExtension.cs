using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
	public static class ExcelExtension
	{

		public static string SerializeToString(this ICell cell)
		{
			if (cell == null)
			{
				return string.Empty;
			}

			switch (cell.CellType)
			{
				case CellType.Unknown:
					return cell.ToString();
				case CellType.Numeric:
					return cell.NumericCellValue.ToString();
				case CellType.String:
					return cell.StringCellValue;
				case CellType.Formula:
					return cell.CellFormula;
				case CellType.Blank:
					return "";
				case CellType.Boolean:
					return cell.BooleanCellValue.ToString().ToLower();
				case CellType.Error:
					return string.Empty;
			}

			return "";
		}

		//public static string ToProtobuf(this ICell cell,string defaultType = null)
		//{
		//	if (cell == null)
		//	{
		//		if (!string.IsNullOrEmpty(defaultType))
		//		{
		//			switch (defaultType)
		//			{
		//				case "double":
		//				case "float":
		//					return "0.0f";
		//				case "int32":
		//				case "int64":
		//				case "uint32":
		//				case "uint64":
		//				case "sint32":
		//				case "sint64":
		//				case "fixed32":
		//				case "fixed64":
		//				case "sfixed32":
		//				case "sfixed64":
		//				case "bool":
		//					return "0";
		//			}
		//		}
		//		return "";
		//	}

		//	switch (cell.CellType)
		//	{
		//		case CellType.Numeric:
		//			return cell.NumericCellValue.ToString();
		//		case CellType.String:
		//			string stringCellValue = cell.StringCellValue.Replace("\"", "\\\"").Replace("\\", "");
		//			return stringCellValue;
		//		case CellType.Formula:
		//			return cell.CellFormula.ToString();
		//		case CellType.Blank:
		//			return "";
		//		case CellType.Boolean:
		//			return cell.BooleanCellValue.ToString();
		//		default:
		//			return cell.ToString();
		//	}

		//	return "";
		//}
	}
}
