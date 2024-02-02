using NPOI.SS.UserModel;
using SimpleJSON;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
    public class ExcelVarType
	{
		private static ExcelScalarValueTypes s_scalarValueTypes = new ExcelScalarValueTypes();

		private string m_excelType;
		public string ProtobufType { get; private set; }
		public bool IsBaseType { get; private set; }
		public bool IsMap { get; private set; }
		public string MapTypeKey { get; private set; }
		public string MapTypeValue { get; private set; }
		public bool IsList { get; private set; }
		public string MessageType { get; private set; }
		private char m_splitChar;

		public ExcelVarType(string excelType)
		{
			m_excelType = excelType;
			ParseExcelType();
		}

		private void ParseExcelType()
		{
			try
			{
				ProtobufType = m_excelType;
				m_splitChar = ',';
				if (m_excelType.Contains("#"))
				{
					var typeArgs = m_excelType.Split('#');
					switch (typeArgs[0])
					{
						case "map":
							var mapArgs = typeArgs[1].Split(':');
							ProtobufType = $"map<{mapArgs[0]},{mapArgs[1]}>";
							MapTypeKey = mapArgs[0];
							MapTypeValue = mapArgs[1];
							bool keyIsBaseType = s_scalarValueTypes.IsBaseType(MapTypeKey);
							bool valueIsBaseType = s_scalarValueTypes.IsBaseType(MapTypeValue);
							IsBaseType = keyIsBaseType && valueIsBaseType;
							IsMap = true;
							break;
						case "list":
							ProtobufType = $"repeated {typeArgs[1]}";
							break;
					}

					if (typeArgs.Length > 3)
					{
						var split = typeArgs[2].Split('=');
						m_splitChar = split[1].ToCharArray()[0];
					}
				}

				IsList = ProtobufType.Contains("repeated");
				MessageType = ProtobufType.Replace("repeated ", "");
				if (!IsBaseType)
				{
					IsBaseType = s_scalarValueTypes.IsBaseType(MessageType);
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e.ToString());
			}
		}

		public string ToJson(ICell cell)
		{
			if (cell != null)
			{
				if (cell.CellType == CellType.String && !string.IsNullOrEmpty(cell.StringCellValue))
				{
					var args = cell.StringCellValue.Split(m_splitChar);

					if (IsList)
					{
						return s_scalarValueTypes.ToJson(MessageType, args);
					}
					else if (IsMap)
					{
						StringBuilder mapJsoBuilder = new StringBuilder();
						mapJsoBuilder.Append("{");
						for (int i = 0; i < args.Length; i++)
						{
							var mapArgs = args[i].Split(':');
							mapJsoBuilder.Append(s_scalarValueTypes.ToJson(MapTypeKey, mapArgs[0]));
							mapJsoBuilder.Append(":");
							mapJsoBuilder.Append(s_scalarValueTypes.ToJson(MapTypeValue, mapArgs[1]));
							if (i < args.Length - 1)
							{
								mapJsoBuilder.Append(",");
							}
						}
						mapJsoBuilder.Append("}");
						return mapJsoBuilder.ToString();
					}
				}
			}
			return s_scalarValueTypes.ToJson(MessageType, cell);
		}

		public string ToTextProto(ICell cell,string varName)
		{
			if (IsMap)
			{
				if (cell.CellType == CellType.String 
					&& !string.IsNullOrEmpty(cell.StringCellValue))
				{
					var args = cell.StringCellValue.Split(m_splitChar);
					StringBuilder mapJsoBuilder = new StringBuilder();
					for (int i = 0; i < args.Length; i++)
					{
						var mapArgs = args[i].Split(':');

						mapJsoBuilder.Append(varName);
						mapJsoBuilder.Append(":{");
						mapJsoBuilder.Append("key:");
						mapJsoBuilder.Append(s_scalarValueTypes.ToJson(MapTypeKey, mapArgs[0]));
						mapJsoBuilder.Append(",value:");
						mapJsoBuilder.Append(s_scalarValueTypes.ToJson(MapTypeValue, mapArgs[1]));
						mapJsoBuilder.Append("}");
						if (i < args.Length - 1)
						{
							mapJsoBuilder.Append(",");
						}
					}
					return mapJsoBuilder.ToString();
				}
			}

			return ToJson(cell);
		}
	}
}
