using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
	public interface IExcelTemplate
	{
	}

	public class ExcelEnmuTemplate : IExcelTemplate
	{
		public string Name = string.Empty;
		public string Desc = string.Empty;
		public List<string> VarList = new List<string>();
		public List<string> VarDescList = new List<string>();
		public List<int> VarValueList = new List<int>();

		private ExcelVarType m_excelVarType;
		public ExcelEnmuTemplate()
		{
			m_excelVarType = new ExcelVarType("int32");
		}

		public string ToProtobuf()
		{
			StringBuilder stringBuilder = new StringBuilder();
			if (!string.IsNullOrEmpty(Desc))
			{
				stringBuilder.Append("// ");
				stringBuilder.AppendLine(Desc);
			}
			stringBuilder.Append("enum ");
			stringBuilder.AppendLine(Name);
			stringBuilder.AppendLine("{");
			for (int i = 0; i < VarList.Count; i++)
			{
				string varDesc = string.Empty;
				if (i < VarDescList.Count)
				{
					varDesc = VarDescList[i];
				}
				if (!string.IsNullOrEmpty(varDesc))
				{
					stringBuilder.Append("\t");
					stringBuilder.Append("// ");
					stringBuilder.AppendLine(varDesc);
				}
				stringBuilder.Append("\t");
				stringBuilder.Append(VarList[i]);
				stringBuilder.Append(" = ");
				stringBuilder.Append(VarValueList[i]);
				stringBuilder.AppendLine(";");
			}
			stringBuilder.AppendLine("}");

			return stringBuilder.ToString();
		}

		public string ToJson(ICell cell)
		{
			return m_excelVarType.ToJson(cell);
		}
	}

	public class ExcelMessageTemplate : IExcelTemplate
	{
		public string Name = string.Empty;
		public List<string> VarList = new List<string>();
		public List<ExcelVarType> VarTypeList = new List<ExcelVarType>();
		public List<string> VarDescList = new List<string>();
		public string ConfigName = string.Empty;

		public ISheet Sheet;
		public int Row;
		public int Column;

		public string ToProtobuf()
		{
			StringBuilder stringBuilder = new StringBuilder();
			stringBuilder.Append("message ");
			stringBuilder.AppendLine(Name);
			stringBuilder.AppendLine("{");
			for (int i = 0; i < VarList.Count; i++)
			{
				string varDesc = string.Empty;
				if (i < VarDescList.Count)
				{
					varDesc = VarDescList[i];
				}
				if (!string.IsNullOrEmpty(varDesc))
				{
					stringBuilder.Append("\t");
					stringBuilder.Append("// ");
					stringBuilder.AppendLine(varDesc);
				}
				stringBuilder.Append("\t");
				stringBuilder.Append(VarTypeList[i].ProtobufType);
				stringBuilder.Append(" ");
				stringBuilder.Append(VarList[i]);
				stringBuilder.Append(" = ");
				stringBuilder.Append(i + 1);
				stringBuilder.AppendLine(";");
			}
			stringBuilder.AppendLine("}");
			return stringBuilder.ToString();
		}

		public string ToJson(Func<string, IExcelTemplate> findMessage, bool readList = false)
		{
			List<string> messageJsonList = new List<string>();
			try
			{
				bool readWhile = readList;
				int rowIndex = 4;
				do
				{
					StringBuilder messageJson = new StringBuilder();
					messageJson.Append("{");
					bool readError = false;
					for (int i = 0; i < VarList.Count; i++)
					{
						ExcelVarType varType = VarTypeList[i];
						string jsonNode = string.Empty;

						IRow cellRow = null;
						ICell cell = null;
						cellRow = Sheet.GetRow(Row + rowIndex);
						if (cellRow != null)
						{
							cell = cellRow.GetCell(i + 1);
						}

						if (varType.IsBaseType)
						{
							if (cellRow == null)
							{
								Console.WriteLine("[ToJson] BaseType get row is null, But it could be normal.");
								readError = true;
								break;
							}
							jsonNode = varType.ToJson(cell);
						}
						else
						{
							var messageTemplate = findMessage(varType.MessageType);
							if (messageTemplate == null)
							{
								Console.WriteLine("Can't find message type: " + varType.MessageType);
								readError = true;
								break;
							}
							else
							{
								if (messageTemplate is ExcelMessageTemplate)
								{
									jsonNode = (messageTemplate as ExcelMessageTemplate).ToJson(findMessage, varType.IsList);
								}
								else if (messageTemplate is ExcelEnmuTemplate)
								{
									jsonNode = (messageTemplate as ExcelEnmuTemplate).ToJson(cell);
								}
							}
						}
						if (string.IsNullOrEmpty(jsonNode))
						{
							Console.WriteLine("To jsonNode is null.");
							readError = true;
							break;
						}

						messageJson.Append("\"");
						messageJson.Append(VarList[i]);
						messageJson.Append("\"");
						messageJson.Append(":");
						messageJson.Append(jsonNode);
						if (i < VarList.Count - 1)
						{
							messageJson.Append(",");
						}
					}

					messageJson.Append("}");

					if (readError)
					{
						readWhile = false;
					}
					else
					{
						messageJsonList.Add(messageJson.ToString());
					}
					rowIndex++;
				} while (readWhile);


				if (readList)
				{
					StringBuilder listBuilder = new StringBuilder();
					listBuilder.Append("[");
					for (int i = 0; i < messageJsonList.Count; i++)
					{
						listBuilder.Append(messageJsonList[i]);
						if (i < messageJsonList.Count - 1)
						{
							listBuilder.Append(",");
						}
					}
					listBuilder.Append("]");
					return listBuilder.ToString();
				}
				else
				{
					if (messageJsonList != null && messageJsonList.Count > 0)
					{
						return messageJsonList[0];
					}
				}
			}
			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);
			}

			return null;
		}

		public string ToTextProto(Func<string, IExcelTemplate> findMessage, bool readList = false)
		{
			List<string> messageList = new List<string>();
			try
			{
				bool readWhile = readList;
				int rowIndex = 4;
				do
				{
					StringBuilder messageBuilder = new StringBuilder();
					messageBuilder.Append("{");
					bool readError = false;
					for (int i = 0; i < VarList.Count; i++)
					{
						ExcelVarType varType = VarTypeList[i];
						string jsonNode = string.Empty;

						IRow cellRow = null;
						ICell cell = null;
						cellRow = Sheet.GetRow(Row + rowIndex);
						if (cellRow != null)
						{
							cell = cellRow.GetCell(i + 1);
						}

						if (varType.IsBaseType)
						{
							if (cellRow == null)
							{
								Console.WriteLine("[ToTextProto] BaseType get row is null, But it could be normal.");
								readError = true;
								break;
							}
							jsonNode = varType.ToTextProto(cell, VarList[i]);
						}
						else
						{
							var messageTemplate = findMessage(varType.MessageType);
							if (messageTemplate == null)
							{
								Console.WriteLine("Can't find message type: " + varType.MessageType);
								readError = true;
								break;
							}
							else
							{
								if (messageTemplate is ExcelMessageTemplate)
								{
									jsonNode = (messageTemplate as ExcelMessageTemplate).ToTextProto(findMessage, varType.IsList);
								}
								else if (messageTemplate is ExcelEnmuTemplate)
								{
									jsonNode = (messageTemplate as ExcelEnmuTemplate).ToJson(cell);
								}
							}
						}
						if (string.IsNullOrEmpty(jsonNode))
						{
							Console.WriteLine("To jsonNode is null.");
							readError = true;
							break;
						}
						if (!varType.IsMap)
						{
							messageBuilder.Append(VarList[i]);
							messageBuilder.Append(":");
						}
						messageBuilder.Append(jsonNode);
						if (i < VarList.Count - 1)
						{
							messageBuilder.Append(",");
						}
					}

					messageBuilder.Append("}");

					if (readError)
					{
						readWhile = false;
					}
					else
					{
						messageList.Add(messageBuilder.ToString());
					}
					rowIndex++;
				} while (readWhile);


				if (readList)
				{
					StringBuilder listBuilder = new StringBuilder();
					listBuilder.Append("[");
					for (int i = 0; i < messageList.Count; i++)
					{
						listBuilder.Append(messageList[i]);
						if (i < messageList.Count - 1)
						{
							listBuilder.Append(",");
						}
					}
					listBuilder.Append("]");
					return listBuilder.ToString();
				}
				else
				{
					if (messageList != null && messageList.Count > 0)
					{
						return messageList[0];
					}
				}
			}
			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);
			}

			return null;
		}

	}

}
