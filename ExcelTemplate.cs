using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Crypto.Encodings;
using SimpleJSON;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Excel2Config
{
	public class ExcelTemplate
	{
		List<ExcelEnmuTemplate> m_excelEnmuList;
		List<ExcelMessageTemplate> m_excelMessageList;
		string m_packageName;

		XSSFWorkbook m_workBook;
		string m_excelPath;
		string m_outputDir;

		public ExcelTemplate(string excelPath,string outputDir)
		{
			m_excelEnmuList = new List<ExcelEnmuTemplate>();
			m_excelMessageList = new List<ExcelMessageTemplate>();

			try
			{
				if (!File.Exists(excelPath))
				{
					Console.WriteLine($"{excelPath} exists is false.");
				}	

				using (var stream = new FileStream(excelPath, FileMode.Open,FileAccess.Read,FileShare.ReadWrite))
				{
					stream.Position = 0;
					m_workBook = new XSSFWorkbook(stream);
					for (int i = 0; i < m_workBook.NumberOfSheets; i++)
					{
						var sheet = m_workBook.GetSheetAt(i);
						ReadSheet(sheet);
					}
				}

				m_excelPath = excelPath;
				if (string.IsNullOrEmpty(outputDir))
				{
					outputDir = Path.GetDirectoryName(excelPath);
				}
				m_outputDir = outputDir;
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}

		public ExcelTemplate WriteProtobuf()
		{
			try
			{
				string protobufPath = Path.Combine(m_outputDir, Path.GetFileNameWithoutExtension(m_excelPath) + ".proto");
				string protobuf = ToProtobuf();
				if (File.Exists(protobufPath))
				{
					File.Delete(protobufPath);
				}
				File.WriteAllText(protobufPath, protobuf);
				Console.WriteLine($"{protobufPath}  write success.");
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}
			return this;
		}

		public string ToProtobuf()
		{
			StringBuilder stringBuilder = new StringBuilder();

			stringBuilder.AppendLine("syntax = \"proto3\";\n");

			if (!string.IsNullOrEmpty(m_packageName))
			{
				stringBuilder.Append("package ");
				stringBuilder.Append(m_packageName);
				stringBuilder.AppendLine(";\n");
			}

			foreach (var item in m_excelEnmuList)
			{
				stringBuilder.AppendLine(item.ToProtobuf());
			}

			foreach (var item in m_excelMessageList)
			{
				stringBuilder.AppendLine(item.ToProtobuf());
			}

			return stringBuilder.ToString();
		}

		public ExcelTemplate WriteJson()
		{
			foreach (var item in m_excelMessageList)
			{
				if (!string.IsNullOrEmpty(item.ConfigName))
				{
					string jsonPath = Path.Combine(m_outputDir, item.ConfigName + ".json");
					if (File.Exists(jsonPath))
					{
						File.Delete(jsonPath);
					}
					string json = ToJson(item);
					File.WriteAllText(jsonPath, json);
					Console.WriteLine($"{jsonPath}  write success.");
				}
			}
			return this;
		}

		public string ToJson(ExcelMessageTemplate messageTemplate)
		{
			try
			{
				var jsonNode = messageTemplate.ToJson(FindExcelTemplate);
				return jsonNode;
			}
			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);
			}
			return string.Empty;
		}

		public ExcelTemplate WriteTextProto()
		{
			foreach (var item in m_excelMessageList)
			{
				if (!string.IsNullOrEmpty(item.ConfigName))
				{
					string textProtoPath = Path.Combine(m_outputDir, item.ConfigName + ".textproto");
					if (File.Exists(textProtoPath))
					{
						File.Delete(textProtoPath);
					}
					string textProto = ToTextProto(item);
					File.WriteAllText(textProtoPath, textProto);
					Console.WriteLine($"{textProtoPath}  write success.");
				}
			}
			return this;
		}

		public ExcelTemplate WriteBinaryProto(string protocPath,string shellPath)
		{
			try
			{
				Thread.Sleep(200);
				string protoFileName = Path.GetFileNameWithoutExtension(m_excelPath) + ".proto";
				foreach (var item in m_excelMessageList)
				{
					if (!string.IsNullOrEmpty(item.ConfigName))
					{
						protocPath = protocPath.Replace("\\", "/");
						Process.Start(protocPath, "--version").WaitForExit();
						string binaryProtoPath = Path.Combine(m_outputDir, item.ConfigName + ".bytes");
						if (File.Exists(binaryProtoPath))
						{
							File.Delete(binaryProtoPath);
						}
						string messageType = item.Name;
						if (!string.IsNullOrEmpty(m_packageName))
						{
							messageType = $"{m_packageName}.{messageType}";
						}
						string textProtoPath = Path.Combine(m_outputDir, item.ConfigName + ".textproto");
						Console.WriteLine($"textProtoPath is exists: {File.Exists(textProtoPath)}");


						string outDir = m_outputDir;// Path.GetFullPath(m_outputDir);
						string arguments = $"--proto_path=\"{outDir}/\" {protoFileName} --encode={messageType} < \"{Path.GetFullPath(textProtoPath)}\" > \"{Path.GetFullPath(binaryProtoPath)}\"";
						arguments = arguments.Replace("\\", "/");

						var shArguments = $"\"{protocPath}\" {arguments}";
						var p = Process.Start($"\"{shellPath}\" -c \"{shArguments}\"");
						p.WaitForExit();
						
						string exitCode = p.ExitCode == 0 ? "success" : "fail";
						Console.WriteLine($"[WriteBinaryProto] Process.Start {shellPath} -c {shArguments} \n-->{p.ExitCode} {exitCode}.");
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"WriteBinaryProto Exception:{ex.Message}");
			}
			
			return this;
		}


		public string ToTextProto(ExcelMessageTemplate messageTemplate)
		{
			try
			{
				var textProto = messageTemplate.ToTextProto(FindExcelTemplate);
				//去掉最外层的大括号
				textProto = textProto.Substring(1, textProto.Length - 2);
				return textProto;
			}
			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);
			}
			return string.Empty;
		}

		public ExcelTemplate RunProtocCmd(string protocPath, string shellPath, string cmd)
		{
			try
			{
				Thread.Sleep(100);
				string protoFileName = Path.GetFileNameWithoutExtension(m_excelPath) + ".proto";
				foreach (var item in m_excelMessageList)
				{
					if (!string.IsNullOrEmpty(item.ConfigName))
					{
						protocPath = protocPath.Replace("\\", "/");
						Process.Start(protocPath, "--version").WaitForExit();

						string outDir = m_outputDir;// Path.GetFullPath(m_outputDir);
						string arguments = $"--proto_path=\"{outDir}/\" {protoFileName} {cmd}";
						arguments = arguments.Replace("\\", "/");

						var shArguments = $"\"{protocPath}\" {arguments}";
						var p = Process.Start($"\"{shellPath}\" -c \"{shArguments}\"");
						p.WaitForExit();

						string exitCode = p.ExitCode == 0 ? "success" : "fail";
						Console.WriteLine($"[RunProtocCmd] Process.Start {shellPath} -c {shArguments} \n-->{p.ExitCode} {exitCode}.");
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"RunProtocCmd Exception:{ex.Message}");
			}

			return this;
		}


		private IExcelTemplate FindExcelTemplate(string messageType)
		{
			var messageTemp = m_excelMessageList.Find(x => x.Name.Equals(messageType));
			if (messageTemp != null)
			{
				return messageTemp;
			}

			var enmuTemp = m_excelEnmuList.Find(x => x.Name.Equals(messageType));
			if (enmuTemp != null)
			{
				return enmuTemp;
			}

			return null;
		}

		private void ReadSheet(ISheet sheet)
		{
			if (sheet == null)
			{
				return;
			}

			for (int i = 0; i < sheet.LastRowNum; i++)
			{
				var row = sheet.GetRow(i);
				if (row == null)
				{
					continue;
				}

				var cell = row.GetCell(0);
				if (cell != null)
				{
					if ("#message".Equals(cell.StringCellValue))
					{
						ParseMessageTemplate(sheet, i, 0);
					}
					else if ("#enmu".Equals(cell.StringCellValue))
					{
						ParseEnumTemplate(sheet, i, 0);
					}
					else if ("#package".Equals(cell.StringCellValue))
					{
						ParsePackageName(sheet, i, 0);
					}
				}
			}
		}

		private void ParseMessageTemplate(ISheet sheet,int row,int column)
		{
			try
			{
				ExcelMessageTemplate messageTemplate = new ExcelMessageTemplate();
				messageTemplate.Sheet = sheet;
				messageTemplate.Row = row;
				messageTemplate.Column = column;	

				var sheetRoow = sheet.GetRow(row);
				string name = sheetRoow.GetCell(1).StringCellValue;
				if (string.IsNullOrEmpty(name))
				{
					return;
				}
				messageTemplate.Name = name;
				var typeRow = sheet.GetRow(row+1);
				if (typeRow == null)
				{
					return;
				}

				var configRow = sheet.GetRow(row - 1);
				if (configRow != null)
				{
					var configTagCell = configRow.GetCell(0);
					if (configTagCell != null && configTagCell.CellType == CellType.String)
					{
						if (configTagCell.StringCellValue.Equals("#config"))
						{
							var configCell = configRow.GetCell(1);
							if (configCell != null && configCell.CellType == CellType.String)
							{
								messageTemplate.ConfigName = configCell.StringCellValue;
							}
						}
					}
				}

				column = 0;
				while (true)
				{
					var cell = typeRow.GetCell(column++);
					if (cell == null)
					{
						break;
					}
					string cellValue = cell.StringCellValue;
					if (string.IsNullOrEmpty(cellValue))
					{ break; }
					if (!cellValue.StartsWith("#"))
					{
						ExcelVarType excelVarType = new ExcelVarType(cellValue);
						messageTemplate.VarTypeList.Add(excelVarType);
					}
				}

				var varRow = sheet.GetRow(row+2);
				if (varRow == null)
				{
					return;
				}

				column = 0;
				while (true)
				{
					var cell = varRow.GetCell(column++);
					if (cell == null)
					{
						break;
					}
					string cellValue = cell.StringCellValue;
					if (string.IsNullOrEmpty(cellValue))
					{ break; }
					if (!cellValue.StartsWith("#"))
					{
						messageTemplate.VarList.Add(cellValue);
					}
				}

				var descRow = sheet.GetRow(row+2);
				if (descRow == null)
				{
					return;
				}
				column = 0;
				while (true)
				{
					var cell = descRow.GetCell(column++);
					if (cell == null)
					{
						break;
					}
					string cellValue = cell.SerializeToString();
					if (!cellValue.StartsWith("#"))
					{
						messageTemplate.VarDescList.Add(cellValue);
					}
				}

				if (messageTemplate.VarList.Count == 0)
				{
					return;
				}

				if (messageTemplate.VarList.Count != messageTemplate.VarTypeList.Count)
				{
					return;
				}

				m_excelMessageList.Add(messageTemplate);
			}
			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}

		private void ParseEnumTemplate(ISheet sheet, int row, int column)
		{
			try
			{
				ExcelEnmuTemplate enmuTemplate = new ExcelEnmuTemplate();
				var sheetRoow = sheet.GetRow(row);
				string name = sheetRoow.GetCell(1).StringCellValue;
				if (string.IsNullOrEmpty(name))
				{
					return;
				}
				enmuTemplate.Name = name;
				enmuTemplate.Desc = sheet.GetRow(row + 1).GetCell(1).StringCellValue;
				int varRowIndex = 2;
				int lastVarValue = -1;
				while (true)
				{
					var varRow = sheet.GetRow(row + varRowIndex);
					if (varRow == null)
					{
						break;
					}
					var varCell = varRow.GetCell(1);
					if (varCell == null)
					{
						break;
					}
					string varCellName = varCell.StringCellValue;
					if (string.IsNullOrEmpty(varCellName))
					{
						break;
					}
					if (!varCellName.StartsWith("#"))
					{
						var varDescCell = varRow.GetCell(2);
						var varValueCell = varRow.GetCell(3);
						int varValue = 0;
						if (varValueCell != null)
						{
							varValueCell.SetCellType(CellType.Numeric);
							varValue = (int)varValueCell.NumericCellValue;
						}
						if (varValue == 0)
						{
							varValue = lastVarValue + 1;
						}
						enmuTemplate.VarList.Add(varCellName);
						enmuTemplate.VarDescList.Add(varDescCell.SerializeToString());
						enmuTemplate.VarValueList.Add(varValue);
						lastVarValue = varValue;
					}
					varRowIndex++;
				}

				if (enmuTemplate.VarList.Count > 0)
				{
					m_excelEnmuList.Add(enmuTemplate);
				}
			}
			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}

		private void ParsePackageName(ISheet sheet, int row, int column)
		{
			try
			{
				var rowCell = sheet.GetRow(row);
				if (rowCell == null)
				{
					return;
				}
				var columnCell = rowCell.GetCell(1);
				if (columnCell == null)
				{
					return;
				}

				if (columnCell.CellType != CellType.String)
				{
					return;
				}

				m_packageName = columnCell.StringCellValue;
			}
			catch (System.Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}
	}
}
