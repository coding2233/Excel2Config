using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Config
{
	public class ExcelProgram
	{
		public ExcelProgram(string excelPath, string outputDir, string protocPath, bool recursive, bool toJson, bool toProto, bool toTextProto, bool toBinaryProto,string shellPath,string protocCmd)
		{
			//查找excel文件
			List<string> excelPathList = new List<string>();
			if (File.Exists(excelPath))
			{
				excelPathList.Add(excelPath);
			}
			else if (Directory.Exists(excelPath))
			{
				excelPathList.AddRange(GetExcelFiles(excelPath, recursive));
			}

			if(excelPathList.Count == 0)
			{
				Console.WriteLine("Excel path no files or folders found.");
				return;
			}

			//输出目录
			try
			{
				if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
				{
					Directory.CreateDirectory(outputDir);
				}
			}
			catch(Exception ex)
			{
				Console.WriteLine($"CreateDirectory Exception: {ex.Message}");
			}

			//excel转换
			try
			{
				foreach (string file in excelPathList)
				{
					var excelTemplate = new ExcelTemplate(file, outputDir);
					if (toJson)
					{
						excelTemplate.WriteJson();
					}
					if (toProto)
					{
						excelTemplate.WriteProtobuf();
					}
					if (toTextProto)
					{
						excelTemplate.WriteTextProto();
					}
					if (toBinaryProto)
					{
						excelTemplate.WriteBinaryProto(protocPath, shellPath);
					}
					if (!string.IsNullOrEmpty(protocCmd))
					{
						excelTemplate.RunProtocCmd(protocPath, shellPath, protocCmd);
					}
				}
			}
			catch (Exception ex) 
			{
				Console.WriteLine($"ExcelTemplate To Conifg Exception: {ex.Message}");
			}


		}

		private List<string> GetExcelFiles(string dir, bool recursive)
		{
			List<string> files = new List<string>();
			try
			{
				foreach (var item in Directory.GetFiles(dir))
				{
					if (item.EndsWith(".xlsx") || item.EndsWith(".xls"))
					{
						files.Add(item);
					}
				}

				if (recursive)
				{
					foreach (var item in Directory.GetDirectories(dir))
					{
						var dirFiles = GetExcelFiles(item, recursive);
						files.AddRange(dirFiles);
					}
				}
			}
			catch (System.Exception e)
			{
				Console.WriteLine($"GetExcelFiles Exception :{e.Message}");
			}
			return files;
		}
	}
}
