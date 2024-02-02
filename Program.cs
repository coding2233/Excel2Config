// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");

using Excel2Config;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

class Program
{
    const string Version = "0.1.0";
    static void Main(string[] args)
    {
        string excelPath = string.Empty;
        string outputDir = string.Empty;
        string protocPath = "protoc";
        string shellPath = "sh";
        string protocCmd = string.Empty;
		bool recursive = false;
		bool toJson = false;
        bool toProto = false;
        bool toTextProto = false;
        bool toBinaryProto = false;
        for (int i = 0; i < args.Length; i++)
        {
            string argItem = args[i];
            if (string.IsNullOrEmpty(argItem))
            {
                continue;
            }

            if (argItem.Equals("--version"))
            {
                Console.WriteLine(Version);
            }
            else if (argItem.StartsWith("--excel_path="))
            {
                excelPath = argItem.Split('=')[1];
            }
            else if (argItem.Equals("-R") || argItem.Equals("--recursive"))
            {
                recursive = true;
            }
            else if (argItem.StartsWith("--output_path="))
            {
                outputDir = argItem.Split('=')[1];
            }
            else if ("--to_json".Equals(args[i]))
            {
                toJson = true;
            }
            else if (argItem.StartsWith("--to_protobuf="))
            {
                string toProtobufArg = argItem.Split('=')[1];
                var toProtobufArgs = toProtobufArg.Split('|');
                foreach (var itemArg in toProtobufArgs)
                {
					if ("proto".Equals(itemArg))
					{
						toProto = true;
					}
					else if ("textproto".Equals(itemArg))
					{
						toTextProto = true;
					}
					//all
					else
					{
						toProto = true;
						toTextProto = true;
						toBinaryProto = true;
					}
				}
            }
            else if (argItem.StartsWith("--protoc="))
            {
				protocPath = argItem.Split("=")[1];
			}
			else if (argItem.StartsWith("--shell="))
			{
				shellPath = argItem.Split("=")[1];
			}
			else if (argItem.StartsWith("--protoc_cmd="))
			{
				toProto = true;
                int cmdIndex = argItem.IndexOf('=') + 1;
                protocCmd = argItem.Substring(cmdIndex, argItem.Length - cmdIndex);
			}
			else if (argItem.Equals("--help"))
			{
                StringBuilder helpBuilder = new StringBuilder();
                AddHelpCommand(helpBuilder,"--help", "Show this text.");
				AddHelpCommand(helpBuilder, "--version", $"Show version info. {Version}.");
				AddHelpCommand(helpBuilder, "--excel_path=", "The path to the excel file or folder.");
				AddHelpCommand(helpBuilder, "--recursive,-R", "Traverse all the subfolders of the excel folder.");
				AddHelpCommand(helpBuilder, "--output_path=", "Setting the output directory. If it is not set, it is the folder path of excel.");
				AddHelpCommand(helpBuilder, "--to_json", "Convert to a json configuration file.");
				AddHelpCommand(helpBuilder, "--to_protobuf=", "Convert to a protobuf configuration file. Input parameter proto|textproto|binaryproto|all, all is recommended.");
				AddHelpCommand(helpBuilder, "--protoc=", "Set the path to the protoc execution file.Environment variables are used by default protoc.");
				AddHelpCommand(helpBuilder, "--shell=", "Set the path to the shell execution file.Environment variables are used by default sh.");
				AddHelpCommand(helpBuilder, "--protoc_cmd=", "By default, the output file path of proto is set, and other protoc commands that need to be executed are added.");
				Console.WriteLine(helpBuilder.ToString());
			}
		}

        if (!string.IsNullOrEmpty(excelPath))
        {
            new ExcelProgram(excelPath, outputDir, protocPath, recursive, toJson, toProto, toTextProto, toBinaryProto, shellPath, protocCmd);
		}
    }

    static void AddHelpCommand(StringBuilder stringBuilder, string command,string desc)
    {
		stringBuilder.Append(command);
        stringBuilder.Append(' ', 20 - command.Length);
		stringBuilder.AppendLine(desc);
	}
}

