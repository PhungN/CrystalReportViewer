using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace GmsReportViewer
{
    public class ReportParser
    {
        public ReportParser(string scriptFile)
        {
            Data = new List<TemplateData>();
            using (FileStream fs = new FileStream(scriptFile, System.IO.FileMode.Open))
            {
                using (System.IO.StreamReader sr = new System.IO.StreamReader(fs))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        TemplateData td = new TemplateData(line);
                        if (td.Command != TemplateCommand.Unknown)
                        {
                            if (td.Command == TemplateCommand.SetSQLQuery)
                                Data.Insert(1, td);
                            else
                                Data.Add(td);
                        }
                    }
                }
            }
        }

        public List<TemplateData> Data { get; private set; }

        public static string HexToAsciiString(string hexString) => HexToString(hexString, Encoding.ASCII.GetString);

        public static string HexToUnicodeString(string hexString) => HexToString(hexString, Encoding.Unicode.GetString);

        private static string HexToString(string hexString, Func<byte[], string> ToString)
        {
            hexString = hexString.Trim().Remove(0, 2);  // Remove the 0x
            var bytes = new byte[hexString.Length / 2];
            for (var i = 0; i < bytes.Length; i++)
            {
                bytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            }
            return ToString(bytes).Trim();
        }

        public static List<TemplateData> GetTemplateData(string scriptFile) => new ReportParser(scriptFile).Data;
    }

    public enum TemplateCommand { Unknown, SetJob, SetNthTableLocation, SetTableLocation, MapNthTable, SetNthTableLogonInfo,
        LogonServer, SetSQLQuery, SetSelectionFormula, SetSubReportSelectionFormula, SetNthParameterField, SetNthSortField,
        DeleteNthSortField, SetNthGroupSortField, DeleteNthGroupSortField, SetFormula, SetReportTitle, OutputToWindow
    }

    public class CRParamFields
    {
        public int ParmNumber { get; private set; } = 0;
        public string CurrentValue { get; private set; } = "";
        public bool CurrentValueSet { get; private set; } = false;
        public string DefaultValue { get; private set; } = "";
        public bool DefaultValueSet { get; private set; } = false;
        public int Direction { get; private set; } = 0;
        public string EditMask { get; private set; } = "";
        public bool IsLimited { get; private set; } = false;
        public double MaxSize { get; private set; } = 0;
        public double MinSize { get; private set; } = 0;
        public string Name { get; private set; } = "";
        public bool NeedsCurrentValue { get; private set; } = false;
        public string Prompt { get; private set; } = "";
        public string ReportName { get; private set; } = "";

        public static CRParamFields FromStrings(string[] lstFields)
        {
            CRParamFields paramFields = new CRParamFields();
            if (lstFields.Length >= 2)
            {
                try
                {
                    paramFields.ParmNumber = int.Parse(lstFields[0]) - 1;
                    string[] fields = lstFields[1].Split(';');
                    if (fields.Length >= 14)
                    {
                        paramFields.CurrentValue = ReportParser.HexToUnicodeString(fields[0]);
                        paramFields.CurrentValueSet = int.Parse(fields[1]) > 0 ? true : false;
                        paramFields.DefaultValue = ReportParser.HexToUnicodeString(fields[2]);
                        paramFields.DefaultValueSet = int.Parse(fields[3]) > 0 ? true : false;
                        paramFields.Direction = int.Parse(fields[4]);
                        paramFields.EditMask = ReportParser.HexToUnicodeString(fields[5]);
                        paramFields.IsLimited = int.Parse(fields[6]) > 0 ? true : false;
                        paramFields.MaxSize = double.Parse(fields[7]);
                        paramFields.MinSize = double.Parse(fields[8]);
                        paramFields.Name = ReportParser.HexToUnicodeString(fields[9]);
                        paramFields.NeedsCurrentValue = int.Parse(fields[10]) > 0 ? true : false;
                        paramFields.Prompt = ReportParser.HexToUnicodeString(fields[11]);
                        paramFields.ReportName = ReportParser.HexToUnicodeString(fields[12]);
                    }
                }
                catch (ArgumentNullException) { }
                catch(FormatException) { }
                catch(OverflowException) { }
            }
            
            return paramFields;
        }
    }

    public class LogonInfo
    {
        public string ServerName { get; set; }
        public string DatabaseName { get; set; }
        public string UserID { get; set; }
        public string Password { get; set; }
    }

    public class TemplateData
    {
        public TemplateCommand Command { get; private set; }
        public string[] Data { get; private set; }

        public TemplateData(string line)
        {
            int index = line.IndexOf('<');
            Command = GetTemplateCommand(line.Substring(0, index));
            Data = line.Substring(index + 1, line.Length - index - 2).Split(',');
        }

        private TemplateCommand GetTemplateCommand(string sCommand)
        {
            TemplateCommand cmd = TemplateCommand.Unknown;
            switch (sCommand)
            {
                case "SetJob": cmd = TemplateCommand.SetJob; break;
                case "SetTableLocation": cmd = TemplateCommand.SetTableLocation; break;
                case "SetNthTableLocation": cmd = TemplateCommand.SetNthTableLocation; break;
                case "MapNthTable": cmd = TemplateCommand.MapNthTable; break;
                case "SetNthTableLogonInfo": cmd = TemplateCommand.SetNthTableLogonInfo; break;
                case "LogonServer": cmd = TemplateCommand.LogonServer; break;
                case "SetSQLQuery": cmd = TemplateCommand.SetSQLQuery; break;
                case "SetSelectionFormula": cmd = TemplateCommand.SetSelectionFormula; break;
                case "SetSubReportSelectionFormula": cmd = TemplateCommand.SetSubReportSelectionFormula; break;
                case "SetNthParameterField": cmd = TemplateCommand.SetNthParameterField; break;
                case "SetNthSortField": cmd = TemplateCommand.SetNthSortField; break;
                case "DeleteNthSortField": cmd = TemplateCommand.DeleteNthSortField; break;
                case "SetNthGroupSortField": cmd = TemplateCommand.SetNthGroupSortField; break;
                case "DeleteNthGroupSortField": cmd = TemplateCommand.DeleteNthGroupSortField; break;
                case "SetFormula": cmd = TemplateCommand.SetFormula; break;
                case "SetReportTitle": cmd = TemplateCommand.SetReportTitle; break;
                case "OutputToWindow": cmd = TemplateCommand.OutputToWindow; break;
            }
            return cmd;
        }
    }
}
