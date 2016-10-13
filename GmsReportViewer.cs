using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Data.OleDb;

namespace GmsReportViewer
{
    public partial class GmsReportViewer : Form
    {
        private ReportDocument _reportDocument = null;

        public GmsReportViewer(string reportScript)
        {
            InitializeComponent();
            ShowReport(reportScript);
        }

        private void ShowReport(string reportScript)
        {
            _reportDocument = new ReportDocument();
            List<TemplateData> data = ReportParser.GetTemplateData(reportScript);
            for (int i = 0; i < data.Count; i++)
            {
                TemplateData td = data[i];
                switch (td.Command)
                {
                    case TemplateCommand.SetJob:
                        SetJob(td.Data);
                        break;
                    case TemplateCommand.SetNthTableLocation:
                        SetNthTableLocation(td.Data);
                        break;
                    case TemplateCommand.SetTableLocation:
                        SetTableLocation(td.Data);
                        break;
                    case TemplateCommand.MapNthTable: MapNthTable(td.Data); break;
                    case TemplateCommand.SetNthTableLogonInfo: SetNthTableLogonInfo(td.Data); break;
                    case TemplateCommand.LogonServer: LogonServer(td.Data); break;
                    case TemplateCommand.SetSQLQuery:
                        SetSQLQuery(td.Data);
                        break;
                    case TemplateCommand.SetSelectionFormula:
                        SetSelectionFormula(ReportParser.HexToAsciiString(td.Data[0]));
                        break;
                    case TemplateCommand.SetSubReportSelectionFormula: break;
                    case TemplateCommand.SetNthParameterField:
                        SetNthParameterField(td.Data);
                        break;
                    case TemplateCommand.SetNthSortField:
                        SetNthSortField(td.Data);
                        break;
                    case TemplateCommand.DeleteNthSortField: break;
                    case TemplateCommand.SetNthGroupSortField: break;
                    case TemplateCommand.DeleteNthGroupSortField: break;
                    case TemplateCommand.SetFormula: SetFormula(td.Data); break;
                    case TemplateCommand.SetReportTitle:
                        SetReportTitle(ReportParser.HexToUnicodeString(td.Data[0]));
                        break;
                    case TemplateCommand.OutputToWindow:
                        OutputToWindow(td.Data);
                        break;
                    default: break;
                }
            }
        }

        private void OutputToWindow(string[] data)
        {
            string title = _reportDocument.SummaryInfo?.ReportTitle ?? data?[0] ?? null;
            if (!string.IsNullOrEmpty(title))
                this.Text = title;
        }

        private void SetJob(string[] data)
        {
            if (data.Length <= 0)
                return;
            _reportDocument.Load(data[0]);
            crystalReportViewer.ReportSource = _reportDocument;
        }

        private void SetNthParameterField(string[] data)
        {
            CRParamFields crParamFields = CRParamFields.FromStrings(data);
            var paramField = crystalReportViewer.ParameterFieldInfo?[crParamFields.ParmNumber] ?? null;
            if(paramField != null)
            {
                paramField.CurrentValues.Clear();
                paramField.CurrentValues.AddValue(crParamFields.CurrentValue);
                paramField.DefaultValues.Clear();
                paramField.DefaultValues.AddValue(crParamFields.DefaultValue);
            }
        }

        private bool SetTableLocation(string oldLocation, string newLocation, ReportDocument reportDucument)
        {
            foreach(Table table in reportDucument.Database.Tables)
            {
                if(string.Compare(table.Name, oldLocation, true) == 0)
                {
                    table.Location = newLocation;
                    foreach (ReportDocument rp in reportDucument.Subreports)
                    {
                        foreach (Table t in rp.Database.Tables)
                        {
                            if(string.Compare(t.Name, oldLocation, true) == 0)
                                t.Location = newLocation;
                        }
                    }
                    return true;
                }
            }

            return false;
        }

        private void SetTableLocation(string[] data)
        {
            if (data.Length < 3)
                return;
            string oldLocation = data[0].Trim();
            string newLocation = data[1].Trim();
            if (!SetTableLocation(oldLocation, newLocation, _reportDocument))
            {
                int includeSubReport = 0;
                if (int.TryParse(data[2].Trim(), out includeSubReport) && includeSubReport > 0)
                {
                    foreach (Section section in _reportDocument.ReportDefinition.Sections)
                    {
                        foreach (ReportObject reportObject in section.ReportObjects)
                        {
                            if (reportObject.Kind == ReportObjectKind.SubreportObject)
                            {
                                SubreportObject subReportObject = reportObject as SubreportObject;
                                if (subReportObject != null)
                                {
                                    ReportDocument rd = subReportObject.OpenSubreport(subReportObject.SubreportName);
                                    SetTableLocation(oldLocation, newLocation, rd);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void MapNthTable(string[] data)
        {
            if (data.Length < 2)
                return;
            int tableNumber = 0;
            if(int.TryParse(data[0].Trim(), out tableNumber) && tableNumber >= 1)
            {
                var table = _reportDocument.Database?.Tables?[tableNumber - 1] ?? null;
                if(table != null)
                {
                    table.Location = data[1].Trim();
                }
            }
        }

        private void SetNthTableLogonInfo(string[] data)
        {
            int tableNumber = 0;
            int propagate = 0;
            if(int.TryParse(data[0].Trim(), out tableNumber) && int.TryParse(data[2].Trim(), out propagate))
            {
                string[] logonInfos = data?[1].Split('$');
                if (logonInfos == null || logonInfos.Length < 2)
                    return;

                var table = _reportDocument.Database?.Tables?[tableNumber - 1] ?? null;
                if(table != null)
                {
                    var logonInfo = GetLogonInfo(data[1]);
                    if(logonInfo != null)
                    {
                        table.LogOnInfo.ConnectionInfo = GetConnectionInfo(logonInfo);
                        var tableLocation = table.Location;
                        var index = tableLocation.IndexOf('.');
                        if (index > 0)
                        {
                            table.Location = tableLocation.Substring(index + 1);
                        }
                        table.TestConnectivity();
                    }
                }
            }
        }

        private void LogonServer(string[] data)
        {
            if (data.Length <= 0)
                return;
            var logonInfo = GetLogonInfo(data[0]);
            string userID = logonInfo.UserID != "" ? logonInfo.UserID : "dbo";
            //string connString = $@"DSN={logonInfo.ServerName};Database={logonInfo.DatabaseName};Uid={userID}; Pwd={logonInfo.Password};UserDSNProperties=0;";
            _reportDocument.SetDatabaseLogon(userID, logonInfo.Password, logonInfo.ServerName, logonInfo.DatabaseName);
            ConnectionInfo connectInfo = GetConnectionInfo(logonInfo);
            foreach(Table table in _reportDocument.Database.Tables)
            {
                table.LogOnInfo.ConnectionInfo = connectInfo;
                if (!table.TestConnectivity())
                {
                    break;
                }
            }
        }

        private void SetReportTitle(string title)
        {
            _reportDocument.SummaryInfo.ReportTitle = title;
        }

        private void SetSelectionFormula(string formula)
        {
            crystalReportViewer.SelectionFormula = formula;
        }

        private void SetNthSortField(string[] data)
        {
            if (data.Length < 3)
                return;
            int fieldNumber = 0, fieldDirection = 0;
            if(int.TryParse(data[0], out fieldNumber) && int.TryParse(data[2], out fieldDirection))
            {
                string fieldData = data[1].Trim();
                SortDirection sortDirection = fieldDirection == 0 ? SortDirection.AscendingOrder : SortDirection.DescendingOrder;
                var dataDefinition = _reportDocument.DataDefinition;
                if (fieldNumber <= dataDefinition.SortFields.Count)
                {
                    SortField sortField = dataDefinition.SortFields[fieldNumber - 1];
                    if(sortField.Field.Name == fieldData)
                    {
                        sortField.SortDirection = sortDirection;
                    }
                    else if (fieldData[0] == '{')
                    {
                        string[] arrFieldData = fieldData.Substring(1, fieldData.Length - 2).Split('.');
                        if (arrFieldData.Length >= 2)
                        {
                            string tableName = arrFieldData[0], fieldName = arrFieldData[1];
                            sortField.Field = _reportDocument.Database.Tables?[tableName]?.Fields?[fieldName];
                            sortField.SortDirection = sortDirection;
                        }
                    }
                }
            }
        }

        private void SetFormula(string[] data)
        {
            if (data.Length < 2)
                return;
            string formulaName = data[0].Trim();
            string formula = ReportParser.HexToAsciiString(data[1]);
            FormulaFieldDefinition ffd = _reportDocument.DataDefinition.FormulaFields[formulaName];
            if (ffd.FormulaName == formulaName)
                ffd.Text = formula;
        }

        private void SetNthTableLocation(string[] data)
        {
            if (data.Length < 4)
                return;
            int tableNumber = 0;
            if(int.TryParse(data[0], out tableNumber))
            {
                string connBuffer = data[1];
                string location = data[2];
                string subLocation = data[3];
                if (tableNumber < _reportDocument.Database.Tables.Count)
                {
                    _reportDocument.Database.Tables[tableNumber].Location = location;
                }
            }
        }

        private List<string> GetTableNames(string queryText)
        {
            List<string> tableNames = new List<string>();
            int fromIndex = queryText.IndexOf("FROM");
            int whereIndex = queryText.IndexOf("WHERE");
            foreach(string name in queryText.Substring(fromIndex + 4, whereIndex - fromIndex - 4).Trim().Split(','))
            {
                string[] ns = name.Trim().Split(' ');
                if(ns.Length > 0)
                {
                    tableNames.Add(ns[0].Replace("\"", ""));
                }
            }
            return tableNames;
        }

        private void SetSQLQuery(string[] data)
        {
            if (data.Length == 0)
                return;
            string queryText = ReportParser.HexToAsciiString(data[0]);
            List<string> tableNames = GetTableNames(queryText);
            if (tableNames.Count > 0)
            {
                string connString = GetConnectionString(crystalReportViewer.LogOnInfo[0].ConnectionInfo);
                using (var sqlConnection = new OdbcConnection(connString))
                {
                    sqlConnection.Open();
                    using (var sqlCommand = new OdbcCommand(queryText, sqlConnection))
                    {
                        using (var dataAdapter = new OdbcDataAdapter(queryText, sqlConnection))
                        {
                            using (var dataSet = new DataSet())
                            {
                                dataAdapter.Fill(dataSet, tableNames[0]);
                                _reportDocument.SetDataSource(dataSet);
                            }
                        }
                    }
                    sqlConnection.Close();
                }
            }
        }

        private LogonInfo GetLogonInfo(string logonString)
        {
            string[] logonInfos = logonString.Split('$');
            if (logonInfos == null || logonInfos.Length < 2)
                return null;
            var logonInfo = new LogonInfo();
            logonInfo.ServerName = logonInfos[0].Trim();
            logonInfo.DatabaseName = logonInfos[1].Trim();
            if (logonInfos.Length >= 4)
            {
                logonInfo.UserID = logonInfos[2];
                logonInfo.Password = logonInfos[3];
            }

            return logonInfo;
        }

        private ConnectionInfo GetConnectionInfo(LogonInfo logonInfo)
        {
            var connInfo = new ConnectionInfo(crystalReportViewer.LogOnInfo[0].ConnectionInfo);
            connInfo.ServerName = logonInfo.ServerName;
            connInfo.DatabaseName = logonInfo.DatabaseName;
            connInfo.UserID = logonInfo.UserID;
            connInfo.Password = logonInfo.Password;
            return connInfo;
        }

        string GetConnectionString(ConnectionInfo connInfo)
        {
            return $@"DSN={connInfo.ServerName};Uid={connInfo.UserID}; Pwd={connInfo.Password};";
        }
    }
}
