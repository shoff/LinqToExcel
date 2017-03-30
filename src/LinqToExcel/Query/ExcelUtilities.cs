using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;
using LinqToExcel.Domain;
using LinqToExcel.Extensions;

namespace LinqToExcel.Query
{
    internal static class ExcelUtilities
    {
        internal static string GetConnectionString(ExcelQueryArgs args)
        {
            var connString = "";
            var fileNameLower = args.FileName.ToLower();

            if (fileNameLower.EndsWith("xlsx") ||
                fileNameLower.EndsWith("xlsm"))
            {
                connString =
                    $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={args.FileName};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""";
            }
            else if (fileNameLower.EndsWith("xlsb"))
            {
                connString =
                    $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={args.FileName};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""";
            }
            else if (fileNameLower.EndsWith("csv"))
            {
                if (args.DatabaseEngine == DatabaseEngine.Jet)
                {
                    connString =
                        $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={Path.GetDirectoryName(args.FileName)};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1""";
                }
                else
                {
                    connString =
                        $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Path.GetDirectoryName(args.FileName)};Extended Properties=""text;Excel 12.0;HDR=YES;IMEX=1""";
                }
            }
            else
            {
                if (args.DatabaseEngine == DatabaseEngine.Jet)
                {
                    connString =
                        $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={args.FileName};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""";
                }
                else
                {
                    connString =
                        $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={args.FileName};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""";
                }
            }

            if (args.NoHeader)
            {
                connString = connString.Replace("HDR=YES", "HDR=NO");
            }

            if (args.ReadOnly)
            {
                connString = connString.Replace("IMEX=1", "IMEX=1;READONLY=TRUE");
            }

            return connString;
        }

        internal static IEnumerable<string> GetWorksheetNames(string fileName)
        {
	        return GetWorksheetNames(fileName, new ExcelQueryArgs());
        }

		internal static IEnumerable<string> GetWorksheetNames(string fileName, ExcelQueryArgs args)
		{
			args.FileName = fileName;
            args.ReadOnly = true;
			return GetWorksheetNames(args);
		} 

		internal static OleDbConnection GetConnection(ExcelQueryArgs args)
		{
			if (args.UsePersistentConnection)
			{
			    return args.PersistentConnection ?? (args.PersistentConnection = new OleDbConnection(GetConnectionString(args)));
			}

            return new OleDbConnection(GetConnectionString(args));
		}

        internal static IEnumerable<string> GetWorksheetNames(ExcelQueryArgs args)
        {
            var worksheetNames = new List<string>();

	        var conn = GetConnection(args);
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                var excelTables = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });

                if (excelTables != null)
                {
                    worksheetNames.AddRange(
                        from DataRow row in excelTables.Rows
                        where IsTable(row)
                        let tableName = row["TABLE_NAME"].ToString()
                            .Replace("$", "")
                            .RegexReplace("(^'|'$)", "")
                            .Replace("''", "'")
                        where IsNotBuiltinTable(tableName)
                        select tableName);

                    excelTables.Dispose();
                }
            }
            finally
            {
                if (!args.UsePersistentConnection)
                {
                    conn.Dispose();
                }
            }
			
            return worksheetNames;
        }

        internal static bool IsTable(DataRow row)
        {
            return row["TABLE_NAME"].ToString().EndsWith("$") || (row["TABLE_NAME"].ToString().StartsWith("'") && row["TABLE_NAME"].ToString().EndsWith("$'"));
        }

        internal static bool IsNamedRange(DataRow row)
        {
            return (row["TABLE_NAME"].ToString().Contains("$") && !row["TABLE_NAME"].ToString().EndsWith("$") && !row["TABLE_NAME"].ToString().EndsWith("$'")) || !row["TABLE_NAME"].ToString().Contains("$");
        }

        internal static bool IsWorkseetScopedNamedRange(DataRow row)
        {
            return IsNamedRange(row) && row["TABLE_NAME"].ToString().Contains("$");
        }

        internal static bool IsNotBuiltinTable(string tableName)
        {
            return !tableName.Contains("FilterDatabase") && !tableName.Contains("Print_Area");
        }

        internal static IEnumerable<string> GetColumnNames(string worksheetName, string fileName)
        {
            var args = new ExcelQueryArgs
            {
                WorksheetName = worksheetName,
                FileName = fileName
            };
            return GetColumnNames(args);
        }

        internal static IEnumerable<string> GetColumnNames(string worksheetName, string namedRange, string fileName)
        {
            var args = new ExcelQueryArgs
            {
                WorksheetName = worksheetName,
                NamedRangeName = namedRange,
                FileName = fileName
            };
            return GetColumnNames(args);
        }

        internal static IEnumerable<string> GetColumnNames(ExcelQueryArgs args)
        {
            var columns = new List<string>();
            var conn = GetConnection(args);
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                using (var command = conn.CreateCommand())
                {
                    // ReSharper disable UseStringInterpolation
                    command.CommandText = string.Format( "SELECT TOP 1 * FROM [{0}{1}]", string.Format("{0}{1}", args.WorksheetName, "$"), args.NamedRangeName);
                    // ReSharper restore UseStringInterpolation


                    var data = command.ExecuteReader();
                    columns.AddRange(GetColumnNames(data));
                }
            }
            finally
            {
                if (!args.UsePersistentConnection)
                {
                    conn.Dispose();
                }
            }

            return columns;
        }

        internal static IEnumerable<string> GetColumnNames(IDataReader data)
        {
            var columns = new List<string>();
            var sheetSchema = data.GetSchemaTable();
            if (sheetSchema != null)
            {
                foreach (DataRow row in sheetSchema.Rows)
                {
                    columns.Add(row["ColumnName"].ToString());
                }
            }

            return columns;
        }

        internal static DatabaseEngine DefaultDatabaseEngine()
        {
            return Is64BitProcess() ? DatabaseEngine.Ace : DatabaseEngine.Jet;
        }

        internal static bool Is64BitProcess()
        {
            return (IntPtr.Size == 8);
        }

        internal static IEnumerable<string> GetNamedRanges(string fileName, string worksheetName)
        {
            return GetNamedRanges(fileName, worksheetName, new ExcelQueryArgs());
        }

        internal static IEnumerable<string> GetNamedRanges(string fileName)
        {
            return GetNamedRanges(fileName, new ExcelQueryArgs());
        }

        internal static IEnumerable<string> GetNamedRanges(string fileName, ExcelQueryArgs args)
        {
            args.FileName = fileName;
            args.ReadOnly = true;
            return GetNamedRanges(args);
        }

        internal static IEnumerable<string> GetNamedRanges(string fileName, string worksheetName, ExcelQueryArgs args)
        {
            args.FileName = fileName;
            args.WorksheetName = worksheetName;
            args.ReadOnly = true;
            return GetNamedRanges(args);
        }

        internal static IEnumerable<string> GetNamedRanges(ExcelQueryArgs args)
        {
            var namedRanges = new List<string>();

            var conn = GetConnection(args);
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                var excelTables = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });

                if (excelTables != null)
                {
                    namedRanges.AddRange(
                        from DataRow row in excelTables.Rows
                        where IsNamedRange(row)
                        && (!string.IsNullOrEmpty(args.WorksheetName) ? row["TABLE_NAME"].ToString().StartsWith(args.WorksheetName) : !IsWorkseetScopedNamedRange(row))
                        let tableName = row["TABLE_NAME"].ToString()
                            .Replace("''", "'")
                        where IsNotBuiltinTable(tableName)
                        select tableName.Split('$').Last());

                    excelTables.Dispose();
                }
            }
            finally
            {
                if (!args.UsePersistentConnection)
                {
                    conn.Dispose();
                }
            }

            return namedRanges;
        }

    }
}
