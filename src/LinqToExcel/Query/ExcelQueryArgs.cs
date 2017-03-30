namespace LinqToExcel.Query
{
    using System;
    using System.Collections.Generic;
    using System.Data.OleDb;
    using System.Linq;
    using System.Text;
    using Domain;

    internal class ExcelQueryArgs
    {
        internal ExcelQueryArgs()
            : this(new ExcelQueryConstructorArgs {DatabaseEngine = ExcelUtilities.DefaultDatabaseEngine()})
        {
        }

        internal ExcelQueryArgs(ExcelQueryConstructorArgs args)
        {
            this.FileName = args.FileName;
            this.DatabaseEngine = args.DatabaseEngine;
            this.ColumnMappings = args.ColumnMappings ?? new Dictionary<string, string>();
            this.Transformations = args.Transformations ?? new Dictionary<string, Func<string, object>>();
            this.StrictMapping = args.StrictMapping ?? StrictMappingType.None;
            this.UsePersistentConnection = args.UsePersistentConnection;
            this.TrimSpaces = args.TrimSpaces;
            this.ReadOnly = args.ReadOnly;
        }

        public override string ToString()
        {
            var columnMappingsString = new StringBuilder();
            foreach (var kvp in this.ColumnMappings)
            {
                columnMappingsString.AppendFormat("[{0} = '{1}'] ", kvp.Key, kvp.Value);
            }
            var transformationsString = string.Join(", ", this.Transformations.Keys.ToArray());

            return
                string.Format(
                    "FileName: '{0}'; WorksheetName: '{1}'; WorksheetIndex: {2}; StartRange: {3}; EndRange: {4}; Named Range: {11}; NoHeader: {5}; ColumnMappings: {6}; Transformations: {7}, StrictMapping: {8}, UsePersistentConnection: {9}, TrimSpaces: {10}",
                    this.FileName, this.WorksheetName, this.WorksheetIndex, this.StartRange, this.EndRange, this.NoHeader,
                    columnMappingsString, transformationsString, this.StrictMapping, this.UsePersistentConnection, this.TrimSpaces,
                    this.NamedRangeName);
        }

        internal Dictionary<string, string> ColumnMappings { get; set; }
        internal DatabaseEngine DatabaseEngine { get; set; }
        internal string EndRange { get; set; }
        internal string FileName { get; set; }
        internal string NamedRangeName { get; set; }
        internal bool NoHeader { get; set; }
        internal OleDbConnection PersistentConnection { get; set; }
        internal bool ReadOnly { get; set; }
        internal string StartRange { get; set; }
        internal StrictMappingType? StrictMapping { get; set; }
        internal Dictionary<string, Func<string, object>> Transformations { get; }
        internal TrimSpacesType TrimSpaces { get; set; }
        internal bool UsePersistentConnection { get; set; }
        internal int? WorksheetIndex { get; set; }
        internal string WorksheetName { get; set; }
    }
}