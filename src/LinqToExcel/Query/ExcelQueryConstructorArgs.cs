namespace LinqToExcel.Query
{
    using System;
    using System.Collections.Generic;
    using Domain;

    internal class ExcelQueryConstructorArgs
    {
        internal Dictionary<string, string> ColumnMappings { get; set; }
        internal DatabaseEngine DatabaseEngine { get; set; }
        internal string FileName { get; set; }
        internal bool ReadOnly { get; set; }
        internal StrictMappingType? StrictMapping { get; set; }
        internal Dictionary<string, Func<string, object>> Transformations { get; set; }
        internal TrimSpacesType TrimSpaces { get; set; }
        internal bool UsePersistentConnection { get; set; }
    }
}