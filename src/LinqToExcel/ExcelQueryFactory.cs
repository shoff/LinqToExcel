namespace LinqToExcel
{
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;
    using Domain;
    using Logging;
    using NLog;
    using Query;

    public class ExcelQueryFactory : IExcelQueryFactory
    {
        private static readonly ILogger log = LogManager.GetCurrentClassLogger();
        private readonly Dictionary<string, string> columnMappings = new Dictionary<string, string>();
        private readonly Dictionary<string, Func<string, object>> transformations = new Dictionary<string, Func<string, object>>();
        private ExcelQueryArgs queryArgs;

        public ExcelQueryFactory()
            : this(null)
        {
        }

        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="logManagerFactory">
        /// Factory that facilitates the creation of an external log manager (i.e. log4net) to 
        /// allow internal methods of LinqToExcel to perform diagnostic logging.
        /// </param>
        public ExcelQueryFactory(string fileName)
        {
            this.FileName = fileName;
            this.DatabaseEngine = ExcelUtilities.DefaultDatabaseEngine();
        }

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column)
        {
            this.AddMapping(this.GetPropertyName(property), column);
        }

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <param name="propertyName">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        public void AddMapping(string propertyName, string column)
        {
            this.columnMappings[propertyName] = column;
        }

        /// <summary>
        /// Add a column to property mapping with a transformation operation
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        /// <param name="transformation">Lambda expression that transforms the original string value to the desired property value</param>
        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column,
            Func<string, object> transformation)
        {
            this.AddMapping(property, column);
            this.AddTransformation(property, transformation);
        }

        /// <summary>
        /// Transforms a cell value in the spreadsheet to the desired property value
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property value to transform</param>
        /// <param name="transformation">Lambda expression that transforms the original string value to the desired property value</param>
        /// <example>
        /// AddTransformation{Person}(p => p.IsActive, x => x == "Y");
        /// AddTransformation{Person}(p => p.IsYoung, x => DateTime.Parse(x) > new DateTime(2000, 1, 1));
        /// </example>
        public void AddTransformation<TSheetData>(Expression<Func<TSheetData, object>> property, Func<string, object> transformation)
        {
            this.transformations.Add(this.GetPropertyName(property), transformation);
        }

        /// <summary>
        /// Sets the database engine to use 
        /// (Spreadsheets ending in xlsx, xlsm, and xlsb must use the Ace database engine)
        /// (If running 64 bit this defaults to ACE (JET doesn't work anyway), if running 32 bit this detaults to JET)
        /// </summary>
        public DatabaseEngine DatabaseEngine { get; set; }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Full path to the Excel spreadsheet
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Returns a list of columns names that a worksheet contains
        /// </summary>
        /// <param name="worksheetName">Worksheet name to get the list of column names from</param>
        public IEnumerable<string> GetColumnNames(string worksheetName)
        {
            if (string.IsNullOrEmpty(this.FileName))
            {
                throw new NullReferenceException("FileName property is not set");
            }

            return ExcelUtilities.GetColumnNames(worksheetName, this.FileName);
        }

        /// <summary>
        /// Returns a list of columns names that a worksheet contains
        /// </summary>
        /// <param name="worksheetName">Worksheet name to get the list of column names from</param>
        /// <param name="namedRangeName">Named Range name to get the list of column names from</param>
        public IEnumerable<string> GetColumnNames(string worksheetName, string namedRange)
        {
            if (string.IsNullOrEmpty(this.FileName))
            {
                throw new NullReferenceException("FileName property is not set");
            }

            return ExcelUtilities.GetColumnNames(worksheetName, namedRange, this.FileName);
        }

        /// <summary>
        /// Returns a list of workbook-scope named ranges that the spreadsheet contains
        /// </summary>
        public IEnumerable<string> GetNamedRanges()
        {
            if (string.IsNullOrEmpty(this.FileName))
            {
                throw new NullReferenceException("FileName property is not set");
            }

            return ExcelUtilities.GetNamedRanges(this.FileName);
        }

        /// <summary>
        /// Returns a list of worksheet-scope named ranges that the worksheet contains
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public IEnumerable<string> GetNamedRanges(string worksheetName)
        {
            if (string.IsNullOrEmpty(this.FileName))
            {
                throw new NullReferenceException("FileName property is not set");
            }

            return ExcelUtilities.GetNamedRanges(this.FileName, worksheetName);
        }

        /// <summary>
        /// Returns a list of worksheet names that the spreadsheet contains
        /// </summary>
        public IEnumerable<string> GetWorksheetNames()
        {
            if (string.IsNullOrEmpty(this.FileName))
            {
                throw new NullReferenceException("FileName property is not set");
            }

            return ExcelUtilities.GetWorksheetNames(this.FileName);
        }

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        public ExcelQueryable<TSheetData> NamedRange<TSheetData>(string namedRangeName)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<TSheetData> NamedRange<TSheetData>(string worksheetName, string namedRangeName)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<TSheetData> NamedRange<TSheetData>(int worksheetIndex, string namedRangeName)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range
        /// </summary>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        public ExcelQueryable<Row> NamedRange(string namedRangeName)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<Row> NamedRange(string worksheetName, string namedRangeName)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<Row> NamedRange(int worksheetIndex, string namedRangeName)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range that does not have a header row
        /// </summary>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        public ExcelQueryable<RowNoHeader> NamedRangeNoHeader(string namedRangeName)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true,
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range that does not have a header row
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<RowNoHeader> NamedRangeNoHeader(int worksheetIndex, string namedRangeName)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetIndex = worksheetIndex,
                    NamedRangeName = namedRangeName
                }));
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range that does not have a header row
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<RowNoHeader> NamedRangeNoHeader(string worksheetName, string namedRangeName)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(new ExcelQueryArgs(this.GetConstructorArgs())
            {
                NoHeader = true,
                WorksheetName = worksheetName,
                NamedRangeName = namedRangeName
            }));
        }

        /// <summary>
        /// Confirms all the worksheet columns are mapped to a property, and if not, throws a StrictMappingException
        /// </summary>
        public StrictMappingType? StrictMapping { get; set; }

        /// <summary>
        /// Indicates how to treat leading and trailing spaces in string values.
        /// </summary>
        public TrimSpacesType TrimSpaces { get; set; }

        /// <summary>
        /// If true, uses a single, persistent connection for the lifetime of the factory.
        /// If false, a new connection is created for each query
        /// Default is false
        /// </summary>
        public bool UsePersistentConnection { get; set; }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>()
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(new ExcelQueryArgs(this.GetConstructorArgs())));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        public ExcelQueryable<Row> Worksheet()
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetName = worksheetName
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> Worksheet(string worksheetName)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetName = worksheetName
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> Worksheet(int worksheetIndex)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader()
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(string worksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetName = worksheetName
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(int worksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetIndex = worksheetIndex
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange"></param>
        /// <param name="endRange"></param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    StartRange = startRange,
                    EndRange = endRange
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    StartRange = startRange,
                    EndRange = endRange
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    StartRange = startRange,
                    EndRange = endRange
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<Row>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    StartRange = startRange,
                    EndRange = endRange
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    StartRange = startRange,
                    EndRange = endRange
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<TSheetData>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    StartRange = startRange,
                    EndRange = endRange
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange,
                    WorksheetName = worksheetName
                }));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(this.PersistQueryArgs(
                new ExcelQueryArgs(this.GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange,
                    WorksheetIndex = worksheetIndex
                }));
        }

        private string GetPropertyName<TSheetData>(Expression<Func<TSheetData, object>> property)
        {
            var exp = (LambdaExpression)property;

            //exp.Body has 2 possible types
            //If the property type is native, then exp.Body == typeof(MemberExpression)
            //If the property type is not native, then exp.Body == typeof(UnaryExpression) in which 
            //case we can get the MemberExpression from its Operand property
            var mExp = exp.Body.NodeType == ExpressionType.MemberAccess ?
                (MemberExpression)exp.Body :
                (MemberExpression)((UnaryExpression)exp.Body).Operand;
            return mExp.Member.Name;
        }

        internal ExcelQueryConstructorArgs GetConstructorArgs()
        {
            return new ExcelQueryConstructorArgs
            {
                FileName = this.FileName,
                DatabaseEngine = this.DatabaseEngine,
                StrictMapping = this.StrictMapping,
                ColumnMappings = this.columnMappings,
                Transformations = this.transformations,
                UsePersistentConnection = this.UsePersistentConnection,
                TrimSpaces = this.TrimSpaces,
                ReadOnly = this.ReadOnly
            };
        }

        private ExcelQueryArgs PersistQueryArgs(ExcelQueryArgs args)
        {
            // We want to keep the persistent connection if there is one
            if (this.queryArgs != null)
            {
                args.PersistentConnection = this.queryArgs.PersistentConnection;
            }

            return this.queryArgs = args;
        }

        ~ExcelQueryFactory()
        {
            this.Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (queryArgs?.PersistentConnection != null)
                {
                    try
                    {
                        this.queryArgs.PersistentConnection.Dispose();
                        this.queryArgs.PersistentConnection = null;
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex,  ex.Message);
                    }
                }
            }

            this.queryArgs = null;
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName })
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName })
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName })
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName })
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName,
            Dictionary<string, string> columnMappings,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName,
            Dictionary<string, string> columnMappings,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName,
            Dictionary<string, string> columnMappings,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName,
            Dictionary<string, string> columnMappings,
            ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Opens the excel file in read-only mode
        /// </summary>
        public bool ReadOnly { get; set; }
    }
}