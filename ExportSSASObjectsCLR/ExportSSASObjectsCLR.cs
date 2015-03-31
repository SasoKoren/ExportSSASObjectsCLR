using System;
using System.Data;
using System.Collections;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using AMO = Microsoft.AnalysisServices;

public partial class UserDefinedFunctions
{

    private class ResultRow
    {
        public SqlString ServerName;
        public SqlString DatabaseName;
        public SqlString AnalyticalObject;
        public SqlString ElementName;
        public SqlString TransPath;
        public SqlString TransKey;
        public SqlString TransObject;
        public SqlInt32 TransLanguage;
        public SqlString TransValue;

        public ResultRow(SqlString ServerName_, SqlString DatabaseName_, SqlString AnalyticalObject_, SqlString ElementName_, SqlString TransPath_, SqlString TransKey_, SqlString TransObject_, SqlInt32 TransLanguage_, SqlString TransValue_)
        {
            ServerName = ServerName_;
            DatabaseName = DatabaseName_;
            AnalyticalObject = AnalyticalObject_;
            ElementName = ElementName_;
            TransPath = TransPath_;
            TransKey = TransKey_;
            TransObject = TransObject_;
            TransLanguage = TransLanguage_;
            TransValue = TransValue_;
        }
    }


    [Microsoft.SqlServer.Server.SqlFunction(
            FillRowMethodName = "OutputResultRow",
            TableDefinition = "ServerName nvarchar(255), DatabaseName nvarchar(255), AnalyticalObject nvarchar(255), ElementName nvarchar(255), TransPath nvarchar(255), TransKey nvarchar(255), TransObject nvarchar(255), TransLanguage int, TransValue nvarchar(255)"
        )
    ]
    public static IEnumerable GetSSASObjects(string ServerName, string DatabaseName)
    {

        AMO.Server Server = new AMO.Server();
        AMO.Database db;
        ArrayList results = new ArrayList();

        // connect to server
        try
        {
            Server.Connect(@"Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" + DatabaseName + ";Data Source=" + ServerName);
        }
        catch
        {
            throw new Exception("Cannot connect to the SSAS server");
        }

        // find the database
        db = Server.Databases.FindByName(DatabaseName);
        if (db == null)
            throw new Exception("Cannot connect to the database");

        // MAIN

        // dimensions
        foreach (AMO.Dimension dim in db.Dimensions)
        {
                
            results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, dim.Name, "Dimension", "[" + dim.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
            foreach (AMO.Translation t in dim.Translations)
                results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, dim.Name, "Dimension", "[" + dim.ID + "]", "Caption", t.Language, t.Caption));                

            // dimension attributes
            foreach (AMO.DimensionAttribute attr in dim.Attributes)
            {                    
                results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, attr.Name, "Dimension.Attribute", "[" + dim.ID + "].[" + attr.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
                foreach (AMO.AttributeTranslation t in attr.Translations)
                    results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, attr.Name, "Dimension.Attribute", "[" + dim.ID + "].[" + attr.ID + "]", "Caption", t.Language, t.Caption));
            }

            // dimension hierarchies
            foreach (AMO.Hierarchy h in dim.Hierarchies)
            {
                results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, h.Name, "Dimension.Hierarchy", "[" + dim.ID + "].[" + h.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
                foreach (AMO.Translation t in h.Translations)
                    results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, h.Name, "Dimension.Hierarchy", "[" + dim.ID + "].[" + h.ID + "]", "Caption", t.Language, t.Caption));

                // dimension hierarchy levels
                foreach (AMO.Level l in h.Levels)
                {
                    results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, l.Name, "Dimension.Hierarchy.Level", "[" + dim.ID + "].[" + h.ID + "].[" + l.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
                    foreach (AMO.Translation t in l.Translations)
                        results.Add(new ResultRow(ServerName, DatabaseName, dim.ID, l.Name, "Dimension.Hierarchy.Level", "[" + dim.ID + "].[" + h.ID + "].[" + l.ID + "]", "Caption", t.Language, t.Caption));
                }
            }

        }

        // cubes
        foreach (AMO.Cube cube in db.Cubes)
        {
            results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, cube.Name, "Cube", "[" + cube.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
            foreach (AMO.Translation t in cube.Translations)
                results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, cube.Name, "Cube", "[" + cube.ID + "]", "Caption", t.Language, t.Caption));

            // cube dimensions
            foreach (AMO.CubeDimension d in cube.Dimensions)
            {
                results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, d.Name, "Cube.Dimension", "[" + cube.ID + "].[" + d.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
                foreach (AMO.Translation t in d.Translations)
                    results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, d.Name, "Cube.Dimension", "[" + cube.ID + "].[" + d.ID + "]", "Caption", t.Language, t.Caption));
            }

            // measure groups
            foreach (AMO.MeasureGroup mg in cube.MeasureGroups)
            {
                results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, mg.Name, "Cube.MeasureGroup", "[" + cube.ID + "].[" + mg.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
                foreach (AMO.Translation t in mg.Translations)
                    results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, mg.Name, "Cube.MeasureGroup", "[" + cube.ID + "].[" + mg.ID + "]", "Caption", t.Language, t.Caption));

                // measures
                foreach (AMO.Measure m in mg.Measures)
                {
                    results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
                    foreach (AMO.Translation t in m.Translations)
                        results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]", "Caption", t.Language, t.Caption));

                    // display folders for measures
                    if (!string.IsNullOrEmpty(m.DisplayFolder))
                        results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]", "DisplayFolder", SqlInt32.Null, m.DisplayFolder));

                    foreach (AMO.Translation t in m.Translations)
                        if (!string.IsNullOrEmpty(t.DisplayFolder))
                            results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, m.Name, "Cube.MeasureGroup.Measure", "[" + cube.ID + "].[" + mg.ID + "].[" + m.ID + "]", "DisplayFolder", t.Language, t.DisplayFolder));

                }

            }

            // calculated measures (mdx)
            foreach (AMO.MdxScript mdx in cube.MdxScripts)
            {
                foreach (AMO.CalculationProperty cp in mdx.CalculationProperties)
                {
                    results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]", SqlString.Null, SqlInt32.Null, SqlString.Null));
                    foreach (AMO.Translation t in cp.Translations)
                        results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]", "Caption", t.Language, t.Caption));

                    // display folders for calculated measures
                    if (!string.IsNullOrEmpty(cp.DisplayFolder))
                        results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]", "DisplayFolder", SqlInt32.Null, cp.DisplayFolder));

                    foreach (AMO.Translation t in cp.Translations)
                        if (!string.IsNullOrEmpty(t.DisplayFolder))
                            results.Add(new ResultRow(ServerName, DatabaseName, cube.ID, cp.CalculationReference, "Cube.MdxScript.CalculationProperty", "[" + cube.ID + "].[MdxScript].[" + cp.CalculationReference + "]", "DisplayFolder", t.Language, t.DisplayFolder));

                }
            }

        }

        return results;
    }


    public static void OutputResultRow(
        object resultObj, 
        out SqlString oServerName, 
        out SqlString oDatabaseName, 
        out SqlString oAnalyticalObject,
        out SqlString oElementName,
        out SqlString oTransPath,
        out SqlString oTransKey,
        out SqlString oTransObject,
        out SqlInt32 oTransLanguage,
        out SqlString oTransValue)
    {
        ResultRow result = (ResultRow)resultObj;

        oServerName = result.ServerName;
        oDatabaseName = result.DatabaseName;
        oAnalyticalObject = result.AnalyticalObject;
        oElementName = result.ElementName;
        oTransPath = result.TransPath;
        oTransKey = result.TransKey;
        oTransObject = result.TransObject;
        oTransLanguage = result.TransLanguage;
        oTransValue = result.TransValue;

    }
}
