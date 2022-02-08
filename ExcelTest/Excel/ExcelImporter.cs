using ExcelDataReader;
using ExcelTest.API_Classes.Body_Elements;
using ExcelTest.API_Classes.Body_Elements.Types;
using ExcelTest.Exceptions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelTest.Excel
{
    class ExcelImporter
    {
        //Important: To add a new type you need to add it to method GetFormFieldElement, CheckFieldType and class ExcelColumn 
        readonly string FilePath;

        public ExcelImporter(string filePath)
        {
            FilePath = filePath;
        }

        private static Type CheckFieldType(string value)
        {
            return typeof(String);
            if (Int32.TryParse(value, out int int32))
                return typeof(Int32);
            if (Double.TryParse(value, out double doubleValue))
                return typeof(Double);
            if (DateTime.TryParse(value, out DateTime dateTime))
            {
                if (value.Contains("00:00:00"))
                    return typeof(Date);
                return typeof(DateTime);

            }
            if (Boolean.TryParse(value, out Boolean boolean))
                return typeof(Boolean);
            if (value.Contains("http://") || value.Contains("https://"))
                return typeof(HyperLink);

            return typeof(String);
        }

        private object GetFormFieldElement(string guid, ExcelColumn<object> excelColumn, string name, object value)
        {
            Type type = excelColumn.ColumnType;
            if (type == typeof(Int32))
                return new FormFieldElement<int>(guid, excelColumn.ColumnSType, name, Int32.Parse(value.ToString()), value.ToString());
            if (type == typeof(Double))
                return new FormFieldElement<double>(guid, excelColumn.ColumnSType, name, Double.Parse(value.ToString()), value.ToString());
            if (type == typeof(DateTime))
                return new FormFieldElement<DateTime>(guid, excelColumn.ColumnSType, name, DateTime.Parse(value.ToString()), value.ToString());
            if (type == typeof(Date))
                return new FormFieldElement<Date>(guid, excelColumn.ColumnSType, name, new Date(value.ToString()), value.ToString().Substring(0, 10));
            if (type == typeof(HyperLink))
                return new FormFieldElement<HyperLink>(guid, excelColumn.ColumnSType, name, new HyperLink(value.ToString()), value.ToString());

            return new FormFieldElement<string>(guid, excelColumn.ColumnSType, name, value.ToString(), value.ToString());
        }


        private List<ExcelColumn<object>> LoadColumnTypes(DataTable table)
        {
            Console.WriteLine("Loading colums types");
            List<ExcelColumn<object>> columns = new List<ExcelColumn<object>>();

            //Checking all columns
            for (int column = 0; column < table.Columns.Count; column++)
            {

                //Get type of first row
                Type columnType = CheckFieldType(table.Rows[0][column].ToString());

                //Checking if all rows are of the same type 
                for (int row = 1; row < table.Rows.Count; row++)
                {
                    Type fieldType = CheckFieldType(table.Rows[row][column].ToString());
                    if (columnType != fieldType || columnType == typeof(string))
                    {

                        //One int in double column
                        if (columnType == typeof(Double) && fieldType == typeof(Int32))
                            continue;
                        //One double in int column
                        if (columnType == typeof(Int32) && fieldType == typeof(Double))
                        {
                            columnType = typeof(Double);
                            continue;
                        }

                        //One Date in DateTime column
                        if (columnType == typeof(DateTime) && columnType == typeof(Date))
                            continue;

                        //One DateTime in date column
                        if (columnType == typeof(Date) && columnType == typeof(DateTime))
                        {
                            columnType = typeof(DateTime);
                            continue;
                        }

                        columnType = typeof(string);
                        break;
                    }
                }
                try
                {
                    columns.Add(new ExcelColumn<object>(table.Columns[column].ColumnName, columnType));
                }
                catch (UnknownTypeException e)
                {
                    Console.WriteLine(e.Message);
                }


            }
            Console.WriteLine("Success - loaded " + columns.Count + " columns");
            return columns;
        }

        DataTable RemoveDuplicatesFromDataTable(DataTable table, List<string> keyColumns)
        {
            Console.Write("Removing duplications in");
            foreach (var item in keyColumns)
            {
                Console.Write(" " + item.ToString());
            }
            Console.WriteLine(" column");
            int oldRowsCout = table.Rows.Count;
            Dictionary<string, string> uniquenessDict = new Dictionary<string, string>(table.Rows.Count);
            StringBuilder stringBuilder = null;
            int rowIndex = 0;
            DataRow row;
            DataRowCollection rows = table.Rows;
            while (rowIndex < rows.Count)
            {
                row = rows[rowIndex];
                stringBuilder = new StringBuilder();
                foreach (string colname in keyColumns)
                {
                    //stringBuilder.Append(((double)row[colname]));
                    stringBuilder.Append(row[colname]);
                }
                if (uniquenessDict.ContainsKey(stringBuilder.ToString()))
                {
                    rows.Remove(row);
                }
                else
                {
                    uniquenessDict.Add(stringBuilder.ToString(), string.Empty);
                    rowIndex++;
                }
            }
            Console.WriteLine("Success - removed " + (oldRowsCout - table.Rows.Count).ToString() + " elements");

            return table;
        }

        public DataTable RemoveDuplicateRows(DataTable table, string DistinctColumn, string[] columns)
        {
            Console.WriteLine("Removing duplications in " + DistinctColumn + " column");
            try
            {
                ArrayList UniqueRecords = new ArrayList();
                ArrayList DuplicateRecords = new ArrayList();

                // Check if records is already added to UniqueRecords otherwise,
                // Add the records to DuplicateRecords
                foreach (DataRow dRow in table.Rows)
                {
                    if (UniqueRecords.Contains(dRow[DistinctColumn]))
                        DuplicateRecords.Add(dRow);
                    else
                        UniqueRecords.Add(dRow[DistinctColumn]);
                }

                // Remove duplicate rows from DataTable added to DuplicateRecords
                foreach (DataRow dRow in DuplicateRecords)
                {
                    table.Rows.Remove(dRow);
                }

                Console.WriteLine("Success - removed " + DuplicateRecords.Count + " elements");

                // Return the clean DataTable which contains unique records.
                return table;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public List<FormFieldList> LoadData(List<string> guids, int tableID = 0, List<string> columnsToCheckDuplicateRows = null, List<string> columnsToRemove = null)
        {

            List<FormFieldList> formFieldLists = new List<FormFieldList>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    //Export data to DataTable
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration() { UseColumnDataType = false, ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration() { UseHeaderRow = true } });
                    var table = result.Tables[tableID];

                    Console.WriteLine("Processing " + table.TableName + " table");

                    //Remove duplications
                    if (columnsToCheckDuplicateRows.Count > 0)
                        table = RemoveDuplicatesFromDataTable(table, columnsToCheckDuplicateRows);

                    if (columnsToRemove != null)
                    {
                        foreach (var column in columnsToRemove)
                        {
                            table.Columns.Remove(column);
                        }
                    }

                    //Get columns type
                    var columns = LoadColumnTypes(table);

                    //Checking if the number of GUIDs is consistent with the number of columns.
                    if (columns.Count != guids.Count)
                        throw new IncorrectGuidSizeException("The number of guids differs from the number of columns.\nGuids - " + guids.Count + " Columns - " + columns.Count);

                    for (int row = 0; row < table.Rows.Count; row++)
                    {
                        //Adding all data from row to FormFieldList
                        FormFieldList formFieldList = new FormFieldList();
                        for (int colum = 0; colum < table.Columns.Count; colum++) //test
                        {
                            try
                            {
                                formFieldList.Add(GetFormFieldElement(guids[colum], columns[colum], columns[colum].Name, table.Rows[row][colum]));
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Loading error: column - " + colum + " row - " + row + "\n" + e.Message);
                            }
                        }
                        formFieldLists.Add(formFieldList);
                    }
                }
            }
            Console.WriteLine("All data loaded to memory");
            return formFieldLists;
        }



    }
}
