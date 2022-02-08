using ExcelTest.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest.Excel
{
    class ExcelColumn<T>
    {

        public string Name { get; set; }
        public string ColumnSType { get; set; }

        public Type ColumnType;

        public ExcelColumn(string name, Type columnType)
        {
            this.ColumnType = columnType;
            this.Name = name;
            setSType();
        }

        private void setSType()
        {
            switch (this.ColumnType.Name)
            {
                case "String":
                    this.ColumnSType = "SingleLine";
                    break;
                case "Double":
                    this.ColumnSType = "Decimal";
                    break;
                case "Int32":
                    this.ColumnSType = "Int";
                    break;
                case "DateTime":
                    this.ColumnSType = "DateTime";
                    break;
                case "HyperLink":
                    this.ColumnSType = "HyperLink";
                    break;
                case "Date":
                    this.ColumnSType = "Date";
                    break;
                case "DBNull":
                    this.ColumnSType = "SingleLine";
                    break;
                default:
                    this.ColumnSType = this.ColumnType.Name;
                    throw new UnknownTypeException("Unknown type for element " +this.Name + ": " + this.ColumnSType);
            }
        }

    }
}
