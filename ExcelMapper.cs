using ExcelMapping.Extenstions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace ExcelMapping
{
    public abstract class ExcelMapper<T> where T : class, new()
    {
        public ExcelMapper()
        {
            ColumnList = new List<PropertyColumn>();
            ListOfObjects = new List<T>();
        }

        public List<T> ListOfObjects { get; set; }

        private List<PropertyColumn> ColumnList { get; set; }

        protected void MappingColumn(string excelColumnName, Expression<Func<T, object>> property)
        {
            ColumnList.Add(new PropertyColumn
            {
                ColumnName = excelColumnName,
                ColumnNumber = 0,
                PropertyName = GetPropertyName(property)
            });
        }

        protected void MappingColumn(string excelColumnName, Expression<Func<T, object>> property, IParsingDataProperty parsingData, params IValidationDataProperty[] validationDataList)
        {
            ColumnList.Add(new PropertyColumn
            {
                ColumnName = excelColumnName,
                ColumnNumber = 0,
                Parser = parsingData,
                Validators = validationDataList.ToList(),
                PropertyName = GetPropertyName(property)
            });
        }

        protected void MappingColumn(string excelColumnName, Expression<Func<T, object>> property, params IValidationDataProperty[] validationDataList)
        {
            ColumnList.Add(new PropertyColumn
            {
                ColumnName = excelColumnName,
                ColumnNumber = -1,
                Validators = validationDataList.ToList(),
                PropertyName = GetPropertyName(property)
            });
        }

        public List<T> Map(ExcelWorksheet workSheet)
        {
            var columnHeaders = workSheet.GettingHeaders();
            GenerateIndexInColumnList(columnHeaders);
            var numberOfRows = workSheet.Dimension.Rows;
            for (int rowNumber = 2; rowNumber <= numberOfRows; rowNumber++)
            {
                var element = new T();
                var isNotValid = false;
                for (int colIndex = 0; colIndex < ColumnList.Count; colIndex++)
                {
                    var colNumber = ColumnList[colIndex].ColumnNumber;
                    var cellText = workSheet.Cells[rowNumber, colNumber].Text;
                    var isCellMerged = workSheet.Cells[rowNumber, colNumber].Merge;
                    if (cellText.IsEmpty())
                        if (isCellMerged)
                            cellText = workSheet.Cells[1, colNumber, rowNumber, colNumber].Where(a => a.Text.IsNotEmpty()).Select(a => a.Text).LastOrDefault();

                    // Validation Here
                    var validators = ColumnList[colIndex].Validators;
                    for (int validatorIndex = 0; validatorIndex < validators.Count && !isNotValid; validatorIndex++)
                        isNotValid = !(validators[validatorIndex].CheckPropertyValue(cellText));

                    if (isNotValid) break;

                    // Parsing Here
                    var propertyInfo = typeof(T).GetProperty(ColumnList[colIndex].PropertyName);
                    object propertyValue = cellText;
                    if (ColumnList[colIndex].Parser != null) propertyValue = ColumnList[colIndex].Parser.ParsePropertyValue(cellText);

                    try
                    {
                        var valueToStore = Convert.ChangeType(propertyValue, propertyInfo.PropertyType);
                        propertyInfo.SetValue(element, valueToStore);
                    }
                    catch (Exception)
                    {
                        isNotValid = true;
                        break;
                    }
                }

                if (isNotValid) continue;
                ListOfObjects.Add(element);
            }
            return ListOfObjects;
        }

        private void GenerateIndexInColumnList(List<string> columns)
        {
            for (int i = 0; i < ColumnList.Count; i++)
            {
                ColumnList[i].ColumnNumber = columns.IndexOf(ColumnList[i].ColumnName) + 1;
            }
            ColumnList = new List<PropertyColumn>(ColumnList.Where(a => a.ColumnNumber != 0));
        }

        private string GetPropertyName(Expression<Func<T, object>> property)
        {
            LambdaExpression lambda = property;
            MemberExpression memberExpression;

            if (lambda.Body is UnaryExpression)
            {
                UnaryExpression unaryExpression = (UnaryExpression)(lambda.Body);
                memberExpression = (MemberExpression)(unaryExpression.Operand);
            }
            else
            {
                memberExpression = (MemberExpression)(lambda.Body);
            }

            return ((PropertyInfo)memberExpression.Member).Name;
        }
    }
}
