using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.Data;
using LicenseContext = System.ComponentModel.LicenseContext;

namespace Api.Helpers
{
    public static class ExcelExportHelper
    {
        public static string ExcelContentType
        {
            get
            { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; }
        }

        public static byte[] ExportExcel(DataTable dataTable, string heading = "", string heading1 = "", string heading2 = "", string heading3 = "", bool showSrNo = true)
        {

            byte[] result = null;
            using (var package = new ExcelPackage())
            {       
                int startRowFrom = 0;
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(String.Format("Sheet1"));
                if (!String.IsNullOrEmpty(heading3))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 6;
                }
                else if (!String.IsNullOrEmpty(heading2))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 5;
                }
                else if (!String.IsNullOrEmpty(heading1))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 4;
                }
                else if (!String.IsNullOrEmpty(heading))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 3;
                }

                // add the content into the Excel file  
                workSheet.Cells["A" + startRowFrom].LoadFromDataTable(dataTable, true);

                // autofit width of cells with small content  
                int columnIndex = 1;
                var b = workSheet.Dimension;
                if (b != null)
                {
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        ExcelRange columnCells = workSheet.Cells[workSheet.Dimension.Start.Row, columnIndex, workSheet.Dimension.End.Row, columnIndex];
                        workSheet.Column(columnIndex).AutoFit();

                        if (column.ColumnName.Contains("Date"))
                        {
                            workSheet.Column(columnIndex).Style.Numberformat.Format = "mm/dd/yyyy";
                        }
                        columnIndex++;
                    }
                }
                else
                {
                    ExcelWorksheet worksheet = new ExcelPackage().Workbook.Worksheets.Add("Sheet1");
                }
                // format header - bold, yellow on black  
                using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom, dataTable.Columns.Count])
                {
                    r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    r.Style.Font.Bold = true;
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#1fb5ad"));
                }

                // format cells - add borders  
                if (dataTable.Rows.Count != 0)
                {
                    using (ExcelRange r = workSheet.Cells[startRowFrom + 1, 1, startRowFrom + dataTable.Rows.Count, dataTable.Columns.Count])
                    {
                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    }
                }

                if (!String.IsNullOrEmpty(heading))
                {

                    if (!String.IsNullOrEmpty(heading1))
                    {
                        workSheet.Cells["A2:O2"].Merge = true;
                        workSheet.Cells["A2:O2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        workSheet.Cells["A2:O2"].Style.Font.Bold = true;
                        workSheet.Cells["A2:O2"].Style.Font.Size = 15;
                        workSheet.Cells["A2"].Value = heading1;

                    }
                    if (!String.IsNullOrEmpty(heading2))
                    {
                        workSheet.Cells["A3:O3"].Merge = true;
                        workSheet.Cells["A3:O3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        workSheet.Cells["A3:O3"].Style.Font.Bold = true;
                        workSheet.Cells["A3"].Value = heading2;
                    }
                    if (!String.IsNullOrEmpty(heading3))
                    {
                        workSheet.Cells["A4:O4"].Merge = true;
                        workSheet.Cells["A4:O4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        workSheet.Cells["A4:O4"].Style.Font.Bold = true;
                        workSheet.Cells["A4:O4"].Style.Font.Size = 12;
                        workSheet.Cells["A4"].Value = heading3;
                    }

                    workSheet.InsertRow(1, 1);

                }



                result = package.GetAsByteArray();
            }

            return result;
        }

        public static byte[] ExportExcel<T>(DataTable data, string Heading = "", string Heading1 = "", string Heading2 = "", string Heading3 = "", bool showSlno = false)
        {
            return ExportExcel(data, Heading, Heading1, Heading2, Heading3, showSlno);
        }

        public static DataTable ListToDataTable<T>(List<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable dataTable = new DataTable();

            for (int i = 0; i < properties.Count; i++)
            {
                PropertyDescriptor property = properties[i];
                dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
            }

            object[] values = new object[properties.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = properties[i].GetValue(item);
                }

                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

        public static byte[] ExportExcel1(DataTable dataTable, string heading = "", string heading1 = "", string heading2 = "", string heading3 = "", bool showSrNo = true, params string[] columnsToTake)
        {

            byte[] result = null;
            using (ExcelPackage package = new ExcelPackage())
            {
                int startRowFrom = 0;
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(String.Format("Sheet1"));
                if (!String.IsNullOrEmpty(heading3))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 6;
                }
                else if (!String.IsNullOrEmpty(heading2))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 5;
                }
                else if (!String.IsNullOrEmpty(heading1))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 4;
                }
                else if (!String.IsNullOrEmpty(heading))
                {
                    startRowFrom = String.IsNullOrEmpty(heading) ? 1 : 3;
                }


                if (showSrNo)
                {
                    DataColumn dataColumn = dataTable.Columns.Add("SN", typeof(int));
                    dataColumn.SetOrdinal(0);
                    int index = 1;
                    foreach (DataRow item in dataTable.Rows)
                    {
                        item[0] = index;
                        index++;
                    }
                }


                // add the content into the Excel file  
                workSheet.Cells["A" + startRowFrom].LoadFromDataTable(dataTable, true);

                // autofit width of cells with small content  
                int columnIndex = 1;
                var b = workSheet.Dimension;
                if (b != null)
                {
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        ExcelRange columnCells = workSheet.Cells[workSheet.Dimension.Start.Row, columnIndex, workSheet.Dimension.End.Row, columnIndex];
                        //int maxLength = columnCells.Where(x => x.Value != null).Max(cell => cell.Value.ToString().Count());
                        workSheet.Column(columnIndex).AutoFit();
                        //if (maxLength < 150)
                        //{
                        //    workSheet.Column(columnIndex).AutoFit();
                        //}
                        if (column.ColumnName.Contains("Date"))
                        {
                            workSheet.Column(columnIndex).Style.Numberformat.Format = "mm/dd/yyyy";
                        }
                        columnIndex++;
                    }
                }
                else
                {
                    ExcelWorksheet worksheet = new ExcelPackage().Workbook.Worksheets.Add("Sheet1");
                }
                // format header - bold, yellow on black  
                using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom, dataTable.Columns.Count])
                {
                    r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    r.Style.Font.Bold = true;
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#0c7899"));
                }

                // format cells - add borders  
                if (dataTable.Rows.Count != 0)
                {
                    using (ExcelRange r = workSheet.Cells[startRowFrom + 1, 1, startRowFrom + dataTable.Rows.Count, dataTable.Columns.Count])
                    {
                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                    }
                }


                // removed ignored columns  
                for (int i = dataTable.Columns.Count - 1; i >= 0; i--)
                {
                    if (i == 0 && showSrNo)
                    {
                        continue;
                    }
                    if (!columnsToTake.Contains(dataTable.Columns[i].ColumnName))
                    {
                        workSheet.DeleteColumn(i + 1);
                    }
                }

                if (!String.IsNullOrEmpty(heading))
                {
                    workSheet.Cells["A1"].Value = heading;
                    workSheet.Cells["A1"].Style.Font.Size = 10;
                    workSheet.Cells["A1"].Style.Font.Bold = true;
                    if (!String.IsNullOrEmpty(heading1))
                    {
                        workSheet.Cells["A2"].Value = heading1;
                        workSheet.Cells["A2"].Style.Font.Size = 10;
                        workSheet.Cells["A2"].Style.Font.Bold = true;
                    }
                    if (!String.IsNullOrEmpty(heading2))
                    {
                        workSheet.Cells["A3"].Value = heading2;
                        workSheet.Cells["A3"].Style.Font.Size = 10;
                        workSheet.Cells["A3"].Style.Font.Bold = true;
                    }
                    if (!String.IsNullOrEmpty(heading3))
                    {
                        workSheet.Cells["A4"].Value = heading3;
                        workSheet.Cells["A4"].Style.Font.Size = 10;
                        workSheet.Cells["A4"].Style.Font.Bold = true;
                    }
                    //workSheet.Cells["A1:L3"].Merge = true;
                    //workSheet.Cells["A1"].Style.WrapText = true;
                    //workSheet.Cells["A1:L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //workSheet.Column(1).AutoFit();
                    //workSheet.InsertColumn(1, 1);
                    workSheet.InsertRow(1, 1);
                    //workSheet.Column(1).Width = 50;
                }



                result = package.GetAsByteArray();
            }

            return result;
        }

        public static byte[] ExportExcel1<T>(List<T> data, string Heading = "", string Heading1 = "", string Heading2 = "", string Heading3 = "", bool showSlno = false, params string[] ColumnsToTake)
        {
            return ExportExcel1(ListToDataTable<T>(data), Heading, Heading1, Heading2, Heading3, showSlno, ColumnsToTake);
        }

    }

 }

