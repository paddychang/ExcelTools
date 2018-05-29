using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelTools
{
    public class DataSetModify
    {
        public DataTable ReadCsvFile(string filePath)
        {
            string Fulltext;
            DataTable dtCsv = new DataTable();

            using (StreamReader sr = new StreamReader(filePath))
            {
                while (!sr.EndOfStream)
                {
                    Fulltext = sr.ReadToEnd().ToString(); // Read full file text  
                    string[] rows = Fulltext.Split('\n'); // Split full file text into rows
                    for (int i = 0; i < rows.Count() - 1; i++)
                    {
                        string[] rowValues = rows[i].Split(','); // Split each row by comma

                        if (i == 0)
                        {
                            DataRow dr = dtCsv.NewRow();
                            for (int j = 0; j < rowValues.Count(); j++)
                            {
                                dtCsv.Columns.Add();
                            }
                            for (int k = 0; k < rowValues.Count(); k++)
                            {
                                if (rowValues[k].ToString() != "")
                                {
                                    dtCsv.Columns[k].ColumnName = rowValues[k].ToString();
                                    dr[k] = rowValues[k].ToString();
                                }
                                else
                                {
                                    dtCsv.Columns[k].ColumnName = "Coulumn" + k;
                                    dr[k] = "Coulumn" + k;
                                }
                            }
                        }
                        else
                        {
                            DataRow dr = dtCsv.NewRow();
                            int count = 0;
                            for (int k = 0; k < rowValues.Count(); k++)
                            {
                                if (rowValues[k].ToString() == null)
                                    dr[k] = DBNull.Value;
                                else if (rowValues[k].ToString().Contains("\""))
                                {
                                    string tmp = rowValues[k].ToString();
                                    for (int m = k + 1; m < rowValues.Count(); m++)
                                    {
                                        if (rowValues[m].ToString().Contains("\""))
                                        {
                                            tmp = tmp + "," + rowValues[m].ToString();
                                            tmp = tmp.Replace("\"", "");
                                            count++;
                                            k = m;
                                            break;
                                        }
                                        else
                                        {
                                            tmp = tmp + "," + rowValues[m].ToString();
                                            count++;
                                        }
                                    }
                                    dr[k - count] = tmp;
                                }
                                else
                                {
                                    dr[k - count] = rowValues[k].ToString();
                                }
                            }
                            dtCsv.Rows.Add(dr);
                        }
                    }
                }
                Console.WriteLine(dtCsv.Columns.Count);
                return dtCsv;
            }
        }

        public DataSet ReadExcelFile(string filePath)
        {
            DataSet ds = new DataSet();
            FileStream input = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = null;

            if (filePath.EndsWith(".xlsx") || filePath.EndsWith(".xlsm"))
            {
                // Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(input);
            }
            else
            {
                // Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(input);
            }
            ds = excelReader.AsDataSet();
            excelReader.Close();
            return ds;
        }

        public void PrintDataset(DataSet ds)
        {
            Console.WriteLine("Tables in '{0}' DataSet.\n", ds.DataSetName);
            foreach (DataTable dt in ds.Tables)
            {
                Console.WriteLine("{0} Table.\n", dt.TableName);
                for (int curCol = 0; curCol < dt.Columns.Count; curCol++)
                {
                    Console.Write(dt.Columns[curCol].ColumnName.Trim() + "\t");
                }
                for (int curRow = 0; curRow < dt.Rows.Count; curRow++)
                {
                    for (int curCol = 0; curCol < dt.Columns.Count; curCol++)
                    {
                        Console.Write(dt.Rows[curRow][curCol].ToString().Trim() + "\t");
                    }
                    Console.WriteLine();
                }
            }
        }

        public void PrintTbalesName(DataSet ds)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                Console.WriteLine(ds.Tables[i].TableName.ToString());
            }
        }

        public void PrintDataTable(DataTable table)
        {
            foreach (DataRow dataRow in table.Rows)
            {
                foreach (var row in dataRow.ItemArray)
                {
                    Console.WriteLine(row);
                }
            }
        }

        public void PrintList<T>(IEnumerable<T> list)
        {
            foreach (var item in list)
                Console.WriteLine(item);
        }

        public DataSet EmptyRowsRemoving(DataSet ds, int nonEmptyColumnIndex)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                for (int j = ds.Tables[i].Rows.Count - 1; j > -1; j--)
                {
                    if (ds.Tables[i].Rows[j][nonEmptyColumnIndex].ToString() == "")
                    {
                        ds.Tables[i].Rows.RemoveAt(j);
                        ds.Tables[i].AcceptChanges();
                    }
                }
                for (int k = ds.Tables[i].Columns.Count - 1; k > -1; k--)
                {
                    if (ds.Tables[i].Rows.Count > 1)
                    {
                        if (ds.Tables[i].Rows[1][k].ToString() == "")
                        {
                            ds.Tables[i].Columns.RemoveAt(k);
                            ds.Tables[i].AcceptChanges();
                        }
                    }
                }
            }
            return ds;
        }

        public DataSet EmptyColunmsRemoving(DataSet ds)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                for (int k = ds.Tables[i].Columns.Count - 1; k > -1; k--)
                {
                    if (ds.Tables[i].Rows.Count > 1)
                    {
                        if (ds.Tables[i].Rows[1][k].ToString() == "")
                        {
                            ds.Tables[i].Columns.RemoveAt(k);
                            ds.Tables[i].AcceptChanges();
                        }
                    }
                }
            }
            return ds;
        }

        public DataTable SetColumnName(DataTable dt)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = dt.Rows[0][i].ToString();
                dt.AcceptChanges();
            }
            return dt;
        }

        public DataSet SetColumnName(DataSet ds)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                DataColumnCollection col = ds.Tables[i].Columns;
                if (col.Count != 0)
                {
                    for (int j = 0; j < ds.Tables[i].Columns.Count; j++)
                    {
                        if (ds.Tables[i].Columns[j].ToString() == "")
                        {
                            ds.Tables[i].Columns[j].ColumnName = "Column" + i;
                            ds.AcceptChanges();
                        }
                        else
                        {
                            Console.WriteLine(ds.Tables[i].Rows[0][j].ToString());
                            ds.Tables[i].Columns[j].ColumnName = ds.Tables[i].Rows[0][j].ToString();
                            ds.AcceptChanges();
                        }
                    }
                }
            }
            return ds;
        }

        public DataSet SetTableName(DataSet ds, string tableName, int index)
        {
            ds.Tables[index].TableName = tableName;
            ds.AcceptChanges();
            return ds;
        }

        public DataTable SetColumnTypes(DataTable dt, Type type, int columnIndex)
        {
            DataTable dtNew = new DataTable();

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (columnIndex == i)
                    dtNew.Columns.Add(dt.Columns[i].ColumnName, type);
                else
                    dtNew.Columns.Add(dt.Columns[i].ColumnName, typeof(string));
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dtNew.Rows.Add();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j].ToString() == "")
                    {
                        if (type == typeof(int) && columnIndex == j)
                            dtNew.Rows[i][j] = 0;
                        else if (type == typeof(float) && columnIndex == j)
                            dtNew.Rows[i][j] = 0.0;
                        else
                            dtNew.Rows[i][j] = DBNull.Value;
                    }
                    else
                        dtNew.Rows[i][j] = dt.Rows[i][j];
                }
            }
            dt = null;
            dt = dtNew.Copy();
            dt.AcceptChanges();
            return dt;
        }

        public DataSet SetColumnTypes(DataSet ds, Type type, int tableIndex, int columnIndex)
        {
            DataTable dtNew = new DataTable();
            DataSet dsNew = new DataSet();

            for (int i = 0; i < ds.Tables[tableIndex].Columns.Count; i++)
            {
                if (columnIndex == i)
                    dtNew.Columns.Add(ds.Tables[tableIndex].Columns[i].ColumnName, type);
                else
                    dtNew.Columns.Add(ds.Tables[tableIndex].Columns[i].ColumnName, typeof(string));
            }
            for (int i = 0; i < ds.Tables[tableIndex].Rows.Count; i++)
            {
                dtNew.Rows.Add();
                for (int j = 0; j < ds.Tables[tableIndex].Columns.Count; j++)
                {
                    if (ds.Tables[tableIndex].Rows[i][j].ToString() == "")
                    {
                        if (ds.Tables[tableIndex].Columns[j].DataType == typeof(int) && columnIndex == j)
                            dtNew.Rows[i][j] = 0;
                        else if (ds.Tables[tableIndex].Columns[j].DataType == typeof(float) && columnIndex == j)
                            dtNew.Rows[i][j] = 0.0;
                        else
                            dtNew.Rows[i][j] = DBNull.Value;
                    }
                    else
                        dtNew.Rows[i][j] = ds.Tables[tableIndex].Rows[i][j];
                }
            }
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                if (i == tableIndex)
                {
                    dtNew.TableName = ds.Tables[i].TableName.ToString();
                    dsNew.Tables.Add(dtNew);
                }
                else
                {
                    dsNew.Tables.Add(ds.Tables[i]);
                }
            }
            dsNew.AcceptChanges();
            ds = null;
            ds = dsNew.Copy();
            ds.AcceptChanges();
            return ds;
        }

        public DataTable SetColumnTypes(DataTable dt, Type type, List<int> columnList)
        {
            DataTable dtNew = new DataTable();

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                foreach (var item in columnList)
                {
                    if (item == i)
                        dtNew.Columns.Add(dt.Columns[i].ColumnName, type);
                    else
                        dtNew.Columns.Add(dt.Columns[i].ColumnName, typeof(string));
                }
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dtNew.Rows.Add();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j].ToString() == "")
                    {
                        foreach (var item in columnList)
                        {
                            if (dt.Columns[j].DataType == typeof(int) && item == j)
                                dtNew.Rows[i][j] = 0;
                            else if (dt.Columns[j].DataType == typeof(float) && item == j)
                                dtNew.Rows[i][j] = 0.0;
                            else
                                dtNew.Rows[i][j] = DBNull.Value;
                        }
                    }
                    else
                        dtNew.Rows[i][j] = dt.Rows[i][j];
                }
            }
            dt = null;
            dt = dtNew.Copy();
            dt.AcceptChanges();
            return dt;
        }

        public DataSet SetColumnTypes(DataSet ds, Type type, int tableIndex, List<int> columnList)
        {
            DataTable dtNew = new DataTable();
            DataSet dsNew = new DataSet();
            bool flag = true;

            for (int i = 0; i < ds.Tables[tableIndex].Columns.Count; i++)
            {
                flag = true;
                foreach (int item in columnList)
                {
                    if (item == i)
                    {
                        dtNew.Columns.Add(ds.Tables[tableIndex].Columns[i].ColumnName, type);
                        flag = false;
                    }
                }
                if (flag)
                {
                    dtNew.Columns.Add(ds.Tables[tableIndex].Columns[i].ColumnName, typeof(string));
                    flag = true;
                }

            }

            for (int i = 0; i < ds.Tables[tableIndex].Rows.Count; i++)
            {
                dtNew.Rows.Add();
                for (int j = 0; j < ds.Tables[tableIndex].Columns.Count; j++)
                {
                    if (ds.Tables[tableIndex].Rows[i][j].ToString() == "")
                    {
                        if (ds.Tables[tableIndex].Columns[j].DataType == typeof(int))
                            dtNew.Rows[i][j] = 0;
                        else if (ds.Tables[tableIndex].Columns[j].DataType == typeof(float))
                            dtNew.Rows[i][j] = 0.0;
                        else
                            dtNew.Rows[i][j] = DBNull.Value;
                    }
                    else
                    {
                        dtNew.Rows[i][j] = ds.Tables[tableIndex].Rows[i][j];
                    }
                }
            }

            for (int i = 0; i < ds.Tables.Count; i++)
            {
                if (i == tableIndex)
                {
                    dtNew.TableName = ds.Tables[i].TableName.ToString();
                    dsNew.Tables.Add(dtNew);
                }
                else
                {
                    dsNew.Tables.Add(ds.Tables[i]);
                }
            }
            dsNew.AcceptChanges();
            ds = null;
            ds = dsNew.Copy();
            ds.AcceptChanges();
            return ds;
        }

        public List<string> GetListColumnName(DataTable dt)
        {
            List<string> colNmaeList = new List<string>();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                colNmaeList.Add(dt.Columns[i].ToString());
            }
            return colNmaeList;
        }

        public string GetTableName(DataSet ds, int index)
        {
            return ds.Tables[index].TableName.ToString();
        }

        public List<string> GetAllTableName(DataSet ds)
        {
            List<string> list = new List<string>();
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                list.Add(ds.Tables[i].TableName.ToString());
            }
            return list;
        }

        public DataSet RemoveTable(DataSet ds, string tableName)
        {
            if (ds.Tables.Contains(tableName) && ds.Tables.CanRemove(ds.Tables[tableName]))
                ds.Tables.Remove(ds.Tables[tableName]);
            ds.AcceptChanges();
            return ds;
        }

        public DataSet RemoveTable(DataSet ds, int tableIndex)
        {
            ds.Tables.RemoveAt(tableIndex);
            ds.AcceptChanges();
            return ds;
        }

        public DataSet RemoveColumn(DataSet ds, int tableIndex, int columnIndex)
        {
            ds.Tables[tableIndex].Columns.RemoveAt(columnIndex);
            ds.AcceptChanges();
            return ds;
        }

        public DataSet RemoveColumn(DataSet ds, string tableName, string columnNmae)
        {
            ds.Tables[tableName].Columns.Remove(columnNmae);
            ds.AcceptChanges();
            return ds;
        }

        public DataTable RemoveColumn(DataTable dt, int columnIndex)
        {
            dt.Columns.RemoveAt(columnIndex);
            return dt;
        }

        public DataTable RemoveColumn(DataTable dt, string columnNmae)
        {
            dt.Columns.Remove(columnNmae);
            return dt;
        }

        public DataSet RemoveRow(DataSet ds, string tableName, int rowIndex)
        {
            ds.Tables[tableName].Rows[rowIndex].Delete();
            ds.AcceptChanges();
            return ds;
        }

        public DataSet RemoveRow(DataSet ds, int tableIndex, int rowIndex)
        {
            ds.Tables[tableIndex].Rows[rowIndex].Delete();
            ds.AcceptChanges();
            return ds;
        }

        public DataTable InsertColumn(DataTable dt, string columnName, Type type, int columnIndex)
        {
            DataColumn Col = dt.Columns.Add(columnName, type);
            Col.SetOrdinal(columnIndex);
            dt.AcceptChanges();
            return dt;
        }

        public DataSet InsertColumn(DataSet ds, string columnName, Type type, int tableIndex, int columnIndex)
        {
            DataColumn Col = ds.Tables[tableIndex].Columns.Add(columnName, type);
            Col.SetOrdinal(columnIndex);
            ds.AcceptChanges();
            return ds;
        }

        public DataSet InsertRow(DataSet ds, DataRow dr, int tableIndex, int rowIndex)
        {
            ds.Tables[tableIndex].Rows.InsertAt(dr, rowIndex);
            ds.AcceptChanges();
            return ds;
        }

        public DataTable InsertRow(DataTable dt, DataRow dr, int rowIndex)
        {
            dt.Rows.InsertAt(dr, rowIndex);
            dt.AcceptChanges();
            return dt;
        }

        public void GenerateExcelFile(DataSet ds, string paramFileFullPath, bool printCoulumnName)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    DataColumnCollection col = ds.Tables[i].Columns;
                    if (col.Count > 0)
                    {
                        ExcelWorksheet ws = package.Workbook.Worksheets.Add(ds.Tables[i].TableName.ToString());
                        ws.Cells["A1"].LoadFromDataTable(ds.Tables[i], printCoulumnName);
                    }
                }
                package.SaveAs(new FileInfo(paramFileFullPath));
            }
        }

    }
}