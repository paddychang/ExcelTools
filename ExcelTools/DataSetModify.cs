﻿using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelTools
{
    public static class DataSetModify
    {
        public static DataTable ReadCsvFile(string filePath)
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
                        {
                            if (i == 0)
                            {
                                DataRow dr = dtCsv.NewRow();
                                for (int j = 0; j < 37; j++)
                                {
                                    dtCsv.Columns.Add();
                                }
                                for (int j = 0; j < rowValues.Count(); j++)
                                {
                                    if (rowValues[j].ToString() != "")
                                    {
                                        dtCsv.Columns[j].ColumnName = rowValues[j].ToString();
                                        dr[j] = rowValues[j].ToString();
                                    }
                                    else
                                        dtCsv.Columns[j].ColumnName = "Coulumn" + j;
                                    dr[j] = "Coulumn" + j;
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
                }
                return dtCsv;
            }
        }

        public static DataSet ReadExcelFile(string filePath)
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

        public static void PrintDataset(DataSet ds)
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

        public static void PrintDataTable(DataTable table)
        {
            foreach (DataRow dataRow in table.Rows)
            {
                foreach (var row in dataRow.ItemArray)
                {
                    Console.WriteLine(row);
                }
            }
        }

        public static void PrintList<T>(IEnumerable<T> list)
        {
            foreach (var item in list)
                Console.WriteLine(item);
        }

        public static DataSet EmptyRowsRemoving(DataSet ds, int nonEmptyColumnIndex)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                for (int j = ds.Tables[i].Rows.Count - 1; j > 0; j--)
                {
                    if (ds.Tables[i].Rows[j][nonEmptyColumnIndex].ToString() == "")
                    {
                        ds.Tables[i].Rows.RemoveAt(j);
                        ds.Tables[i].AcceptChanges();
                    }
                }
                for (int k = ds.Tables[i].Columns.Count - 1; k > 0; k--)
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

        public static DataSet EmptyColunmsRemoving(DataSet ds)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                for (int k = ds.Tables[i].Columns.Count - 1; k > 0; k--)
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

        public static DataTable SetColumnName(DataTable dt)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = dt.Rows[0][i].ToString();
            }
            return dt;
        }

        public static DataSet SetColumnName(DataSet ds)
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                DataColumnCollection col = ds.Tables[i].Columns;
                if (col.Count != 0)
                {
                    for (int j = 0; j < ds.Tables[i].Columns.Count; j++)
                    {
                        ds.Tables[i].Columns[j].ColumnName = ds.Tables[i].Rows[0][j].ToString();
                    }
                }
            }
            return ds;
        }

        public static DataSet SetTableName(DataSet ds, string tableName, int index)
        {
            ds.Tables[index].TableName = tableName;
            return ds;
        }

        public static DataTable SetColumnTypes(DataTable dt, Type type, int columnIndex)
        {
            DataTable dtNew = new DataTable();

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (columnIndex == i)
                    dtNew.Columns.Add(dt.Columns[i].ColumnName, type);
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dtNew.Rows.Add();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j].ToString() == "")
                    {
                        if (type == typeof(int))
                            dtNew.Rows[i][j] = 0;
                        else if (type == typeof(float))
                            dtNew.Rows[i][j] = 0.0;
                        else if (type == typeof(string))
                            dtNew.Rows[i][j] = DBNull.Value;
                    }
                    else
                        dtNew.Rows[i][j] = dt.Rows[i][j];
                }
            }
            dt = null;
            dt = dtNew.Copy();
            return dt;
        }

        public static DataTable SetColumnTypes(DataTable dt, Type type, List<int> columnList)
        {
            DataTable dtNew = new DataTable();

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                foreach (var item in columnList)
                {
                    if (item == i)
                        dtNew.Columns.Add(dt.Columns[i].ColumnName, type);
                }
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dtNew.Rows.Add();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j].ToString() == "")
                    {
                        if (type == typeof(int))
                            dtNew.Rows[i][j] = 0;
                        else if (type == typeof(float))
                            dtNew.Rows[i][j] = 0.0;
                        else if (type == typeof(string))
                            dtNew.Rows[i][j] = DBNull.Value;
                    }
                    else
                        dtNew.Rows[i][j] = dt.Rows[i][j];
                }
            }
            dt = null;
            dt = dtNew.Copy();
            return dt;
        }

        public static List<string> GetListColumnName(DataTable dt)
        {
            List<string> colNmaeList = new List<string>();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                colNmaeList.Add(dt.Columns[i].ToString());
            }
            return colNmaeList;
        }

        public static string GetTableName(DataSet ds, int index)
        {
            return ds.Tables[index].TableName.ToString();
        }

        public static List<string> GetAllTableName(DataSet ds)
        {
            List<string> list = new List<string>();
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                list.Add(ds.Tables[i].TableName.ToString());
            }
            return list;
        }

        public static DataSet RemoveTable(DataSet ds, string tableName)
        {
            if (ds.Tables.Contains(tableName) && ds.Tables.CanRemove(ds.Tables[tableName]))
                ds.Tables.Remove(ds.Tables[tableName]);
            return ds;
        }

        public static DataSet RemoveColumn(DataSet ds, int tableIndex, int columnIndex)
        {
            ds.Tables[tableIndex].Columns.RemoveAt(columnIndex);
            return ds;
        }

        public static DataSet RemoveColumn(DataSet ds, string tableName, string columnNmae)
        {
            ds.Tables[tableName].Columns.Remove(columnNmae);
            return ds;
        }

        public static DataTable RemoveColumn(DataTable dt, int tableIndex, int columnIndex)
        {
            dt.Columns.RemoveAt(columnIndex);
            return dt;
        }

        public static DataTable RemoveColumn(DataTable dt, string tableName, string columnNmae)
        {
            dt.Columns.Remove(columnNmae);
            return dt;
        }

        public static DataTable InsertColumn(DataTable dt, string columnName, Type type, int columnIndex)
        {
            DataColumn Col = dt.Columns.Add(columnName, type);
            Col.SetOrdinal(columnIndex);
            return dt;
        }

        public static DataSet InsertColumn(DataSet ds, string columnName, Type type, int columnIndex, int tableIndex)
        {
            DataColumn Col = ds.Tables[tableIndex].Columns.Add(columnName, type);
            Col.SetOrdinal(columnIndex);
            return ds;
        }

        public static void GenerateExcelFile(DataSet ds, string paramFileFullPath, bool printCoulumnName)
        {
            Console.WriteLine(ds.Tables[0].Columns[0].ColumnName.ToString());
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

