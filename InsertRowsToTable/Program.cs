using System;
using System.Collections.Generic;
using System.Linq;

namespace InsertRowsToTable
{
    class Program
    {
        public static void Main(string[] args)
        {
            var vm = new ViewModel
                     {
                         固定類型s = new Row[]
                                 {
                                     new Row
                                     {
                                         欄位標題 = "欄位值 A",
                                         類型s = new 類型[]
                                               {
                                                   new 類型 { Name = "低", 數據 = new 數據 { 欄位標題1 = "欄位值 A1", 欄位標題2 = "欄位值 A2", 欄位標題3 = "欄位值 A3" } },
                                                   new 類型 { Name = "中", 數據 = new 數據 { 欄位標題1 = "欄位值 A4", 欄位標題2 = "欄位值 A5", 欄位標題3 = "欄位值 A6" } },
                                                   new 類型 { Name = "高", 數據 = new 數據 { 欄位標題1 = "欄位值 A7", 欄位標題2 = "欄位值 A8", 欄位標題3 = "欄位值 A9" } },
                                               }
                                     },
                                     new Row
                                     {
                                         欄位標題 = "欄位值 B",
                                         類型s = new 類型[]
                                               {
                                                   new 類型 { Name = "低", 數據 = new 數據 { 欄位標題1 = "欄位值 B1", 欄位標題2 = "欄位值 B2", 欄位標題3 = "欄位值 B3" } },
                                                   new 類型 { Name = "中", 數據 = new 數據 { 欄位標題1 = "欄位值 B4", 欄位標題2 = "欄位值 B5", 欄位標題3 = "欄位值 B6" } },
                                                   new 類型 { Name = "高", 數據 = new 數據 { 欄位標題1 = "欄位值 B7", 欄位標題2 = "欄位值 B8", 欄位標題3 = "欄位值 B9" } },
                                               }
                                     },
                                     new Row
                                     {
                                         欄位標題 = "欄位值 C",
                                         類型s = new 類型[]
                                               {
                                                   new 類型 { Name = "低", 數據 = new 數據 { 欄位標題1 = "欄位值 C1", 欄位標題2 = "欄位值 C2", 欄位標題3 = "欄位值 C3" } },
                                                   new 類型 { Name = "中", 數據 = new 數據 { 欄位標題1 = "欄位值 C4", 欄位標題2 = "欄位值 C5", 欄位標題3 = "欄位值 C6" } },
                                                   new 類型 { Name = "高", 數據 = new 數據 { 欄位標題1 = "欄位值 C7", 欄位標題2 = "欄位值 C8", 欄位標題3 = "欄位值 C9" } },
                                               }
                                     },
                                 },
                         手輸類型s = new Row[]
                                 {
                                     new Row
                                     {
                                         欄位標題 = "手輸欄位值 A",
                                         類型s = new 類型[]
                                               {
                                                   new 類型 { Name = "低", 數據 = new 數據 { 欄位標題1 = "手輸欄位值 A1", 欄位標題2 = "手輸欄位值 A2", 欄位標題3 = "手輸欄位值 A3" } },
                                                   new 類型 { Name = "中", 數據 = new 數據 { 欄位標題1 = "手輸欄位值 A4", 欄位標題2 = "手輸欄位值 A5", 欄位標題3 = "手輸欄位值 A6" } },
                                                   new 類型 { Name = "高", 數據 = new 數據 { 欄位標題1 = "手輸欄位值 A7", 欄位標題2 = "手輸欄位值 A8", 欄位標題3 = "手輸欄位值 A9" } },
                                               },
                                     },
                                     new Row
                                     {
                                         欄位標題 = "手輸欄位值 B",
                                         類型s = new 類型[]
                                               {
                                                   new 類型 { Name = "低", 數據 = new 數據 { 欄位標題1 = "手輸欄位值 B1", 欄位標題2 = "手輸欄位值 B2", 欄位標題3 = "手輸欄位值 B3" } },
                                                   new 類型 { Name = "中", 數據 = new 數據 { 欄位標題1 = "手輸欄位值 B4", 欄位標題2 = "手輸欄位值 B5", 欄位標題3 = "手輸欄位值 B6" } },
                                                   new 類型 { Name = "高", 數據 = new 數據 { 欄位標題1 = "手輸欄位值 B7", 欄位標題2 = "手輸欄位值 B8", 欄位標題3 = "手輸欄位值 B9" } },
                                               },
                                     },
                                 }
                     };

            var table1 = Table塞資料方式_由上而下(GenerateTemplateTable(), vm);
            PrintTable(table1);
            Console.WriteLine("----------------------------------------------------");
            var table2 = Table塞資料方式_由下而上(GenerateTemplateTable(), vm);
            PrintTable(table2);
        }

        private static Table GenerateTemplateTable()
        {
            return new Table
                   {
                       TableRows = new List<TableRow>
                                   {
                                       // Title
                                       new TableRow { 欄位標題 = "欄位標題", 類型 = "類型", 欄位標題1 = "欄位標題1", 欄位標題2 = "欄位標題2", 欄位標題3 = "欄位標題3", },

                                       // 中間為資料列：先放 固定類型s 再放 手輸類型s

                                       // 最下面為備註列
                                       new TableRow { 欄位標題 = "備註", 類型 = String.Empty, 欄位標題1 = String.Empty, 欄位標題2 = String.Empty, 欄位標題3 = String.Empty, },
                                   }
                   };
        }

        private static Table Table塞資料方式_由上而下(Table templateTable, ViewModel vm)
        {
            var rowDataStartIndex = 1;

            foreach (var row in vm.固定類型s)
            {
                for (int index = 0; index < row.類型s.Length; index++)
                {
                    var 類型 = row.類型s[index];

                    templateTable.TableRows.Insert(rowDataStartIndex++, ToTableRow(row, 類型));
                }
            }

            foreach (var row in vm.手輸類型s)
            {
                for (int index = 0; index < row.類型s.Length; index++)
                {
                    var 類型 = row.類型s[index];

                    templateTable.TableRows.Insert(rowDataStartIndex++, ToTableRow(row, 類型));
                }
            }

            return templateTable;
        }

        private static Table Table塞資料方式_由下而上(Table templateTable, ViewModel vm)
        {
            var rowDataStartIndex = 1;

            foreach (var row in vm.手輸類型s.Reverse())
            {
                for (int index = row.類型s.Length - 1; index >= 0; index--)
                {
                    var 類型 = row.類型s[index];

                    templateTable.TableRows.Insert(rowDataStartIndex, ToTableRow(row, 類型));
                }
            }

            foreach (var row in vm.固定類型s.Reverse())
            {
                for (int index = row.類型s.Length - 1; index >= 0; index--)
                {
                    var 類型 = row.類型s[index];

                    templateTable.TableRows.Insert(rowDataStartIndex, ToTableRow(row, 類型));
                }
            }

            return templateTable;
        }

        private static TableRow ToTableRow(Row row, 類型 類型)
        {
            return new TableRow
                   {
                       欄位標題  = row.欄位標題,
                       類型    = 類型.Name,
                       欄位標題1 = 類型.數據.欄位標題1,
                       欄位標題2 = 類型.數據.欄位標題2,
                       欄位標題3 = 類型.數據.欄位標題3,
                   };
        }

        private static void PrintTable(Table templateTable)
        {
            foreach (var row in templateTable.TableRows)
            {
                Console.WriteLine($"{row.欄位標題}\t{row.類型}\t{row.欄位標題1}\t{row.欄位標題2}\t{row.欄位標題3}");
            }
        }
    }

    public class ViewModel
    {
        public Row[] 固定類型s { get; set; }

        public Row[] 手輸類型s { get; set; }
    }

    public class Row
    {
        public string 欄位標題 { get; set; }

        public 類型[] 類型s { get; set; }
    }

    public class 類型
    {
        public string Name { get; set; }

        public 數據 數據 { get; set; }
    }

    public class 數據
    {
        public string 欄位標題1 { get; set; }

        public string 欄位標題2 { get; set; }

        public string 欄位標題3 { get; set; }
    }

    public class Table
    {
        public List<TableRow> TableRows { get; set; }
    }

    public class TableRow
    {
        public string 欄位標題 { get; set; }

        public string 類型 { get; set; }

        public string 欄位標題1 { get; set; }

        public string 欄位標題2 { get; set; }

        public string 欄位標題3 { get; set; }
    }
}
