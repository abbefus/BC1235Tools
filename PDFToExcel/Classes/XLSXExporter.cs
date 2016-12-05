using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ABUtils
{
    public static class Utils
    {
        public static int GetNextEmptyRow(ExcelWorksheet sheet)
        {
            var row = sheet.Dimension.End.Row;
            while (row >= 1)
            {
                var range = sheet.Cells[row, 1, row, sheet.Dimension.End.Column];
                if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
                {
                    break;
                }
                row--;
            }
            return (row + 1);
        }

        public static IEnumerable<T> AsEnumerable<T>(T obj)
        {
            yield return obj;
        }

    }

    public class OpenWorkbook : IDisposable
    {
        private ExcelPackage package;
        public ExcelWorksheet ActiveWorksheet { get; private set; }
        public bool Disposed { get; private set; }

        public OpenWorkbook(string path)
        {
            Disposed = false;
            FileInfo fileinfo = new FileInfo(path);
            if (!fileinfo.Exists)
            {
                package = new ExcelPackage(fileinfo);
            }
        }

        public bool AddWorksheet(string wsname)
        {
            if (Disposed || package == null) return false;
            if (!ListWorksheets().Contains(wsname))
            {
                ActiveWorksheet = package.Workbook.Worksheets.Add(wsname);
                return true;
            }
            return false;
        }

        public bool SetActiveWorksheet(string wsname)
        {
            if (Disposed) return false;

            if (ListWorksheets().Contains(wsname)) 
            {
                ActiveWorksheet = package.Workbook.Worksheets[wsname];
                return true;
            }
            return false;
        }

        public bool UpdateRow(int rowindex, object[] row, int columnindex=1)
        {
            if (!Disposed && ActiveWorksheet != null)
            {
                ExcelRange range = ActiveWorksheet.Cells[rowindex, columnindex];
                range.LoadFromArrays(Utils.AsEnumerable(row));
                return true;
            }
            return false;
        }
        public bool UpdateRows(int rowindex, IEnumerable<object[]> row, int columnindex = 1)
        {
            if (!Disposed && ActiveWorksheet != null)
            {
                ExcelRange range = ActiveWorksheet.Cells[rowindex, columnindex];
                range.LoadFromArrays(row);
                return true;
            }
            return false;
        }

        public string[] ListWorksheets()
        {
            return package.Workbook.Worksheets.Select(x => x.Name).ToArray();
        }

        public void Save()
        {
            package.Save();

        } 

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            Disposed = true;
            if (disposing)
            {
                if (package != null) { package.Dispose(); }
            }
        }

        public ExcelNamedRangeCollection NamedRanges
        {
            get { return package.Workbook.Names; }
        }


    }
}
