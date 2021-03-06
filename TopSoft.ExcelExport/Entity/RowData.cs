﻿using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TopSoft.ExcelExport.Entity
{
    public class RowData
    {
        public Row Row { get; set; }
        public List<ExcelCell> Cells { get; set; }

        public RowData()
        {
            Row = new Row();
            Cells = new List<ExcelCell>();
        }

        public RowData(Row row, List<ExcelCell> cells)
        {
            if(row == null)
            {
                throw new ArgumentNullException("row");
            }

            if(cells == null)
            {
                throw new ArgumentNullException("cells");
            }

            Row = row;
            Cells = cells;
        }
    }
}
