using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleSpreadsheet = Google.Apis.Sheets.v4.Data.Spreadsheet;
using GoogleSheet = Google.Apis.Sheets.v4.Data.Sheet;

namespace Sheeeets
{
    public class Sheet
    {
        public static Tuple<int, int> RowCol(int row, int col)
        {
            return new Tuple<int, int>(row, col);
        }

        public static void AsIntAddress(string addr, out int row, out int col)
        {
            var abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var rxcolpat = new Regex(@"^(\w+?)\d");
            var rxrowpat = new Regex(@"\w(\d+?)$");
            row = int.Parse(rxrowpat.Match(addr).Groups[0].Value) - 1;

            var cols = rxcolpat.Match(addr).Groups[0].Value;
            int digit = 0;
            col = 0;
            foreach (var c in cols)
            {
                var pos = abc.IndexOf(c, 0);
                col += pos * (int)Math.Pow(abc.Length, digit);
                digit++;
            }
        }

        public Spreadsheet ParentSpreadsheet { get; private set; }
        public GoogleSheet GoogleSheet { get; private set; }

        public Sheet(Spreadsheet pss, GoogleSheet gss)
        {
            ParentSpreadsheet = pss;
            GoogleSheet = gss;
        }

        private readonly Dictionary<Tuple<int, int>, CellData> _cellCache = new Dictionary<Tuple<int, int>, CellData>();

        private int _occupiedGridWidth = -1;
        private int _occupiedGridHeight = -1;
        public int OccupiedGridWidth
        {
            get
            {
                if (!ParentSpreadsheet.DataAvailable) return -1;
                if (_occupiedGridWidth != -1) return _occupiedGridWidth;
                DiscoverOccupiedArea();
                return _occupiedGridWidth;
            }
        }

        public int OccupiedGridHeight
        {
            get
            {
                if (!ParentSpreadsheet.DataAvailable) return -1;
                if (_occupiedGridHeight != -1) return _occupiedGridHeight;
                DiscoverOccupiedArea();
                return _occupiedGridHeight;
            }
        }

        private void DiscoverOccupiedArea()
        {
            DiscoverCells();
            _occupiedGridHeight = _cellCache.Keys.Select(k => k.Item1).Concat(new[] { 0 }).Max() + 1;
            _occupiedGridWidth = _cellCache.Keys.Select(k => k.Item2).Concat(new[] { 0 }).Max() + 1;
        }

        private bool _cellsDiscovered;
        private void DiscoverCells()
        {
            if(_cellsDiscovered) return;
            foreach (var gd in GoogleSheet.Data)
            {
                int? rr = 0, cc = 0;
                rr = gd.StartRow;
                if (!rr.HasValue) rr = 0;
                foreach (var rd in gd.RowData)
                {
                    cc = gd.StartColumn;
                    if (!cc.HasValue) cc = 0;
                    foreach (var cd in rd.Values)
                    {
                        _cellCache.Add(RowCol(rr.Value, cc.Value), cd);
                        cc++;
                    }
                    rr++;
                }
            }
            _cellsDiscovered = true;
        }
        private CellData GetCell(int row, int col)
        {
            var cellid = new Tuple<int, int>(row, col);
            if (_cellCache.ContainsKey(cellid)) return _cellCache[cellid];
            DiscoverCells();
            return _cellCache.ContainsKey(cellid) ? _cellCache[cellid] : null;
        }

        public CellData this[Tuple<int, int> rowcol] => GetCell(rowcol.Item1, rowcol.Item2);
        public CellData this[int row, int col] => GetCell(row, col);

        public CellData this[string addr]
        {
            get
            {
                int row, col;
                AsIntAddress(addr, out row, out col);
                return GetCell(row, col);
            }
        }

        public CellData[][] this[string topleft, string bottomright]
        {
            get
            {
                int tlr, tlc, brr, brc;
                AsIntAddress(topleft, out tlr, out tlc);
                AsIntAddress(bottomright, out brr, out brc);
                return this[tlr, tlc, brr, brc];
            }
        }
        public CellData[][] this[int tlr, int tlc, int brr, int brc]
        {
            get
            {
                var res = new CellData[brr - tlr][];
                int ii = 0, jj = 0;
                for (int i = tlr; i < brr; i++)
                {
                    res[ii] = new CellData[brc - tlc];
                    jj = 0;
                    for (int j = tlc; i < brc; i++)
                    {
                        res[ii][jj] = GetCell(i, j);
                        jj++;
                    }
                    ii++;
                }
                return res;
            }
        }
    }
}
