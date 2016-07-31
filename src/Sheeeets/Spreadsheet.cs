using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleSpreadsheet = Google.Apis.Sheets.v4.Data.Spreadsheet;
using GoogleSheet = Google.Apis.Sheets.v4.Data.Sheet;

namespace Sheeeets
{
    public class Spreadsheet : IEnumerable<Sheet>
    {
        public string ID { get; private set; }

        public Session ParentSession { get; private set; }

        public SpreadsheetsResource.GetRequest Request { get; private set; }

        public GoogleSpreadsheet GoogleSpreadsheet { get; private set; }

        public bool Requesting { get; private set; }
        public bool DataAvailable { get; private set; }
        public bool RequestError { get; private set; }

        public int Count => GoogleSpreadsheet.Sheets.Count;

        private readonly Dictionary<int, Sheet> _sheetCache = new Dictionary<int, Sheet>();

        private Sheet GetSheet(int i)
        {
            if (_sheetCache.ContainsKey(i)) return _sheetCache[i];
            var sheet = new Sheet(this, GoogleSpreadsheet.Sheets[i]);
            _sheetCache.Add(i, sheet);
            return sheet;
        }

        public Sheet this[int index]
        {
            get { return GetSheet(index); }
            set { }
        }

        public Spreadsheet(string id, Session psession)
        {
            ID = id;
            ParentSession = psession;
            Request = psession.Service.Spreadsheets.Get(id);
            Request.IncludeGridData = true;
            Requesting = false;
        }

        public event EventHandler OnRequestCompleted;
        public event EventHandler<OnErrorEventArgs> OnRequestError;

        public void RequestSpreadsheetAsync()
        {
            Requesting = true;
            RequestError = false;
            Task.Run(() =>
            {
                try
                {
                    GoogleSpreadsheet = Request.Execute();
                    _sheetCache.Clear();
                    Requesting = false;
                    DataAvailable = true;
                    OnRequestCompleted?.Invoke(this, EventArgs.Empty);
                }
                catch (Exception e)
                {
                    RequestError = true;
                    Requesting = false;
                    var args = new OnErrorEventArgs
                    {
                        Error = e
                    };
                    OnRequestError?.Invoke(this, args);
                }
            });
        }

        public IEnumerator<Sheet> GetEnumerator()
        {
            if(!DataAvailable) yield break;
            for (int i = 0; i < GoogleSpreadsheet.Sheets.Count; i++)
            {
                yield return GetSheet(i);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        //public Sheet this[int i] => GoogleSpreadsheet.Sheets[i];
    }
}
