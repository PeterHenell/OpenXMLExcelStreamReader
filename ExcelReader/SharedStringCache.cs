using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTest
{
    public class SharedStringCache : IDisposable
    {
        private SharedStringItem[] _sharedStrings;
        private IEnumerator<SharedStringItem> _sharedStringEnumerator;
        private int _currentSharedStringIndex;

        public SharedStringCache(WorkbookPart _workbookPart)
        {
            this._sharedStrings = new SharedStringItem[_workbookPart.SharedStringTablePart.SharedStringTable.Count];
            this._sharedStringEnumerator = _workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().GetEnumerator();
            this._currentSharedStringIndex = 0;
        }

        private void ReadSharedStringsUpTo(int shindex)
        {
            while (_currentSharedStringIndex <= shindex)
            {
                _sharedStringEnumerator.MoveNext();
                _sharedStrings[_currentSharedStringIndex++] = _sharedStringEnumerator.Current;
            }
        }

        public void Dispose()
        {
            if (_sharedStringEnumerator != null)
            {
                _sharedStringEnumerator.Dispose();
            }
        }
    }
}
