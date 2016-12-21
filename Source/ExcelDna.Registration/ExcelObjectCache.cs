using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelDna.Integration;

namespace ExcelDna.Registration
{
    /// <summary>
    /// Any types decorated with this attribute will be marshalled to Excel by reference through the ParameterConversionConfiguration.ReferenceMashaller object
    /// </summary>
    [AttributeUsage(AttributeTargets.Class|AttributeTargets.Interface, Inherited = true)]
    public class ExcelMarshalByRefAttribute : Attribute
    {
        public ExcelMarshalByRefAttribute()
        {
        }
    }

    /// <summary>
    /// This class implements an object cache to marshall object identities back and forth the Excel workbook.
    /// It was deisgned as a singleton because it uses thread local storage to support multi-threaded scenarios in an efficient way.
    /// </summary>
    public class ExcelObjectCache : IReferenceMarshaller
    {
        private static volatile ExcelObjectCache _instance;
        private static readonly object SyncRoot = new Object();

        private ExcelObjectCache() { }

        /// <summary>
        /// The singleton instance of this class
        /// </summary>
        public static ExcelObjectCache Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (SyncRoot)
                    {
                        if (_instance == null)
                            _instance = new ExcelObjectCache();
                    }
                }

                return _instance;
            }
        }

        /// <summary>
        /// This tracks which object id's have been create din which cell
        /// </summary>
        private Dictionary<ExcelReference, HashSet<int>> BirthCell = new Dictionary<ExcelReference, HashSet<int>>();

        /// <summary>
        /// This is where we actually look objects up
        /// </summary>
        private Dictionary<int, object> IdLookup = new Dictionary<int, object>();

        /// <summary>
        /// This acculates all objects created during the possibly nested evaluations that happen in a cell.
        /// It is cleared every time the thread sees a different cell.
        /// </summary>
        [ThreadStatic]
        private static Dictionary<int, object> TempObjects = new Dictionary<int, object>();

        /// <summary>
        /// This is the cell being handled by the current thread
        /// </summary>
        [ThreadStatic]
        private static ExcelReference _currentCell;

        /// <summary>
        /// An atomically incremented counted that provides the object identifiers.
        /// </summary>
        private static int _idCounter = 0;

        /// NullXlRef is what we get for calls from VBA
        private static readonly ExcelReference NullXlRef = new ExcelReference(-1, -1);

        private static readonly char Separator = '@';

        /// <inheritdoc/>
        public object Lookup(string idString)
        {
            int idPos = idString.IndexOf('@');
            if (idPos<0)
                throw new ArgumentException($"Object id '{idString}' is not in the format <name>{Separator}<id>");
            int id = int.Parse(idString.Substring(idPos + 1));
            object result;
            if (!TempObjects.TryGetValue(id, out result))
            {
                lock (BirthCell)
                {
                    IdLookup.TryGetValue(id, out result);
                }
            }
            return result;
        }

        /// <inheritdoc/>
        public string Store(object o)
        {
            int id = Interlocked.Increment(ref _idCounter);
            lock (BirthCell)
            {
                TempObjects[id] = o;
                HashSet<int> objectsBornInTheCurrentCell = BirthCell[_currentCell]; // must succeed if SetCurrentCell has been called
                objectsBornInTheCurrentCell.Add(id);
                IdLookup[id] = o;
            }
            return $"{o.GetType().Name}{Separator}{id}";
        }

        /// <inheritdoc/>
        public void SetCurrentCell()
        {
            ExcelReference previousCell = _currentCell;
            ExcelReference thisCell = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            thisCell = thisCell ?? NullXlRef;
            lock (BirthCell)
            {
                // If we are now processing a different cell, then we get rid of any old temp objects
                if (thisCell != previousCell)
                {
                    HashSet<int> objectsCreatedAtThePreviousCell = null;
                    if (previousCell != null)
                        BirthCell.TryGetValue(previousCell, out objectsCreatedAtThePreviousCell);
                    foreach (int tempId in TempObjects.Keys)
                        if (objectsCreatedAtThePreviousCell==null || !objectsCreatedAtThePreviousCell.Contains(tempId))
                            IdLookup.Remove(tempId);
                    TempObjects.Clear();
                    _currentCell = thisCell;
                }

                HashSet<int> objectsCreatedAtTheCurrentCell = null;
                BirthCell.TryGetValue(thisCell, out objectsCreatedAtTheCurrentCell);

                if (objectsCreatedAtTheCurrentCell == null)
                {
                    objectsCreatedAtTheCurrentCell = new HashSet<int>();
                    BirthCell[thisCell] = objectsCreatedAtTheCurrentCell;
                }
                else
                {
                    foreach (var id in objectsCreatedAtTheCurrentCell)
                        IdLookup.Remove(id);
                    objectsCreatedAtTheCurrentCell.Clear();
                }
            }
        }
    }
}
