using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelDna.Integration;

namespace ExcelDna.Registration
{
    [AttributeUsage(AttributeTargets.Class|AttributeTargets.Interface, Inherited = true)]
    public class ExcelMarshalByRefAttribute : Attribute
    {
        public ExcelMarshalByRefAttribute()
        {
        }
    }

    /// <summary>
    /// This class implements an object cache to marshall object identities back and forth the Excel s/s.
    /// The algorithm works as follows:
    /// 1. The new object is assigned an id and is places in a thread-local dictionary for temporary objects. This dictionary is cleaned every time
    /// </summary>
    public class ExcelObjectCache : IReferenceMarshaller
    {
        private static volatile ExcelObjectCache _instance;
        private static readonly object SyncRoot = new Object();

        private ExcelObjectCache() { }

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

        private struct CacheItem
        {
            public CacheItem(int theId, object theObject)
            {
                Id = theId;
                Object = theObject;
            }

            public readonly int Id;
            public readonly object Object;
        }

        private Dictionary<ExcelReference, CacheItem> Cache = new Dictionary<ExcelReference, CacheItem>();
        private Dictionary<int, object> IdLookup = new Dictionary<int, object>();

        [ThreadStatic]
        private static Dictionary<int, object> TempObjects = new Dictionary<int, object>();
        [ThreadStatic]
        private static ExcelReference _currentCaller;
        private static int _idCounter = 0;
        private static readonly ExcelReference NullXlRef = new ExcelReference(-1, -1);

        private static readonly char Separator = '@';

        public object Lookup(string idString)
        {
            int idPos = idString.IndexOf('@');
            if (idPos<0)
                throw new ArgumentException($"Object id '{idString}' is not in the format <name>{Separator}<id>");
            int id = int.Parse(idString.Substring(idPos + 1));
            object result;
            if (!TempObjects.TryGetValue(id, out result))
            {
                lock (Cache)
                {
                    IdLookup.TryGetValue(id, out result);
                }
            }
            return result;
        }

        public string Store(object o)
        {
            int id;
            // First of all, we assign an id and place object in thread-local dictionary for temporary objects.
            // This is where we support function calls like objectConsumerFunc(objectFactory1(...), ..., objectFactoryN(...))
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            caller = caller ?? NullXlRef; // NullXlRef is what we get for calls from VBA
            lock (Cache)
            {
                // If we are now processing a different cell, then we get rid of any old temp objects
                if (caller != _currentCaller)
                {
                    foreach (int tempId in TempObjects.Keys)
                        IdLookup.Remove(tempId);
                    TempObjects.Clear();
                    _currentCaller = caller;
                }

                id = Interlocked.Increment(ref _idCounter);
                TempObjects[id] = o;
                
                // If the cache had any object at this cell, then drop it from the lookup dictionary
                CacheItem oldObject;
                Cache.TryGetValue(caller, out oldObject);
                IdLookup.Remove(oldObject.Id);

                // Store the new object in the cache
                Cache[caller] = new CacheItem(id, o);
                IdLookup[id] = o;
            }
            return $"{o.GetType().Name}{Separator}{id}";
        }
    }
}
