using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("1980308C-6000-49EC-8A6A-4E5D04EB8D08")]
    public interface IeListe
    {
        int Add(object value, String Key = null);

        int AddWithKey(object value, String Key);

        void Clear();

        bool Contains(object value);

        bool ContainKey(String Key);

        int IndexOf(object value);

        String KeyOf(object value);

        void Move(int startindex, int endindex);

        void Insert(int index, object value);

        void InsertWithKey(int index, object value, String Key);

        bool IsFixedSize { get; }

        bool IsReadOnly { get; }

        void Remove(object value);

        void RemoveAt(int index);

        void RemoveAtKey(String Key);

        object this[dynamic index] { get; set; }

        object ItemByKey(String Key);

        void CopyTo(Array array, int index);

        int Count { get; }

        bool IsSynchronized { get; }

        object SyncRoot { get; }

        IEnumerator GetEnumerator();

        object Clone();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6C05A3DE-4AE5-4935-B3B8-B15EB03A3EEA")]
    [ProgId("Frameworks.eListe")]
    public class eListe : ArrayList, IeListe
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eListe).Name;

        private Dictionary<String, Object> _Dic;

        #endregion

        #region "Constructeur\Destructeur"

        public eListe()
            : base()
        {
            _Dic = new Dictionary<string, object>();
        }

        public eListe(ICollection Collection)
            : base(Collection)
        {
            _Dic = new Dictionary<string, object>();
        }

        #endregion

        #region "Propriétés"

        public object this[dynamic index]
        {
            get
            {
                if (index is String)
                    return _Dic[index];
                else
                    return base[(int)index];
            }
            set
            {
                if (index is String)
                {
                    base[base.IndexOf(_Dic[(String)index])] = value;
                    _Dic[(String)index] = value;
                }
                else
                    base[(int)index] = value;
            }
        }

        #endregion

        #region "Méthodes"

        public int Add(object value, string Key = null)
        {
            if (Key != null)
                return AddWithKey(value, Key);

            return base.Add(value);
        }

        public int AddWithKey(object value, string Key)
        {
            if (!_Dic.ContainsKey(Key))
            {
                _Dic.Add(Key, value);
                return base.Add(value);
            }

            return -1;
        }

        public bool ContainKey(string Key)
        {
            return _Dic.ContainsKey(Key);
        }

        public string KeyOf(object value)
        {
            try
            {
                return _Dic.First(x => x.Value == value).Key;
            }
            catch
            {
                return null;
            }

        }

        public void Move(int startindex, int endindex)
        {
            object pObj = base[startindex];
            base.Remove(pObj);
            base.Insert(endindex, pObj);
        }

        public void InsertWithKey(int index, object value, string Key)
        {
            if (!_Dic.ContainsKey(Key))
            {
                _Dic.Add(Key, value);
                base.Insert(index, value);
            }
        }

        public void RemoveAtKey(string Key)
        {
            if (_Dic.ContainsKey(Key))
            {
                object pObj = _Dic[Key];
                _Dic.Remove(Key);
                base.Remove(pObj);
            }
        }

        public object ItemByKey(string Key)
        {
            if (_Dic.ContainsKey(Key))
            {
                return _Dic[Key];
            }

            return null;
        }

        #endregion

        

        
    }
}

