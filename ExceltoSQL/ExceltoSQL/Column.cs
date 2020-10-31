using System.ComponentModel;

namespace ExceltoSQL
{
    public class Column : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public int Index { get; set; }
        public string Type { get; set; }
        private bool _include;
        public bool Include { get { return _include; } set{ _include = value; OnChanged("Include"); } }

        public Column(string name, string type, int index)
        {
            Name = name;
            Type = type;
            Include = true;
            Index = index;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnChanged(string prop)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
