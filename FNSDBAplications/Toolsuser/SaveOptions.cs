using System;
using System.ComponentModel;

namespace FNSDBAplications.Toolsuser
{
    [Serializable]
    public class SaveOptions : INotifyPropertyChanged
    {
        private bool _SaveUser = false;
        public bool SaveUser
        {
            get { return _SaveUser; }
            set
            {
                _SaveUser = value;
                OnPropertyChanged(nameof(SaveUser));
            }
        }

        private string _SaveLoginID;
        public string SaveLoginID
        {
            get { return _SaveLoginID; }
            set
            {
                _SaveLoginID = value;
                OnPropertyChanged(nameof(SaveLoginID));
            }
        }

        private string _SaveLoginPSW;
        public string SaveLoginPSW
        {
            get { return _SaveLoginPSW; }
            set
            {
                _SaveLoginPSW = value;
                OnPropertyChanged(nameof(SaveLoginPSW));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
