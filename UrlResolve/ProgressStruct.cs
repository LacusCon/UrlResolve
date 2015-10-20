using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UrlResolve
{
    class ProgressStruct : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }

        private int _progress = 0;

        public int ProgressNow
        {
            get { return this._progress; }
            set
            {
                this._progress = value;
                OnPropertyChanged("ProgressNow");
            }
        }
    }
}
