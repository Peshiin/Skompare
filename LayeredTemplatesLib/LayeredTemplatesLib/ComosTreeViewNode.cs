using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Comos.Controls;
using Comos.Global;
using Comos.Global.AppControls;
using Plt;

namespace LayeredTemplatesLib
{
    public class ComosTreeViewNode : INotifyPropertyChanged
    {
        //Vlastnosti třídy
        //########################################################
        private ObservableCollection<ComosTreeViewNode> mChildren;
        public IList<ComosTreeViewNode> Children { get { return mChildren; } }

        public ComosTreeViewNode Parent;

        public object ComosObject;
        public string Name
        {
            get
            {
                if (ComosObject is IComosDDevice)
                {
                    return ((IComosDDevice)ComosObject).FullName();
                }
                if (ComosObject is IComosDDocument)
                {
                    return ((IComosDDocument)ComosObject).FullName();
                }
                else
                {
                    return "Not assigned";
                }
            }
            set
            {
                if (description != value)
                {
                    description = value;
                    NotifyPropertyChanged(nameof(Description));
                }
            }
        }

        private string description;
        public string Description
        {
            get
            {
                if(ComosObject is IComosDDevice)
                {
                    return ((IComosDDevice)ComosObject).Description;
                }
                if(ComosObject is IComosDDocument)
                {
                    return ((IComosDDocument)ComosObject).Description;
                }
                else
                {
                    return description;
                }
            }
            set
            {
                if(description != value)
                {
                    description = value;
                    NotifyPropertyChanged(nameof(Description));
                }
            }
        }

        private string dn;
        public string DN
        {
            get { return dn; }
            set
            {
                if (dn != value)
                {
                    dn = value;
                    NotifyPropertyChanged(nameof(DN));
                }
            }
        }

        private string ps;
        public string PS
        {
            get { return ps; }
            set
            {
                if (ps != value)
                {
                    ps = value;
                    NotifyPropertyChanged(nameof(PS));
                }
            }
        }

        private string ts;
        public string TS
        {
            get { return ts; }
            set
            {
                if (ts != value)
                {
                    ts = value;
                    NotifyPropertyChanged(nameof(TS));
                }
            }
        }


        //Události třídy
        //########################################################
        public event PropertyChangedEventHandler PropertyChanged;
        /// <summary>
        /// Hlášení změny ve vlastnotech třídy prezentační vrstvě.
        /// </summary>
        protected void NotifyPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if(handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }

        //Konstruktory třídy
        //########################################################
        public ComosTreeViewNode(ComosTreeViewNode Parent)
        {
            this.Parent = Parent;
            mChildren = new ObservableCollection<ComosTreeViewNode>();
        }

        //Metody třídy
        //########################################################
        public override string ToString()
        {
            if(this.ComosObject is IComosDDevice)
                return ((IComosDDevice)this.ComosObject).Name;
            else if (this.ComosObject is IComosDDocument)
                return ((IComosDDocument)this.ComosObject).Name;
            else
                return base.ToString();
        }
    }
}
