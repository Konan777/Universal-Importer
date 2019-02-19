using System;
using System.Configuration;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Windows;
using UniversalImporter.DAL;
using System.Threading.Tasks;
using System.IO;

namespace UniversalImporter.ViewModel
{
    public class MainViewModel : ViewModelBase
    {
        #region private members
        //private string _fileName = @"D:\excel_1_100k.xlsx";
        private string _fileName = @"D:\excel_1_300k.xlsx";
        private string _connectionString = ConfigurationSettings.AppSettings["ConnectionString"].ToString();
        public ObservableCollection<string> _tables;
        private string _selectedTable;
        private DateTime _dateBeg = DateTime.Parse("2018-12-17");
        private DateTime _dateEnd = DateTime.Parse("2018-12-18");
        private int _progressBarValue;
        private int _progressBarMaximum;
        private string _progressBarText;
        private bool _statuBarVisiblity;
        private bool _buttEnabled;
        #endregion

        #region public members
        public string FileName
        {
            get { return _fileName; }
            set
            {
                _fileName = value;
                NotifyPropertyChanged();
            }
        }
        public string ConnectionString
        {
            get { return _connectionString; }
            set
            {
                _connectionString = value;
                NotifyPropertyChanged();
            }
        }
        public ObservableCollection<string> Tables
        {
            get { return _tables; }
            set
            {
                _tables = value;
                NotifyPropertyChanged();
            }
        }
        public string SelectedTable
        {
            get { return _selectedTable; }
            set
            {
                _selectedTable = value;
                NotifyPropertyChanged();
            }
        }
        public DateTime DateBeg
        {
            get { return _dateBeg; }
            set
            {
                _dateBeg = value;
                NotifyPropertyChanged();
            }
        }
        public DateTime DateEnd
        {
            get { return _dateEnd; }
            set
            {
                _dateEnd = value;
                NotifyPropertyChanged();
            }
        }
        public RelayCommand CmdSelectFile { get; }
        public RelayCommand CmdRefreshTables { get; }
        public RelayCommand CmdSaveXml { get; }
        public RelayCommand CmdSaveBulk { get; }
        public bool StatuBarVisiblity
        {
            get
            {
                return _statuBarVisiblity;
            }

            private set
            {
                _statuBarVisiblity = value;
                NotifyPropertyChanged();
            }
        }
        public int ProgressBarValue
        {
            get
            {
                return _progressBarValue;
            }

            private set
            {
                _progressBarValue = value;
                NotifyPropertyChanged();
            }
        }
        public int ProgressBarMaximum
        {
            get
            {
                return _progressBarMaximum;
            }

            private set
            {
                _progressBarMaximum = value;
                NotifyPropertyChanged();
            }
        }
        public string ProgressBarText
        {
            get
            {
                return _progressBarText;
            }

            private set
            {
                _progressBarText = value;
                NotifyPropertyChanged();
            }
        }
        public bool ButtEnabled
        {
            get
            {
                return _buttEnabled;
            }

            private set
            {
                _buttEnabled = value;
                NotifyPropertyChanged();
            }
        }

        #endregion

        public MainViewModel()
        {
            CmdSelectFile = new RelayCommand(o => { SelectFile(); });
            CmdRefreshTables = new RelayCommand(o => { RefreshTables(); }); 
            CmdSaveXml = new RelayCommand(o => { Save(false); });
            CmdSaveBulk = new RelayCommand(o => { Save(true); });
            Tables = DataAccessLayer.RefreshTables(ConnectionString);
            ButtEnabled = true;
            RefreshTables();
        }


        #region Commands handlers
        private void SelectFile()
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
                FileName = openFileDialog.FileName;
        }
        private void RefreshTables()
        {
            Tables = DataAccessLayer.RefreshTables(ConnectionString);
        }
        private async void Save(bool Bulk)
        {
            if (string.IsNullOrEmpty(FileName))
            {
                MessageBox.Show("На выбран файл для обработки!");
                return;
            }
            if (string.IsNullOrEmpty(SelectedTable))
            {
                MessageBox.Show("На выбрана таблица для вставки данных!");
                return;
            }
            // для того чтобы выбрались данные за весь выбранный период добавим к конченой дате время 
            // Т.к. у нас контрол только с датами а в файле "excel_1.xls" есть еще и время
            TimeSpan ts = new TimeSpan(23, 59, 59);
            DateEnd = DateEnd.Date + ts;
            ProgressBarValue = 0;
            ProgressBarMaximum = 100;
            StatuBarVisiblity = true;
            try
            {
                var dal = new DataAccessLayer(SelectedTable, ConnectionString);
                var progress = new Progress<int>();
                progress.ProgressChanged += (sender, args) =>
                {
                    ProgressBarValue = args;
                    ProgressBarText = string.Format("Parsed {0} from {1} rows", args, ProgressBarMaximum);
                };
                var task = Task.Run(() =>
                {
                    ButtEnabled = false;
                    IReader reader = null;
                    // Для больших файлов другой класс
                    if (Path.GetExtension(FileName).ToUpper() == ".XLSX")
                        reader = new ReaderEPPlus();
                    else
                        reader = new ReaderOleDb();


                    reader.Init(FileName, dal.GetSqlTableSchema());
                    ProgressBarMaximum = reader.RowsCount;
                    dal.Save(reader, DateBeg, DateEnd, progress, Bulk);
                });
                    await task;
            } catch (Exception ex)
            {
                MessageBox.Show("Ошибка:"+ex.Message);
            }
            //dal.SaveXML(reader, DateBeg, DateEnd, progress);
            StatuBarVisiblity = false;
            ProgressBarValue = 0;
            ProgressBarText = "";
            ButtEnabled = true;
        }
        #endregion


    }
}
