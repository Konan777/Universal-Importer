using System;
using System.Configuration;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Windows;
using UniversalImporter.DAL;

namespace UniversalImporter.ViewModel
{
    public class MainViewModel : ViewModelBase
    {
        #region private members
        //private string _fileName = @"D:\excel_1_100k.xlsx";
        private string _fileName = @"D:\excel_1.xls";
        private string _connectionString = ConfigurationSettings.AppSettings["ConnectionString"].ToString();
        public ObservableCollection<string> _tables;
        private string _selectedTable;
        private DateTime _dateBeg = DateTime.Parse("2018-12-17");
        private DateTime _dateEnd = DateTime.Parse("2018-12-18");
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
        public RelayCommand CmdSave { get; }
        #endregion

        public MainViewModel()
        {
            CmdSelectFile = new RelayCommand(o => { SelectFile(); });
            CmdRefreshTables = new RelayCommand(o => { RefreshTables(); }); 
            CmdSave = new RelayCommand(o => { Save(); });
            Tables = DataAccessLayer.RefreshTables(ConnectionString);
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
        private void Save()
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

            var dal = new DataAccessLayer(SelectedTable, FileName, ConnectionString);
            dal.SaveXML(DateBeg, DateEnd);
        }
        #endregion


    }
}
