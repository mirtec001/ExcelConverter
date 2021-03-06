﻿using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace AdminPanel.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    { 
        private Microsoft.Office.Interop.Word.Application app;
        private string _openFilePath;
        private string _saveFilePath;
        private string _sheet_name;
        private int _currentProgress;
        private string _statusMessage;
        private int ticker = 0;
        private decimal percentComplete;
        private BackgroundWorker backgroundworker;
        public ICommand FormatReport_Click
        {
            get
            {
                return new RelayCommand(RunApp);
            }
        }
        public ICommand OpenBrowse_Click
        {
            get
            {
                return new RelayCommand(OpenBrowse);
            }
        }

        public ICommand SaveBrowse_Click
        {
            get
            {
                return new RelayCommand(SaveBrowse);
            }
        }

        public int CurrentProgress
        {
            get
            {
                return _currentProgress;
            }
            set
            {
                if (_currentProgress != value)
                {
                    _currentProgress = value;
                    RaisePropertyChanged("CurrentProgress");
                }
            }
        }

        public string OpenFilePath
        {
            get
            {
                return _openFilePath;
            }
            set
            {
                if (_openFilePath != value)
                {
                    _openFilePath = value;
                    RaisePropertyChanged("OpenFilePath");
                }
            }
        }

        public string SaveFilePath
        {
            get
            {
                return _saveFilePath;
            }
            set
            {
                if (_saveFilePath != value)
                {
                    _saveFilePath = value;
                    RaisePropertyChanged("SaveFilePath");
                }
            }
        }

        public string SheetName
        {
            get
            {
                return _sheet_name;
            }
            set
            {
                if (_sheet_name != value)
                {
                    _sheet_name = value;
                    RaisePropertyChanged("SheetName");
                }
            }
        }

        public string StatusMessage
        {
            get
            {
                return _statusMessage;
            }
            set
            {
                _statusMessage = value;
                RaisePropertyChanged("StatusMessage");
            }
        }
        public MainWindowViewModel()
        {
            backgroundworker = new BackgroundWorker();

            backgroundworker.WorkerReportsProgress = true;
            backgroundworker.WorkerSupportsCancellation = true;

        }

        private void OpenBrowse()
        {
            // Open a xlsx
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            if (ofd.ShowDialog() == true)
                OpenFilePath = ofd.FileName;
        }

        private void SaveBrowse()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word Document Files (*.docx)|*.docx|All Files (*.*)|*.*";
            if (sfd.ShowDialog() == true)
                SaveFilePath = sfd.FileName;
        }

        private void RunApp()
        {
            if (backgroundworker.IsBusy != true)
            {
                StatusMessage = "Converting document";
                backgroundworker.DoWork += new DoWorkEventHandler(backgroundworker_DoWork);
                backgroundworker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundwoker_RunWorkerCompleted);
                backgroundworker.ProgressChanged += ProgressChanged;
                backgroundworker.RunWorkerAsync();
            }
        }

        private void backgroundwoker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                StatusMessage = "Error " + e.Error.Message;
            }
            else
            {
                StatusMessage = "Done"; 
            }
        }

        private void backgroundworker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker localWorker = sender as BackgroundWorker;

            try
            {
                app = new Microsoft.Office.Interop.Word.Application();
                var doc = app.Documents.Add();
                int writeup = 1;
                System.Data.DataTable excelTable = LoadWorksheetInDataTable(OpenFilePath, SheetName);
                int total_rows = excelTable.Rows.Count;
                
                foreach (DataRow row in excelTable.Rows)
                {
                    var paragraph = doc.Paragraphs.Add();
                    paragraph.Range.Text = "Writeup " + writeup.ToString() + " of " + excelTable.Rows.Count.ToString() + "\n------------------------------";
                    foreach (DataColumn cols in excelTable.Columns)
                    {
                        string data = cols.ColumnName + ": " + row[cols];
                        // Console.WriteLine(data);
                        paragraph.Range.Text = paragraph.Range.Text + data;
                    }
                    writeup++;
                    ticker++;
                    percentComplete = Decimal.Divide(ticker, total_rows);
                    localWorker.ReportProgress(Convert.ToInt32(percentComplete * 100));
                    Console.WriteLine(percentComplete * 100);
                    doc.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
                    
                }

                foreach(Section wordSection in doc.Sections)
                {
                    Range headerRange = wordSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    headerRange.Fields.Add(headerRange, WdFieldType.wdFieldNumPages);
                    Paragraph p4 = headerRange.Paragraphs.Add();
                    p4.Range.Text = " of ";
                    headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                    Paragraph p1 = headerRange.Paragraphs.Add();
                    p1.Range.Text = "Page ";
                    headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                }

                app.ActiveDocument.SaveAs(SaveFilePath, WdSaveFormat.wdFormatXMLDocument);
                doc.Close();
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                    Marshal.FinalReleaseComObject(app);
                }
            }
        }

        private System.Data.DataTable LoadWorksheetInDataTable(string fileName, string sheetName)
        {
            System.Data.DataTable sheetData = new System.Data.DataTable();
            using (OleDbConnection conn = this.returnConnection(fileName))
            {
                conn.Open();
                // retrieve the data using data adapter
                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter("select * from [" + sheetName + "$]", conn);
                sheetAdapter.Fill(sheetData);
                conn.Close();
            }
            return sheetData;
        }

        private OleDbConnection returnConnection(string fileName)
        {
            string extension = fileName.Substring(fileName.Length - 4);
            if (extension == "xlsx")
            {
                return new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;");
            }
            else
            {
                return new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + "; Jet OLEDB:Engine Type=5;Extended Properties=\"Excel 8.0;\"");
            }
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            CurrentProgress = e.ProgressPercentage;
        }

    }
}
