using ExcelApp.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelApp.Model;
using System.Windows;

namespace ExcelApp.ViewModel
{
    class MainVM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        void Notify(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }


        ObservableCollection<string> _emailList;
        string _email;
        ExcelModel _excelModel;
        ExcelGenerator _excelGenerator;
        MailMessanger messanger;
        string _wordText;
        WordGenerator _wordGenerator;

        public ButtonCommand SendButtonClick
        {
            get
            {
                return new ButtonCommand(
                    () => 
                    {
                        _excelGenerator.Generate();
                        _wordGenerator = new WordGenerator(_wordText);
                        _wordGenerator.Generate();
                        messanger.SendMail();
                        MessageBox.Show("Отправлено!");
                    },
                    () =>
                    {
                        return EmailList.Count != 0 && _excelModel.CellsCount != 0
                        && _excelModel.RandomMax != 0;
                    });
            }
        }

        public ExcelModel ExcelModelTemplate
        {
            get { return _excelModel; }
            set
            {
                _excelModel = value;
                Notify("ExcelModelTemplate");
            }
        }

        public ButtonCommand AddButtonClick
        {
            get
            {
                return new ButtonCommand(() => 
                {
                    EmailList.Add(Email);
                    Email = "";
                }, ()=> 
                { 
                    return Email.Contains('@') && Email.Contains('.'); 
                });
            }
        }

        public MainVM()
        {
            EmailList = new ObservableCollection<string>();
            Email = "";
            ExcelModelTemplate = new ExcelModel();
            _excelGenerator = new ExcelGenerator(_excelModel);
            messanger = new MailMessanger(EmailList);
        }

        public ObservableCollection<string> EmailList
        {
            get { return _emailList; }
            set
            {
                _emailList = value;
                Notify("EmailList");
            }
        }
        public string WordText
        {
            get { return _wordText; }
            set
            {
                _wordText = value;
                Notify("WordText");
            }
        }

        public string Email
        {
            get { return _email; }
            set
            {
                _email = value;
                Notify("Email");
            }
        }
    }
}
