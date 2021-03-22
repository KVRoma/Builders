using Builders.Commands;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class DetailsViewModel : ViewModel
    {
        private User company;
        private bool isOk;
        public string WindowName { get; set; } = "Company Details";

        public User Company
        {
            get { return company; }
            set
            {
                company = value;
                OnPropertyChanged(nameof(Company));
            }
        }
        public bool IsOk
        {
            get { return isOk; }
            set
            {
                isOk = value;
                OnPropertyChanged(nameof(IsOk));
            }
        }
        
        public DetailsViewModel(User user)
        {
            Company = user;
            IsOk = false;
        }

        private Command close;

        public Command Close => close ?? (close = new Command(obj=> 
        {            
            if (obj is System.Windows.Window)
            {
                IsOk = true;
                (obj as System.Windows.Window).Close();
            }
        }));

        
    }
}
