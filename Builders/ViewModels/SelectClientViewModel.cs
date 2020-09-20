using Builders.Commands;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class SelectClientViewModel : ViewModel
    {
        //private BuilderContext db;
        public string WindowName { get; } = "Select Client";

        private Client clientSelect;
        private IEnumerable<Client> clients;
        private bool pressButton;
        private bool enabledButton;
        private int? lastId;

        public Client ClientSelect
        {
            get { return clientSelect; }
            set
            {
                clientSelect = value;
                OnPropertyChanged(nameof(ClientSelect));                
            }
        }
        public IEnumerable<Client> Clients
        {
            get { return clients; }
            set
            {
                clients = value;
                OnPropertyChanged(nameof(Clients));
            }
        }
        public bool PressButton
        {
            get { return pressButton; }
            set
            {
                pressButton = value;
                OnPropertyChanged(nameof(PressButton));
            }
        }
        public bool EnabledButton
        {
            get { return enabledButton; }
            set
            {
                enabledButton = value;
                OnPropertyChanged(nameof(EnabledButton));
            }
        }
        public int? LastId
        {
            get { return lastId; }
            set 
            {
                lastId = value;
                OnPropertyChanged(nameof(LastId));
            }
        }
            


        private Command okCommand;

        public Command OkCommand => okCommand ?? (okCommand = new Command(obj=> 
        {
            if (obj is System.Windows.Window)
            {
                PressButton = true;
                (obj as System.Windows.Window).Close();
            }
        }));

        public SelectClientViewModel()
        {
                        
            EnabledButton = true;
            PressButton = false;
        }
            
    }
}
