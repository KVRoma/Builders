using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class MessageViewModel : ViewModel
    {
        public string WindowName { get; } = "Message....";
        private bool pressOk;
        private string text;
        private int height;
        private int width;

        public bool PressOk
        {
            get { return pressOk; }
            set
            {
                pressOk = value;
                OnPropertyChanged(nameof(PressOk));
            }
        }
        public string Text
        {
            get { return text; }
            set
            {
                text = value;
                OnPropertyChanged(nameof(Text));
            }
        }
        public int Height
        {
            get { return height; }
            set
            {
                height = value;
                OnPropertyChanged(nameof(Height));
            }
        }
        public int Width
        {
            get { return width; }
            set
            {
                width = value;
                OnPropertyChanged(nameof(Width));
            }
        }

        private Command okCommand;
        private Command cancelCommand;

        public Command OkCommand => okCommand ?? (okCommand = new Command(obj =>
        {

            if (obj is System.Windows.Window)
            {
                PressOk = true;
                (obj as System.Windows.Window).Close();
            }

        }));
        public Command CancelCommand => cancelCommand ?? (cancelCommand = new Command(obj =>
        {

            if (obj is System.Windows.Window)
            {
                PressOk = false;
                (obj as System.Windows.Window).Close();
            }

        }));

        public MessageViewModel(int heightWindow, int widthwindow, string message)
        {
            Height = heightWindow;
            Width = widthwindow;
            Text = message;
            PressOk = false;
        }
    }
}
