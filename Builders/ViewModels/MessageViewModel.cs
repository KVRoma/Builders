using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Foreground = System.Windows.Media.Brush;



namespace Builders.ViewModels
{
    public class MessageViewModel : ViewModel
    {
        public string WindowName { get; } = "Message....";
        private bool pressOk;
        private string text;
        private int height;
        private int width;
        private Foreground brush;

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
        public Foreground Brush
        {
            get { return brush; }
            set
            {
                brush = value;
                OnPropertyChanged(nameof(Brush));
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

        

        public MessageViewModel(int heightWindow, int widthwindow, string message, Foreground color)
        {
            Height = heightWindow;
            Width = widthwindow;
            Text = message;
            PressOk = false;
            Brush = (color != null) ? color : System.Windows.Media.Brushes.Blue;
        }
    }
}
