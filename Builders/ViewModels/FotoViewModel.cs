using Builders.Commands;
using Builders.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace Builders.ViewModels
{
    public class FotoViewModel : ViewModel
    {
        public string NameWindow { get; } = "Foto View";
        public string GoButton { get; } = ">>";
        public string BackButton { get; } = "<<";
        private string sheetsCount;

        private int? clientSelectId;
        private string maskName;
        private BitmapImage image;
        private string nameFile;
        private List<string> images;
        private int imgCounter;

        public string SheetsCount
        {
            get { return sheetsCount; }
            set
            {
                sheetsCount = value;
                OnPropertyChanged(nameof(SheetsCount));
            }
        }
        public int? ClientSelectId
        {
            get { return clientSelectId; }
            set
            {
                clientSelectId = value;
                OnPropertyChanged(nameof(ClientSelectId));
            }
        }
        public string MaskName
        {
            get { return maskName; }
            set 
            {
                maskName = value;
                OnPropertyChanged(nameof(MaskName));
            }
        }
        public BitmapImage Image
        {
            get { return image; }
            set
            {
                image = value;
                OnPropertyChanged(nameof(Image));
            }
        }
        public string NameFile
        {
            get { return nameFile; }
            set
            {
                nameFile = value;
                OnPropertyChanged(nameof(NameFile));
                Image = GetImage(NameFile);
            }
        }

        private Command _goFile;
        private Command _backFile;
        private Command _delFile;

        public Command GoFile => _goFile ?? (_goFile = new Command(obj=> 
        {
            if (NameFile != null)
            { 
                if(imgCounter < images.Count -1)
                {
                    imgCounter++;
                    NameFile = images[imgCounter];
                }
                SheetsCount = imgCounter + 1 + " in " + images.Count;
            }
        }));
        public Command BackFile => _backFile ?? (_backFile = new Command(obj=> 
        {
            if (NameFile != null)
            {
                if (imgCounter > 0)
                {
                    imgCounter--;
                    NameFile = images[imgCounter];
                }
                SheetsCount = imgCounter + 1 + " in " + images.Count;
            }
        }));
        public Command DelFile => _delFile ?? (_delFile = new Command(obj=> 
        {
            if (NameFile != null)
            {
                try
                {                  
                    
                    File.Delete(NameFile);
                    images.RemoveAt(imgCounter);

                    if (imgCounter > 0)
                    {
                        imgCounter--;
                        NameFile = images[imgCounter];
                        SheetsCount = imgCounter + 1 + " in " + images.Count;
                    }
                    else if (imgCounter == 0 && images.Count >= 1)
                    {                        
                        NameFile = images[imgCounter];
                        SheetsCount = imgCounter + 1 + " in " + images.Count;
                    }
                    else
                    {
                        NameFile = null;
                        SheetsCount = 0 + " in " + images.Count;
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                
            }
        }));

        public FotoViewModel(int? selectId, string mask)
        {
            ClientSelectId = selectId;
            MaskName = mask;
            images = GetFotoNameArray(Directory.GetCurrentDirectory() + "\\Foto");
            NameFile = (images.Count > 0) ? (images[0]) : (null);
            imgCounter = 0;
            
            if (images.Count > 0)
            {
                SheetsCount = imgCounter + 1 + " in " + images.Count;
            }
            else
            {
                SheetsCount = imgCounter + " in " + images.Count;
            }
        }


       
        /// <summary>
        /// Метод повертає масив файлів з обраного каталогу
        /// </summary>
        /// <param name="nameFile"></param>        
        /// <returns></returns>
        public List<string> GetFotoNameArray(string pathDir)
        {
            string[] images = Directory.GetFiles(pathDir, ClientSelectId + MaskName);    // зберігаємо масив фото з каталогу  
            return new List<string>(images);
        }

        /// <summary>
        /// Перевіряє файл, якщо не фото то показує заставку Error.png
        /// </summary>
        /// <param name="nameFile"></param>
        public BitmapImage GetImage(string nameFile)
        {            
            try
            {
                if (nameFile != null)
                {
                    //create new stream and create bitmap frame
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = new FileStream(nameFile, FileMode.Open, FileAccess.Read);
                    //load the image now so we can immediately dispose of the stream
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    //clean up the stream to avoid file access exceptions when attempting to delete images
                    bitmapImage.StreamSource.Dispose();

                    return bitmapImage;
                }
                else
                {
                    return null;
                }                
            }
            catch
            {
                return null; //new BitmapImage(new Uri(AppDomain.CurrentDomain.BaseDirectory + "Error.png"));        //Якщо це не фото то показуєм картинку
            }            
        }


    }
}
