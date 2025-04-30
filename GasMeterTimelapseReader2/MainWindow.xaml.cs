using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Tesseract;

namespace GasMeterTimelapseReader2
{
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = @"E:\\temp\\gasmeter_timelapse_images";
            var numbers = ReadNumbersFromImages(folderPath);
            dataGrid.ItemsSource = numbers;
        }

        private List<string> ReadNumbersFromImages(string folderPath)
        {
            var numbers = new List<string>();
            var files = Directory.GetFiles(folderPath, "*.jpg");

            using (var engine = new TesseractEngine(@"E:\temp\gasmeter_timelapse_images\tessdata", "eng", EngineMode.Default))
            {
                foreach (var file in files)
                {
                    try
                    {
                        Mat image = Cv2.ImRead(file);
                        Mat cropped = new Mat(image, new OpenCvSharp.Rect(638, 457, 206, 103));
                        Cv2.ImWrite("cropped.jpg", cropped);

                        using (var img = Pix.LoadFromFile("cropped.jpg"))
                        {
                            using (var page = engine.Process(img))
                            {
                                var text = page.GetText();

                                if (!string.IsNullOrEmpty(text))
                                {
                                    numbers.Add(text);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error reading image {file}: {ex.Message}");
                        return null;
                    }
                }
            }
            return numbers;
        }
    }
}