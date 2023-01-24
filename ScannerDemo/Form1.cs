using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WIA;

namespace ScannerDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ScannedDocuments");
            Directory.CreateDirectory(folderPath);
            
            ListScanners();

            // Set start output folder TMP
            textBox1.Text = folderPath;
            // Set JPEG as default
            comboBox1.SelectedIndex = 1;

        }

        private void ListScanners()
        {
            // Clear the ListBox.
            listBox1.Items.Clear();

            // Create a DeviceManager instance
            var deviceManager = new DeviceManager();

            // Loop through the list of devices and add the name to the listbox
            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                // Add the device only if it's a scanner
                if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                {
                    continue;
                }

                // Add the Scanner device to the listbox (the entire DeviceInfos object)
                // Important: we store an object of type scanner (which ToString method returns the name of the scanner)
                listBox1.Items.Add(
                    new Scanner(deviceManager.DeviceInfos[i])
                );
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(StartScanning).ContinueWith(result => TriggerScan());
        }

        private void TriggerScan()
        {
            Console.WriteLine("Image succesfully scanned");
        }

        public void StartScanning()
        {
            bool hasPage = true;
            
            do
            {
                Scanner device = null;

                this.Invoke(new MethodInvoker(delegate ()
                {
                    device = listBox1.SelectedItem as Scanner;
                }));

                if (device == null)
                {
                    MessageBox.Show("You need to select first an scanner device from the list",
                                    "Warning",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                
                ImageFile image = new ImageFile();
                string imageExtension = "";

                this.Invoke(new MethodInvoker(delegate ()
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            image = device.ScanImage(WIA.FormatID.wiaFormatPNG);
                            imageExtension = ".png";
                            break;
                        case 1:
                            image = device.ScanImage(WIA.FormatID.wiaFormatJPEG);
                            imageExtension = ".jpeg";
                            break;
                        case 2:
                            image = device.ScanImage(WIA.FormatID.wiaFormatBMP);
                            imageExtension = ".bmp";
                            break;
                        case 3:
                            image = device.ScanImage(WIA.FormatID.wiaFormatGIF);
                            imageExtension = ".gif";
                            break;
                        case 4:
                            image = device.ScanImage(WIA.FormatID.wiaFormatTIFF);
                            imageExtension = ".tiff";
                            break;
                    }
                }));

                string tempFileName = Guid.NewGuid().ToString();
                var path = Path.Combine(textBox1.Text, tempFileName + imageExtension);
                if (image == null)
                    {
                        hasPage = false;
                        break;
                    }
                // Save the image

                if (File.Exists(path))
                {
                    File.Delete(path);
                }


                image.SaveFile(path);


                //pictureBox1.Image = new Bitmap(path);

                
                
            }while (hasPage);
        }

        
        public string SetNewName(string Dir, string OldFileName)
        {

            System.Collections.ArrayList fileList = new System.Collections.ArrayList();
            string NewFileName = string.Empty;
            string[] files = Directory.GetFiles(Dir);

            for (int i = 0; i < files.Length; i++)
            {
                if (files[i].Contains(OldFileName))
                {
                    fileList.Add(files[i]);
                }
            }
            NewFileName = Dir + "\\" + OldFileName + "_" + fileList.Count + ".pdf";


            return NewFileName;
        }

        public void ImageForPDF(string ImagemCaminhoOrigem, string caminhoSaidaPDF)
        {
            string[] caminhoImagens = GetImageFiles(ImagemCaminhoOrigem);
            {
                foreach (var item in caminhoImagens)
                {
                    
                    string pdfpath = caminhoSaidaPDF + ImagemCaminhoOrigem.Substring(ImagemCaminhoOrigem.LastIndexOf("\\")) + ".pdf";
                    if (File.Exists(pdfpath))
                    {
                        //File.Delete(pdfpath);
                        pdfpath = SetNewName(caminhoSaidaPDF, ImagemCaminhoOrigem.Substring(ImagemCaminhoOrigem.LastIndexOf("\\")));
                    }
                    using (var doc = new iTextSharp.text.Document())
                    {
                        iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new FileStream(pdfpath, FileMode.Create));
                        doc.Open();
                        iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(item);
                        image.ScalePercent(24, 24);
                        image.Alignment = iTextSharp.text.Image.ALIGN_MIDDLE;

                        doc.Add(image);

                    }
                    //File.Delete(item);
                }
            }
            
        }

        public string[] GetImageFiles(string ImageSource)
        {
            System.Collections.ArrayList Files = new System.Collections.ArrayList();
            string[] FilesArray = Directory.GetFiles(ImageSource);
            foreach (string file in FilesArray)
            {
                string extension = file.Substring(file.LastIndexOf(".")).ToUpper();
                if (extension.CompareTo(".JPG") == 0 || extension.CompareTo(".JPEG") == 0 || extension.CompareTo(".GIF") == 0)
                {
                    Files.Add(file);
                }
            }

            string[] returnFiles = new string[Files.Count];
            for (int i = 0; i < Files.Count; i++)
            {
                returnFiles[i] = Files[i].ToString();
            }

            return returnFiles;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            DialogResult result = folderDlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ImageForPDF(textBox1.Text, textBox1.Text);
        }
    }
}
