using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ConvertToImage;

namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                textBox1.Text = path;
            }

        }

        ConvertToImageHelper cti = new ConvertToImageHelper();
        private void button1_Click(object sender, EventArgs e)
        {

            bool result = false;
            string[] filter = { ".doc", ".docx", ".ppt", ".pptx", ".pdf", ".jpg", ".bmp", ".dib", ".jpeg", ".gif", ".xls", ".xlsx" };
            string FilePath = textBox1.Text;
            string extension = Path.GetExtension(FilePath.Trim()).ToLower();
            string fileName = Path.GetFileName(FilePath.Trim());
            if (!filter.Contains(extension))
            {
                throw new Exception("上传文件类型错误");
            }
            if (extension == ".docx" || extension == ".doc")
            {
                result = cti.ConvertWordToImage(FilePath, Directory.GetCurrentDirectory() + "\\DOC", 0, 0, ImageFormat.Png, 256);
            }
            else
            if (extension == ".xls" || extension == ".xlsx")
            {
                result = cti.ConvertExcelToImage(FilePath, Directory.GetCurrentDirectory() + "\\XLS", 0, 0, ImageFormat.Png, 256);
            }
            else
            if (extension == ".pdf")
            {
                result = cti.ConvertPDFToImage(FilePath, Directory.GetCurrentDirectory() + "\\PDF", 0, 0, ImageFormat.Png, 256);
            }
            else
            if (extension == ".pptx" || extension == ".ppt")
            {
                result = cti.ConvertPPTToImage(FilePath, Directory.GetCurrentDirectory() + "\\PPT", 0, 0, ImageFormat.Png, 256);
            }
            else
            if (extension == ".jpg" || extension == ".jpeg" || extension == ".bmp" || extension == ".dib" || extension == ".gif")
            {
                result = cti.ConvertJPGToImage(FilePath, Directory.GetCurrentDirectory() + "\\PNG");
            }
            if (result)
            {
                MessageBox.Show("成功");
            }
            else
            {
                MessageBox.Show("失败");
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {           
            List<Stream> streams = cti.ConvertWordToImage(textBox1.Text, 0, 0, ImageFormat.Png, 256);
            //foreach (Stream stream in streams)
            //{
            //    imageList1.Images.Add(Bitmap.FromStream(stream));
            //    stream.Close();
            //}
            //pictureBox1.BackgroundImage = imageList1.Images[0];

            pictureBox1.BackgroundImage = Bitmap.FromStream(streams[0]);
            this.Refresh();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<Stream> streams = cti.ConvertExcelToImage(textBox1.Text, 0, 0, ImageFormat.Png, 256);
            foreach (Stream stream in streams)
            {
                imageList1.Images.Add(Bitmap.FromStream(stream));
                stream.Close();
            }
            pictureBox1.Image = imageList1.Images[0];
            this.Refresh();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            List<Stream> streams = cti.ConvertPDFToImage(textBox1.Text, 0, 0, ImageFormat.Png, 256);
            foreach (Stream stream in streams)
            {
                imageList1.Images.Add(Bitmap.FromStream(stream));
                stream.Close();
            }
            pictureBox1.Image = imageList1.Images[0];
            this.Refresh();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            List<Stream> streams = cti.ConvertPPTToImage(textBox1.Text, 0, 0, ImageFormat.Png, 256);
            foreach (Stream stream in streams)
            {
                imageList1.Images.Add(Bitmap.FromStream(stream));
                stream.Close();
            }
            pictureBox1.Image = imageList1.Images[0];
            this.Refresh();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Stream stream = cti.ConvertJPGToImage(textBox1.Text);
            pictureBox1.Image = Bitmap.FromStream(stream);
            stream.Close();
            this.Refresh();
        }
    }
}
