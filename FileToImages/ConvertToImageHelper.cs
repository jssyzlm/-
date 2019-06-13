using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Imaging;
using System.Collections;
using Aspose.Words.Saving;
using Aspose.Words;
using Spire.Xls;
using System.Drawing;

namespace ConvertToImage
{
    public class ConvertToImageHelper
    {
        //private readonly string FilePath;
        
        //public ConvertToImageHelper(string FilePath)
        //{
        //    this.FilePath = FilePath;
        //}

        //public bool CheckPath(string filePath)
        //{
        //    bool result = false;
        //    string[] filter = {".doc",".docx",".ppt",".pptx",".pdf",".jpg",".bmp", ".dib",".jpeg",".gif",".xls",".xlsx"};
        //    string extension = Path.GetExtension(FilePath.Trim()).ToLower();
        //    string fileName = Path.GetFileName(FilePath.Trim());
        //    if (!filter.Contains(extension))
        //    {
        //        throw new Exception("上传文件类型错误");
        //    }
        //    if(extension == ".docx" || extension == ".doc")
        //    {
        //        result = ConvertWordToImage(FilePath, Directory.GetCurrentDirectory() + "\\DOC",0,0, ImageFormat.Png,256);
        //    }
        //    else
        //    if (extension == ".xls" || extension == ".xlsx")
        //    {
        //        result = ConvertExcelToImage(FilePath, Directory.GetCurrentDirectory() + "\\XLS", 0, 0, ImageFormat.Png, 256);
        //    }
        //    else
        //    if (extension == ".pdf")
        //    {
        //        result = ConvertPDFToImage(FilePath, Directory.GetCurrentDirectory() + "\\PDF", 0, 0, ImageFormat.Png, 256);
        //    }
        //    else
        //    if (extension == ".pptx" || extension == ".ppt")
        //    {
        //        result = ConvertPPTToImage(FilePath, Directory.GetCurrentDirectory() + "\\PPT", 0, 0, ImageFormat.Png, 256);
        //    }
        //    else
        //    if (extension == ".jpg" || extension == ".jpeg" || extension == ".bmp" || extension == ".dib" || extension == ".gif")
        //    {
        //        result = ConvertJPGToImage(FilePath, Directory.GetCurrentDirectory() + "\\PNG");
        //    }
        //    return result;
        //}

        /// <summary>
        /// word转图片到目录
        /// </summary>
        /// <param name="originFilePath">word路径</param>
        /// <param name="imageOutputDirPath">目录路径</param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat">转换的图片格式</param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public bool ConvertWordToImage(string originFilePath, string imageOutputDirPath, int startPageNum, int endPageNum, ImageFormat imageFormat, float resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".docx" && extension != ".doc")
            {
                throw new Exception("文件类型错误");
            }
            ArrayList listimagename = new ArrayList();  
            try
            {
                /* open word file */
                Aspose.Words.Document doc = new Aspose.Words.Document(originFilePath);
                /* validate parameter */
                if (doc == null)
                {
                    throw new Exception("Word文件无效或者Word文件被加密！");
                }
                if (imageOutputDirPath.Trim().Length == 0)
                {
                    imageOutputDirPath = System.IO.Path.GetDirectoryName(originFilePath);
                }
                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }
                string imageName = System.IO.Path.GetFileNameWithoutExtension(originFilePath);
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }
                if (endPageNum > doc.PageCount || endPageNum <= 0)
                {
                    endPageNum = doc.PageCount;
                }
                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }
                if (imageFormat == null)
                {
                    imageFormat = ImageFormat.Png;
                }
                if (resolution <= 0)
                {
                    resolution = 128;
                }
                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(GetSaveFormat(imageFormat));
                imageSaveOptions.Resolution = resolution;
                /* start to convert each page */
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    imageSaveOptions.PageIndex = i - 1;
                    doc.Save(System.IO.Path.Combine(imageOutputDirPath, imageName) + "_" + i.ToString("000") + "." + imageFormat.ToString(), imageSaveOptions);
                    listimagename.Add(System.IO.Path.Combine(imageOutputDirPath, imageName) + "_" + i.ToString("000") + "." + imageFormat.ToString());                    
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }
       
        /// <summary>
        /// word转图片到数据流
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public List<Stream> ConvertWordToImage(string originFilePath, int startPageNum, int endPageNum, ImageFormat imageFormat, float resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".docx" && extension != ".doc")
            {
                throw new Exception("文件类型错误");
            }
            List<Stream> listImageStream = new List<Stream>();
            try
            {
                /* open word file */
                Aspose.Words.Document doc = new Aspose.Words.Document(originFilePath);
                
                /* validate parameter */
                if (doc == null)
                {
                    throw new Exception("Word文件无效或者Word文件被加密！");
                }                
                string imageName = Path.GetFileNameWithoutExtension(originFilePath);
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }
                if (endPageNum > doc.PageCount || endPageNum <= 0)
                {
                    endPageNum = doc.PageCount;
                }
                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }
                if (imageFormat == null)
                {
                    imageFormat = ImageFormat.Png;
                }
                if (resolution <= 0)
                {
                    resolution = 128;
                }
                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(GetSaveFormat(imageFormat));
                imageSaveOptions.Resolution = resolution;
                /* start to convert each page */
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    Stream stream = new MemoryStream();
                    imageSaveOptions.PageIndex = i - 1;
                    doc.Save(stream, imageSaveOptions);
                    listImageStream.Add(stream);                   
                }
                
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            return listImageStream;
        }

        /// <summary>
        /// Excel转图片到目录
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="imageOutputDirPath"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public bool ConvertExcelToImage(string originFilePath, string imageOutputDirPath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".xls" && extension != ".xlsx")
            {
                throw new Exception("文件类型错误");
            }
            try
            {
                if (imageOutputDirPath.Trim().Length == 0)
                { 
                    imageOutputDirPath = System.IO.Path.GetDirectoryName(originFilePath);
                }
                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }

                Workbook workbook = new Workbook();
                workbook.LoadFromFile(originFilePath);
                int sheetCount = workbook.Worksheets.Count();              
                for (int i = 0; i < sheetCount; i++)
                {
                    Worksheet worksheet = workbook.Worksheets[i];
                    string filePath = originFilePath + "_" + worksheet.Name + ".png";
   
                    worksheet.SaveToImage(filePath, imageFormat);
                    if(File.Exists(filePath))
                    {
                        string destfilePath = imageOutputDirPath + "\\" + Path.GetFileName(filePath);
                        if (File.Exists(destfilePath))
                        {
                            File.Delete(destfilePath);
                        }
                        File.Move(filePath, destfilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Excel转图片到数据流
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public List<Stream> ConvertExcelToImage(string originFilePath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".xls" && extension != ".xlsx")
            {
                throw new Exception("文件类型错误");
            }
            List<Stream> listImages = new List<Stream>();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(originFilePath);
                int sheetCount = workbook.Worksheets.Count();
                for (int i = 0; i < sheetCount; i++)
                {
                    Stream stream = new MemoryStream();
                    Worksheet worksheet = workbook.Worksheets[i];

                    string filePath = originFilePath + "_" + worksheet.Name + ".png";

                    worksheet.SaveToImage(filePath, imageFormat);
                    if (File.Exists(filePath))
                    {
                        Bitmap bitmap = new Bitmap(filePath);                      
                        bitmap.Save(stream,ImageFormat.Png);
                        bitmap.Dispose();
                        File.Delete(filePath);
                    }

                    listImages.Add(stream);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            return listImages;
        }

        /// <summary>
        /// PDF转图片到目录
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="imageOutputDirPath"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public bool ConvertPDFToImage(string originFilePath, string imageOutputDirPath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".pdf")
            {
                throw new Exception("文件类型错误");
            }
            ArrayList listimagename = new ArrayList();
            try
            {
                Aspose.Pdf.Document doc = new Aspose.Pdf.Document(originFilePath);
                if (doc == null)
                {
                    throw new Exception("pdf文件无效或者pdf文件被加密！");
                }
                if (imageOutputDirPath.Trim().Length == 0)
                {
                    imageOutputDirPath = System.IO.Path.GetDirectoryName(originFilePath);
                }
                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }
                if (endPageNum > doc.Pages.Count || endPageNum <= 0)
                {
                    endPageNum = doc.Pages.Count;
                }
                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }
                if (resolution <= 0)
                {
                    resolution = 128;
                }
                string imageNamePrefix = System.IO.Path.GetFileNameWithoutExtension(originFilePath);
                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(GetSaveFormat(imageFormat));
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    MemoryStream stream = new MemoryStream();
                    string imgPath = System.IO.Path.Combine(imageOutputDirPath, imageNamePrefix) + "_" + i.ToString("000") + "." + imageFormat.ToString();
                    Aspose.Pdf.Devices.Resolution reso = new Aspose.Pdf.Devices.Resolution(resolution);
                    Aspose.Pdf.Devices.JpegDevice jpegDevice = new Aspose.Pdf.Devices.JpegDevice(reso, 100);
                    jpegDevice.Process(doc.Pages[i], stream);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(stream);
                    img.Save(imgPath);
                    stream.Close();
                    listimagename.Add(System.IO.Path.Combine(imageOutputDirPath, imageNamePrefix) + "_" + i.ToString("000") + "." + imageFormat.ToString());
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// PDF转图片到数据流
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public List<Stream> ConvertPDFToImage(string originFilePath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".pdf")
            {
                throw new Exception("文件类型错误");
            }
            List<Stream> listImage = new List<Stream>();
            try
            {
                Aspose.Pdf.Document doc = new Aspose.Pdf.Document(originFilePath);
                if (doc == null)
                {
                    throw new Exception("pdf文件无效或者pdf文件被加密！");
                }
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }
                if (endPageNum > doc.Pages.Count || endPageNum <= 0)
                {
                    endPageNum = doc.Pages.Count;
                }
                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }
                if (resolution <= 0)
                {
                    resolution = 128;
                }
                string imageNamePrefix = System.IO.Path.GetFileNameWithoutExtension(originFilePath);
                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(GetSaveFormat(imageFormat));
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    MemoryStream stream = new MemoryStream();
                    Aspose.Pdf.Devices.Resolution reso = new Aspose.Pdf.Devices.Resolution(resolution);
                    Aspose.Pdf.Devices.JpegDevice jpegDevice = new Aspose.Pdf.Devices.JpegDevice(reso, 100);
                    jpegDevice.Process(doc.Pages[i], stream);
                    listImage.Add(stream);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            return listImage;
        }

        /// <summary>
        /// PPT转图片到目录
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="imageOutputDirPath"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public bool ConvertPPTToImage(string originFilePath, string imageOutputDirPath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".ppt" && extension != ".pptx")
            {
                throw new Exception("文件类型错误");
            }
            try
            {
                Aspose.Slides.Presentation doc = new Aspose.Slides.Presentation(originFilePath);
                if (doc == null)
                {
                    throw new Exception("ppt文件无效或者ppt文件被加密！");
                }
                if (imageOutputDirPath.Trim().Length == 0)
                {
                    imageOutputDirPath = System.IO.Path.GetDirectoryName(originFilePath);
                }
                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }
                if (endPageNum > doc.Slides.Count || endPageNum <= 0)
                {
                    endPageNum = doc.Slides.Count;
                }
                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }
                if (resolution <= 0)
                {
                    resolution = 128;
                }
                /* 先将ppt转换为pdf临时文件 */
                string tmpPdfPath = originFilePath + ".pdf";
                doc.Save(tmpPdfPath, Aspose.Slides.Export.SaveFormat.Pdf);
                /* 再将pdf转换为图片 */
                ConvertPDFToImage(tmpPdfPath, imageOutputDirPath, 0, 0, imageFormat, 200);
                /*删除pdf临时文件 */
                File.Delete(tmpPdfPath);

            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// PPT转图片到数据流
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imageFormat"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public List<Stream> ConvertPPTToImage(string originFilePath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".ppt" && extension != ".pptx")
            {
                throw new Exception("文件类型错误");
            }
            List<Stream> listImages = null;
            try
            {
                Aspose.Slides.Presentation doc = new Aspose.Slides.Presentation(originFilePath);
                if (doc == null)
                {
                    throw new Exception("ppt文件无效或者ppt文件被加密！");
                }
                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }
                if (endPageNum > doc.Slides.Count || endPageNum <= 0)
                {
                    endPageNum = doc.Slides.Count;
                }
                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }
                if (resolution <= 0)
                {
                    resolution = 128;
                }
                /* 先将ppt转换为pdf临时文件 */
                string tmpPdfPath = originFilePath + ".pdf";
                doc.Save(tmpPdfPath, Aspose.Slides.Export.SaveFormat.Pdf);
                /* 再将pdf转换为图片 */
                listImages = ConvertPDFToImage(tmpPdfPath , 0, 0, imageFormat, 200);
                /*删除pdf临时文件 */
                File.Delete(tmpPdfPath);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            return listImages;
        }

        /// <summary>
        /// 其它图片转PNG格式图片到目录
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="imageOutputDirPath"></param>
        /// <returns></returns>
        public bool ConvertJPGToImage(string originFilePath, string imageOutputDirPath)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".jpg" && extension != ".jpeg" && extension != ".bmp" && extension != ".dib" && extension != ".gif")
            {
                throw new Exception("文件类型错误");
            }
            try
            {
                string destinationPath = System.IO.Path.GetDirectoryName(imageOutputDirPath);
                //// 判断目标目录是否存在如果不存在则新建
                if (!System.IO.Directory.Exists(imageOutputDirPath))
                {
                    System.IO.Directory.CreateDirectory(imageOutputDirPath);
                }
                string filePath = imageOutputDirPath + "\\" + Path.GetFileName(originFilePath);
                System.IO.File.Copy(originFilePath, filePath, true);
                
                string newPath = Path.ChangeExtension(filePath, "Png");
                File.Move(filePath, newPath);
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 其它图片转PNG格式到数据流
        /// </summary>
        /// <param name="originFilePath"></param>
        /// <param name="imageFormat"></param>
        /// <returns></returns>
        public Stream ConvertJPGToImage(string originFilePath)
        {
            string extension = Path.GetExtension(originFilePath.Trim()).ToLower();
            if (extension != ".jpg" && extension != ".jpeg" && extension != ".bmp" && extension != ".dib" && extension != ".gif")
            {
                throw new Exception("文件类型错误");
            }
            Stream stream = new MemoryStream();
            try
            {
                Bitmap bitmap = new Bitmap(originFilePath);
                bitmap.Save(stream,ImageFormat.Png);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
            return stream;
        }

        private Aspose.Words.SaveFormat GetSaveFormat(ImageFormat imageFormat)
        {
            Aspose.Words.SaveFormat sf = Aspose.Words.SaveFormat.Unknown;
            if (imageFormat.Equals(ImageFormat.Png))
                sf = Aspose.Words.SaveFormat.Png;
            else if (imageFormat.Equals(ImageFormat.Jpeg))
                sf = Aspose.Words.SaveFormat.Jpeg;
            else if (imageFormat.Equals(ImageFormat.Tiff))
                sf = Aspose.Words.SaveFormat.Tiff;
            else if (imageFormat.Equals(ImageFormat.Bmp))
                sf = Aspose.Words.SaveFormat.Bmp;
            else
                sf = Aspose.Words.SaveFormat.Unknown;
            return (sf);
        }


    }    
}
