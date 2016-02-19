using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleCode.NewPage
{
    class Program
    {
        //此範例使用 iTextSharp 建立兩頁尺寸為 A4 的 PDF 檔案
        static void Main(string[] args)
        {
            string _fileNameError;
            string _fileName;

            _fileNameError = "newPageError.pdf";
            _fileName = "newPage.pdf";

            byte[] _buffer;

            Document _document;
            PdfWriter _writer;

            #region 錯誤案例：使用 NewPage() 方法後沒有在第二個頁面上增加任何物件，PDF檔案並不會有第二頁

            //使用記憶體資料串流處理新建立的 PDF 檔案
            //沒有搭配 using 會出現 exception
            using (MemoryStream _inputStream = new MemoryStream())
            {
                _document = new Document(PageSize.A4);
                _writer = PdfWriter.GetInstance(_document, _inputStream);

                _document.Open();
                _document.Add(new Chunk("Page 1/1"));

                //新增空白頁面
                _document.NewPage();
                _document.Close();

                _buffer = _inputStream.ToArray();
            }

            //將產生的 byte[] 儲存成檔案
            File.WriteAllBytes(_fileNameError, _buffer);

            //開啟新建立的 PDF 檔案
            System.Diagnostics.Process.Start(_fileNameError);

            #endregion

            #region 案例：使用 NewPage() 方法在第二個頁面上有增加物件，PDF檔案會出現第二頁

            //使用記憶體資料串流處理新建立的 PDF 檔案
            //沒有搭配 using 會出現 exception
            using (MemoryStream _inputStream = new MemoryStream())
            {
                _document = new Document(PageSize.A4);
                _writer = PdfWriter.GetInstance(_document, _inputStream);

                _document.Open();
                _document.Add(new Chunk("Page 1/2"));

                //新增空白頁面
                _document.NewPage();
                //增加一個顯示物件：Chunk 就類似 Html 的 span
                _document.Add(new Chunk(string.Empty));
                _document.Close();

                _buffer = _inputStream.ToArray();
            }

            //將產生的 byte[] 儲存成檔案
            File.WriteAllBytes(_fileName, _buffer);

            //開啟新建立的 PDF 檔案
            System.Diagnostics.Process.Start(_fileName);

            #endregion
        }
    }
}
