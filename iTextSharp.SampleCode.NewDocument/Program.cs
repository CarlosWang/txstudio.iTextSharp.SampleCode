using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleCode.NewDocument
{
    class Program
    {
        //此範例為使用 iTextSharp 建立一個 A4 尺寸的 PDF 檔案
        //並在檔案中插入當下時間
        static void Main(string[] args)
        {
            string _timeNow;
            string _fileName;

            _timeNow = string.Format("{0:yyyy/MM/dd HH:mm:ss}", DateTime.Now);
            _fileName = "document.pdf";


            byte[] _buffer;

            Document _document;
            PdfWriter _writer;


            //使用記憶體資料串流處理新建立的 PDF 檔案
            //沒有搭配 using 會出現 exception
            using (MemoryStream _inputStream = new MemoryStream())
            {
                //建立 A4 尺寸的 PDF 文件
                _document = new Document(PageSize.A4);
                _writer = PdfWriter.GetInstance(_document, _inputStream);

                //加入當下時間的文字內容
                _document.Open();
                _document.Add(new Chunk(_timeNow));
                _document.Close();

                _buffer = _inputStream.ToArray();
            }


            //將產生的 byte[] 儲存成檔案
            File.WriteAllBytes(_fileName, _buffer);

            //開啟新建立的 PDF 檔案
            System.Diagnostics.Process.Start(_fileName);

        }
    }
}
