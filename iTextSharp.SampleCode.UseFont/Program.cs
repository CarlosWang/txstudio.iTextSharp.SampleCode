using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleCode.UseFont
{
    class Program
    {
        //此範例使用 iTextSharp 透過設定中文字型的方式建立包含中文字的 PDF 檔案
        static void Main(string[] args)
        {
            string _fileName;

            _fileName = "UseFont.pdf";

            byte[] _buffer;


            #region 準備自訂的字型物件
            string _fontFolder;

            _fontFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

            BaseFont _msjhBaseFont;
            Font _msjhFont;
            Font _msjhBoldFont;

            //設定基礎字型物件：此設定為 "微軟正黑體"
            _msjhBaseFont = BaseFont.CreateFont(Path.Combine(_fontFolder, "msjh.ttf")
                , BaseFont.IDENTITY_H
                , BaseFont.EMBEDDED);

            //指定字型物件使用微軟正黑體，字型大小為 16
            _msjhFont = new Font(_msjhBaseFont, 16);

            //指定字型物件使用微軟正黑體，字型大小為 16，粗體字
            _msjhBoldFont = new Font(_msjhBaseFont, 16, Font.BOLD);
            #endregion


            Document _document;
            PdfWriter _writer;

            using (MemoryStream _inputStream = new MemoryStream())
            {
                _document = new Document(PageSize.A4);
                _writer = PdfWriter.GetInstance(_document, _inputStream);

                _document.Open();

                _document.Add(new Paragraph("use font support zh-tw character 指定字型內容就可以顯示中文字", _msjhFont));
                _document.Add(new Paragraph("use font support zh-tw character 指定字型內容就可以顯示中文字-粗體", _msjhBoldFont));

                //新增空白頁面
                _document.NewPage();
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
