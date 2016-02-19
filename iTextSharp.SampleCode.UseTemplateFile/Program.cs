using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleCode.UseTemplateFile
{
    class CardInfo
    {
        public string Name { get; set; }
        public DateTime? Birthday { get; set; }
        public string Location { get; set; }
        public string EmployeeID { get; set; }
        public string Notify { get; set; }

        public byte[] ProfileImage { get; set; }
    }

    class Program
    {
        //此範例使用 iTextSharp 將文字內容加入到指定範本 PDF 檔案

        /*
         * 模擬情境：
         *  使用員工資料（包含圖片）輸入到資料卡範本 PDF 檔案的指定位置
         *  並將檔案儲存為 byte[]
         */
        static void Main(string[] args)
        {
            string _fileName;
            string _templateName;

            _templateName = "template-card.pdf";
            _fileName = "Card.pdf";

            byte[] _templateBuffer;
            byte[] _buffer;

            //模擬物件
            CardInfo _cardInfo;

            _cardInfo = new CardInfo();
            _cardInfo.Name = "未命名";
            _cardInfo.Birthday = new DateTime(9999, 12, 31);
            _cardInfo.Location = "工程師之家";
            _cardInfo.Notify = "請緊握扶手站穩踏階";
            _cardInfo.EmployeeID = "Z9999";
            _cardInfo.ProfileImage = File.ReadAllBytes("person-empty.png");


            #region 準備自訂的字型物件
            string _fontFolder;

            _fontFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

            BaseFont _msjhBaseFont;
            Font _msjhFont;
            Font _msjhBoldFont;
            Font _msjhSmallFont;

            //設定基礎字型物件：此設定為 "微軟正黑體"
            _msjhBaseFont = BaseFont.CreateFont(Path.Combine(_fontFolder, "msjh.ttf")
                , BaseFont.IDENTITY_H
                , BaseFont.EMBEDDED);

            //指定字型物件使用微軟正黑體，字型大小為 12
            _msjhFont = new Font(_msjhBaseFont, 12);

            //指定字型物件使用微軟正黑體，字型大小為 12，粗體字
            _msjhBoldFont = new Font(_msjhBaseFont, 12, Font.BOLD);

            //指定字型物件使用微軟正黑體，字型大小為 8
            _msjhSmallFont = new Font(_msjhBaseFont, 8);
            #endregion


            Document _document;
            PdfWriter _writer;
            PdfReader _reader;

            PdfContentByte _contentByte;
            PdfImportedPage _importPage;

            //讀取範本 PDF 為 PdfReader 物件
            _templateBuffer = File.ReadAllBytes(_templateName);
            _reader = new PdfReader(_templateBuffer);

            using (MemoryStream _inputStream = new MemoryStream())
            {
                _document = new Document(PageSize.A4);
                _writer = PdfWriter.GetInstance(_document, _inputStream);

                _document.Open();

                #region 將範本 PDF 匯入到正在編輯的 PdfWriter 物件
                //取得範本 PDF 第一頁內容
                _importPage = _writer.GetImportedPage(_reader, 1);

                //將範本的第一頁內容加入到 PdfContentByte
                _contentByte = _writer.DirectContent;
                _contentByte.AddTemplate(_importPage, 0, 0);
                #endregion


                #region 將指定文字或圖片加入到 PDF
                ColumnText _columnText;
                Rectangle _rectangle;
                Phrase _phrase;

                _columnText = new ColumnText(_contentByte);


                _rectangle = new Rectangle(210, 680, 340, 700);
                _phrase = new Phrase("名　　字："+_cardInfo.Name,_msjhFont);
                _columnText.SetSimpleColumn(_rectangle);
                _columnText.SetText(_phrase);
                _columnText.Go();

                _rectangle = new Rectangle(210, 660, 340, 680);
                _phrase = new Phrase(string.Format("出生日期：{0:yyyy/MM/dd}", _cardInfo.Birthday), _msjhFont);
                _columnText.SetSimpleColumn(_rectangle);
                _columnText.SetText(_phrase);
                _columnText.Go();

                _rectangle = new Rectangle(210, 640, 340, 660);
                _phrase = new Phrase("服務地點：" + _cardInfo.Location, _msjhFont);
                _columnText.SetSimpleColumn(_rectangle);
                _columnText.SetText(_phrase);
                _columnText.Go();

                _rectangle = new Rectangle(210, 620, 340, 640);
                _phrase = new Phrase("員工編號："+ _cardInfo.EmployeeID, _msjhFont);
                _columnText.SetSimpleColumn(_rectangle);
                _columnText.SetText(_phrase);
                _columnText.Go();


                
                _rectangle = new Rectangle(270, 405, 350, 425);
                _phrase = new Phrase(_cardInfo.Notify, _msjhSmallFont);
                _columnText.SetSimpleColumn(_rectangle);
                _columnText.SetText(_phrase);
                _columnText.Go();


                //透過繪製矩形取得絕對位置擺放的相對位置，取消註解後自行服用
                //_rectangle = new Rectangle(190, 405, 350, 425);
                //_rectangle.BackgroundColor = BaseColor.LIGHT_GRAY;
                //_contentByte.Rectangle(_rectangle);

                //_columnText.SetSimpleColumn(_phrase, 190, 385, 350, 425, 12, Element.ALIGN_RIGHT);
                //_columnText.Go();

                //_columnText.SetSimpleColumn(_phrase, 190, 385, 350, 425, 24, Element.ALIGN_LEFT);
                //_columnText.Go();

                //_columnText.SetSimpleColumn(_phrase, 190, 385, 350, 425, 24, Element.ALIGN_CENTER);
                //_columnText.Go();


                Image _profileImage;

                _profileImage = Image.GetInstance(_cardInfo.ProfileImage);
                _profileImage.SetAbsolutePosition(125, 605);

                _contentByte.AddImage(_profileImage);


                #endregion

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
