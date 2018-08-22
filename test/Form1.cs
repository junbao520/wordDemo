
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;

namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //LastTest 最后一次提交测试 这个只是一个Demo 里面包含了word的基本操作，可以看看
            //todo:this is  creatDocument
            //WriteIntoWord wiw = null;
            //wiw = new WriteIntoWord();
            //var path = System.IO.Directory.GetCurrentDirectory() + "\\" + "test.dotx";
            //wiw.CreateNewDocument(path);
            //var picpath = System.IO.Directory.GetCurrentDirectory() + "\\" + "capture.png";
            ////测试生成word
            //string[] strBookmarks = { "machine", "feedback", "macname", "path", "question", "repairdate", "solution", "trouble" };
            //wiw.WriteIntoDocument("machine", "firstTest" + DateTime.Now.ToString());
            //wiw.WritePicIntoDocument("question", picpath);
            //path = System.IO.Directory.GetCurrentDirectory() + "\\" + "test" + ".doc";
            //wiw.Save_CloseDocument(path);


           // 第一步：组件安装后，创建一个C#控制台项目，添加引用及命名空间如下：


           // Document doc = new Document();

           // doc.LoadFromFile("sample.doc");


           // 第三步：实例化一个PrintDialog的对象，设置相关属性。关联doc.PrintDialog属性和PrintDialog对象:



           //  PrintDialog dialog = new PrintDialog();

           // dialog.AllowPrintToFile = true;

           // dialog.AllowCurrentPage = true;

           // dialog.AllowSomePages = true;

           // dialog.UseEXDialog = true;

           // doc.PrintDialog = dialog;


           // 第四步: 后台打印。使用默认打印机打印出所有页面。这段代码也可以用于网页后台打印:

           //PrintDocument printDoc = doc.PrintDocument;

           // printDoc.Print();


           // 第五步: 如要显示打印对话框，就调用ShowDialog方法，根据打印预览设置选项，打印word文档:

           //   if (dialog.ShowDialog() == DialogResult.OK)

           // {

           //     printDoc.Print();

           // }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = new Document();
            var path = System.IO.Directory.GetCurrentDirectory() + "\\" + "test" + ".doc";

            var tt1= DateTime.Now;
            doc.LoadFromFile(path);

          
            //doc.SaveToFile(System.IO.Directory.GetCurrentDirectory() + "\\"+"111.pdf", Spire.Doc.FileFormat.PDF);

           
            System.Drawing.Image image = doc.SaveToImages(0, Spire.Doc.Documents.ImageType.Metafile);
            //todo:保存成为图片不理想 //
            image.Save("sample.jpg", ImageFormat.Jpeg);
            var tt2 = DateTime.Now;
            var ts2 = (tt2 - tt1).TotalMilliseconds;

            MessageBox.Show("ok"+ts2);

            DateTime t1 = new DateTime();
            WriteIntoWord word = new WriteIntoWord();
             var bitmap= word.WordtoImage(path);
             bitmap[0].Save("sample1.jpg", ImageFormat.Jpeg);
             DateTime t2 = new DateTime();
            var ts = (t2 - t1).TotalMilliseconds;



            MessageBox.Show("ok"+ts);


            //Spire.Pdf.PdfDocument pdf = new Spire.Pdf.PdfDocument();
            //pdf.LoadFromFile(path);
            //pdf.SaveToFile(a.AppPath() + @"\111.doc", Spire.Pdf.FileFormat.DOC);
            //pdf = null;





            //PrintDialog dialog = new PrintDialog();
            //dialog.AllowPrintToFile = true;
            //dialog.AllowCurrentPage = true;
            //dialog.AllowSomePages = true;
            //dialog.UseEXDialog = true;

            //if (dialog.ShowDialog() == DialogResult.OK)
            //{
            //    printDoc.Print();
            //}

        }
    }
}
