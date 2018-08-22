
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
            

            //

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = new Document();
            var path = System.IO.Directory.GetCurrentDirectory() + "\\" + "test" + ".doc";

            var tt1= DateTime.Now;
            doc.LoadFromFile(path);
            //doc.SaveToFile(System.IO.Directory.GetCurrentDirectory() + "\\"+"111.pdf", Spire.Doc.FileFormat.PDF);

            System.Drawing.Image image = doc.SaveToImages(0, Spire.Doc.Documents.ImageType.Metafile);
            //todo:保存成为图片不理想
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
