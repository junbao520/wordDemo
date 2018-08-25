
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Web;
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

            //git hub demo 学习git的基本操作 tes

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

        //todo:目前最核心的就是支付宝里面的钱 可以考虑修改到我的电信卡上面进行绑定
        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = new Document();
            var path = System.IO.Directory.GetCurrentDirectory() + "\\" + "test" + ".doc";

            var tt1 = DateTime.Now;
            doc.LoadFromFile(path);


            //doc.SaveToFile(System.IO.Directory.GetCurrentDirectory() + "\\"+"111.pdf", Spire.Doc.FileFormat.PDF);


            System.Drawing.Image image = doc.SaveToImages(0, Spire.Doc.Documents.ImageType.Metafile);
            //todo:保存成为图片不理想 //
            image.Save("sample.jpg", ImageFormat.Jpeg);
            var tt2 = DateTime.Now;
            var ts2 = (tt2 - tt1).TotalMilliseconds;

            MessageBox.Show("ok" + ts2);

            DateTime t1 = new DateTime();
            WriteIntoWord word = new WriteIntoWord();
            var bitmap = word.WordtoImage(path);
            bitmap[0].Save("sample1.jpg", ImageFormat.Jpeg);
            DateTime t2 = new DateTime();
            var ts = (t2 - t1).TotalMilliseconds;



            MessageBox.Show("ok" + ts);


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

        private void button2_Click(object sender, EventArgs e)
        {

            var path = "D:\\123456.JPG";
            //if else else if 
            //if else else if 
            //if else else if 
            //I dot know how to do something 
            //I dot knwo how to do something 
            //if else else if 
            //if else else if 
             //
            //udp 通讯工具
            //udp 通讯工具
            //udp 通讯工具

            var userName = txtServerName.Text.Trim();
            var pwd = txtPwd.Text.Trim();
            var Ip = txtIP.Text.Trim();
            var ServerName = txtServerName.Text.Trim();
            var IdCard = txtIDCard.Text.Trim();
            try
            {
                //var con = SLKM2.Conn.ConnKm2(userName, pwd, Ip, ServerName);
                //if (con)
                //{
                //    MessageBox.Show("连接成功");
                //}
                //else
                //{
                //    MessageBox.Show("连接失败");
                //}
                var msg = string.Empty;
                bool result = SLKM2.Km2Oracle.Km2KsStartBd(IdCard, "capture", ref msg);
                if (result)
                {
                    MessageBox.Show("签到成功");
                }
                else
                {
                    MessageBox.Show("签到失败");
                }
                MessageBox.Show(msg);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        //if else if else  if else if else if wlaw i
        //.......................................
        //.......................................
        //......................................
        //ifelseelseifelseifelseifelseifelseif elseifelseifelseifelseifif
        //ifelseifelseifelseifelseifelseifelseifelseifelseifelseifelseifslse
        //ifelseifelseifelseelseifelseifelseifelseifelseifelseifelseifelseifelseifelseifelseifelseifelseif
        //ifelseifelseifelseifelseififelseifelseififelseififelseifelseifeleelse
        //if else else i
        public string PostPics(string url, List<string> picName, List<string> picPathList, Dictionary<string, string> param)
        {
            string boundary = "-----------------------------7e2292f20603";
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Method = "POST";
            request.ContentType = "multipart/form-data; boundary=" + boundary;
            // 设置参数
            // request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
            // request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36";
            request.AllowAutoRedirect = false;
            // request.ContentType = "multipart/form-data; boundary=----WebKitFormBoundaryKKSsFFnaCFSR8ic7";
            byte[] itemBoundaryBytes = Encoding.UTF8.GetBytes("\r\n--" + boundary + "\r\n");
            byte[] endBoundaryBytes = Encoding.UTF8.GetBytes("\r\n--" + boundary + "--\r\n");
            Stream postStream = request.GetRequestStream();
            postStream.Write(itemBoundaryBytes, 0, itemBoundaryBytes.Length);
            //files
            for (int i = 0; i < picPathList.Count; i++)
            {
                string FilePath = picPathList[i];
                int pos = FilePath.LastIndexOf("\\");
                string fileName = FilePath.Substring(pos + 1);
                //请求头部信息 
                //
                string postDataName = picName[i];
                StringBuilder sbHeader = new StringBuilder(string.Format("Content-Disposition:form-data;name=\"{0}\";filename=\"{1}\"\r\n\r\nContent-Type: image/jpeg\r\n\r\n", postDataName, fileName));
                byte[] postHeaderBytes = Encoding.UTF8.GetBytes(sbHeader.ToString());


                FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
                byte[] bArr = new byte[fs.Length];
                fs.Read(bArr, 0, bArr.Length);
                fs.Close();

                //file
                postStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
                postStream.Write(bArr, 0, bArr.Length);
                postStream.Write(itemBoundaryBytes, 0, itemBoundaryBytes.Length);
            }

           // postStream.Write(endBoundaryBytes, 0, endBoundaryBytes.Length);
           // postStream.Close();

            //错误处理
            StreamReader sr;
            HttpStatusCode hs = new HttpStatusCode();
            try
            {
                //发送请求并获取相应回应数据
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                //直到request.GetResponse()程序才开始向目标网页发送Post请求
                Stream instream = response.GetResponseStream();
                hs = response.StatusCode;
                sr = new StreamReader(instream, Encoding.UTF8);

            }
            catch (WebException ex)
            {
                Stream responseStream = ex.Response.GetResponseStream();
                sr = new StreamReader(responseStream, Encoding.UTF8);
            }

            //返回结果网页（html）代码
            string content = sr.ReadToEnd();
            // Debug.Log(content);
            // Debug.Log("ylj=" + hs.ToString());
            return content;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var url = "http://2176mf7449.51mypc.cn:57268/upload";
            var fileNames = new List<string>() { "myfile2", "myfile" };

            var dir = System.Environment.CurrentDirectory;
            var picPathList = new List<string>() { dir + "\\" + "test11.jpg", dir + "\\" + "test11.jpg" };
            var param = new Dictionary<string, string>();


            var result = PostPics(url, fileNames, picPathList, param);
            // UploadImage(System.Environment.CurrentDirectory + "\\" + "capture.jpg");
            //  postData();


        }



        /// <summary>
        /// 通过http上传图片及传参数
        /// </summary>
        /// <param name="imgPath">图片地址(绝对路径：D:\demo\img\123.jpg)</param>
        public void UploadImage(string imgPath)
        {
            var uploadUrl = "http://2176mf7449.51mypc.cn:57268/upload";
            //var dic = new Dictionary<string, string>() {
            //    {"para1",1.ToString() },
            //    {"para2",2.ToString() },
            //    {"para3",3.ToString() },
            //};
            //var postData = Utils.BuildQuery(dic);//转换成：para1=1&para2=2&para3=3
            var postUrl = uploadUrl;
            HttpWebRequest request = WebRequest.Create(postUrl) as HttpWebRequest;
            request.AllowAutoRedirect = true;
            request.Method = "POST";
            request.Timeout = 50000;

            string boundary = DateTime.Now.Ticks.ToString("X"); // 随机分隔线
            request.ContentType = "multipart/form-data;charset=utf-8;boundary=" + boundary;
            byte[] itemBoundaryBytes = Encoding.UTF8.GetBytes("\r\n--" + boundary + "\r\n");
            byte[] endBoundaryBytes = Encoding.UTF8.GetBytes("\r\n--" + boundary + "--\r\n");

            int pos = imgPath.LastIndexOf("\\");
            string fileName = imgPath.Substring(pos + 1);

            //请求头部信息 
            StringBuilder sbHeader = new StringBuilder(string.Format("Content-Disposition:form-data;name=\"myfile\";filename=\"{0}\"\r\n\r\n Content-Type:application/octet-stream\r\n\r\n", fileName));
            byte[] postHeaderBytes = Encoding.UTF8.GetBytes(sbHeader.ToString());

            FileStream fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read);
            byte[] bArr = new byte[fs.Length];
            fs.Read(bArr, 0, bArr.Length);
            fs.Close();

            Stream postStream = request.GetRequestStream();
            postStream.Write(itemBoundaryBytes, 0, itemBoundaryBytes.Length);
            postStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
            postStream.Write(bArr, 0, bArr.Length);
            postStream.Write(endBoundaryBytes, 0, endBoundaryBytes.Length);
            postStream.Close();

            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            Stream instream = response.GetResponseStream();
            StreamReader sr = new StreamReader(instream, Encoding.UTF8);
            string content = sr.ReadToEnd();
        }



        public void postData()
        {
            FileStream fs = new FileStream("capture.jpg", FileMode.Open, FileAccess.Read);
            byte[] byteFile = new byte[fs.Length];
            fs.Read(byteFile, 0, Convert.ToInt32(fs.Length));
            fs.Close();

            string postString = "ip=192.168.0.1&idcard=500227199111294612";//这里即为传递的参数，可以用工具抓包分析，也可以自己分析，主要是form里面每一个name都要加进来  
            postString = string.Format("myfile={0}&myfile2={1}", HttpUtility.UrlEncode(Convert.ToBase64String(byteFile)), HttpUtility.UrlEncode(Convert.ToBase64String(byteFile)));
            byte[] postData = Encoding.UTF8.GetBytes(postString);//编码，尤其是汉字，事先要看下抓取网页的编码方式  
            WebClient webClient = new WebClient();
            webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");//采取POST方式必须加的header，如果改为GET方式的话就去掉这句话即可
            //webClient.Headers.Add("Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;");

            byte[] responseData = webClient.UploadData("http://2176mf7449.51mypc.cn:57268/upload", "POST", postData);//得到返回字符流  
            string srcString = Encoding.UTF8.GetString(responseData);//解码 

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string sendString ="hellword";//要发送的字符串 
            byte[] sendData = null;//要发送的字节数组 
            UdpClient client = null;
            var ip = txtIP.Text.Trim();
            IPAddress remoteIP = IPAddress.Parse(ip); //假设发送给这个IP
            int remotePort = Convert.ToInt32(txtPort.Text.Trim());
            IPEndPoint remotePoint = new IPEndPoint(remoteIP, remotePort);//实例化一个远程端点 


            sendString = Console.ReadLine();
            sendData = Encoding.Default.GetBytes("hellword");

            client = new UdpClient();
            client.Send(sendData, sendData.Length, remotePoint);//将数据发送到远程端点 
            client.Close();//关闭连接 


        }

   
            public  bool GetPicThumbnail(string sFile, string dFile, int dHeight, int dWidth, int flag)
            {
                System.Drawing.Image iSource = System.Drawing.Image.FromFile(sFile);
                ImageFormat tFormat = iSource.RawFormat;
                int sW = 0, sH = 0;

                //按比例缩放
                Size tem_size = new Size(iSource.Width, iSource.Height);

                if (tem_size.Width > dHeight || tem_size.Width > dWidth)
                {
                    if ((tem_size.Width * dHeight) > (tem_size.Width * dWidth))
                    {
                        sW = dWidth;
                        sH = (dWidth * tem_size.Height) / tem_size.Width;
                    }
                    else
                    {
                        sH = dHeight;
                        sW = (tem_size.Width * dHeight) / tem_size.Height;
                    }
                }
                else
                {
                    sW = tem_size.Width;
                    sH = tem_size.Height;
                }

                Bitmap ob = new Bitmap(dWidth, dHeight);
                Graphics g = Graphics.FromImage(ob);

                g.Clear(Color.WhiteSmoke);
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                g.DrawImage(iSource, new Rectangle((dWidth - sW) / 2, (dHeight - sH) / 2, sW, sH), 0, 0, iSource.Width, iSource.Height, GraphicsUnit.Pixel);

                g.Dispose();
                //以下代码为保存图片时，设置压缩质量  
                EncoderParameters ep = new EncoderParameters();
                long[] qy = new long[1];
                qy[0] = flag;//设置压缩的比例1-100  
                EncoderParameter eParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qy);
                ep.Param[0] = eParam;
                try
                {
                    ImageCodecInfo[] arrayICI = ImageCodecInfo.GetImageEncoders();
                    ImageCodecInfo jpegICIinfo = null;
                    for (int x = 0; x < arrayICI.Length; x++)
                    {
                        if (arrayICI[x].FormatDescription.Equals("JPEG"))
                        {
                            jpegICIinfo = arrayICI[x];
                            break;
                        }
                    }
                    if (jpegICIinfo != null)
                    {
                        ob.Save(dFile, jpegICIinfo, ep);//dFile是压缩后的新路径  
                    }
                    else
                    {
                        ob.Save(dFile, tFormat);
                    }
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    iSource.Dispose();
                    ob.Dispose();
                }
            }

        private void button5_Click(object sender, EventArgs e)
        {
            GetPicThumbnail("tests.jpg", "new.jpg", 130, 160, 100);
            MessageBox.Show("ok");
        }
    }
}
