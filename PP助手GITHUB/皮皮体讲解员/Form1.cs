using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
//using Microsoft.CognitiveServices.Speech;
//using ss= System.Speech.Synthesis;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using AlphaForm;
using System.Xml;

namespace 皮皮助手
{
    public partial class Form1 : Form
    {
        List<bool> isSpeak = new List<bool>();//用于记录每一页的Note是否被阅读过。
        int currenSpeakIndex = -1;
        PowerPoint.Presentation ppr = null;
        PowerPoint.Application papp = null;
        List<string> notes = new List<string>();
        //SpeechSynthesizer synthesizer = null;
        bool AudioCompleted = false;
        //获取PPT的页码。
        private Thread thPage;
        private int pg = 0;//全局PPT当前页码
        bool exitThread = false;
        Thread thSpeak = null;
        Thread thStopSpeak = null;
        Thread incWeb = null;
       ExampleServer server = null;

        private void init()
        {
            textBox1.Text = "";
            currenSpeakIndex = -1;
            isSpeak = new List<bool>();
            notes = new List<string>();
            papp = null;
            //if (thPage!=null) thPage.Abort();
            if (thSpeak != null) thSpeak.Abort();
            if (incWeb != null) incWeb.Abort();
            isClose = false;
            //Application.ExitThread();

          
        }
        public void LoadConfig()
        {
            using (StreamReader fs = new StreamReader("config.ini"))
            {
                Global.spkName = fs.ReadLine();
                Global.edgePath = fs.ReadLine();
                Global.autoPlay = Convert.ToBoolean(fs.ReadLine());
                fs.Close();
            }
        }

        public Form1()
        {
            InitializeComponent();
            this.Location = new Point(20, System.Windows.Forms.SystemInformation.WorkingArea.Height - 430);
            Control.CheckForIllegalCrossThreadCalls = false;
            LoadConfig();
            if (Global.autoPlay) chkAuto.Checked = true;
            //speakEdge("");//初始化文件
            ////启动朗读的WEBSERVER
            //Thread webTh = new Thread(startWeb);
            //webTh.IsBackground = true;
            //webTh.Start();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            notes = new List<string>();//清空上一个PPT的备注内容
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "C# Corner Open File Dialog";
            //fdlg.InitialDirectory = @"d:\";   //@是取消转义字符的意思
            fdlg.Filter = "幻灯片(*.ppt)|*.ppt|幻灯片高版本(*.pptx)|*.pptx";
            /*
             * FilterIndex 属性用于选择了何种文件类型,缺省设置为0,系统取Filter属性设置第一项
             * ,相当于FilterIndex 属性设置为1.如果你编了3个文件类型，当FilterIndex ＝2时是指第2个.
             */
            fdlg.FilterIndex = 2;
            /*
             *如果值为false，那么下一次选择文件的初始目录是上一次你选择的那个目录，
             *不固定；如果值为true，每次打开这个对话框初始目录不随你的选择而改变，是固定的  
             */
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = System.IO.Path.GetFileNameWithoutExtension(fdlg.FileName);
                txtPath.Text = fdlg.FileName;

            }

        }

        PowerPoint.Slide sl = null;

        private void readNotes()
        {
            string fullname = txtPath.Text;
            papp = new PowerPoint.Application();

            ppr = papp.Presentations.Open(fullname, Microsoft.Office.Core.MsoTriState.msoCTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            PowerPoint.SlideShowSettings slideShow = ppr.SlideShowSettings;

            slideShow.Run();
            ppr.SlideShowWindow.View.GotoSlide(1);//打开第一页

            string doc = "";
            int i = 0;
            foreach (PowerPoint.Slide slide in ppr.Slides)
            {
                sl = slide;
                i++;
                doc += "\r\n第" + i.ToString() + "页的备注：";
                string nt = "";
                foreach (PowerPoint.Shape shape in slide.NotesPage.Shapes)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            nt = shape.TextFrame.TextRange.Text;
                            doc += nt + "\r\n";
                        }
                    }
                }
                notes.Add(nt);
                isSpeak.Add(false);
            }
            //papp = null;
            textBox1.Text += "一共发现PPT备注：" + notes.Count.ToString() + "\r\n";
            textBox1.Text += doc;
           

        }
        private async void speakEdge(string text)
        {
            StreamWriter wt = new StreamWriter("readContent.txt");
            text = comboBox1.Text + "@*@" + text;
            wt.Write(text);
            wt.Close();
        }

        //public  async Task SynthesisToSpeakerAsync(string text)
        //{
        //    if (text == "")
        //    {
        //        textBox1.Text += $"识别读取文本： [{text}]\r\n";
        //        return;
        //    }
        //    var config = SpeechConfig.FromSubscription("6274bb247f494139a9d9c8a5866f0889", "westus2");
        //    config.SpeechRecognitionLanguage = "zh-CN";
        //    config.SpeechSynthesisLanguage = "zh-CN";
        //    //https://docs.microsoft.com/en-us/azure/cognitive-services/speech-service/language-support#neural-voices
        //    config.SpeechSynthesisVoiceName = comboBox1.Text;
        //    synthesizer = null;
        //    using (synthesizer = new SpeechSynthesizer(config))
        //    {
        //        using (var result = await synthesizer.SpeakTextAsync(text))
        //        {

        //            if (result.Reason == ResultReason.SynthesizingAudioCompleted)
        //            {
        //                AudioCompleted = true;
        //                textBox1.Text += $"识别读取文本： [{text}]\r\n";
        //            }
        //            else if (result.Reason == ResultReason.Canceled)
        //            {
        //                var cancellation = SpeechSynthesisCancellationDetails.FromResult(result);
        //                textBox1.Text += ($"CANCELED: Reason={cancellation.Reason}\r\n");

        //                if (cancellation.Reason == CancellationReason.Error)
        //                {
        //                    textBox1.Text += ($"CANCELED: ErrorCode={cancellation.ErrorCode}\r\n");
        //                    textBox1.Text += ($"CANCELED: ErrorDetails=[{cancellation.ErrorDetails}]\r\n");
        //                    textBox1.Text += ($"CANCELED: Did you update the subscription info?\r\n");
        //                }
        //            }
        //        }
        //        // This is to give some time for the speaker to finish playing back the audio
        //    }
        //}

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                init();
                //minsizeForm();//最小化窗体
                if (txtPath.Text != "")
                {
                    readNotes();
                }
                //注册绑定幻灯片的事件
                registAllEvents();
                speakFirst();//朗读第一页
                //thPage = new Thread(new ThreadStart(getPage));
                //thPage.IsBackground = true;
                //thPage.Start();
                //MessageBox.Show("1");

                tabControl1.SelectedIndex = 1;

                //thSpeak = new Thread(new ThreadStart(speaker));
                //thSpeak.IsBackground = true;
                //thSpeak.Start();//开始讲话线程

                //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "--start-maximized --app=http://127.0.0.1:5050/spk.html");//pplication\msedge.exe" --start-maximized --app=http://127.0.0.1:4050/spk.html
                Process.Start(txtEDGE.Text, " --app=http://127.0.0.1:5050/spk.html");//pplication\msedge.exe" --start-maximized --app=http://127.0.0.1:4050/spk.html

                //incWeb = new Thread(new ThreadStart(includeWebForm));
                //incWeb.IsBackground = true;
                //incWeb.Start();//开始讲话线程


                isClose = false;
                //pg = 0;

                if (chkAuto.Checked) this.timer1.Enabled = true;//自动播放
            }
            catch (Exception)
            {
                ;
            }

        }
        /// <summary>
        /// 将网站页面包装到窗体上,或者是最小化
        /// </summary>
        private void includeWebForm()
        {
            //隐藏播放窗体
            int spkForm = 0;
            while (spkForm == 0)
            {
                Thread.Sleep(200);
                spkForm = FindWindow(null, "speak");//一直获取窗口句柄
                //break;
            }
            if (!isSeted)
            {
                //ShowWindow(spkForm, SW_HIDE);
                //子窗口句柄,父窗口句柄（我这里用的Winform里的panel控件的句柄，这样就会将我的子窗口嵌入到panel里面）
                //SetParent(spkForm, panel1.Handle);
                //Int32 wndStyle = GetWindowLong(spkForm, -16);
                //wndStyle &= ~WS_BORDER;
                //wndStyle &= ~WS_THICKFRAME;
                //SetWindowLong(spkForm, -16, wndStyle);
                //MoveWindow(spkForm, 0, 0, panel1.Width, panel1.Height, true);

                ShowWindow(spkForm,2);//1    正常大小显示窗口//2    最小化窗口3    最大化窗口
                isSeted = true;
            }

        }
        private void startWeb()
        {
            if (server == null)
            {
                server = new ExampleServer("127.0.0.1", 5050);
                server.SetRoot("./");
                server.Logger = new ConsoleLogger();
                server.Start();
            }

        }

        bool isRead = false;//判断是否朗读过一次，因为第二次需要重启EDGE浏览器进程
        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
        const int WM_CLOSE = 0x0010;



        /// <summary>
        /// 朗读输入的文本
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        { 
            if (isRead)
            {
                //关闭浏览器进程
                closeEdge();
            }

            //SynthesisToSpeakerAsync(txtInput.Text);
            speakEdge(txtInput.Text);
            Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", " --app=http://127.0.0.1:5050/spk.html");//pplication\msedge.exe" --start-maximized --app=http://127.0.0.1:4050/spk.html   --no-startup-window --incognito

            isRead = true;
        }


        void minSizeEdge()
        {
            try
            {
                Thread.Sleep(200);
                int spkForm = FindWindow(null, "speak");
            }
            catch (Exception)
            {
                ;
            }
        }

        private void closeEdge()
        {
            try
            {
                Thread.Sleep(200);
                int spkForm = FindWindow(null, "speak");
                if (spkForm == 0) return;
                SendMessage(new IntPtr(spkForm), WM_CLOSE, new IntPtr(0), new IntPtr(0));
            }
            catch (Exception)
            {
                ;
            }
        }
        private void txtInput_Click(object sender, EventArgs e)
        {
            if (txtInput.Text == "请输入朗读的内容。")
            {
                txtInput.Text = "";
            }
        }
        /// <summary>
        /// 结束播放
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void btnEnd_Click(object sender, EventArgs e)
        {
            //if (papp == null) return;
            //await synthesizer.StopSpeakingAsync().ConfigureAwait(false);
            try
            {
                if (ppr != null)
                {
                    ppr.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ppr);
                }
                ppr = null;
                if (papp != null)
                {
                    papp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(papp);
                }
                KillProcess("wpp");
                KillProcess("POWERPNT");
            }
            catch (Exception)
            {
                ;
            }
        }


        //private void set_lableText(string s) //主线程调用的函数
        //{
        //    lblPage.Text = "当前第" + s + "页";
        //    pg = Convert.ToInt32(s);
        //}
        //delegate void set_Text(string s); //定义委托

        //set_Text Set_Text; //定义委托


        bool isClose = false;//一旦PPT退出，则标记pg=-1,
        private void getPage()
        {

            while (true)
            {
                //if (exitThread) break;//线程退出标志
                try
                {
                    if (ppr == null)
                    {
                        Thread.Sleep(750);
                        continue;
                    }
                    string page = ppr.SlideShowWindow.View.Slide.SlideIndex.ToString();  //这里是获取当前播放的幻灯片的页码

                    //不阻塞主线程
                    this.BeginInvoke(new EventHandler(delegate
                    {
                        lblPage.Text = "当前第" + page.ToString() + "页";
                        pg = Convert.ToInt32(page);
                    }));
                    Thread.Sleep(100);
                   
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Object does not exist"))
                    {
                        if (!isClose)
                        {
                            closeEdge();//用于处理PPT播放关闭后WEB页面没有关的问题
                            pg = -1;//一旦PPT退出，则标记pg=-1,
                            restoreForm();//恢复窗体大小
                        }
                        isClose = true;//表示本异常只执行一次,如果开起新的朗读，则标记为false;
                    }
                    Thread.Sleep(750);
                    continue;//PPT没打开的时候才会触发
                }
            }
        }

        private  void speaker()
        {

            while (true)
            {

                if (isSpeak.Count == 0)
                {
                    Thread.Sleep(100);
                    continue;
                }
                if (notes.Count == 0)
                {
                    Thread.Sleep(100);
                    continue;
                }
                if (pg == -1) break;//表示PPT已经关闭，所以退出线程循环
                try
                {
                    if (pg >= 1)
                    {
                        if (isSpeak[pg - 1] == false)//表示没有读过
                        {
                            currenSpeakIndex = pg;//记录当前正在阅读的索引
                            //await SynthesisToSpeakerAsync(notes[pg - 1]);//阻塞直到播放完成,这是第一次播放
                            speakEdge(notes[pg - 1]);
                            isSpeak[currenSpeakIndex - 1] = true;//标记已读,避免重复阅读
                            textBox1.Text += "阅读完成\r\n";

                        }
                    }
                    //if (papp == null) break;
                    Thread.Sleep(100);
                    if (currenSpeakIndex == notes.Count)
                    {
                        //storeForm();//表示读取完毕
                        closeEdge();//关闭WEB

                        break;//线程退出标志
                    }
                }
                catch (Exception ex)
                {

                    textBox1.Text += ex.Message;
                    continue;
                }
            }
        }


        /// <summary>
        /// 界面窗体恢复原始大小和透明度
        /// </summary>
        private void restoreForm()
        {
            this.Size = new Size(704, 408);//界面恢复
            this.Opacity = 1;
        }




        private void SynthesisCompleted(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (server != null) server.Stop();
            closeEdge();
            KillProcess("wpp");
            KillProcess("POWERPNT");
            try
            {
                Application.ExitThread();
                System.Environment.Exit(0);
            }
            catch (Exception)
            {
                ;
            }
            //Application.ExitThread();
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            txtPath.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
          
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else e.Effect = DragDropEffects.None;
        }

        private void KillProcess(string processName)
        {
            //获得进程对象，以用来操作  
            System.Diagnostics.Process myproc = new System.Diagnostics.Process();
            //得到所有打开的进程   
            try
            {
                //获得需要杀死的进程名  
                foreach (Process thisproc in Process.GetProcessesByName(processName))
                {
                    //立即杀死进程  
                    thisproc.Kill();
                }
            }
            catch (Exception Exc)
            {
                //throw new Exception("", Exc);
            }
        }


        private void button5_Click_1(object sender, EventArgs e)
        {
            SplicPPT();
        }
        public void SplicPPT()
        {
            progressBar1.Visible = true;
            string name = Path.GetFileNameWithoutExtension(txtPath.Text);
            try
            {
                PowerPoint.Application ppt1 = new PowerPoint.Application();
                PowerPoint.Presentation pptFile = ppt1.Presentations.Open(txtPath.Text, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

                for (int i = 1; i <= pptFile.Slides.Count; i++)
                {
                    string spptName = Application.StartupPath + "/" + i.ToString() + ".pptx";
                    // string spptPath = Application.StartupPath+"/split/" + spptName + ".pptx";
                    //新建一个PPT
                    PowerPoint.Presentation new1 = ppt1.Presentations.Add(MsoTriState.msoFalse);
                    //把第i页插入新的PPT中
                    new1.Slides.InsertFromFile(txtPath.Text, 0, i, i);
                    //设置PPT比例16:9
                    //new1.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen16x9;
                    //保存
                    new1.SaveAs(spptName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoFalse);
                    new1.Close();
                    this.progressBar1.Value = i / pptFile.Slides.Count * 100;
                }
                pptFile.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }



   
        private const int WS_THICKFRAME = 262144;
        private const int WS_BORDER = 8388608;

        bool isSeted = false;
        private const int SW_HIDE = 0;  //隐藏任务栏
        private const int SW_RESTORE = 9;//显示任务栏
        [DllImport("user32.dll")]
        public static extern int ShowWindow(int hwnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern int FindWindow(string lpClassName, string lpWindowName);

        //设置窗口的父窗体
        [DllImport("user32.dll", SetLastError = true)]
        private static extern long SetParent(int hWndChild, IntPtr hWndNewParent); //该api用于嵌入到窗口中运行

        //获取窗口样式
        [DllImport("user32.dll")]
        public static extern int GetWindowLong(int hWnd, int nIndex);

        //设置窗口样式
        [DllImport("user32.dll")]
        public static extern int SetWindowLong(int hWnd, int nIndex, int dwNewLong);

        //设置窗口位置
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int MoveWindow(int hWnd, int x, int y, int nWidth, int nHeight, bool BRePaint);

        private void label3_Click(object sender, EventArgs e)
        {
            //if (label3.Text == "R")
            //{
            //    restoreForm();
            //    label3.Text = "M";
            //}
            //else if (label3.Text == "M")
            //{
            //    minsizeForm();//最小化
            //    label3.Text = "R";
            //}
        }

        /// <summary>
        ///最小化窗体
        /// </summary>
        private void minsizeForm()
        {
            //最小化窗体
            this.Size = new Size(80, 70);
            //this.label3.Location = new Point(0, 0);
            //this.label3.Text = "R";
            this.Opacity = 0.7;
            this.TopMost = true;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            alphaFormTransformer1.TransformForm(255);


            this.BeginInvoke(new EventHandler(delegate
            {
                Reg rg = new Reg();
                txtNum.Text = rg.GetMNum();
                if (Global.isreg)
                {
                    txtSerial.Text = "已经注册";
                    txtSerial.ReadOnly = true;
                    button3.Enabled = false;
                }
            }));
        }

        private void btnEDGE_Click(object sender, EventArgs e)
        {
            OpenFileDialog edgeFile = new OpenFileDialog();
            edgeFile.InitialDirectory= @"C:\Program Files (x86)\Microsoft\";
            edgeFile.Title = "选择EDGE浏览器位置";
            //fdlg.InitialDirectory = @"d:\";   //@是取消转义字符的意思
            edgeFile.Filter = "EDGE(*.exe)|*.exe";
            /*
             * FilterIndex 属性用于选择了何种文件类型,缺省设置为0,系统取Filter属性设置第一项
             * ,相当于FilterIndex 属性设置为1.如果你编了3个文件类型，当FilterIndex ＝2时是指第2个.
             */
            edgeFile.FilterIndex = 2;
            /*
             *如果值为false，那么下一次选择文件的初始目录是上一次你选择的那个目录，
             *不固定；如果值为true，每次打开这个对话框初始目录不随你的选择而改变，是固定的  
             */
            edgeFile.RestoreDirectory = true;
            if (edgeFile.ShowDialog() == DialogResult.OK)
            {
                txtEDGE.Text = System.IO.Path.GetFileNameWithoutExtension(edgeFile.FileName);
                txtEDGE.Text = edgeFile.FileName;
            }

        }
        void speakFirst()
        {
            //因为第一次打开的无法触发App_SlideShowBegin事件，所以在这里初始化第一页的朗读
            pg = 1;
            lblPage.Text = "当前第" + pg.ToString() + "页";//记录显示页码
            speak();
        }
        private void registAllEvents()
        {
            
            papp.SlideShowBegin += App_SlideShowBegin; ;//启动放映
            papp.SlideShowEnd += App_SlideShowEnd;//结束放映
            papp.SlideShowNextSlide += App_SlideShowNextSlide;//放映改变
            papp.SlideShowNextBuild += App_SlideShowNextBuild;//下一个动画
            papp.SlideShowOnPrevious += App_SlideShowOnPrevious;//上一个动画
        }
        int animation = 0;//某一页幻灯片的动画索引

        /// <summary>
        /// 开始播放
        /// </summary>
        /// <param name="Wn"></param>
        private void App_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            pg = 1;
            lblPage.Text = "当前第" + pg.ToString() + "页";//记录显示页码
        }
        /// <summary>
        /// 播放结束
        /// </summary>
        /// <param name="Pres"></param>
        private void App_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            pg = -1;
            lblPage.Text = "当前第" + pg.ToString() + "页";//记录显示页码
            this.timer1.Enabled = false;//
            closeEdge();//关闭WEB

        }
        /// <summary>
        /// 幻灯片切了下一页
        /// </summary>
        /// <param name="Wn"></param>
        private void App_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            pg = Wn.View.Slide.SlideIndex;
            lblPage.Text = "当前第" + pg.ToString() + "页";//记录显示页码
            animation = 0;
            lblDh.Text = animation.ToString();//显示动画索引
            //启动朗读代码
            speak();
        }

        /// <summary>
        /// 下一个动画
        /// </summary>
        /// <param name="Wn"></param>
        private void App_SlideShowNextBuild(PowerPoint.SlideShowWindow Wn)
        {
            animation++;
            lblDh.Text = animation.ToString();//显示动画索引            
            speakAnimation();//启动下一个动画的朗读代码
        }

        private void App_SlideShowOnPrevious(PowerPoint.SlideShowWindow Wn)
        {
            //上一个动画
            animation--;
        }

        private void speakAnimation()
        {
            try
            {
                //如果已读没有，如果备注没有，如果页码没有，则不朗读
                if (isSpeak.Count == 0 || notes.Count == 0 || pg <= 0)
                {
                    return;
                }
                string spk = notes[pg - 1];//当前要朗读的内容
                string[] animationNotes = null;
                if (spk.Contains('\r') || spk.Contains('\n'))
                {
                    //如果有回车，表示有动画
                    animationNotes = spk.Split(new char[] { '\r', '\n' });
                    spk = animationNotes[animation];//先朗读第一个动画备注
                    speakEdge(spk);
                    textBox1.Text += "动画备注"+animation+"阅读完成\r\n";
                    Global.end = false;//JS标记未阅读完成
                }
                
            }
            catch (Exception ex)
            {
                textBox1.Text += ex.Message;
                return;
            }
        }
        private void speak()
        {
            try
            {
                //如果已读没有，如果备注没有，如果页码没有，则不朗读
                if (isSpeak.Count == 0 || notes.Count == 0 || pg <= 0)
                {
                    return;
                }
                string spk = notes[pg - 1];//当前要朗读的内容
                string[] animationNotes = null;
                if (spk.Contains('\r')||spk.Contains('\n'))
                {
                    //如果有回车，表示有动画
                    animationNotes = spk.Split(new char[] { '\r', '\n' });
                    spk=animationNotes[0];//先朗读第一段
                }
                if (isSpeak[pg - 1] == false)//表示没有读过
                {
                    currenSpeakIndex = pg;//记录当前正在阅读的索引
                    speakEdge(spk);
                    isSpeak[currenSpeakIndex - 1] = true;//标记已读,避免重复阅读
                    textBox1.Text += "阅读完成\r\n";
                    Global.end = false;//JS标记未阅读完成
                }
            }
            catch (Exception ex)
            {
                textBox1.Text += ex.Message;
                return;
            }

        }
        /// <summary>
        /// PPT下一页。
        /// </summary>
        public void NextSlide()
        {
            if (this.papp != null)
                this.ppr.SlideShowWindow.View.Next();
        }
        /// <summary>
        /// PPT上一页。
        /// </summary>
        public void PreviousSlide()
        {
            if (this.papp != null)
                this.ppr.SlideShowWindow.View.Previous();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //如果阅读完成，则下一页
            if (Global.end)   NextSlide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveConfig(comboBox1.Text,txtEDGE.Text,chkAuto.Checked.ToString());
            Global.spkName = comboBox1.Text;
            Global.edgePath = txtEDGE.Text;
            Global.autoPlay = chkAuto.Checked;
        }
        private void SaveConfig(string spkName ,string edgePath,string autoPlay)
        {
            using (StreamWriter fs = new StreamWriter("config.ini", false))
            {
                fs.WriteLine(spkName);
                fs.WriteLine(edgePath);
                fs.WriteLine(autoPlay);
                fs.Close();
            }
        }

        private void label3_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void alphaFormMarker1_Load(object sender, EventArgs e)
        {
           // alphaFormTransformer1.TransformForm(255);
        }

        public new virtual bool Visible
        {
            get
            {
                return base.Visible;
            }

            set
            {
                // This will set this form's Visible property as well
                // as the attached alpha (layered) window which hosts the window
                // frame.
                alphaFormTransformer1.SetAlphaAndParentFormVisible(value);
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            alphaFormTransformer1.Fade(FadeType.FadeOut, true, false, 100);
            Form own = Owner;
            if (Owner != null)
            {
                Owner = null;
                own.Close();
            }
            base.OnClosing(e);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Owner != null)
                (Owner as FrmPP).DrawerButton_Click(this, EventArgs.Empty);
        }
        public bool isReg = false;   //是否注册过
        private void button3_Click_1(object sender, EventArgs e)
        {

            //string path1 =  //Assembly.GetExecutingAssembly().Location;
            Process pro = new Process();
            Reg sr = new Reg();

            if (txtSerial.Text == sr.GetRNum())
            {
                MessageBox.Show("注册成功！重启软件后生效！", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //string filePath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;    //获取到bin目录的下层路径：bin\Debug\

                DirectoryInfo info = new DirectoryInfo(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase);
                string filePath = info.Parent.FullName;

                FileStream fs = new FileStream(filePath + "Reg.ini", FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fs);
                //开始写入
                sw.Write(sr.GetRNum());
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();

                this.isReg = true;
                //System.Environment.Exit(0);
                this.Close();
            }
            else
            {
                this.isReg = false;
                MessageBox.Show("注册码错误！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
               
                txtSerial.Focus();
                //System.Environment.Exit(0);
            }
        }
     
    }
}
