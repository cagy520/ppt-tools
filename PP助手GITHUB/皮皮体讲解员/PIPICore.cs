using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;

namespace 皮皮助手
{
    public class PIPICore
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
        public void Start()
        {

           

            speakEdge("");//初始化文件
            //启动朗读的WEBSERVER
            Thread webTh = new Thread(startWeb);
            webTh.IsBackground = true;
            webTh.Start();

        }

        PowerPoint.Slide sl = null;

        public void readNotes(string fullname)
        {

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
            XLog.LogPPT("一共发现PPT备注：" + notes.Count.ToString());
            XLog.LogPPT(doc);
        }


        public void LoadConfig()
        {
            using (StreamReader fs = new StreamReader("config.ini"))
            {
                Global.spkName = fs.ReadLine();
                Global.edgePath = fs.ReadLine();
                Global.autoPlay =Convert.ToBoolean(fs.ReadLine());
                fs.Close();
            }
        }

        public  void speakEdge(string text)
        {
            StreamWriter wt = new StreamWriter("readContent.txt");
            text = Global.spkName + "@*@" + text;
            wt.Write(text);
            wt.Close();
        }

     

        public void start()
        {
            try
            {
                init();
                //minsizeForm();//最小化窗体

                //注册绑定幻灯片的事件
                registAllEvents();
                speakFirst();//朗读第一页
                //thPage = new Thread(new ThreadStart(getPage));
                //thPage.IsBackground = true;
                //thPage.Start();
                //MessageBox.Show("1");

            //    tabControl1.SelectedIndex = 1;

                //thSpeak = new Thread(new ThreadStart(speaker));
                //thSpeak.IsBackground = true;
                //thSpeak.Start();//开始讲话线程

                //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "--start-maximized --app=http://127.0.0.1:5050/spk.html");//pplication\msedge.exe" --start-maximized --app=http://127.0.0.1:4050/spk.html
                Process.Start(Global.edgePath, " --app=http://127.0.0.1:5050/spk.html");//pplication\msedge.exe" --start-maximized --app=http://127.0.0.1:4050/spk.html

                //incWeb = new Thread(new ThreadStart(includeWebForm));
                //incWeb.IsBackground = true;
                //incWeb.Start();//开始讲话线程


                isClose = false;
                //pg = 0;

                //if (chkAuto.Checked) this.timer1.Enabled = true;//自动播放
            }
            catch (Exception)
            {
                ;
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

                        pg = Convert.ToInt32(page);

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
                        }
                        isClose = true;//表示本异常只执行一次,如果开起新的朗读，则标记为false;
                    }
                    Thread.Sleep(750);
                    continue;//PPT没打开的时候才会触发
                }
            }
        }

        private void speaker()
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
                            speakEdge(notes[pg - 1]);
                            isSpeak[currenSpeakIndex - 1] = true;//标记已读,避免重复阅读

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

                    //textBox1.Text += ex.Message;
                    continue;
                }
            }
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


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        void speakFirst()
        {
            //因为第一次打开的无法触发App_SlideShowBegin事件，所以在这里初始化第一页的朗读
            pg = 1;
          
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

        }
        /// <summary>
        /// 播放结束
        /// </summary>
        /// <param name="Pres"></param>
        private void App_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            pg = -1;

            closeEdge();//关闭WEB

        }
        /// <summary>
        /// 幻灯片切了下一页
        /// </summary>
        /// <param name="Wn"></param>
        private void App_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            pg = Wn.View.Slide.SlideIndex;

            animation = 0;

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

                    Global.end = false;//JS标记未阅读完成
                }

            }
            catch (Exception ex)
            {

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
                if (spk.Contains('\r') || spk.Contains('\n'))
                {
                    //如果有回车，表示有动画
                    animationNotes = spk.Split(new char[] { '\r', '\n' });
                    spk = animationNotes[0];//先朗读第一段
                }
                if (isSpeak[pg - 1] == false)//表示没有读过
                {
                    currenSpeakIndex = pg;//记录当前正在阅读的索引
                    speakEdge(spk);
                    isSpeak[currenSpeakIndex - 1] = true;//标记已读,避免重复阅读

                    Global.end = false;//JS标记未阅读完成
                }
            }
            catch (Exception ex)
            {

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
            if (Global.end) NextSlide();
        }

    }
}
