using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AlphaForm;
using System.Speech.Synthesis;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace 皮皮助手
{
    public partial class FrmPP : Form
    {


        public Form1 m_drawer;
        Point m_drawerLocOld;
        Point m_mainLocOld;
        bool m_drawerOpen;
        bool m_lockLoc;


        public FrmPP()
        {
            InitializeComponent();

            this.Location = new Point(20, 0);//设置窗体位置System.Windows.Forms.SystemInformation.WorkingArea.Height - 300
            Control.CheckForIllegalCrossThreadCalls = false;

            this.BackgroundImage = Image.FromFile(Application.StartupPath + "/a.png");
            this.alphaFormTransformer1.BackgroundImage = Image.FromFile(Application.StartupPath + "/b.png");


            m_drawer = new Form1();
            m_drawer.ShowInTaskbar = false;

            m_drawerOpen = false;
            m_lockLoc = false;

            LocationChanged += new EventHandler(MainFormLocationChanged);
            m_drawer.LocationChanged += new EventHandler(DrawerLocationChanged);

            //systemSpk("皮皮来了");
            speakEdge("");//初始化文件
            //启动朗读的WEBSERVER
            Thread webTh = new Thread(startWeb);
            webTh.IsBackground = true;
            webTh.Start();

            System.Media.SoundPlayer player = new System.Media.SoundPlayer("hello.wav");
            player.PlaySync();//另起线程播放
            timer2.Enabled = true;
        }

        protected override void OnShown(EventArgs e)
        {
            alphaFormTransformer1.Fade(FadeType.FadeIn, false, false, 500);
            base.OnShown(e);
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // We want to close both the drawer window and the main form if either
            // receives a close event. So we use the Owner property as a flag to 
            // manage this. Also see OnClosing in SlideOut.cs.
            if (m_drawer.Owner != null)
            {
                m_drawer.Owner = null;
                m_drawer.Close();
            }
            alphaFormTransformer1.Fade(FadeType.FadeOut, true, false, 300);
            base.OnClosing(e);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // It's going to fade in, so we must set opacity to 0.
            alphaFormTransformer1.TransformForm(0);
           // IsReg();

            //ttMsg.SetToolTip(pictureBox1, "点击圆点打开设置，或者将PPT拖到PP身体上。");//ttMsg为ToolTip控件,txtLoginName为文本框
           // ttMsg.Show("该登录名已存在", txtLoginName);
        }
        bool isReg = false;//判断是否注册
        private void IsReg()
        {
            DirectoryInfo info = new DirectoryInfo(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase);
            String path = info.Parent.FullName;
            //string filePath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;    //获取到bin目录的下层路径：bin\Debug\
            if (File.Exists(path + "Reg.ini"))
            {
                string regC = "";
                FileInfo fi = new FileInfo(path + "Reg.ini");
                StreamReader srd = fi.OpenText();
                regC = srd.ReadToEnd();
                srd.Close();
                isReg = regC == new Reg().GetRNum().Trim();//判断是否注册

                Global.isreg = isReg;
            }
            if (!isReg)//提示用户注册软件
            {
                MessageBox.Show("软件未注册！只能阅读5页PPT");

                //DialogResult result = MessageBox.Show("您还未注册，是否需要注册", "信息", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                //if (result == DialogResult.Yes)
                //{
                //    fr.ShowDialog();
                //}
                //else
                //{
                //    System.Environment.Exit(0);
                //}
            }
        }

        public void DrawerButton_Click(object sender, EventArgs e)
        {
            //Form1 fm = new Form1();
            //fm.Show();
            //return;

            int ledge;

            // See docs in AlphaForm on this class
            Impulse imp = new Impulse(7.0);

            int frames;
            int duration;

            bool SaveTop = TopMost;

            // Suppress location changed event handlers for drawer and main form
            m_lockLoc = true;

            if (m_drawerOpen)
            {
                duration = 400;

                // owner must be null to guarantee drawer window will appear under
                // the main form's layered window.
                m_drawer.Owner = null;

                // Force main form to topmost so that drawer is underneath
                TopMost = true;

                ledge = m_drawer.Left;

                int op = 0;
                int uop = 0;

                // 30-40 fps is a reasonable median starting value (it will
                // be adjusted, on-the-fly, up or down to give 
                // the greatest frame rate for the requested duration)
                // The only risk to setting an initially high fps is on slower
                // graphics displays where this routine will make a larger
                // reduction in frame rate to meet the desired duration. Such
                // on-the-fly frame rate adjustments may be perceived as jitteryness.
                // The inverse is also true - setting it initially too low will force
                // a fast graphics system to make large increases to the frame rate.
                // (Convergence in either case is very fast though :-))

                frames = Math.Max((40 * duration) / 1000, 1);
                int inc = Math.Max(m_drawer.Width / frames, 1);
                int frameDur = duration / frames;

                // 0.92 figured by trial & error - because the frame of the drawer
                // window image includes some empty space, and I want it to slide in/out
                // right to the edge. If I wasn't so lazy, I'd crop the image in 
                // Photoshop instead of doing this :-) However this illustrates an
                // important point. If you need to rely on some fraction of the drawer
                // window width or height as an animation parameter, don't use an
                // absolute pixel value as that won't scale for different
                // DPI screens.
                int slideW = (int)(0.92 * m_drawer.Width);

                while (uop != slideW)
                {
                    uop = (int)Math.Round(imp.Evaluate(op / (double)slideW) * (double)slideW);

                    if (uop > slideW)
                        uop = slideW;
                    DateTime sf = DateTime.Now;
                    m_drawer.SetDesktopLocation(ledge - uop, m_drawer.Location.Y);
                    TimeSpan ts = DateTime.Now - sf;
                    int wait = frameDur - ts.Milliseconds;

                    m_drawer.Update(); // force contents to update

                    // We either wait or speed up to maintain requested duration
                    if (wait > 0)
                    {
                        System.Threading.Thread.Sleep(wait);
                        if (wait <= frameDur / 2 && inc > 1)
                        {
                            inc = Math.Max(1, inc / 2);
                            frameDur /= 2;
                        }
                    }
                    else if (Math.Abs(wait) >= frameDur)
                    {
                        inc *= 2;
                        frameDur *= 2;
                    }

                    op += inc;
                }

                m_drawerOpen = false;

                // Hide the drawer window (it's closed - behind the main form now).
                // Note: We need to hide the composite window which includes the layered window.
                // The Visible property hides the base class property in SlideOut, and a utility function 
                // is called to set the property for both the form and layered window (see
                // SlideOut.cs).
                m_drawer.Visible = false;

                //DrawerButton.Text = "Open Drawer ->";

            }
            else
            {
                // Setting TopMost on this form won't guarantee the drawer window will be shown 
                // underneath, and we want it to appear behind the main form when made visible.
                // An easy way to achieve this is to just show the drawer window off-screen first, display, then 
                // set TopMost on the main form to ensure the drawer is behind.
                m_drawer.Location = new Point(-1000000, -1000000);

                // Note: We need to show the composite window which includes the layered window.
                // The Visible property hides the base class property in SlideOut, and a utility function 
                // is called to set the property for both the form and layered window (see
                // SlideOut.cs).
                m_drawer.Visible = true;

                TopMost = true;

                // The drawer window is behind (z-order) and off-screen now, so move it into position
                // for animating.
                m_drawer.Location = new Point(Location.X + Width - m_drawer.Width, Location.Y + (Height - m_drawer.Height) / 2);

                ledge = m_drawer.Left;
                duration = 600; // arbitrary

                int op = 0;
                int uop = 0;

                // 30-40 fps is a reasonable median starting value (it will
                // be adjusted, on-the-fly, up or down to give 
                // the greatest frame rate for the requested duration)
                // The only risk to setting an initially high fps is on slower
                // graphics displays where this routine will make a larger
                // reduction in frame rate to meet the desired duration. Such
                // on-the-fly frame rate adjustments may be perceived as jitteryness.
                // The inverse is also true - setting it initially too low will force
                // a fast graphics system to make large increases to the frame rate.
                // (Convergence in either case is very fast though :-))

                frames = Math.Max((40 * duration) / 1000, 1);
                int inc = Math.Max(m_drawer.Width / frames, 1);
                int frameDur = duration / frames;

                // See note above on where this 0.92 comes from.
                int slideW = (int)(0.92 * m_drawer.Width);

                while (uop != slideW)
                {
                    uop = (int)Math.Round(imp.Evaluate(op / (double)slideW) * (double)slideW);

                    if (uop > slideW)
                        uop = slideW;
                    DateTime sf = DateTime.Now;
                    m_drawer.SetDesktopLocation(ledge + uop, m_drawer.Location.Y);
                    TimeSpan ts = DateTime.Now - sf;
                    int wait = frameDur - ts.Milliseconds;

                    m_drawer.Update(); // force contents to update

                    // We either wait or speed up to maintain requested duration
                    if (wait > 0)
                    {
                        System.Threading.Thread.Sleep(wait);

                        if (wait <= frameDur / 2 && inc > 1)
                        {
                            inc = Math.Max(1, inc / 2);
                            frameDur /= 2;
                        }
                    }
                    else if (Math.Abs(wait) >= frameDur)
                    {
                        inc *= 2;
                        frameDur *= 2;
                    }

                    op += inc;
                }

                // Once drawer is displayed, we set ownership to main form so that
                // the two are in sync with respect to activation and close events.
                m_drawer.Owner = this;
                m_drawerLocOld = m_drawer.Location;
                m_drawerOpen = true;
                //DrawerButton.Text = "<-- Close Drawer";
            }

            TopMost = SaveTop;
            m_lockLoc = false;
        }

        // These location change event handlers enable synchroneous dragging of
        // the main form and the drawer window
        void DrawerLocationChanged(Object sender, EventArgs e)
        {
            bool saveLock = m_lockLoc;
            if (!m_lockLoc)
            {
                m_lockLoc = true;
                SetDesktopLocation(Location.X + m_drawer.Location.X - m_drawerLocOld.X, Location.Y + m_drawer.Location.Y - m_drawerLocOld.Y);
                m_mainLocOld = Location;
            }
            m_drawerLocOld = m_drawer.Location;
            m_lockLoc = saveLock;
        }

        void MainFormLocationChanged(Object sender, EventArgs e)
        {
            bool saveLock = m_lockLoc;
            if (!m_lockLoc)
            {
                m_lockLoc = true;
                m_drawer.SetDesktopLocation(m_drawer.Location.X + Location.X - m_mainLocOld.X, m_drawer.Location.Y + Location.Y - m_mainLocOld.Y);
                m_drawerLocOld = m_drawer.Location;
            }
            m_mainLocOld = Location;
            m_lockLoc = saveLock;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmPP_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else e.Effect = DragDropEffects.None;
        }

        /// <summary>
        /// 拖动后开始的地方
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FrmPP_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            if (!path.Contains(".ppt"))
            {
                systemSpk("文件不对！");
                return;
            }
            startPPT(path);

        }


        private void systemSpk(string txt)
        {
            SpeechSynthesizer speech = new SpeechSynthesizer();
            speech.Speak(txt);
        }

        List<bool> isSpeak = new List<bool>();//用于记录每一页的Note是否被阅读过。
        int currenSpeakIndex = -1;
        PowerPoint.Presentation ppr = null;
        PowerPoint.Application papp = null;
        List<string> notes = new List<string>();
        bool AudioCompleted = false;
        private Thread thPage;
        private int pg = 0;//全局PPT当前页码
        bool exitThread = false;
        Thread thSpeak = null;
        Thread thStopSpeak = null;
        Thread incWeb = null;
        ExampleServer server = null;
        PowerPoint.Slide sl = null;

        private void init()
        {

            currenSpeakIndex = -1;
            isSpeak = new List<bool>();
            notes = new List<string>();
            papp = null;//幻灯片对象
            //if (thPage!=null) thPage.Abort();
            if (thSpeak != null) thSpeak.Abort();
            if (incWeb != null) incWeb.Abort();
            LoadConfig();//加载全局配置
        }
        public void Start()
        {
            speakEdge("");//初始化文件
            //启动朗读的WEBSERVER
            Thread webTh = new Thread(startWeb);
            webTh.IsBackground = true;
            webTh.Start();

        }


        /// <summary>
        /// 读取PPT备注信息
        /// </summary>
        /// <param name="fullname"></param>
        private void readNotes(string fullname)
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
                Global.autoPlay = Convert.ToBoolean(fs.ReadLine());
                fs.Close();
            }
        }

        public void speakEdge(string text)
        {
            try
            {
                if (!isReg&&pg>=5)
                {
                    StreamWriter wt1 = new StreamWriter("readContent.txt");
                    text = Global.spkName + "@*@" + "软件没有注册，只能帮您阅读到第五页了。";
                    wt1.Write(text);
                    wt1.Close();
                }
                else
                {
                    StreamWriter wt = new StreamWriter("readContent.txt");
                    text = Global.spkName + "@*@" + text;
                    wt.Write(text);
                    wt.Close();
                }
            }
            catch (Exception ex)
            {
                XLog.LogSystem(ex.Message);
            }
        }


        void startPPT(string ppt)
        {
            try
            {
                init();
                //minsizeForm();//最小化窗体
                readNotes(ppt);
                //注册绑定幻灯片的事件
                registAllEvents();

                speakFirst();//朗读第一页

                //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "--start-maximized --app=http://127.0.0.1:5050/spk.html");//pplication\msedge.exe" --start-maximized --app=http://127.0.0.1:4050/spk.html
                Process.Start(Global.edgePath, " --app=http://127.0.0.1:5050/spk.html");//pplication\msedge.exe" --start-maximized --app=http://127.0.0.1:4050/spk.html

                incWeb = new Thread(new ThreadStart(minEdge));
                incWeb.IsBackground = true;
                incWeb.Start();//最小化窗体进程

                if (Global.autoPlay)
                {
                    XLog.LogSystem("自动播放代码");
                    this.timer1.Enabled = true;//自动播放
                }

            }
            catch (Exception ex)
            {
                XLog.LogPPT(ex.Message);
            }

        }

        private void startWeb()
        {
            try
            {
                if (server == null)
                {
                    server = new ExampleServer("127.0.0.1", 5050);
                    server.SetRoot("./");
                    server.Logger = new ConsoleLogger();
                    server.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("端口5050异常:" + ex.Message);
            }

        }


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
        private void minEdge()
        {
            try
            {
                for (int i = 0; i < 200; i++)
                {
                    Thread.Sleep(100);
                    int spkForm = FindWindow(null, "speak");
                    if (spkForm == 0) continue;//循环200次一直找到窗体为止
                    ShowWindow(spkForm, 2);
                    break;
                }
            }
            catch (Exception ex)
            {
                XLog.LogSystem("最小化EDGE窗体异常:" + ex.Message);
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
            timer1.Enabled = false;
            Global.end = true;
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
        /// <summary>
        /// 朗读当前页面的第一段
        /// </summary>
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
                    // 如果回车比较多，则没问题，如果回车数量少于动画数量

                    spk = animationNotes[animation];//先朗读第一个动画备注
                    speakEdge(spk);

                    Global.end = false;//JS标记未阅读完成
                }

            }
            catch (Exception ex)
            {
                speakEdge("备注里面的回车数量少于动画的数量！");
                XLog.LogSpk("备注里面的换行数量少于动画的数量！" + ex.Message);
            }
        }
        /// <summary>
        /// 朗读当前页面的第一段
        /// </summary>
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
            try
            {
                if (Global.end) NextSlide(); //如果阅读完成，则下一页
            }
            catch (Exception ex)
            {
                XLog.LogPPT("预防PPT退出，还在点下一页：" + ex.Message);

            }
        }

        private void 男声ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveConfig("Microsoft Yunyang Online (Natural) - Chinese (Mainland)",Global.edgePath, Global.autoPlay.ToString());
            LoadConfig();
        }


        private void saveConfig(string spkName, string edgePath, string autoPlay)
        {
            using (StreamWriter fs = new StreamWriter("config.ini", false))
            {
                fs.WriteLine(spkName);
                fs.WriteLine(edgePath);
                fs.WriteLine(autoPlay);
                fs.Close();
            }
        }

        private void 女声xiaoxiaoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveConfig("Microsoft Xiaoxiao Online (Natural) - Chinese (Mainland)", Global.edgePath, Global.autoPlay.ToString());
            LoadConfig();
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (server != null) server.Stop();
            closeEdge();
            //KillProcess("wpp");
            //KillProcess("POWERPNT");
            try
            {
                Application.ExitThread();
                System.Environment.Exit(0);
            }
            catch (Exception)
            {
                ;
            }
        }

        private void FrmPP_MouseEnter(object sender, EventArgs e)
        {
            //this.BackgroundImage = 皮皮助手.Properties.Resources.a;
            //Bitmap bmap = 皮皮助手.Properties.Resources.b;
            //alphaFormTransformer1.UpdateSkin(bmap, null, 255);

        }

        private void FrmPP_MouseLeave(object sender, EventArgs e)
        {
            //this.BackgroundImage = 皮皮助手.Properties.Resources._1;
            //Bitmap bmap = 皮皮助手.Properties.Resources._2;
            //alphaFormTransformer1.UpdateSkin(bmap, null, 255);
        }

        private void alphaFormTransformer1_MouseEnter(object sender, EventArgs e)
        {
            //this.BackgroundImage = 皮皮助手.Properties.Resources.a;



        }

        private void alphaFormTransformer1_MouseLeave(object sender, EventArgs e)
        {

            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            //this.BackgroundImage = Image.FromFile(Application.StartupPath + @"\img\1.png");
            //this.alphaFormTransformer1.BackgroundImage = Image.FromFile(Application.StartupPath + @"\img\2.png");
            //Bitmap bmap = 皮皮助手.Properties.Resources._2;
            //alphaFormTransformer1.UpdateSkin(bmap, null, 255);
            //this.BackgroundImage = Image.FromFile(Application.StartupPath + @"\img\1.png");
        }

        private void ppMove()
        {
      
            for (int i = 0; i < System.Windows.Forms.SystemInformation.WorkingArea.Height-300; i++)
            {
                this.Invoke((EventHandler)delegate
                {
                    this.Location = new Point(20, i);//设置窗体位置
                    //Thread.Sleep(10);
                });

             
            }
           
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            ppMove();
            timer2.Enabled = false;
        }
    }
}
