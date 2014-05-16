using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        //declare all global class variables
        public string pAppName = "";
        public string pWebUrlNotice = "";
        public string pWebUrlSuccess = "";
        public string pWebUrlSuccessReboot = "";
        public string pWebUrlFailed = "";
        public Int32 postPoneTime = 0;
        public string doReboot = "";
        public Int32 rebootTime = 0;
        public Int32 closeTime = 0;
        public Int32 installCounter = 0;
        public Int32 installTimeOut = 6000;
        public Int32 count = 0;
        public Int32 programExitCode;
        public string debugMode = "";
        public Int32 ptaskTrayReminder1 = 0;
        public Int32 ptaskTrayReminder2 = 0;
        public Int32 ptaskTrayReminder3 = 0;
        public Int32 ptaskTrayReminder4 = 0;
        public Boolean appControl;
        public string fileName = "settings.cfg";
        public char deLim = '=';
        public string postPoneText = "Postpone Disabled";
        public string pCompanyName = "";
        public string pFileName = "";
        public string installFinished;
        public string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase);
        public Int32 timerIntervalSeconds = 60000;
        delegate void SetTextCallback(string text);
        private Process myProcess = new Process();
        System.Windows.Forms.Timer timer01 = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer timer02 = new System.Windows.Forms.Timer();


        //start reboot imports and declarations
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        internal struct TokPriv1Luid
        {
            public int Count;
            public long Luid;
            public int Attr;
        }

        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern IntPtr GetCurrentProcess();

        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr
        phtok);

        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern bool LookupPrivilegeValue(string host, string name,
        ref long pluid);

        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall,
        ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);

        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool ExitWindowsEx(int flg, int rea);

        internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
        internal const int TOKEN_QUERY = 0x00000008;
        internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;
        internal const string SE_SHUTDOWN_NAME = "SeShutdownPrivilege";
        internal const int EWX_LOGOFF = 0x00000000;
        internal const int EWX_SHUTDOWN = 0x00000001;
        internal const int EWX_REBOOT = 0x00000002;
        internal const int EWX_FORCE = 0x00000004;
        internal const int EWX_POWEROFF = 0x00000008;
        internal const int EWX_FORCEIFHUNG = 0x00000010;
        //end reboot imports and declarations

        public Form1()
        {       
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //need to associate tasktray icon with the right click context menu
            mynotifyIcon1.ContextMenuStrip = contextMenuStrip1;
            //alright lets check for config file and then read some data
            LoadConfigFile(fileName, deLim);
            //determine if we are showing the control menu
            this.ControlBox = appControl;
            //Display debug info if debug is set to on in config file
            DebugStatus(debugMode);
            //change the title bar to match the company
            this.Text = pCompanyName + " Software Deployment";
            //change the webpage to what we found in the config file
            SetUrls("Notice");
            //check the postpone time and change the button accordingly
            if (postPoneTime >= 1)
            {
                this.SetText(Convert.ToString(postPoneTime));
                timer01.Interval = timerIntervalSeconds;
                timer01.Tick += new EventHandler(timerprocpostpone);
                timer01.Enabled = true;
            }
            else
            {
                if (InstallButton.Text == "Restart Computer")
                {
                    this.SetText(Convert.ToString(rebootTime));
                    timer02.Interval = timerIntervalSeconds;
                    timer02.Tick += new EventHandler(timerprocreboot);
                    timer02.Enabled = true;
                }
            }
        }
        //end form load

        //timer for postponement and UI update
        public void timerprocpostpone(object o1, EventArgs e1)
        {
            if (postPoneTime != 0)
            {
                if (InstallButton.Text == "Install Now")
                {
                    postPoneTime--;
                    this.SetText(Convert.ToString(postPoneTime));
                }
            }
            else
            {
                InstallButton.Enabled = false;
                ((System.Windows.Forms.Timer)o1).Tick -= new EventHandler(timerprocpostpone);
                ((System.Windows.Forms.Timer)o1).Stop();
                DoInstall(pFileName);
            }
        }
        
        //timer for reboot and UI update
        public void timerprocreboot(object o1, EventArgs e1)
        {
            if (rebootTime > 1)
            {
                rebootTime--;
                this.SetText(Convert.ToString(rebootTime));
            }
            else
            {
                ((System.Windows.Forms.Timer)o1).Stop();
                ((System.Windows.Forms.Timer)o1).Tick -= new EventHandler(timerprocreboot);
                DoExitWin(EWX_REBOOT);
                Environment.ExitCode = programExitCode;
                Application.Exit();
            }
        }

        public void timerprocClose(object o1, EventArgs e1)
        {
            if (closeTime > 1)
            {
                closeTime--;
                this.SetText(Convert.ToString(closeTime));
            }
            else
            {
                ((System.Windows.Forms.Timer)o1).Stop();
                ((System.Windows.Forms.Timer)o1).Tick -= new EventHandler(timerprocClose);
                Environment.ExitCode = programExitCode;
                Application.Exit();
            }
        }

        //execute the postpone 
        private void postPoneButton_Click(object sender, EventArgs e)
        {
            //set the window to minimized and display task tray icon
            this.WindowState = FormWindowState.Minimized;
            if (FormWindowState.Minimized == this.WindowState)
            {
                this.Hide();
                //add icon to the task tray and hide the window from the taskbar
                mynotifyIcon1.Visible = true;
                if (InstallButton.Text == "Restart Computer")
                {
                    this.SetText(Convert.ToString(rebootTime));
                }
                else
                {
                    this.SetText(Convert.ToString(postPoneTime));
                }
                mynotifyIcon1.ShowBalloonTip(15000);
            }
            else if (FormWindowState.Normal == this.WindowState)
            {
                mynotifyIcon1.Visible = false;
            }
        }
        //end postpone click
        
        //re-open the program from the task tray
        private void mynotifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            mynotifyIcon1.Visible = false;
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();
        }
        //end mynotifyicon1

        private void Install_Click(object sender, EventArgs e)
        {
            if (InstallButton.Enabled == true)
            {
                switch (InstallButton.Text)
                {
                    case "Install Now":
                        {
                           timer01.Enabled = true;
                           DoInstall(pFileName);
                           break;
                        }
                    case "Restart Computer":
                        {
                            //restart the machine
                            DialogResult result = MessageBox.Show("Are you sure you want to restart now?", "Restart Computer?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                            if (result == DialogResult.Yes)
                            {
                                DoExitWin(EWX_REBOOT);
                                Environment.ExitCode = programExitCode;
                                Application.Exit();
                                break;
                            }
                            break;
                        }

                    case "Close Program":
                        {
                            //exit the program with the error code returned from the install script
                            postPoneButton.Enabled = false;
                            Environment.ExitCode = programExitCode;
                            Application.Exit();
                            break;
                        }
                }
            }
        }
        //end install click




        //safely set text fields with new text data
        public void SetText(string text)
        {
            
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.lCountDownTimer.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                //MessageBox.Show("should be changing ui to " + text);
                //set timer label and postpone button
                
                if (InstallButton.Text == "Install Now")
                {
                    this.lCountDownTimer.Text = text + " Minutes until Mandatory Install";
                    this.postPoneButton.Text = "Postpone " + text + " Minutes";
                    mynotifyIcon1.BalloonTipTitle = pCompanyName + " pending software deployment.";
                    mynotifyIcon1.BalloonTipText = pAppName + " will install automatically in " + text + " minutes.";
                }
                if (InstallButton.Text == "Close Program")
                {
                    this.lCountDownTimer.Text = "This window will close in " + text + " minutes";
                    mynotifyIcon1.BalloonTipTitle = pCompanyName + " has finished the Install.";
                    mynotifyIcon1.BalloonTipText = pAppName + " is now installed for your use.";
                }
                if (InstallButton.Text == "Restart Computer")
                {
                    this.lCountDownTimer.Text = text + " Minutes until Mandatory Restart";
                    this.postPoneButton.Text = "Postpone " + text + " Minutes";
                    mynotifyIcon1.BalloonTipTitle = pCompanyName + " pending Restart.";
                    mynotifyIcon1.BalloonTipText = pAppName + " will Restart your Computer automatically in " + text + " minutes.";
                }
                //flash window if countdown hits any of our reminders 
                if (Convert.ToInt32(text) == ptaskTrayReminder1)
                {
                    this.lCountDownTimer.BackColor = System.Drawing.Color.GreenYellow;
                    FlashWindow.Flash(this, 5);
                    mynotifyIcon1.ShowBalloonTip(15000);
                }
                if (Convert.ToInt32(text) == ptaskTrayReminder2)
                {
                    this.lCountDownTimer.BackColor = System.Drawing.Color.Yellow;
                    FlashWindow.Flash(this, 5);
                    mynotifyIcon1.ShowBalloonTip(15000);
                }
                if (Convert.ToInt32(text) == ptaskTrayReminder3)
                {
                    this.lCountDownTimer.BackColor = System.Drawing.Color.Red;
                    FlashWindow.Flash(this, 15);
                    mynotifyIcon1.ShowBalloonTip(30000);
                }
                if (Convert.ToInt32(text) == ptaskTrayReminder4)
                {
                    this.lCountDownTimer.BackColor = System.Drawing.Color.Red;
                    mynotifyIcon1.ShowBalloonTip(30000);
                }
            }
        }

        //method to restart computer
        private void DoExitWin(int flg)
        {
            bool ok;
            TokPriv1Luid tp;
            IntPtr hproc = GetCurrentProcess();
            IntPtr htok = IntPtr.Zero;
            ok = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
            tp.Count = 1;
            tp.Luid = 0;
            tp.Attr = SE_PRIVILEGE_ENABLED;
            ok = LookupPrivilegeValue(null, SE_SHUTDOWN_NAME, ref tp.Luid);
            ok = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            ok = ExitWindowsEx(flg, 0);
        }

        public void DoInstall(object value)
        {
            //add icon to the task tray and hide the window from the taskbar
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
            notifyIcon2.BalloonTipTitle = pCompanyName + " Is Currently Installing Software.";
            notifyIcon2.BalloonTipText = pAppName + " is currently being Installed. Please Wait.";
            notifyIcon2.Visible = true;
            notifyIcon2.ShowBalloonTip(1000);
            //Start the process
            ProcessStartInfo _processStartInfo = new ProcessStartInfo();
            _processStartInfo.WorkingDirectory = Environment.CurrentDirectory;
            _processStartInfo.FileName = value.ToString();
            myProcess.EnableRaisingEvents = true;
            myProcess = Process.Start(_processStartInfo);
            myProcess.WaitForExit();
            programExitCode = myProcess.ExitCode;
            //Update UI with results
            RetrieveResult(programExitCode);
            //Done with the Install, show the UI again
            notifyIcon2.Visible = false;
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();
        }
        
        //safely set text fields with new text data
        public void DoInstallText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.lCountDownTimer.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(DoInstallText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                lCountDownTimer.Text = "Performing Installation Now";
                postPoneButton.Enabled = false;
                mynotifyIcon1.Visible = false;
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.Activate();
                Thread.Sleep(100);
            }
        }


        private void RetrieveResult(int programExitCode)
        {
            switch (programExitCode)
            {
                case 0:
                {
                    SetUrls("Success");
                    InstallButton.Enabled = true;
                    InstallButton.Text = "Close Program";
                    postPoneButton.Enabled = false;
                    this.lCountDownTimer.BackColor = System.Drawing.Color.Empty;
                    lCountDownTimer.Text = "Installation Succeeded";
                    System.Windows.Forms.Timer timer01 = new System.Windows.Forms.Timer();
                    timer01.Interval = timerIntervalSeconds;
                    timer01.Tick += new EventHandler(timerprocClose);
                    timer01.Enabled = true;
                    break;
                }
                case 1:
                {
                    SetUrls("Failed");
                    InstallButton.Enabled = true;
                    InstallButton.Text = "Close Program";
                    lCountDownTimer.Text = "Installation Failed";
                    postPoneButton.Enabled = false;
                    System.Windows.Forms.Timer timer01 = new System.Windows.Forms.Timer();
                    timer01.Interval = timerIntervalSeconds;
                    timer01.Tick += new EventHandler(timerprocClose);
                    timer01.Enabled = true;
                    break;
                }
                case 1310:
                {
                    SetUrls("SuccessReboot");
                    this.SetText(Convert.ToString(rebootTime));
                    InstallButton.Enabled = true;
                    InstallButton.Text = "Restart Computer";
                    postPoneButton.Enabled = true;
                    System.Windows.Forms.Timer timer01 = new System.Windows.Forms.Timer();
                    timer01.Interval = timerIntervalSeconds;
                    timer01.Tick += new EventHandler(timerprocreboot);
                    timer01.Enabled = true;
                    break;
                }
                default:
                {
                    SetUrls("Success");
                    InstallButton.Enabled = true;
                    InstallButton.Text = "Close Program";
                    break;
                }
            }
        }

        public void SetUrls(string setStatus)
        {
            //read status page and update UI according to url type
            switch (setStatus)
            {
                case "Notice":
                    {
                        if (pWebUrlNotice.StartsWith("http://"))
                        {
                            webBrowser1.Navigate(pWebUrlNotice);
                        }
                        else
                        {
                            string filePath = Path.Combine(appPath, pWebUrlNotice);
                            webBrowser1.Navigate(new Uri(filePath));
                        }
                        break;
                    }
                case "Success":
                    {
                        if (pWebUrlSuccess.StartsWith("http://"))
                        {
                            webBrowser1.Navigate(pWebUrlSuccess);
                        }
                        else
                        {
                            string filePathSuccess = Path.Combine(appPath, pWebUrlSuccess);
                            webBrowser1.Navigate(new Uri(filePathSuccess));
                        }
                        break;
                    }
                case "SuccessReboot":
                    {
                        if (pWebUrlSuccessReboot.StartsWith("http://"))
                        {
                            webBrowser1.Navigate(pWebUrlSuccessReboot);
                        }
                        else
                        {
                            string filePathSuccessReboot = Path.Combine(appPath, pWebUrlSuccessReboot);
                            webBrowser1.Navigate(new Uri(filePathSuccessReboot));
                        }
                        break;
                    }
                case "Failed":
                    {
                        if (pWebUrlFailed.StartsWith("http://"))
                        {
                            webBrowser1.Navigate(pWebUrlFailed);
                        }
                        else
                        {
                            string filePathFailed = Path.Combine(appPath, pWebUrlFailed);
                            webBrowser1.Navigate(new Uri(filePathFailed));
                        }
                        break;
                    }
            }
        }


        public void LoadConfigFile(string fileName, char del)
        {
            //lets check to see if our config file exists
            if (File.Exists(fileName))
            {
                // lets open our config file and populate our settings
                FileStream inifile = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                StreamReader reader = new StreamReader(inifile);
                string recordIn;
                string pOption;
                string[] fields;
                recordIn = reader.ReadLine();
                while (recordIn != null)
                {
                    fields = recordIn.Split(deLim);
                    pOption = fields[0];

                    switch (pOption)
                    {
                        case "postpone_time":
                            if (fields[1] != "")
                            {
                                postPoneTime = Convert.ToInt32(fields[1]);
                            }
                            else
                                postPoneTime = 99;
                            break;
                        case "weburlnotice":
                            pWebUrlNotice = fields[1];
                            if (pWebUrlNotice == "")
                            {
                                MessageBox.Show("No Notice URL configured");
                                Application.Exit();
                            }
                            break;
                        case "weburlsuccess":
                            pWebUrlSuccess = fields[1];
                            if (pWebUrlSuccess == "")
                            {
                                MessageBox.Show("No Success URL configured");
                                Application.Exit();
                            }
                            break;
                        case "weburlsuccessreboot":
                            pWebUrlSuccessReboot = fields[1];
                            if (pWebUrlSuccessReboot == "")
                            {
                                MessageBox.Show("No Sucess Reboot URL configured");
                                Application.Exit();
                            }
                            break;
                        case "weburlfailed":
                            pWebUrlFailed = fields[1];
                            if (pWebUrlFailed == "")
                            {
                                MessageBox.Show("No Failed URL configured");
                                Application.Exit();
                            }
                            break;
                        case "appname":
                            pAppName = fields[1];
                            break;
                        case "forcedexitcode":
                            programExitCode = Convert.ToInt32(fields[1]);
                            break;
                        case "debug":
                            debugMode = fields[1];
                            break;
                        case "tasktrayreminder1":
                            ptaskTrayReminder1 = Convert.ToInt32(fields[1]);
                            break;
                        case "tasktrayreminder2":
                            ptaskTrayReminder2 = Convert.ToInt32(fields[1]);
                            break;
                        case "tasktrayreminder3":
                            ptaskTrayReminder3 = Convert.ToInt32(fields[1]);
                            break;
                        case "tasktrayreminder4":
                            ptaskTrayReminder4 = Convert.ToInt32(fields[1]);
                            break;
                        case "companyname":
                            pCompanyName = fields[1];
                            break;
                        case "script":
                            pFileName = fields[1];
                            break;
                        case "reboot":
                            doReboot = fields[1];
                            break;
                        case "reboot_time":
                            rebootTime = Convert.ToInt32(fields[1]);
                            break;
                        case "auto_close_time":
                            closeTime = Convert.ToInt32(fields[1]);
                            break;
                        case "controlbox":
                            appControl = Convert.ToBoolean(fields[1]);
                            break;
                    }
                    recordIn = reader.ReadLine();
                }
                //were done reading our config. close the file and move on.
                reader.Close();
                inifile.Close();
            }
        }
        //expand the UI and include variables read into program.
        public void DebugStatus(string DebugMode)
        {
            if (debugMode == "on")
            {
                this.Width = 1200;
                this.Height = 793;

                textBox1.Text = Convert.ToString(postPoneTime);
                textBox1.Visible = true;
                label1.Text = "Postpone Time";
                label1.Visible = true;

                textBox2.Text = doReboot;
                textBox2.Visible = true;
                label2.Text = "Initiate Reboot";
                label2.Visible = true;

                textBox3.Text = Convert.ToString(rebootTime);
                textBox3.Visible = true;
                label3.Text = "Reboot Time";
                label3.Visible = true;

                textBox4.Text = Convert.ToString(closeTime);
                textBox4.Visible = true;
                label4.Text = "Auto Close Time";
                label4.Visible = true;

                textBox5.Text = pWebUrlNotice;
                textBox5.Visible = true;
                label5.Text = "Web URL Notice";
                label5.Visible = true;

                textBox6.Text = pWebUrlSuccess;
                textBox6.Visible = true;
                label6.Text = "Web URL Success";
                label6.Visible = true;

                textBox7.Text = pWebUrlSuccessReboot;
                textBox7.Visible = true;
                label7.Text = "Web URL Success Reboot";
                label7.Visible = true;

                textBox8.Text = pWebUrlFailed;
                textBox8.Visible = true;
                label8.Text = "Web URL Failed";
                label8.Visible = true;

                textBox9.Text = pAppName;
                textBox9.Visible = true;
                label9.Text = "App Name";
                label9.Visible = true;

                textBox10.Text = Convert.ToString(programExitCode);
                textBox10.Visible = true;
                label10.Text = "Program ExitCode";
                label10.Visible = true;

                textBox11.Text = pCompanyName;
                textBox11.Visible = true;
                label11.Text = "Company Name";
                label11.Visible = true;

                textBox12.Text = Convert.ToString(ptaskTrayReminder1);
                textBox12.Visible = true;
                label12.Text = "Tasktray Reminder 1";
                label12.Visible = true;

                textBox13.Text = Convert.ToString(ptaskTrayReminder2);
                textBox13.Visible = true;
                label13.Text = "Tasktray Reminder 2";
                label13.Visible = true;

                textBox14.Text = Convert.ToString(ptaskTrayReminder3);
                textBox14.Visible = true;
                label14.Text = "Tasktray Reminder 3";
                label14.Visible = true;

                textBox15.Text = Convert.ToString(ptaskTrayReminder4);
                textBox15.Visible = true;
                label15.Text = "Tasktray Reminder 4";
                label15.Visible = true;

                textBox16.Text = pFileName;
                textBox16.Visible = true;
                label16.Text = "Script Name";
                label16.Visible = true;

                textBox17.Text = Convert.ToString(appControl);
                textBox17.Visible = true;
                label17.Text = "Control Box Enabled";
                label17.Visible = true;
            }
        }

        //restores the window to original state and closes the task tray icon
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            mynotifyIcon1.Visible = false;
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();
        }

        //shows time remaining on postpone accessible from task tray icon
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            mynotifyIcon1.ShowBalloonTip(15000);
        }
        
        //shows the about dialog
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Shift | Keys.H))
            {
                ConfigMgrSoftwareDeploymentUI.AboutBox1 about = new ConfigMgrSoftwareDeploymentUI.AboutBox1();
                about.Show();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
