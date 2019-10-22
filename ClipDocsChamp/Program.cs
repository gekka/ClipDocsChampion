//MIT License

//Copyright(c) 2019 gekka

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

namespace ClipDocsChampion
{
    using System;
    using System.Linq;
    using System.Windows.Forms;

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Watch watch = new Watch();
            watch.Initialize();

            using (WatchForm wf = new WatchForm())
            {
                wf.DrawClipBoard += (s, e) => watch.OnClip();

                Application.Run();
            }
            Properties.Settings.Default.Save();
        }

        class Watch
        {
            public void Initialize()
            {
                ToolStripMenuItem menuEnable = new ToolStripMenuItem("Enable");
                menuEnable.Checked = Properties.Settings.Default.IsEnable;
                menuEnable.CheckOnClick = true;
                menuEnable.CheckedChanged += (s, e) =>
                {
                    Properties.Settings.Default.IsEnable = menuEnable.Checked;
                    Properties.Settings.Default.Save();
                };

                ToolStripTextBox menuID = new ToolStripTextBox();
                menuID.Text = Properties.Settings.Default.WT_mc_id;
                menuID.TextChanged += (s, e) => { Properties.Settings.Default.WT_mc_id = menuID.Text; };

                ToolStripMenuItem menuSetting = new ToolStripMenuItem("WT_mc_id=");
                menuSetting.DropDownItems.Add(menuID);

                ToolStripMenuItem menuExit = new ToolStripMenuItem("Exit");
                menuExit.Click += (s, e) => { Application.Exit(); };

                var notify = new System.Windows.Forms.NotifyIcon();
                notify.Icon = Properties.Resources.DocsIcon;
                notify.Text = "Docs Champion";
                notify.Visible = true;
                notify.ContextMenuStrip = new ContextMenuStrip()
                {
                    Items =
                    {
                        menuEnable,
                        new ToolStripSeparator(),
                        menuSetting,
                        menuExit
                    }
                };

            }

            private System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex("^https+://docs.microsoft.com/.*", System.Text.RegularExpressions.RegexOptions.Compiled);

            /// <summary>クリップボードにあるURLにIDを追加する</summary>
            public void OnClip()
            {
                if (string.IsNullOrEmpty(Properties.Settings.Default.WT_mc_id) || !Properties.Settings.Default.IsEnable)
                {
                    return;
                }

                string text = Clipboard.GetText();
                if (!reg.IsMatch(text))
                {
                    return;
                }
                try
                {
                    System.UriBuilder ub = new UriBuilder(text);

                    System.Collections.Specialized.NameValueCollection queries = System.Web.HttpUtility.ParseQueryString(ub.Query);
                    if (!queries.AllKeys.Any(_ => _ == "WT.mc_id"))
                    {
                        {
                            //WT.mc_idが先頭に来るように
                            System.Collections.Specialized.NameValueCollection newQueries = System.Web.HttpUtility.ParseQueryString("");
                            newQueries.Add("WT.mc_id", Properties.Settings.Default.WT_mc_id);
                            newQueries.Add(queries);
                            queries = newQueries;
                        }

                        ub.Query = queries.ToString();
                        string newtext = ub.Uri.GetComponents(UriComponents.AbsoluteUri & ~UriComponents.Port, UriFormat.UriEscaped);
                        Clipboard.SetText(newtext);
                    }
                }
                catch (ArgumentException)
                {
                }
                catch (UriFormatException)
                {
                }
                catch (Exception)
                {
                }
            }
        }
    }
}

namespace ClipDocsChampion
{
    using System;
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    class Win32API
    {
        [DllImport("coredll.dll", SetLastError = true)]
        public static extern Int16 GetAsyncKeyState(int vKey);

        [DllImport("user32", SetLastError = true)]
        public static extern IntPtr SetClipboardViewer(IntPtr hWndNewViewer);

        [DllImport("user32", SetLastError = true)]
        public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport("user32", SetLastError = true)]
        public extern static IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        public enum WM
        {
            WM_DRAWCLIPBOARD = 0x0308,
            WM_CHANGECBCHAIN = 0x030D,
        }
    }

    /// <summary>クリップボード監視</summary>
    class WatchForm : System.Windows.Forms.Form
    {
        private IntPtr next;

        public WatchForm()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.ControlBox = false;
            this.Opacity = 0;
            this.ShowInTaskbar = false;

            next = Win32API.SetClipboardViewer(this.Handle);
        }

        public event EventHandler DrawClipBoard;

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
            case (int)Win32API.WM.WM_DRAWCLIPBOARD:
                DrawClipBoard?.Invoke(this, EventArgs.Empty);

                Win32API.SendMessage(next, m.Msg, m.WParam, m.LParam);
                break;
            case (int)Win32API.WM.WM_CHANGECBCHAIN:
                if (m.WParam == next)
                {
                    next = m.LParam;
                }
                else if (next != IntPtr.Zero)
                {
                    Win32API.SendMessage(next, m.Msg, m.WParam, m.LParam);
                }
                break;
            default:
                break;
            }
            base.WndProc(ref m);
        }

        public void UnHook()
        {
            if (this.Handle != IntPtr.Zero && next != IntPtr.Zero)
            {
                Win32API.ChangeClipboardChain(this.Handle, next);
                next = IntPtr.Zero;
            }
        }

        protected override void Dispose(bool disposing)
        {
            UnHook();
            base.Dispose(disposing);
        }
    }
}
