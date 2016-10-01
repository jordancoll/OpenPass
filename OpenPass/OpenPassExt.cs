using System;
using System.IO;
using System.Windows.Forms;
using KeePass.Plugins;
using KeePass.Resources;
using KeePass.UI;
using KeePass.Util;
using KeePassLib;

namespace OpenPass
{
    public class OpenPassExt : Plugin
    {
        protected IPluginHost host;
        protected StreamWriter log;
        protected ToolStripMenuItem ctxOpenAndType;

        public override bool Initialize(IPluginHost host)
        {
            this.host = host;

            try
            {
                InitLog();
            }
            catch { }

            if (!CreateMenuItem())
            {
                return false;
            }

            RegisterKeyboardShortcut();

            log.WriteLine("inited");

            return true;
        }

        private bool CreateMenuItem()
        {
            var urlMenu = host.MainWindow.EntryContextMenu.Items["m_ctxEntryUrl"] as ToolStripMenuItem;
            if (urlMenu == null)
            {
                log.WriteLine("Couldn't find URL(s)");
                return false;
            }

            ctxOpenAndType = new ToolStripMenuItem(KPRes.OpenCmd + " + " + KPRes.AutoType);

            ctxOpenAndType.Enabled = CanOpenAndType;
            ctxOpenAndType.ShowShortcutKeys = true;
            ctxOpenAndType.Click += delegate(object sender, EventArgs e) { OpenAndAutoType(); };
            urlMenu.DropDownItems.Insert(1, ctxOpenAndType);
            urlMenu.DropDownOpening += delegate(object sender, EventArgs e) { ctxOpenAndType.Enabled = CanOpenAndType; };
            UIUtil.AssignShortcut(ctxOpenAndType, Keys.Control | Keys.Alt | Keys.U);

            return true;
        }

        private void RegisterKeyboardShortcut()
        {
            var listView = host.MainWindow.Controls.Find("m_lvEntries", true);
            if (listView.Length != 1)
            {
                log.WriteLine("Couldn't add keyboard shortcut, m_lvEntries count=" + listView.Length);
                return;
            }

            log.WriteLine("added keydown event");
            listView[0].KeyDown +=
                delegate(object sender, KeyEventArgs e)
                {
                    if (e.KeyData == (Keys.Control | Keys.Y))
                    {
                        OpenAndAutoType();
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                };

        }

        public override void Terminate()
        {
            log.Flush();
            log.Close();
            base.Terminate();
        }

        private void InitLog()
        {
            var codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            var uri = new UriBuilder(codeBase);
            var path = Uri.UnescapeDataString(uri.Path);
            log = new StreamWriter(File.Open(Path.Combine(Path.GetDirectoryName(path), "OpenPass.log"), FileMode.Append, FileAccess.Write, FileShare.Read));
            log.AutoFlush = true;
        }

        protected bool CanOpenAndType
        {
            get
            {
                var mw = host.MainWindow;
                var selectedEntry = mw.GetSelectedEntry(true);
                return (mw.ActiveDatabase.IsOpen) && (mw.GetSelectedEntriesCount() == 1) && (selectedEntry != null) &&
                    (!selectedEntry.Strings.GetSafe(PwDefs.UrlField).IsEmpty) && (selectedEntry.GetAutoTypeEnabled());
            }
        }

        protected void OpenAndAutoType()
        {
            var entry = host.MainWindow.GetSelectedEntry(true);
            string seq = entry.AutoType.DefaultSequence;
            foreach (var association in entry.AutoType.Associations)
            {
                if (association.WindowName == ">OpenPass<")
                {
                    seq = association.Sequence;
                }
            }

            if (CanOpenAndType && entry != null)
            {
                try
                {
                    WinUtil.OpenEntryUrl(entry);
                    AutoType.PerformIntoCurrentWindow(entry, host.MainWindow.ActiveDatabase, seq);
                }
                catch (Exception ex)
                {
                    log.WriteLine("ERROR: " + ex.Message + ";\n" + ex.StackTrace);
                }
            }
        }
    }
}
