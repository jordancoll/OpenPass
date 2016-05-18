using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using KeePass.Ecas;
using KeePass.Plugins;
using KeePass.Resources;
using KeePass.Util;
using KeePassLib;

namespace OpenPass
{
    public class OpenPassExt : Plugin
    {
        private IPluginHost host;
        private StreamWriter log;
        private ToolStripMenuItem ctxOpenAndType;

        public override bool Initialize(IPluginHost host)
        {
            this.host = host;

            try
            {
                InitLog();
            }
            catch { }

            var urlMenu = host.MainWindow.EntryContextMenu.Items["m_ctxEntryUrl"] as ToolStripMenuItem;
            if (urlMenu == null)
            {
                log.WriteLine("Couldn't find URL(s)");
               
                return false;
            }

            log.WriteLine("inited");

            ctxOpenAndType = new ToolStripMenuItem(KPRes.OpenCmd + " + " + KPRes.AutoType);
            ctxOpenAndType.ShortcutKeyDisplayString = KPRes.KeyboardKeyCtrl + "+" + KPRes.KeyboardKeyAlt + "+U";

            ctxOpenAndType.Enabled = CanOpenAndType;
            ctxOpenAndType.ShowShortcutKeys = true;
            ctxOpenAndType.Click += onClick;
            urlMenu.DropDownItems.Insert(1, ctxOpenAndType);
            urlMenu.DropDownOpening += urlMenu_DropDownOpening;

            log.WriteLine("done stuff");

            //KeePass.UI.UIUtil.ConfigureTbButton(openAndTypeMenuItem, "Open and Auto-Type", null);

            return true;
        }

        void urlMenu_DropDownOpening(object sender, EventArgs e)
        {
            ctxOpenAndType.Enabled = CanOpenAndType;
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

        private bool CanOpenAndType
        {
            get
            {
                var mw = host.MainWindow;
                var selectedEntry = mw.GetSelectedEntry(true);
                return (mw.ActiveDatabase.IsOpen) && (mw.GetSelectedEntriesCount() == 1) && (selectedEntry != null) &&
                    (!selectedEntry.Strings.GetSafe(PwDefs.UrlField).IsEmpty);
            }
        }

        private void onClick(object sender, EventArgs e)
        {
            var entry = host.MainWindow.GetSelectedEntry(true);
            if (CanOpenAndType && entry != null)
            {
                try
                {
                    WinUtil.OpenEntryUrl(entry);
                    //System.Threading.Thread.Sleep(1000);
                    AutoType.PerformIntoCurrentWindow(entry, host.MainWindow.ActiveDatabase);
                }
                catch (Exception ex)
                {
                    log.WriteLine("ERROR: " + ex.Message + ";\n" + ex.StackTrace);
                }
            }
            
        }
    }
}
