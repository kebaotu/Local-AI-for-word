using System;
using LocalWordAI.Persistence;
using LocalWordAI.UI;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace LocalWordAI
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private SidebarPane _sidebarControl;

        internal static ThisAddIn Instance { get; private set; }
        internal SettingsService Settings { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;
            Settings = new SettingsService();
            Settings.Load();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Settings?.Save();
            _taskPane = null;
            _sidebarControl = null;
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        internal void ToggleSidebar()
        {
            if (_taskPane == null)
            {
                _sidebarControl = new SidebarPane();
                _taskPane = this.CustomTaskPanes.Add(_sidebarControl, "Local Word AI");
                _taskPane.Width = 360;
                _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = true;
            }
            else
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        internal bool IsSidebarVisible => _taskPane != null && _taskPane.Visible;

        internal Word.Application WordApp => this.Application;

        internal Word.Document ActiveDoc
        {
            get
            {
                try { return this.Application.ActiveDocument; }
                catch { return null; }
            }
        }
    }
}
