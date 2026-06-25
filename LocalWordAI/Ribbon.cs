using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace LocalWordAI
{
    [ComVisible(true)]
    public partial class Ribbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon>
    <tabs>
      <tab id=""tabLocalAI"" label=""Local AI"">
        <group id=""grpMain"" label=""Local Word AI"">
          <toggleButton id=""btnToggleSidebar""
                        label=""AI Sidebar""
                        screentip=""Bật/tắt AI Sidebar""
                        supertip=""Mở hoặc đóng bảng trợ lý AI bên phải.""
                        onAction=""ToggleSidebar""
                        getPressed=""GetSidebarPressed""
                        size=""large""
                        imageMso=""CreateForm"" />
          <button id=""btnQuickReview""
                  label=""Review""
                  screentip=""Review vùng chọn hoặc tài liệu""
                  onAction=""QuickReview""
                  size=""normal""
                  imageMso=""ReviewShowMarkup"" />
          <button id=""btnRewrite""
                  label=""Rewrite""
                  screentip=""Viết lại vùng chọn""
                  onAction=""RewriteSelection""
                  size=""normal""
                  imageMso=""EditingFindAndReplace"" />
          <button id=""btnRedline""
                  label=""Redline""
                  screentip=""Tạo redline với Track Changes""
                  onAction=""CreateRedline""
                  size=""normal""
                  imageMso=""AcceptAllTrackedChanges"" />
          <separator id=""sep1"" />
          <button id=""btnSettings""
                  label=""Cài đặt""
                  screentip=""Cài đặt Local Word AI""
                  onAction=""OpenSettings""
                  size=""normal""
                  imageMso=""PropertySheet"" />
        </group>
        <group id=""grpCheck"" label=""Kiểm tra"">
          <button id=""btnCheckComments""
                  label=""Comments""
                  screentip=""Xử lý comments""
                  onAction=""CheckComments""
                  size=""normal""
                  imageMso=""ReviewNewComment"" />
          <button id=""btnCheckChanges""
                  label=""Tracked Changes""
                  screentip=""Phân tích tracked changes""
                  onAction=""CheckChanges""
                  size=""normal""
                  imageMso=""TrackChangesMenu"" />
          <button id=""btnCheckTerms""
                  label=""Thuật ngữ""
                  screentip=""Kiểm tra thuật ngữ nhất quán""
                  onAction=""CheckTerms""
                  size=""normal""
                  imageMso=""ReviewFinalShowingMarkup"" />
          <button id=""btnFinalCheck""
                  label=""Final Check""
                  screentip=""Kiểm tra toàn diện trước phát hành""
                  onAction=""FinalCheck""
                  size=""large""
                  imageMso=""CheckDocument"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void ToggleSidebar(IRibbonControl control, bool pressed)
        {
            ThisAddIn.Instance.ToggleSidebar();
            _ribbon?.InvalidateControl("btnToggleSidebar");
        }

        public bool GetSidebarPressed(IRibbonControl control)
        {
            return ThisAddIn.Instance?.IsSidebarVisible ?? false;
        }

        public void QuickReview(IRibbonControl control)
        {
            EnsureSidebarAndRun("review_selection");
        }

        public void RewriteSelection(IRibbonControl control)
        {
            EnsureSidebarAndRun("rewrite");
        }

        public void CreateRedline(IRibbonControl control)
        {
            EnsureSidebarAndRun("redline");
        }

        public void CheckComments(IRibbonControl control)
        {
            EnsureSidebarAndRun("check_comments");
        }

        public void CheckChanges(IRibbonControl control)
        {
            EnsureSidebarAndRun("summarize_changes");
        }

        public void CheckTerms(IRibbonControl control)
        {
            EnsureSidebarAndRun("check_terms");
        }

        public void FinalCheck(IRibbonControl control)
        {
            EnsureSidebarAndRun("final_check");
        }

        public void OpenSettings(IRibbonControl control)
        {
            var dlg = new UI.SettingsDialog();
            dlg.ShowDialog();
        }

        private void EnsureSidebarAndRun(string action)
        {
            if (!ThisAddIn.Instance.IsSidebarVisible)
                ThisAddIn.Instance.ToggleSidebar();
            GetSidebar()?.TriggerQuickAction(action);
            _ribbon?.InvalidateControl("btnToggleSidebar");
        }

        private UI.SidebarPane GetSidebar()
        {
            foreach (var pane in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (pane.Control is UI.SidebarPane sidebar)
                    return sidebar;
            }
            return null;
        }

        internal void RefreshRibbon() => _ribbon?.Invalidate();
    }
}
