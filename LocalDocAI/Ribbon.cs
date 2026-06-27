using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace LocalDocAI
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
      <tab id=""tabLocalAI"" getLabel=""GetLabel"">
        <group id=""grpMain"" getLabel=""GetLabel"">
          <toggleButton id=""btnToggleSidebar""
                        getLabel=""GetLabel""
                        getScreentip=""GetScreentip""
                        getSupertip=""GetSupertip""
                        onAction=""ToggleSidebar""
                        getPressed=""GetSidebarPressed""
                        size=""large""
                        imageMso=""CreateForm"" />
          <button id=""btnQuickReview""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""QuickReview""
                  size=""normal""
                  imageMso=""ReviewShowMarkup"" />
          <button id=""btnRewrite""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""RewriteSelection""
                  size=""normal""
                  imageMso=""EditingFindAndReplace"" />
          <button id=""btnRedline""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""CreateRedline""
                  size=""normal""
                  imageMso=""AcceptAllTrackedChanges"" />
          <button id=""btnTranslate""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""TranslateSelection""
                  size=""normal""
                  imageMso=""Globalization"" />
          <separator id=""sep1"" />
          <button id=""btnSettings""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""OpenSettings""
                  size=""normal""
                  imageMso=""PropertySheet"" />
        </group>
        <group id=""grpCheck"" getLabel=""GetLabel"">
          <button id=""btnCheckComments""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""CheckComments""
                  size=""normal""
                  imageMso=""ReviewNewComment"" />
          <button id=""btnCheckChanges""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""CheckChanges""
                  size=""normal""
                  imageMso=""TrackChangesMenu"" />
          <button id=""btnCheckTerms""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
                  onAction=""CheckTerms""
                  size=""normal""
                  imageMso=""ReviewFinalShowingMarkup"" />
          <button id=""btnFinalCheck""
                  getLabel=""GetLabel""
                  getScreentip=""GetScreentip""
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

        public void TranslateSelection(IRibbonControl control)
        {
            EnsureSidebarAndRun("translate");
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
            _ribbon?.Invalidate();
            GetSidebar()?.UpdateLanguage();
        }

        public string GetLabel(IRibbonControl control)
        {
            return LocalDocAI.Persistence.LocalizationService.Get(control.Id);
        }

        public string GetScreentip(IRibbonControl control)
        {
            return LocalDocAI.Persistence.LocalizationService.Get(control.Id + "_tip");
        }

        public string GetSupertip(IRibbonControl control)
        {
            return LocalDocAI.Persistence.LocalizationService.Get(control.Id + "_super");
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
