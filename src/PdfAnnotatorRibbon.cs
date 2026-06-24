using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace RKPdfAnnotator;

public sealed class PdfAnnotatorRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
        => """
           <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
             <ribbon>
               <tabs>
                 <tab id="RKPdfAnnotatorTab" label="PDF Annotator">
                   <group id="RKPdfAnnotatorGroup" label="Annotate">
                     <button id="rkAnnotateButton"
                             label="Annotate PDFs"
                             size="large"
                             getImage="GetButtonImage"
                             screentip="Annotate P&amp;ID PDFs from the current workbook"
                             supertip="Choose a PDF, select the tag and note columns from the active Excel sheet, then create an annotated PDF."
                             onAction="ShowAnnotateDialog"/>
                     <button id="rkPreviewButton"
                             label="Preview Tags"
                             size="large"
                             getImage="GetButtonImage"
                             screentip="Preview tag data"
                             supertip="Show the tag values that will be used from the current Excel sheet."
                             onAction="ShowPreviewDialog"/>
                     <button id="rkUserGuideButton"
                             label="User Guide"
                             size="large"
                             getImage="GetButtonImage"
                             screentip="Open the PDF Annotator user guide"
                             supertip="Open the built-in HTML guide with installation, workflow, and troubleshooting help."
                             onAction="ShowUserGuide"/>
                     <button id="rkAboutButton"
                             label="About"
                             size="large"
                             getImage="GetButtonImage"
                             screentip="About RK PDF-Annotator"
                             supertip="Show version, project information, contact details, and the GitHub repository."
                             onAction="ShowAbout"/>
                   </group>
                 </tab>
               </tabs>
             </ribbon>
           </customUI>
           """;

    public Bitmap GetButtonImage(IRibbonControl control)
        => control.Id switch
        {
            "rkAboutButton" => AddinIcons.CreateAboutBitmap(32),
            "rkUserGuideButton" => AddinIcons.CreateGuideBitmap(32),
            "rkPreviewButton" => AddinIcons.CreatePreviewBitmap(32),
            _ => AddinIcons.CreateAnnotateBitmap(32)
        };

    public void ShowAnnotateDialog(IRibbonControl control)
    {
        try
        {
            using var dialog = new AnnotateDialog();
            dialog.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RK PDF-Annotator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public void ShowPreviewDialog(IRibbonControl control)
    {
        try
        {
            var sheet = ExcelSheetReader.ReadActiveSheet();
            using var dialog = new PreviewDialog(sheet);
            dialog.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "RK PDF-Annotator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    public void ShowUserGuide(IRibbonControl control)
    {
        UserGuideLauncher.Open();
    }

    public void ShowAbout(IRibbonControl control)
    {
        using var dialog = new AboutDialog();
        dialog.ShowDialog();
    }
}

internal static class UserGuideLauncher
{
    private const string ResourceName = "RKPdfAnnotator.USER_GUIDE.html";
    private const string GuideFileName = "USER_GUIDE.html";
    private const string OnlineGuideUrl = "https://github.com/rudikdk/rk-pdf-annotator/blob/main/excel-addin/docs/USER_GUIDE.html";

    public static void Open()
    {
        try
        {
            string guidePath = ExtractGuide();
            Process.Start(new ProcessStartInfo(guidePath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            var result = MessageBox.Show(
                "The built-in user guide could not be opened." + Environment.NewLine +
                "Open the online guide on GitHub instead?" + Environment.NewLine + Environment.NewLine +
                ex.Message,
                "RK PDF-Annotator",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
                Process.Start(new ProcessStartInfo(OnlineGuideUrl) { UseShellExecute = true });
        }
    }

    private static string ExtractGuide()
    {
        string directory = Path.Combine(Path.GetTempPath(), "RKPdfAnnotator");
        Directory.CreateDirectory(directory);

        string guidePath = Path.Combine(directory, GuideFileName);
        using Stream? resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName);
        if (resource is null)
            throw new InvalidOperationException("The embedded user guide was not found in the add-in.");

        using var output = File.Create(guidePath);
        resource.CopyTo(output);
        return guidePath;
    }
}
