using GroupDocs.Editor;
using GroupDocs.Editor.Formats;
using GroupDocs.Editor.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GroupDocsTesting
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void butDownloadHTML_Click(object sender, EventArgs e)
        {
            if (fileupload.HasFile)
            {
                try
                {
                    string myDate = DateTime.Now.ToString("MM-dd-yyyy--H-mm-ss--");
                    string filename = Path.GetFileName(fileupload.FileName);
                    string fullpath = Server.MapPath("/Uploaded/") +myDate + filename;
                    string fullpathExportDocX = fullpath.Replace(".docx", "-GroupDocsExport.docx");
                    string fullpathExportHTML = fullpath.Replace(".docx", "-GroupDocsExport.html");
                    string fullpathExportPDF = fullpath.Replace(".docx", "-GroupDocsExport.pdf");
                    fullpathExportDocX = fullpathExportDocX.Replace("Uploaded", "Exported");
                    fullpathExportHTML = fullpathExportHTML.Replace("Uploaded", "Exported");
                    fullpathExportPDF = fullpathExportPDF.Replace("Uploaded", "Exported");
                    fileupload.SaveAs(fullpath);
                    using (Editor editor = new Editor(fullpath))
                    {
                        WordProcessingEditOptions editOptions = new WordProcessingEditOptions();
                        editOptions.EnablePagination = false;
                        EditableDocument readyToSave = editor.Edit(editOptions);
                        if (!string.IsNullOrEmpty(fullpathExportDocX))
                            editor.Save(readyToSave, fullpathExportDocX, new WordProcessingSaveOptions(WordProcessingFormats.Docx));
                        if (!string.IsNullOrEmpty(fullpathExportPDF))
                            editor.Save(readyToSave, fullpathExportPDF, new PdfSaveOptions());
                        if (!string.IsNullOrEmpty(fullpathExportHTML))
                            readyToSave.Save(fullpathExportHTML);
                        readyToSave.Dispose();
                        editor.Dispose();
                    }
                    lblStatus.Text = "Upload status: File uploaded and converted!";
                }
                catch (Exception ex)
                {
                    lblStatus.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message;
                }
            }
        }
    }
}