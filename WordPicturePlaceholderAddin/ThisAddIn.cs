using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using WordPicturePlaceholderAddin.Helpers;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace WordPicturePlaceholderAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentOpen += Application_DocumentOpen;


            // If a document is already open, initialize placeholders
            if (this.Application.Documents.Count > 0)
            {
                Application_DocumentOpen(this.Application.ActiveDocument);
            }

        }

        private void Application_DocumentOpen(Document doc)
        {
            try
            {
                if (doc != null)
                {
                    HelperMethods.GetPlaceholderCount(doc);

                    // Try to rebuild the placeholders list
                    HelperMethods.RebuildPlaceholdersList(doc);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show($"No active document found: {ex.Message}", "Error");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
