using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordPicturePlaceholderAddin.Helpers
{
    public static class HelperMethods
    {
        public static int placeHodlerCount;
        public static List<string> placeholders = new List<string>();
        public static void GetPlaceholderCount(Document doc)
        {
            try
            {
                foreach (dynamic prop in doc.CustomDocumentProperties)
                {
                    if (prop.Name == "PlaceholderCount")
                    {
                        placeHodlerCount = Convert.ToInt32(prop.Value);
                    }
                }
            }
            catch { /* Property does not exist yet */ }
        }

        public static void SetPlaceholderCount(Document doc, int count)
        {
            try
            {
                dynamic properties = doc.CustomDocumentProperties;

                // Check if the property already exists
                foreach (dynamic prop in properties)
                {
                    if (prop.Name == "PlaceholderCount")
                    {
                        prop.Value = count;
                        return;
                    }
                }

                // If not found, create it
                properties.Add("PlaceholderCount", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeNumber, count);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving placeholder count: {ex.Message}");
            }
        }

        public static void RebuildPlaceholdersList(Document doc)
        {
            placeholders.Clear();

            for (int i = 1; i <= placeHodlerCount; i++)
            {
                string placeholderText = $"(picture {i})";

                // Check if the text actually exists in the document
                Range range = doc.Content;
                range.Find.ClearFormatting();
                range.Find.Text = placeholderText;

                if (range.Find.Execute())
                {
                    placeholders.Add(placeholderText);
                }
            }
        }

    }
}
