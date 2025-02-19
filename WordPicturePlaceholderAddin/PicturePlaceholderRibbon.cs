using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using WordPicturePlaceholderAddin.Helpers;

namespace WordPicturePlaceholderAddin
{
    public partial class PicturePlaceholderRibbon
    {
        private void PicturePlaceholderRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection selection = Globals.ThisAddIn.Application.Selection;

            // Load existing count from document properties
            HelperMethods.GetPlaceholderCount(doc);
            HelperMethods.placeHodlerCount++;

            // Create placeholder text
            string placeholderText = $"(picture {HelperMethods.placeHodlerCount})";
            HelperMethods.placeholders.Add(placeholderText);

            // Insert text at cursor position
            selection.TypeText(placeholderText + " ");

            // Save updated count in document properties
            HelperMethods.SetPlaceholderCount(doc, HelperMethods.placeHodlerCount);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string message = "Current Placeholders:\n" + string.Join("\n", HelperMethods.placeholders);
            MessageBox.Show(message, "Placeholder List");
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            string input = Interaction.InputBox(
                "Enter the placeholder number to remove",
                "Remove Placeholder",
                "1");

            if (int.TryParse(input, out int placeholderNumber))
            {
                // Validate the entered number.
                if (placeholderNumber < 1 || placeholderNumber > HelperMethods.placeholders.Count)
                {
                    MessageBox.Show("Invalid placeholder number.", "Error");
                    return;
                }

                string placeholderText = $"(picture {placeholderNumber})";

                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Range range = doc.Content;

                // Find the placeholder text in the document
                if (range.Find.Execute(FindText: placeholderText))
                {
                    range.Text = "";

                    HelperMethods.placeholders.Remove(placeholderText);
                    MessageBox.Show($"Removed {placeholderText} from the document.", "Success");
                    HelperMethods.placeHodlerCount--;

                    // Update document properties
                    HelperMethods.SetPlaceholderCount(doc, HelperMethods.placeHodlerCount);

                    // Renumber subsequent placeholders.
                    // Loop from the removed number to the end of the list.
                    for (int i = placeholderNumber; i <= HelperMethods.placeHodlerCount; i++)
                    {
                        // The old placeholder had one number higher.
                        string oldPlaceholder = $"(picture {i + 1})";
                        // The new placeholder should be renumbered.
                        string newPlaceholder = $"(picture {i})";

                        // Update our internal list.
                        HelperMethods.placeholders[i - 1] = newPlaceholder;

                        // Update the document: find the old placeholder and replace it.
                        Range r = doc.Content;
                        r.Find.ClearFormatting();
                        r.Find.Text = oldPlaceholder;
                        if (r.Find.Execute())
                        {
                            r.Text = newPlaceholder;
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"Placeholder {placeholderText} not found in the document.", "Not Found");
                }
            }
            else
            {
                MessageBox.Show("Invalid number entered.", "Error");
            }
        }
    }
}
