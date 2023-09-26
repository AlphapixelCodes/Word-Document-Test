

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

void MultiReplacement(string template, Dictionary<string, string> replacements, string newFilePath)
{
    
    // REPLACE NEW LINES WITH A RANDOM STRING AND THEN ONCE REPLACEMENTS ARE DONE, BREAK UP THE TEXT INTO TEXT BREAK TEXT
    File.Copy(template, newFilePath, true);

    using (WordprocessingDocument doc = WordprocessingDocument.Open(newFilePath, true))
    {
        MainDocumentPart mainPart = doc.MainDocumentPart;

        List<Text> texts = mainPart.RootElement.Descendants<Text>().ToList();
        //add header texts
        texts.AddRange(mainPart.HeaderParts.SelectMany(e => e.Header.Descendants<Run>().SelectMany(e => e.Descendants<Text>())));

        //add footer
        texts.AddRange(mainPart.FooterParts.SelectMany(e => e.Footer.Descendants<Run>().SelectMany(e => e.Descendants<Text>())));
        foreach (var text in texts)
        {
            //Console.WriteLine(text.Text);
            foreach (KeyValuePair<string, string> entry in replacements)
            {
                while (text.Text.Contains(entry.Key))
                {
                    if (entry.Value.Contains("\n"))
                    {
                        var index = text.Text.IndexOf(entry.Key);
                        var start = text.Text.Substring(0, index);
                        
                        var end = text.Text.Substring(start.Length + entry.Key.Length, text.Text.Length - (index + entry.Key.Length));

                        var lines = Regex.Split(entry.Value, "[\n\r]{1,}");

                        
                        Paragraph firstParagraph = (Paragraph)text.Parent.Parent;
                        var paraProperties = firstParagraph.ParagraphProperties; // Get the paragraph properties

                        text.Text = start + lines[0];

                        Paragraph lastParagraph = firstParagraph;

                        for (int i = 1; i < lines.Length; i++)
                        {
                            // Create a new paragraph
                            Paragraph p = new Paragraph();

                            // Clone and set the paragraph properties
                            if (paraProperties != null)
                            {
                                p.ParagraphProperties = (ParagraphProperties)paraProperties.CloneNode(true);
                            }

                            // Clone and set the run properties
                            Run run = new Run();
                            RunProperties runProperties = ((Run)text.Parent).RunProperties;

                            if (runProperties != null)
                            {
                                run.RunProperties = (RunProperties)runProperties.CloneNode(true);
                            }

                            // Add text to the run
                            var t = (lines.Length == i + 1) ? lines[i] + end : lines[i];
                            run.AppendChild(new Text(t));

                            // Add the run to the paragraph
                            p.AppendChild(run);

                            // Insert the new paragraph after the lastParagraph
                            lastParagraph.InsertAfterSelf(p);
                            lastParagraph = p;
                        }

                    }
                    else
                    {
                        text.Text = text.Text.Replace(entry.Key, entry.Value);
                    }
                    
                }
            }
        }
    }

}
var filePath = @"C:\Users\Nick\source\repos\Cover Letter Generator\Cover Letter Generator\Templates\OtherTemplates\Basic Template.docx";//path to letterhead simple template in solution
var newFilePath = Path.GetDirectoryName(filePath) + "\\output.docx";
newFilePath = @"C:\Users\Nick\Documents\GitHub\Word-Document-Test\Word Document Test\output.docx";
//File.Copy(filePath,newFilePath,true);

var replacements = new Dictionary<string, string>()
{
    {"%body%","Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Magna fringilla urna porttitor rhoncus dolor purus non enim. Vel facilisis volutpat est velit egestas dui id ornare arcu. Leo vel orci porta non pulvinar neque laoreet suspendisse. Pretium fusce id velit ut tortor pretium viverra suspendisse. Sem nulla pharetra diam sit amet nisl. Elit ut aliquam purus sit amet luctus venenatis. Id donec ultrices tincidunt arcu. Turpis in eu mi bibendum neque egestas congue. Mi tempus imperdiet nulla malesuada pellentesque elit eget gravida cum. Cras adipiscing enim eu turpis egestas pretium. Nisl nisi scelerisque eu ultrices vitae auctor. Egestas diam in arcu cursus euismod quis viverra nibh.\nGravida arcu ac tortor dignissim. Massa massa ultricies mi quis hendrerit dolor magna eget. Id leo in vitae turpis. Massa sed elementum tempus egestas sed sed. Viverra suspendisse potenti nullam ac. Faucibus vitae aliquet nec ullamcorper sit amet. In fermentum et sollicitudin ac orci phasellus egestas tellus. Blandit massa enim nec dui nunc. Nisi porta lorem mollis aliquam ut porttitor leo. Placerat vestibulum lectus mauris ultrices eros in cursus turpis. Turpis tincidunt id aliquet risus. Condimentum vitae sapien pellentesque habitant morbi tristique senectus et. Amet luctus venenatis lectus magna fringilla urna porttitor rhoncus dolor. Facilisis mauris sit amet massa. Blandit libero volutpat sed cras ornare." }
};
MultiReplacement(filePath, replacements, newFilePath);