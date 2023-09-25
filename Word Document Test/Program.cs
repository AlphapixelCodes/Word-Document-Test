

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;



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
            Console.WriteLine(text.Text);
            foreach (KeyValuePair<string, string> entry in replacements)
            {
                while (text.Text.Contains(entry.Key))
                {
                    text.Text = text.Text.Replace(entry.Key, entry.Value);
                }
            }
        }
    }

}
var filePath = @"C:\Users\Nick\source\repos\Word Document Test\Word Document Test\Letterhead Simple Template.docx";//path to letterhead simple template in solution
var newFilePath = Path.GetDirectoryName(filePath) + "\\output.docx";
//File.Copy(filePath,newFilePath,true);

var replacements = new Dictionary<string, string>()
{
    {"%body%","Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Magna fringilla urna porttitor rhoncus dolor purus non enim. Vel facilisis volutpat est velit egestas dui id ornare arcu. Leo vel orci porta non pulvinar neque laoreet suspendisse. Pretium fusce id velit ut tortor pretium viverra suspendisse. Sem nulla pharetra diam sit amet nisl. Elit ut aliquam purus sit amet luctus venenatis. Id donec ultrices tincidunt arcu. Turpis in eu mi bibendum neque egestas congue. Mi tempus imperdiet nulla malesuada pellentesque elit eget gravida cum. Cras adipiscing enim eu turpis egestas pretium. Nisl nisi scelerisque eu ultrices vitae auctor. Egestas diam in arcu cursus euismod quis viverra nibh.\nGravida arcu ac tortor dignissim. Massa massa ultricies mi quis hendrerit dolor magna eget. Id leo in vitae turpis. Massa sed elementum tempus egestas sed sed. Viverra suspendisse potenti nullam ac. Faucibus vitae aliquet nec ullamcorper sit amet. In fermentum et sollicitudin ac orci phasellus egestas tellus. Blandit massa enim nec dui nunc. Nisi porta lorem mollis aliquam ut porttitor leo. Placerat vestibulum lectus mauris ultrices eros in cursus turpis. Turpis tincidunt id aliquet risus. Condimentum vitae sapien pellentesque habitant morbi tristique senectus et. Amet luctus venenatis lectus magna fringilla urna porttitor rhoncus dolor. Facilisis mauris sit amet massa. Blandit libero volutpat sed cras ornare." }
};
MultiReplacement(filePath, replacements, newFilePath);