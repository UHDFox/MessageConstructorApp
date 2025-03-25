using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace MessageConstructorApp.DocumentHandler;

public class DocumentHandler : IDocumentHandler
{ 
    /*private IReadOnlyDictionary<string, string> Headers = new Dictionary<string, string>
    {
        { "{COMPANY_NAME}", CompanyNameTextBox.Text },
        { "{COMPANY_ADDRESS}", Comp.Text },
        { "{RECIPIENT_POST}", RecipientPostTextBox.Text },
        { "{RECIPIENT_NAME}", RecipientNameTextBox.Text },
        { "{LETTER_SUBJECT}", LetterSubjectTextBox.Text },
        { "{LETTER_BODY}", LetterBodyTextBox.Text },
        { "{SENDER_NAME}", SenderNameTextBox.Text },
        { "{SENDER_POSITION}", SenderPositionTextBox.Text },
        { "{CURRENT_DATE}", DateTime.Now.ToString("dd.MM.yyyy") }
    };*/

    public void GenerateDocument(object sender, RoutedEventArgs e)
    {
        string filepath = "document.txt";

        using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
        {
            var body = wordDocument.MainDocumentPart.Document.Body;
        }
    }

    private void ReplacePlaceholders(OpenXmlElement document)
    {
        //document.Elements<>()
    }
}