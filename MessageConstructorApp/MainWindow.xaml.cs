using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Win32;

namespace MessageConstructorApp;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private string DocumentFilePath { get; set; } = "result.docx";
    private string TemplateFilePath { get; set; } = "template.docx";
    public MainWindow()
    {
        InitializeComponent();
    }
    
    protected void ButtonClick(object sender, RoutedEventArgs e)
    {
        
    }

    private void GenerateDocument()
    {
        try
        {

            if (!File.Exists(TemplateFilePath))
            {
                MessageBox.Show("Шаблон документа не найден.");
                return;
            }
            if(File.Exists(DocumentFilePath))
            {
                File.Delete(DocumentFilePath);
            }
            
            File.Copy(TemplateFilePath, DocumentFilePath, true);


            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(DocumentFilePath, true))
            {
                var body = wordDocument.MainDocumentPart.Document.Body;

                ReplacePlaceholders(body);

                wordDocument.MainDocumentPart.Document.Save();
            }
        }

        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
        
        
    }

    private void ReplacePlaceholders(OpenXmlElement document)
    {
        var headers = new Dictionary<string, string>
        {
            { "{COMPANY_NAME}", string.IsNullOrEmpty(CompanyNameTextBox.Text) ? CompanyNameTextBox.Watermark.ToString()! : CompanyNameTextBox.Text },
            { "{COMPANY_ADDRESS}", string.IsNullOrEmpty(CompanyAddress.Text) ? CompanyAddress.Watermark.ToString()! : CompanyAddress.Text },
            { "{CURRENT_DATE}", LetterDate.SelectedDate?.ToString("dd.MM.yyyy") ?? DateTime.Now.ToString("dd.MM.yyyy") },
            { "{LETTER_SUBJECT}", string.IsNullOrWhiteSpace(LetterTopic.Text) ? LetterTopic.Watermark.ToString()! : LetterTopic.Text },
            { "{RECIPIENT_NAME}", string.IsNullOrWhiteSpace(LetterRecipient.Text) ? LetterRecipient.Watermark.ToString()! : LetterRecipient.Text },
            { "{LETTER_BODY}", string.IsNullOrWhiteSpace(LetterText.Text) ? LetterText.Watermark.ToString()! : LetterText.Text },
            { "{SENDER_NAME}", string.IsNullOrWhiteSpace(LetterSender.Text) ? LetterSender.Watermark.ToString()! : LetterSender.Text },
            { "{SENDER_POSITION}", string.IsNullOrWhiteSpace(SenderPosition.Text) ? SenderPosition.Watermark.ToString()! : SenderPosition.Text }
        };
        
        foreach (var textElement in document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
        {
            foreach (var header in headers)
            {
                if (textElement.Text.Contains(header.Key))
                {
                    textElement.Text = textElement.Text.Replace(header.Key, header.Value);
                }
            }
        }
    }
    protected void SaveButtonClick(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog()
        {
            Filter = "Word Documents (*.docx)|*.docx",
            DefaultExt = "docx"
        };
        
        if (dialog.ShowDialog() == true)
        {
            DocumentFilePath = dialog.FileName;
        }
        
        GenerateDocument();
    }
}