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

namespace MessageConstructorApp;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }
    
    protected void ButtonClick(object sender, RoutedEventArgs e)
    {
        
    }

    public void GenerateDocument(object sender, RoutedEventArgs e)
    {
        try
        {
            string templateFilePath = "template.docx";
            string filepath = "capybara.docx";
            File.Copy(templateFilePath, filepath);

            if (!File.Exists(filepath))
            {
                MessageBox.Show("Шаблон документа не найден.");
                return;
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
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
            { "{COMPANY_NAME}", CompanyNameTextBox.Text },
            { "{COMPANY_ADDRESS}", CompanyAddress.Text },
        };

        var text = document.InnerText;
       /* foreach (var header in headers)
        {
            if (text.Contains(header.Key))
            {
                document.InnerText.Replace(header.Key, header.Value);
            }
        }*/
        
        foreach (var textElement in document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
        {
            // Для каждого текста проверяем на наличие заглушки
            foreach (var header in headers)
            {
                if (textElement.Text.Contains(header.Key))
                {
                    // Заменяем текст внутри элемента
                    textElement.Text = textElement.Text.Replace(header.Key, header.Value);
                }
            }
        }
        
        
        
    }
}