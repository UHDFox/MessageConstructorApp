using System.IO;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using Xceed.Wpf.Toolkit;
using MessageBox = System.Windows.MessageBox;

namespace MessageConstructorApp
{
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
            GenerateDocument();
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

                File.Copy(TemplateFilePath, DocumentFilePath, true);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(DocumentFilePath, true))
                {
                    ReplacePlaceholders(wordDocument); // Pass the correct parameter
                    wordDocument.MainDocumentPart.Document.Save();
                }

                MessageBox.Show("Документ успешно сохранен!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReplacePlaceholders(WordprocessingDocument wordDoc)
        {
            var headers = new Dictionary<string, string>
            {
                { "{COMPANY_NAME}", GetTextOrWatermark(CompanyNameTextBox) },
                { "{COMPANY_ADDRESS}", GetTextOrWatermark(CompanyAddress) },
                { "{CURRENT_DATE}", LetterDate.SelectedDate?.ToString("dd.MM.yyyy") ?? DateTime.Now.ToString("dd.MM.yyyy") },
                { "{LETTER_SUBJECT}", GetTextOrWatermark(LetterSubject) },
                { "{RECIPIENT_NAME}", GetTextOrWatermark(LetterRecipientName) },
                { "{RECIPIENT_POSITION}", GetTextOrWatermark(LetterRecipientPosition) },
                { "{LETTER_BODY}", GetTextOrWatermark(LetterText) },
                { "{SENDER_NAME}", GetTextOrWatermark(LetterSender) },
                { "{SENDER_POSITION}", GetTextOrWatermark(SenderPosition) },
                { "{SENDER_MAIL}", GetTextOrWatermark(SenderMail) }
            };

            foreach (var textElement in wordDoc.MainDocumentPart.Document.Descendants<Text>())
            {
                foreach (var header in headers)
                {
                    if (textElement.Text.Contains(header.Key))
                    {
                        textElement.Text = textElement.Text.Replace(header.Key, header.Value);
                    }
                }
            }

            ReplaceFooterText(wordDoc, headers);
        }

        private void ReplaceFooterText(WordprocessingDocument wordDoc, Dictionary<string, string> headers)
        {
            foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
            {
                foreach (var textElement in footerPart.Footer.Descendants<Text>())
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
        }

        private static string GetTextOrWatermark(WatermarkTextBox textBox)
        {
            return string.IsNullOrWhiteSpace(textBox?.Text) ? textBox?.Watermark?.ToString() ?? string.Empty : textBox.Text;
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
                GenerateDocument();
            }
        }
    }
}
