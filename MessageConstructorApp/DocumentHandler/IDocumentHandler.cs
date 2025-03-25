using System.Windows;
using DocumentFormat.OpenXml;

namespace MessageConstructorApp;

public interface IDocumentHandler
{
    public void GenerateDocument(object sender, RoutedEventArgs e);
}