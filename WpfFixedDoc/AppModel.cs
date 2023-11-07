using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Xps.Packaging;
using System.Reflection.Metadata;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Media.Imaging;
using CommunityToolkit.Mvvm.DependencyInjection;
using System.Xml.Linq;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace WpfFixedDoc
{
    public partial class AppModel : ObservableObject
    {

        private XpsDocument _xpsPackage;

        private string _xpsFilename = "c:\\temp\\test.xps";

        [ObservableProperty]
        private IDocumentPaginatorSource _liveDoc;

        [ObservableProperty]
        private int _counter;

        [RelayCommand]
        private void OpenFile()
        {
            using (var fd = new CommonOpenFileDialog())
            {
                if (fd.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    _xpsPackage = new XpsDocument(fd.FileName, FileAccess.Read, CompressionOption.NotCompressed);

                    // Get the FixedDocumentSequence from the package.
                    FixedDocumentSequence fixedDocumentSequence = _xpsPackage.GetFixedDocumentSequence();

                    // Set the new FixedDocumentSequence as
                    // the DocumentViewer's paginator source.
                    LiveDoc = fixedDocumentSequence as IDocumentPaginatorSource;
                }
            }


        }

        [RelayCommand]
        private void CreateDoc()
        {
            LiveDoc = createFixedDoc();
        }

        [RelayCommand]
        private void PrintDoc()
        {
            var pd = new PrintDialog() { PageRangeSelection = PageRangeSelection.UserPages };

            if (pd.ShowDialog() == true)
            {
                pd.PrintDocument(LiveDoc.DocumentPaginator, "Print Job Name");
            }
        }

        [RelayCommand]
        private void SaveXps()
        {
            var fd = new CommonOpenFileDialog() { IsFolderPicker = true };

            if (fd.ShowDialog() == CommonFileDialogResult.Ok)
            {
                saveXpsFile(Path.Combine(fd.FileName, "test.xps"));
            }
        }

        [RelayCommand]
        private void SavePdf()
        {
            var fd = new CommonOpenFileDialog() { IsFolderPicker = true };

            if (fd.ShowDialog() == CommonFileDialogResult.Ok)
            {
                savePdfFile(Path.Combine(fd.FileName, "test.pdf"));
            }
        }

        private FixedDocument createFixedDoc()
        {

            var document = new FixedDocument();

            document.DocumentPaginator.PageSize = new System.Windows.Size(793.7, 1122.5);

            var margin = new Thickness(50);

            // create a page
            FixedPage page1 = new FixedPage();
            page1.Width = document.DocumentPaginator.PageSize.Width;
            page1.Height = document.DocumentPaginator.PageSize.Height;

            var panel = new StackPanel() { Width = page1.Width }; //, Height = page1.Height };

            var logo = new System.Windows.Controls.Image();

            logo.Source = createBitmap("fish");
            logo.Height = 150;
            logo.HorizontalAlignment = HorizontalAlignment.Left;
            logo.Margin = new Thickness(margin.Left, margin.Top / 2, margin.Right, margin.Bottom / 2);

            panel.Children.Add(logo);

            // panel.Orientation = Orientation.Vertical;

            // add title to the page
            var page1Title = new TextBlock();
            page1Title.Text = "This is the first page";
            page1Title.FontSize = 40; // 30pt text
            page1Title.Margin = margin; // 1 inch margin
            // page1.Children.Add(page1Title);

            DockPanel.SetDock(page1Title, Dock.Top);

            // panel.Children.Add(page1Title);


            // add subtitle to the page
            var page1Subtitle = new TextBlock();
            page1Subtitle.Text = "This is also the first page";
            page1Subtitle.FontSize = 30; // 30pt text
            page1Subtitle.Margin = margin; // 1 inch margin

            //  panel.Children.Add(page1Subtitle);


            var grid = new Grid();

            RowDefinition row_def;
            ColumnDefinition col_def;

            var gridBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 255, 0));

            gridBrush = Brushes.Black;

            for (int i = 0; i < 3; i++)
            {
                row_def = new RowDefinition();
                grid.RowDefinitions.Add(row_def);

                col_def = new ColumnDefinition();

                if (i == 1) col_def.Width = new GridLength(0.5, GridUnitType.Star);

                grid.ColumnDefinitions.Add(col_def);

            }

            for (int i = 0; i < 3; i++)
                for (int j = 0; j < 3; j++)
                {
                    var txt = new TextBlock { Text = $"{i},{j}", HorizontalAlignment = HorizontalAlignment.Center };


                    var border = new Border();

                    border.BorderBrush = gridBrush;

                    border.BorderThickness = new Thickness(1);

                    border.Child = txt;

                    Grid.SetRow(border, i);
                    Grid.SetColumn(border, j);

                    txt.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 0, 0));
                    txt.Foreground = Brushes.Black;
                    txt.Margin = new Thickness(5);
                    grid.Children.Add(border);
                }

            // DockPanel.SetDock(grid, Dock.Bottom);

            var outerBorder = new Border();

            outerBorder.BorderBrush = gridBrush;

            outerBorder.BorderThickness = new Thickness(2);

            outerBorder.Child = grid;


            outerBorder.Margin = new Thickness(margin.Left, 0, margin.Right, 0);

            panel.Children.Add(outerBorder);


            page1.Children.Add(panel);

            // add the page to the document
            PageContent page1Content = new PageContent();
            ((IAddChild)page1Content).AddChild(page1);
            document.Pages.Add(page1Content);

            // do the same for the second page
            FixedPage page2 = new FixedPage();
            page2.Width = document.DocumentPaginator.PageSize.Width;
            page2.Height = document.DocumentPaginator.PageSize.Height;
            TextBlock page2Text = new TextBlock();
            page2Text.Text = "This is NOT the first page";
            page2Text.FontSize = 40;
            page2Text.Margin = new Thickness(96);
            page2.Children.Add(page2Text);
            PageContent page2Content = new PageContent();
            ((IAddChild)page2Content).AddChild(page2);

            //document.Pages.Add(page2Content);

            return document;
        }

        private BitmapImage createBitmap(string imagename)
        {
            var name = $"pack://application:,,/images/{imagename}.png";

            var image = new BitmapImage();
            image.BeginInit();
            image.UriSource = new Uri(name, UriKind.RelativeOrAbsolute);
            image.EndInit();

            return image;
        }

        private bool saveXpsFile(string filename)
        {
            var result = true;

            try
            {
                var xps = new XpsDocument(filename, FileAccess.Write, CompressionOption.NotCompressed);

                var writer = XpsDocument.CreateXpsDocumentWriter(xps);
                writer.Write(LiveDoc as FixedDocument);

                xps.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                result = false;
            }

            return result;
        }

        private void savePdfFile(string filename)
        {
            // Print to PDF
            PdfFilePrinter.PrintXpsToPdf(LiveDoc as FixedDocument, filename, "Document Title");
        }
    }

}
