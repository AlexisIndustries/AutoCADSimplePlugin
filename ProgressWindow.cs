using System.Windows;

namespace AutoCADSimplePlugin;
public class ProgressWindow : Window
{
    public ProgressWindow()
    {
        Title = "Loading File";
        Width = 300;
        Height = 100;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        Content = new System.Windows.Controls.ProgressBar
        {
            IsIndeterminate = true,
            Height = 40,
            Margin = new Thickness(20)
        };
    }
}