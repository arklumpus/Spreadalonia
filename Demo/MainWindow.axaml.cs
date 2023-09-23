using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace Demo
{
    public class MainWindow : Window
    {
        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public MainWindow()
        {
            InitializeComponent();
        }
    }
}