using System.Windows;

namespace Group4338
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var window = new _4338_Чибышев();
            window.Show();
        }
    }
}