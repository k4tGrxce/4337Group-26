using System.Windows;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var wind = new _4337_Соловьев();
            wind.ShowDialog();
        }
    }
}