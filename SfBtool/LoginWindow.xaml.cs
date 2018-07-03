using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SfBtool
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {

        public bool Confirmed;

        public LoginWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = false;
            this.Close();
        }

        public string Login
        {
            get { return LoginTextBox.Text; }
            set { }
        }

        public string Password
        {
            get { return PasswordTextBox.Password; }
            set { }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Confirmed = true;
            this.Close();
        }
    }
}
