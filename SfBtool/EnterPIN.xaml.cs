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
    /// Interaction logic for EnterPIN.xaml
    /// </summary>
    public partial class EnterPIN : Window
    {
        public EnterPIN()
        {
            InitializeComponent();
        }

        public bool PinEntered;

        public string Pin
        {
            get { return PinTextBox.Text; }
            set { }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
                if (!String.IsNullOrEmpty(PinTextBox.Text))

                {
                    PinEntered = true;
                    this.Close();
                }

                else
                {
                    ErrorTextBlock.Text = "Enter Pin";
                }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            PinEntered = false;
            this.Close();
        }
    }
}
