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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Win32;

namespace WpfTutorialSamples.Dialogs
{
	public partial class OpenFileDialogSample : Window
	{
		public OpenFileDialogSample()
		{
			InitializeComponent();
		}

		private void btnOpenFile_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			if (openFileDialog.ShowDialog() == true)
				MessageBox.Show("File path: " + openFileDialog, "My App", MessageBoxButton.OK , MessageBoxImage.Information);
				btnSaveFile.IsEnabled = true;
		}

        private void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
			MessageBox.Show("hi");
        }
    }
}
