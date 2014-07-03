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
using VSMacros.Engines;

namespace VSMacros.Dialogs
{
    /// <summary>
    /// Interaction logic for SaveCurrentDialog.xaml
    /// </summary>
    public partial class SaveCurrentDialog : Window
    {
        public ComboBoxItem SelectedItem;
        public int SelectedShortcutNumber;
        public bool shouldRefreshFileSystem;

        public SaveCurrentDialog()
        {
            InitializeComponent();
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void customAssignButton_Click(object sender, RoutedEventArgs e)
        {
        }

        private void shortcutsComboBox_DropDownClosed(object sender, System.EventArgs e)
        {
            string selectedShortcut = this.shortcutsComboBox.Text;
            int index = 0;

            // Reset bool
            this.shouldRefreshFileSystem = false;

            if (selectedShortcut != "None")
            {
                // Get the command number into index                
                index = GetLastCharAsInt(selectedShortcut);

                // Show overwrite message if needed
                if (!string.IsNullOrEmpty(selectedShortcut))
                {
                    bool willOverwrite = Manager.Instance.Shortcuts[index] != string.Empty;

                    if (willOverwrite)
                    {
                        this.shouldRefreshFileSystem = true;
                        this.warningTextBlock.Text = VSMacros.Resources.ShortcutAlreadyUsed;
                    }
                    else
                    {
                        this.warningTextBlock.Text = string.Empty;
                    }
                }
            }

            this.SelectedShortcutNumber = index;
        }

        private int GetLastCharAsInt(string str)
        {
            int number;
            if (!int.TryParse(str[str.Length - 1].ToString(), out number))
            {
                throw new Exception("Could not retrieve command index from selection.");
            }

            return number;
        }
    }
}
