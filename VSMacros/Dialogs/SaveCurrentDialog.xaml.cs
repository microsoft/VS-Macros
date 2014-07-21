//-----------------------------------------------------------------------
// <copyright file="SaveCurrentDialog.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

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
    public partial class SaveCurrentDialog : MacroDialog
    {
        public SaveCurrentDialog()
        {
            this.InitializeComponent();
        }

        private void ShortcutsComboBox_DropDownClosed(object sender, System.EventArgs e)
        {
            // Get selected number as an integer
            int selectedNumber = this.GetSelectedNumber((ComboBoxItem)this.shortcutsComboBox.SelectedItem);

            // Temporary variables
            bool shouldRefresh = false;
            string warningText = string.Empty;

            if (selectedNumber > 0 && selectedNumber <= 9)
            {
                this.CheckOverwrite(selectedNumber, out shouldRefresh, out warningText);
            }

            this.ShouldRefreshFileSystem = shouldRefresh;
            this.WarningTextBlock.Text = warningText;
            this.SelectedShortcutNumber = selectedNumber;
        }

        private void MacroName_Loaded(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;

            textBox.SelectAll();
            textBox.Focus();
        }
    }
}
