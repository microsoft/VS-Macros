//-----------------------------------------------------------------------
// <copyright file="SaveCurrentDialog.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Windows;
using System.Windows.Controls;

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

            this.AddItems(this.shortcutsComboBox);
        }

        private void ShortcutsComboBox_DropDownClosed(object sender, System.EventArgs e)
        {
            // Get selected number as an integer
            int selectedNumber = this.GetSelectedNumber((ComboBoxItem)this.shortcutsComboBox.SelectedItem);

            if (selectedNumber > 0 && selectedNumber <= 9)
            {
               this.WarningTextBlock.Text = this.SetWarningForOverwrite(selectedNumber);
            }

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
