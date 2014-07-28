//-----------------------------------------------------------------------
// <copyright file="AssignShortcutDialog.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Windows.Controls;

namespace VSMacros.Dialogs
{
    /// <summary>
    /// Interaction logic for AssignShortcutDialog.xaml
    /// </summary>
    public partial class AssignShortcutDialog : MacroDialog
    {
        private string oldShortcut;
        private int oldShortcutNumber;

        public AssignShortcutDialog()
        {
            this.InitializeComponent();

            this.AddItems(this.shortcutsComboBox);

            // Set default values for public members
            this.SelectedShortcutNumber = 0;

            // Retrieve old shortcut
            this.oldShortcut = (MacrosControl.Current.SelectedNode).FormattedShortcut;
            this.oldShortcutNumber = (MacrosControl.Current.SelectedNode).Shortcut;

            // Set the text to the previous shortcut, if it exists
            if (!string.IsNullOrEmpty(this.oldShortcut) && this.oldShortcut.Length >= 3)
            {
                // this.oldShortcut has format "(CTRL+M,#)" -> remove the parentheses
                this.shortcutsComboBox.SelectedIndex = oldShortcutNumber - 1;
                
                // last char should be the command number
                this.SelectedShortcutNumber = this.oldShortcutNumber;
            }
        }

        private void ShortcutsComboBox_DropDownClosed(object sender, System.EventArgs e)
        {
            // Get selected number as an integer
            int selectedNumber = this.GetSelectedNumber((ComboBoxItem)this.shortcutsComboBox.SelectedItem);

            // If selected number is between 1 and 9
            if (selectedNumber > 0 && selectedNumber <= 9)
            {
                // Show overwrite message if needed
                if (selectedNumber != this.oldShortcutNumber)
                {
                    this.WarningTextBlock.Text = this.SetWarningForOverwrite(selectedNumber);
                }
            }

            this.SelectedShortcutNumber = selectedNumber;
        }
    }
}
