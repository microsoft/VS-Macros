//-----------------------------------------------------------------------
// <copyright file="AssignShortcutDialog.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.VisualStudio;
using VSMacros.Engines;

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

            // Set default values for public members
            this.SelectedShortcutNumber = 0;
            this.ShouldRefreshFileSystem = false;

            // Retrieve old shortcut
            this.oldShortcut = (MacrosControl.Current.SelectedNode).FormattedShortcut;
            this.oldShortcutNumber = (MacrosControl.Current.SelectedNode).Shortcut;

            // Set the text to the previous shortcut, if it exists
            if (!string.IsNullOrEmpty(this.oldShortcut) && this.oldShortcut.Length >= 3)
            {
                // this.oldShortcut has format "(ALT+Q,#)" -> remove the parentheses
                this.shortcutsComboBox.Text = this.oldShortcut.Substring(1, this.oldShortcut.Length - 2);
                
                // last char should be the command number
                this.SelectedShortcutNumber = this.oldShortcutNumber;
            }
        }

        private void ShortcutsComboBox_DropDownClosed(object sender, System.EventArgs e)
        {
            // Get selected number as an integer
            int selectedNumber = this.GetSelectedNumber((ComboBoxItem)this.shortcutsComboBox.SelectedItem);
            
            // Temporary variables
            bool shouldRefresh = false;
            string warningText = string.Empty;

            // If selected number is between 1 and 9
            if (selectedNumber > 0 && selectedNumber <= 9)
            {
                // Show overwrite message if needed
                if (selectedNumber != this.oldShortcutNumber)
                {
                    this.CheckOverwrite(selectedNumber, out shouldRefresh, out warningText);
                }
            }

            this.ShouldRefreshFileSystem = shouldRefresh;
            this.WarningTextBlock.Text = warningText;
            this.SelectedShortcutNumber = selectedNumber;
        }
    }
}
