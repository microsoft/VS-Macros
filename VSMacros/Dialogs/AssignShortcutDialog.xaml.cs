//-----------------------------------------------------------------------
// <copyright file="AssignShortcutDialog.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using VSMacros.Engines;

namespace VSMacros.Dialogs
{
    /// <summary>
    /// Interaction logic for AssignShortcutDialog.xaml
    /// </summary>
    public partial class AssignShortcutDialog : Window
    {
        public ComboBoxItem SelectedItem { get; set; }
        public int SelectedShortcutNumber { get; set; }
        public bool ShouldRefreshFileSystem { get; set; }

        private string oldShortcut;

        public AssignShortcutDialog()
        {
            this.InitializeComponent();

            // Set the text of the fields
            this.AssignToLabel.Content = VSMacros.Resources.DialogAssignTo;
            this.shortcutsComboBox.Text = VSMacros.Resources.DialogShortcutSelectionText;
            this.CustomShortcutButton.Content = VSMacros.Resources.DialogCustomShortcut;

            // Set default values for public members
            this.SelectedShortcutNumber = 0;
            this.ShouldRefreshFileSystem = false;

            this.oldShortcut = (MacrosControl.Current.SelectedNode).Shortcut;

            // Set the text to the previous shortcut, if it exists
            if (!string.IsNullOrEmpty(this.oldShortcut) && this.oldShortcut.Length >= 3)
            {
                // this.oldShortcut has format "(CTRL+M,#)" -> remove the parentheses
                this.shortcutsComboBox.Text = this.oldShortcut.Substring(1, this.oldShortcut.Length - 2);
                
                // last char should be the command number
                this.SelectedShortcutNumber = this.GetLastCharAsInt(this.shortcutsComboBox.Text);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Don't accept the dialog box if there is invalid data
            if (!this.IsValid(this)) return;

            this.DialogResult = true;
        }

        private bool IsValid(DependencyObject node)
        {
            // Check if dependency object was passed 
            if (node != null)
            {
                // Check if dependency object is valid. 
                // NOTE: Validation.GetHasError works for controls that have validation rules attached  
                bool isValid = !Validation.GetHasError(node);
                if (!isValid)
                {
                    // If the dependency object is invalid, and it can receive the focus, 
                    // set the focus 
                    if (node is IInputElement) Keyboard.Focus((IInputElement)node);
                    return false;
                }
            }

            // If this dependency object is valid, check all child dependency objects 
            foreach (object subnode in LogicalTreeHelper.GetChildren(node))
            {
                if (subnode is DependencyObject)
                {
                    // If a child dependency object is invalid, return false immediately, 
                    // otherwise keep checking 
                    if (!this.IsValid((DependencyObject)subnode)) { return false; }
                }
            }

            // All dependency objects are valid 
            return true;
        }

        private void CustomAssignButton_Click(object sender, RoutedEventArgs e)
        {
        }

        private void ShortcutsComboBox_DropDownClosed(object sender, System.EventArgs e)
        {
            string selectedShortcut = this.shortcutsComboBox.Text;
            int index = 0;
            
            // Reset bool
            this.ShouldRefreshFileSystem = false;

            if (selectedShortcut != "None")
            {
                // Get the command number into index                
                index = this.GetLastCharAsInt(selectedShortcut);

                // Show overwrite message if needed
                if (selectedShortcut != this.oldShortcut && !string.IsNullOrEmpty(selectedShortcut))
                {
                    bool willOverwrite = Manager.Instance.Shortcuts[index] != string.Empty;

                    if (willOverwrite)
                    {
                        this.ShouldRefreshFileSystem = true;
                        this.warningTextBlock.Text = VSMacros.Resources.DialogShortcutAlreadyUsed;
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
                throw new FormatException(VSMacros.Resources.DialogCannotRetrieveCommandIndex);
            }

            return number;
        }
    }
}
