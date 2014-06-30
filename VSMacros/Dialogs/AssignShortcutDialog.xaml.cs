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
        public ComboBoxItem SelectedItem;
        public int SelectedShortcutNumber;
        public bool shouldRefreshFileSystem;

        private string oldShortcut;

        public AssignShortcutDialog()
        {
            InitializeComponent();

            // Set default values for public members
            this.SelectedShortcutNumber = 0;
            this.shouldRefreshFileSystem = false;

            this.oldShortcut = (MacrosControl.Current.SelectedNode).Shortcut;

            // Set the text to the previous shortcut, if it exists
            if (!string.IsNullOrEmpty(oldShortcut))
            {
                this.shortcutsComboBox.Text = oldShortcut.Substring(1, oldShortcut.Length - 2);
                this.SelectedShortcutNumber = GetLastCharAsInt(this.shortcutsComboBox.Text);
            }
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            // Don't accept the dialog box if there is invalid data
            if (!IsValid(this)) return;

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
                    if (IsValid((DependencyObject)subnode) == false) return false;
                }
            }

            // All dependency objects are valid 
            return true;
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
                if (selectedShortcut != this.oldShortcut && !string.IsNullOrEmpty(selectedShortcut))
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
