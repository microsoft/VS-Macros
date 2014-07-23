//-----------------------------------------------------------------------
// <copyright file="MacroDialog.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.VisualStudio;
using VSMacros;
using VSMacros.Engines;

namespace VSMacros.Dialogs
{
    public class MacroDialog : Window
    {
        public int SelectedShortcutNumber { get; set; }

        public bool ShouldRefreshFileSystem { get; set; }

        public MacroDialog()
        {
            // Set default values for public members
            this.SelectedShortcutNumber = 0;
            this.ShouldRefreshFileSystem = false;
        }

        protected int GetSelectedNumber(ComboBoxItem item)
        {
            if (item == null)
            {
                return 0;
            }

            return item.Tag.ToString()[0] - '0';
        }

        protected void CheckOverwrite(int selectedNumber, out bool shouldRefresh, out string warningText)
        {
            bool willOverwrite = Manager.Shortcuts[selectedNumber] != string.Empty;

            if (willOverwrite)
            {
               shouldRefresh = true;
               warningText = VSMacros.Resources.DialogShortcutAlreadyUsed;
            }
            else
            {
                shouldRefresh = false;
                warningText = string.Empty;
            }
        }
        protected void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Don't accept the dialog box if there is invalid data
            if (!this.IsValid(this)) return;

            this.DialogResult = true;
        }

        protected void CustomAssignButton_Click(object sender, RoutedEventArgs e)
        {
            string targetGUID = "BAFF6A1A-0CF2-11D1-8C8D-0000F87570EE";
            var command = new CommandID(VSConstants.GUID_VSStandardCommandSet97, VSConstants.cmdidToolsOptions);
            var mcs = ((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(IMenuCommandService)) as MenuCommandService;
            mcs.GlobalInvoke(command, targetGUID);
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
                    if (!this.IsValid((DependencyObject)subnode))
                    { 
                        return false;
                    }
                }
            }

            // All dependency objects are valid 
            return true;
        }
    }
}
