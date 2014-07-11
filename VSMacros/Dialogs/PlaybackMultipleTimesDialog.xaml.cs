//-----------------------------------------------------------------------
// <copyright file="PlaybackMultipleTimesDialog.xaml.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace VSMacros.Dialogs
{
    /// <summary>
    /// Interaction logic for PlaybackMultipleTimesDialog.xaml
    /// </summary>
    public partial class PlaybackMultipleTimesDialog : Window
    {
        public int Iterations { get; set; }

        public PlaybackMultipleTimesDialog()
        {
            this.InitializeComponent();

            this.IterationsTextbox.Focus();
            this.IterationsTextbox.SelectAll();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.ToInt(this.IterationsTextbox.Text, true) == -1)
            {
                this.DialogResult = false;
                return;
            }
            
            this.DialogResult = true;
        }

        private bool IsValid(string text)
        {
            // Regex that allows numeric input only
            Regex regex = new Regex("[0-9]+");

            int maxLength = int.MaxValue.ToString().Length;
            return regex.IsMatch(text) && this.IterationsTextbox.Text.Length < maxLength;
        }

        private void IterationsTextbox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Preventing the input event if the input is invalid will only allow the user to enter numeric values
            e.Handled = !this.IsValid(e.Text);

            if (e.Handled || this.ToInt(this.IterationsTextbox.Text + e.Text) > 1)
            {
                this.TimesLabel.Content = VSMacros.Resources.DialogTimesPlural;
            }
            else
            {
                this.TimesLabel.Content = VSMacros.Resources.DialogTimesSingular;
            } 
        }

        private int ToInt(string str, bool showMessage = false)
        {
            int ret = -1;

            try
            {
                ret = Convert.ToInt32(str);
            }
            catch (FormatException e)
            {
                if (showMessage)
                {
                    VSMacros.Engines.Manager.Instance.ShowMessageBox(e.Message);
                }
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e)) 
                { 
                    throw; 
                }
            }

            return ret;
        }
    }
}
