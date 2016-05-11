//-----------------------------------------------------------------------
// <copyright file="PlaybackMultipleTimesDialog.xaml.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

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

            this.Owner = Application.Current.MainWindow;

            this.IterationsTextbox.Focus();
            this.IterationsTextbox.SelectAll();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.Iterations = this.ToInt(this.IterationsTextbox.Text, true);

            if (this.Iterations == 0)
            {
                this.DialogResult = false;
            }
            else
            {
                this.DialogResult = true;
            }
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

            int number = this.ToInt(this.IterationsTextbox.Text + e.Text);

            if (e.Handled ||  number != 1)
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
            int ret = 0;

            Int32.TryParse(str, out ret);

            return ret;
        }
    }
}
