using System.ComponentModel.Composition;
using System.Windows.Media;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Utilities;

namespace VSMacros
{
    #region Format definition
    /// <summary>
    /// Defines an editor format for the TextViewCreationListener type that has a purple background
    /// and is underlined.
    /// </summary>
    [Export(typeof(EditorFormatDefinition))]
    [ClassificationType(ClassificationTypeNames = "TextViewCreationListener")]
    [Name("TextViewCreationListener")]
    [UserVisible(true)] //this should be visible to the end user
    [Order(Before = Priority.Default)] //set the priority to be after the default classifiers
    internal sealed class TextViewCreationListenerFormat : ClassificationFormatDefinition
    {
        /// <summary>
        /// Defines the visual format for the "TextViewCreationListener" classification type
        /// </summary>
        public TextViewCreationListenerFormat()
        {
            this.DisplayName = "TextViewCreationListener"; //human readable version of the name
            this.BackgroundColor = Colors.BlueViolet;
            this.TextDecorations = System.Windows.TextDecorations.Underline;
        }
    }
    #endregion //Format definition
}
