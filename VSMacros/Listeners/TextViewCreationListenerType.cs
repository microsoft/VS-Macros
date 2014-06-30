using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Utilities;

namespace VSMacros
{
    internal static class TextViewCreationListenerClassificationDefinition
    {
        /// <summary>
        /// Defines the "TextViewCreationListener" classification type.
        /// </summary>
        [Export(typeof(ClassificationTypeDefinition))]
        [Name("TextViewCreationListener")]
        internal static ClassificationTypeDefinition TextViewCreationListenerType = null;
    }
}
