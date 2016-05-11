//-----------------------------------------------------------------------
// <copyright file="DepthToMarginConverter.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Globalization;
using System.Windows;
using Microsoft.VisualStudio.PlatformUI;

namespace MicrosoftCorporation.VSMacros.Converters
{
    public class DepthToMarginConverter : ValueConverter<int, Thickness>
    {
        protected override Thickness Convert(int depth, object parameter, CultureInfo culture)
        {
            // Indent from the parent to the child in the TreeView
            int indent = 18;

            return new Thickness(indent * (depth - 1), 0, 0, 0);
        }
    }
}
