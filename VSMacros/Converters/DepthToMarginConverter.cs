//-----------------------------------------------------------------------
// <copyright file="DepthToMarginConverter.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.VisualStudio.PlatformUI;

namespace MicrosoftCorporation.VSMacros.Converters
{
    public class DepthToMarginConverter : ValueConverter<int, Thickness>
    {
        protected override Thickness Convert(int value, object parameter, CultureInfo culture)
        {
            return new Thickness(18 * value, 0, 0, 0);
        }
    }
}
