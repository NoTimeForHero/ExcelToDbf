using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace ExcelToDbf.Utils.Converters
{
    internal class VisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Visibility onTrue = Visibility.Visible;
            Visibility onFalse = Visibility.Collapsed;

            if (parameter is bool boolParam && boolParam)
            {
                onTrue = Visibility.Collapsed;
                onFalse = Visibility.Visible;
            }

            if (value is string strValue) return strValue.Length > 0 ? onTrue : onFalse;
            if (value is bool boolValue) return boolValue ? onTrue : onFalse;
            return value != null ? onTrue : onFalse;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
