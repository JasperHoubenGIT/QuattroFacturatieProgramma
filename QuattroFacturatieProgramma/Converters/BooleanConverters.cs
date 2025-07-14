using System.Globalization;

namespace QuattroFacturatieProgramma.Converters
{
    /// <summary>
    /// Inverteert een boolean waarde (true → false, false → true)
    /// </summary>
    public class InvertedBoolConverter : IValueConverter
    {
        public static readonly InvertedBoolConverter Instance = new();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
                return !boolValue;

            return true; // Default als niet bool
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
                return !boolValue;

            return false;
        }
    }

    /// <summary>
    /// Converteert percentage (0-100) naar progress waarde (0.0-1.0)
    /// </summary>
    public class PercentageToProgressConverter : IValueConverter
    {
        public static readonly PercentageToProgressConverter Instance = new();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double percentage)
            {
                return percentage / 100.0; // 50% wordt 0.5
            }

            return 0.0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double progress)
            {
                return progress * 100.0; // 0.5 wordt 50%
            }

            return 0.0;
        }
    }
}