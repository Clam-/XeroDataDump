﻿using System;
using System.Globalization;
using System.Windows.Data;

namespace XeroDataDump
{
	// Based on http://stackoverflow.com/a/18617513
	[ValueConversion(typeof(Enum), typeof(bool))]
	public class EnumToBoolConverter : IValueConverter
	{
		public EnumToBoolConverter() { }

		public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
		{
			if (value == null || parameter == null) return false;
			string enumValue = value.ToString();
			string targetValue = parameter.ToString();
			bool outputValue = enumValue.Equals(targetValue, StringComparison.InvariantCultureIgnoreCase);
			return outputValue;
		}

		public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
		{
			if (value == null || parameter == null) return null;
			bool useValue = (bool)value;
			string targetValue = parameter.ToString();
			if (useValue) return Enum.Parse(targetType, targetValue);
			return Binding.DoNothing;
		}
	}
}
