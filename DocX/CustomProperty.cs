using System;

namespace Novacode
{
    public class CustomProperty
    {
        /// <summary>
        /// The name of this CustomProperty.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// The value of this CustomProperty.
        /// </summary>
        public object Value { get; }

        internal string Type { get; }

        internal CustomProperty(string name, string type, string value)
        {
            object realValue;
            switch (type)
            {
                case "lpwstr": 
                {
                    realValue = value;
                    break;
                }

                case "i4":
                {
                    realValue = int.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                }

                case "r8":
                {
                    realValue = Double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                }

                case "filetime":
                {
                    realValue = DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
                    break;
                }

                case "bool":
                {
                    realValue = bool.Parse(value);
                    break;
                }

                default: throw new Exception();
            }

            Name = name;
            Type = type;
            Value = realValue;
        }

        private CustomProperty(string name, string type, object value)
        {
            Name = name;
            Type = type;
            Value = value;
        }

        /// <summary>
        /// Create a new CustomProperty to hold a string.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, string value) : this(name, "lpwstr", value as object) { }


        /// <summary>
        /// Create a new CustomProperty to hold an int.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, int value) : this(name, "i4", value) { }


        /// <summary>
        /// Create a new CustomProperty to hold a double.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, double value) : this(name, "r8", value) { }


        /// <summary>
        /// Create a new CustomProperty to hold a DateTime.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, DateTime value) : this(name, "filetime", value.ToUniversalTime()) { }

        /// <summary>
        /// Create a new CustomProperty to hold a bool.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, bool value) : this(name, "bool", value) { }
    }
}
