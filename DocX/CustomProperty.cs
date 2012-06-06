using System;

namespace Novacode
{
    public class CustomProperty
    {
        private string name;
        private object value;
        private string type;

        /// <summary>
        /// The name of this CustomProperty.
        /// </summary>
        public string Name { get { return name;} }

        /// <summary>
        /// The value of this CustomProperty.
        /// </summary>
        public object Value { get { return value; } }

        internal string Type { get { return type; } }

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

            this.name = name;
            this.type = type;
            this.value = realValue;
        }

        private CustomProperty(string name, string type, object value)
        {
            this.name = name;
            this.type = type;
            this.value = value;
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
        public CustomProperty(string name, int value) : this(name, "i4", value as object) { }


        /// <summary>
        /// Create a new CustomProperty to hold a double.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, double value) : this(name, "r8", value as object) { }


        /// <summary>
        /// Create a new CustomProperty to hold a DateTime.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, DateTime value) : this(name, "filetime", value.ToUniversalTime() as object) { }

        /// <summary>
        /// Create a new CustomProperty to hold a bool.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, bool value) : this(name, "bool", value as object) { }
    }
}
