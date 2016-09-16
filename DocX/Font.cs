using System;

namespace Novacode
{
    /// <summary>
    /// Represents a font family
    /// </summary>
    public sealed class Font
    {
        /// <summary>
        /// Initializes a new instance of <see cref="Font" />
        /// </summary>
        /// <param name="name">The name of the font family</param>
        public Font(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentNullException(nameof(name));
            }

            Name = name;
        }

        /// <summary>
        /// The name of the font family
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Returns a string representation of an object
        /// </summary>
        /// <returns>The name of the font family</returns>
        public override string ToString()
        {
            return Name;
        }
    }
}