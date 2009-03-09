using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Novacode
{
    /// <summary>
    /// Paragraph edit types
    /// </summary>
    public enum EditType
    {
        /// <summary>
        /// A ins is a tracked insertion
        /// </summary>
        ins,
        /// <summary>
        /// A del is  tracked deletion
        /// </summary>
        del
    }

    /// <summary>
    /// Custom property types.
    /// </summary>
    public enum CustomPropertyType
    {
        /// <summary>
        /// System.String
        /// </summary>
        Text,
        /// <summary>
        /// System.DateTime
        /// </summary>
        Date,
        /// <summary>
        /// System.Int32
        /// </summary>
        NumberInteger,
        /// <summary>
        /// System.Double
        /// </summary>
        NumberDecimal,
        /// <summary>
        /// System.Boolean
        /// </summary>
        YesOrNo
    }

    /// <summary>
    /// Text types in a Run
    /// </summary>
    public enum RunTextType
    {
        /// <summary>
        /// System.String
        /// </summary>
        Text,
        /// <summary>
        /// System.String
        /// </summary>
        DelText,
    }
}
