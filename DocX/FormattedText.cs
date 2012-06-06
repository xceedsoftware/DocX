using System;

namespace Novacode
{
    public class FormattedText: IComparable
    {
        public FormattedText()
        {
        
        }

        public int index;
        public string text;
        public Formatting formatting;

        public int CompareTo(object obj)
        {
            FormattedText other = (FormattedText)obj;
            FormattedText tf = this;

            if (other.formatting == null || tf.formatting == null)
                return -1;

            return tf.formatting.CompareTo(other.formatting);   
        }
    }
}
