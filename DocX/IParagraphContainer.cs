using System.Collections.ObjectModel;

namespace Novacode
{
    public interface IParagraphContainer
    {
        ReadOnlyCollection<Paragraph> Paragraphs { get; }
    }
}