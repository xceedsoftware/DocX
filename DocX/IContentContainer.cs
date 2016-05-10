using System.Collections.ObjectModel;

namespace Novacode
{
    interface IContentContainer
    {
        ReadOnlyCollection<Content> Paragraphs { get; }
    }
}
