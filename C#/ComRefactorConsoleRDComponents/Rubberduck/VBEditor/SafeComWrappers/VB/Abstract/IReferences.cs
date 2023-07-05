using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IReferences : ISafeEventedComWrapper, IComCollection<IReference>, IEquatable<IReferences>
    {
        event EventHandler<ReferenceEventArgs> ItemAdded;
        event EventHandler<ReferenceEventArgs> ItemRemoved;

        //TODO : RD Commented out
        //IVBE VBE { get; }
        //IVBProject _parentMember { get; }

        IReference AddFromGuid(string guid, int major, int minor);
        IReference AddFromFile(string path);
        void Remove(IReference reference);
    }
}