# Refactor-COM-object-to-VBA-COM-wrapper-class

Aim: To refactor COM object to extract VBA COM wrapper class.  This is in regard for the requirement to construct VBA COM wrappers for classes at [DotNetLib]( https://github.com/MarkJohnstoneGitHub/DotNetLib).  As there are "numerous" properties and methods for each COM object to wrap aim to automate.

Suggestion for RubberDuck feature [Adding refactoring of COM objects](https://github.com/rubberduck-vba/Rubberduck/discussions/6111)

First stage obtaining the COM TypeLib info for an COM object.

Investigating methods in VBA and/or C# to obtain the type library info.

From the type library info for the required class obtain the class template to extract to a VBA COM wrapper classs.

- Preferrable utilize [RubberDuck](https://github.com/rubberduck-vba/Rubberduck) COM typelib wrappers
- Alternative using [twinBasic](https://github.com/twinbasic/twinbasic) addin for VBA could parse a reference pseudocode (From my understanding is based on the RD COM typelib wrappers and handlers). See [latest twinBasic IDE and you can see pseudocode](https://github.com/rubberduck-vba/Rubberduck/discussions/6111#discussioncomment-6041980)

Issues
If issues with the TLI reference in [COM Refactoring.accdb](https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class/blob/main/COM%20Refactoring.accdb) see [tlbinf32.dll in a 64bits .Net application](https://stackoverflow.com/questions/42569377/tlbinf32-dll-in-a-64bits-net-application/42581513#42581513).


Development

Currrently investigation stage brainstorming ideas.
