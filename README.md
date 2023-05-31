# Refactor-COM-object-to-VBA-COM-wrapper-class

Aim: To refactor COM object to extract VBA COM wrapper class.  This is in regard for the requirement to construct COM wrappers for classes at [DotNetLib]( https://github.com/MarkJohnstoneGitHub/DotNetLib).  As there are "numerous" properties and methods for each COM object to wrap aim to automate.

Suggestion for RubberDuck feature [Adding refactoring of COM objects](https://github.com/rubberduck-vba/Rubberduck/discussions/6111)

First stage obtaining the COM TypeLib info for an COM object.

Investigating methods in VBA and/or C# to obtain the type library info.

From the type library info for the required class obtain the class template to extract to a VBA COM wrapper classs.


Issues
If issues with the TLI reference in [COM Refactoring.accdb](https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class/blob/main/COM%20Refactoring.accdb) see [tlbinf32.dll in a 64bits .Net application](https://stackoverflow.com/questions/42569377/tlbinf32-dll-in-a-64bits-net-application/42581513#42581513)

