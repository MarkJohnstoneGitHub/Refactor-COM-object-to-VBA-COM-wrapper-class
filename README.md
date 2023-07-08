# Refactor-COM-object-to-VBA-COM-wrapper-class

Aim: To refactor a COM object to extract a VBA COM wrapper class.  This is in regard for the requirement to construct VBA COM wrappers for classes at [DotNetLib]( https://github.com/MarkJohnstoneGitHub/DotNetLib).  As there are "numerous" properties and members for each COM object to wrap with the aim to automate.

Suggestion for RubberDuck feature [Adding refactoring of COM objects](https://github.com/rubberduck-vba/Rubberduck/discussions/6111)


From a type library obtain the COM object required to implement a VBA COM wrapper class.

- Utilizing [Rubberduck ComReflection](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.Parsing/ComReflection) COM TypeLib wrappers

For testing using [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/DotNetLib/blob/main/bin/Release/DotNetLib.tlb) for the COM object selected DateTime 

Sample output: [DateTimeWrapper.cls](https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class/blob/main/C%23/ComRefactorConsoleRDComponents/Output/DateTimeWrapper.cls)


**Development:**

Outline for implemention of refactoring a COM Object to implement a VBA class COM wrapper. 

1) Locate Com type library required and load using ComLibraryProvider class.
2) Find the Com object required by name [ComCoClass.cs](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComCoClass.cs) object.
3) Selecting various options for implementing :
- Intially the default implementation will be template of the ComInterface object required, using early binding, including descriptions and Rubberduck annotations for a predeclared class.
- Selecting various options would require a GUI etc. 
- Select which methods to implement.
- Ordering/grouping of methods/properties.
     - Ie. group as constructor  functions, properties,  methods in alphabetical order or same as Com object template.
     - Issue distinguishing between constructors and other similar factory functions that don't wish  to group as constructor overloads.
     - In Com objects created I'm placing these in the order as constructors, properties, methods.
           - Therefore possible could obtain a list  of constructor functions before the first property found.   
           - This might not be the case for all Com objects obtained to wrap. 
           - Might require an additional property for indentified constructor factory methods, which might require manual inspection to alter.
           - Note similar factory functions not to confuse with constructor factory functions.
     - Include class and member descriptions
     - Add Rubberduck annotations required. eg module and member description, hidden items, default, enumeration, predeclared class etc.
     - Option to implement as predeclared class.
     - Early or late binding implementation of wrapping the Com object in VBA. Currently early binding only.
  
If allowing for various options create a copy of the Com object ComInterface according to selections. 
i.e. The list of methods is copied according to required methods required and ordering. 
Update ComInterface.IsPreDeclared property?
To allow for sorting/grouping of methods maybe have to expose some lists eg the methods list?

4) Parse default interface for the ComCoClass object obtained to build the VBA Com wrapper class.
- Add Class Header hidden attributes [VBA Attributes](https://vbaplanet.com/attributes.php#:~:text=VBA%20code%20modules%20contain%20attributes,module%20in%20a%20text%20editor.)
  - Add module description attribute if required.
  - Add Rubberduck module description annotation if required. '@ModuleDescription("")
  - Add Rubberduck PredeclaredId annotation if required. '@PredeclaredId
  - Add in comments? i.e. Generated by etc?
  - Potential issues if class name already exists in project or reserved names?
- Declarations
   - Add Option Explicit
   - Add Implements if required. i.e. excluding the default interface.  Currently not implemented.
   - Add private variable for Com object being wrapped in the type libary.
   - EG. Private mDateTime As DotNetLib.DateTime
- Add VBA Constructors and Destructors
  - Private Sub Class_Initialize()
  - Private Sub Class_Terminate()
- Add VBA Internal Friend helper properties to access the Com object wrapped.
- Add Class members (constructors/factory methods, properties, methods)
  - Create in the order of the ComInterface template.
  - Add Rubberduck Description, DefaultMember, Enumerator annotation if required.
  - Add Method signature. Eg Public Function CreateFromDate(ByVal year as long, ByVal month as Long, ByVal day as Long) As DateTime
     - Parameters all one line or split?
     - Maybe make as a selection option?
     - Initially all one line. Potiential issue for extremely long member signatures.
     - Issue the ComInterface properties returning IDateTime how to determine if require an interface or object?
  - Add hidden method attributes required eg. method description, default property, NewEnum etc.
  - Add Com object reference being wrapped.
  - Add member End
     - Eg. End Function, End Property, End Sub
5) Write VBA class output to a file with extension ".cls" for the output path obtained. eg. DateTime.cls
6) May also require creating a static helper class for constant field values/objects i.e. DateTime.MaxValue. Eg. DateTimeStatic.cls
     - Possibly redesign DotNetLib type library moving static fields to static helper class for COM object DateTime?
     - May require option to create static helper class for field properties returning constant objects. eg. DateTime.MaxValue
7) Implement inherited interfaces, so far the DateTime example only has a default interface.
- Currently not implemented

Will require to investigate VBA wrappers of objects eg a Collection wrapper to check correct implementation.

Any custom error handling required to be done manually and/or extending the VBA Com wrapper class as required.

**Implemented July 6th, 20223**
- Obtain a type library by path.
- Obtain the ComCoClass required by name
- Created the VBA Com wrapper class for all properties and members, including Rubberduck annotations and attributes.
- Added internal helper properties to access the wrapped Com object.
- To implement:
     - For parameters and return types for implemented object for an interface use qualified name eg. DotNetLib.TimeSpan i.e. of TYPEKIND.TKIND_DISPATCH use qualified name.
     - Static helper class if required? Would require option to select which members required for static fields.
     - Eg.Public Property Get MaxValue() As DateTime should be implemented in a static DateTime helper class.
     - Inherited interfaces i.e. require Implements section and generate private VBA members including references to the COM object being wrapped.
 
- Issues
  - Member names using reserved VBA words. Eg. Date see: DotNetLib.DateTime.Date method
       - Currently low priority to fix, manually fix by renaming member to DateComponent
       - EG. ``` Public Property Get Date() As DateTime ```
  - Parameter names using VBA reserved words.
       -  DotNetLib.TimeSpan parameter input is a reserved word
       -  ```Public Function Parse2(ByVal input As String, ByRef formatProvider As IFormatProvider) As TimeSpan  ```
  - When wrapping the COM object in members where parameters are the object being wrapped. (Fixed)

Expected Ouput

```
Public Function Compare(ByRef t1 As DateTime, ByRef t2 As DateTime) As Long
   Compare = this.DotNetLibDateTime.Compare(t1.ComObject, t2.ComObject)
End Function
```
  - For parameters and return types for references require qualified name.  eg DotNetLib.TimeSpan 
  - Parameters for object being wrapped displayed as interface of the object. Eg. ITimeSpan
     - Issue addressed, require to preferrable use qualified name
  - Return type not converted from interface to object
       - Issue addressed, require to preferrable use qualified name
       - ```Public Property Get TimeOfDay() As TimeSpan ```
       - Expected output ```Public Property Get TimeOfDay() As DotNetLib.TimeSpan ```
  - Return type is an array (Fixed)
       - ```Public Function GetDateTimeFormats() As String ```
       - Expected output ```Public Function GetDateTimeFormats() As String() ```
  - Potiential issuse with member TryParse and TryParse2
       - ``` bool TryParse(string s, out DateTime result); ```
       - require to pass out by reference DateTime result require to check correctly implemented.
  - Parameter is an array. (Fixed)
       - Fixed type library DotNetLib.DateTime.ParseExact3 member for marshalling array
       - Issue with VBA parameter not converted to an array. (Fixed)
       - ``` Public Function ParseExact3(ByVal s As String, ByRef formats As String, ByRef provider As IFormatProvider, ByVal style As DateTimeStyles) As DateTime ```
       - Expected output ``` Public Function ParseExact3(ByVal s As String, ByRef formats() As String, ByRef provider As IFormatProvider, ByVal style As DateTimeStyles) As DateTime ```
- [ComTypeName](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComTypeName.cs)
     - Possible requires refactoring with a GUID property and TYPEKIND replacing GUID property for each TYPEKIND?
     - See [ComParameter](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComParameter.cs)
  
Overall preforming resonable well with outstanding isssues: 
-   Member names using reserved words. Eg. ``` Public Property Get Date() As DateTimeWrapper ```
-   To implement, Implements section for interfaces implemented.


*Implementing linking an interface to its implementations*
- To implement associating an interface to its default implementation or implementations requires a list keyed by interface GUID  and implementation GUID?
- Require parent GUID of the type library?
- [how-to-get-type-library-from-progid-or-clsid-without-loading-the-com-object](https://stackoverflow.com/questions/12975329/how-to-get-type-library-from-progid-or-clsid-without-loading-the-com-object)
- I.e. An interface may have many implementations.
- Where an interface or implementation may be located in an external type library.
- Therefore require a list of references/dependencies where a type library required to be located by GUID/CLSID?

i.e. For the DotNetLib.tlb example IFormatProvider is referenced from \Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb type library.
In this example is the interface is required as there are multiple implementations. 

```
Public Function ToString4(ByVal format As String, ByRef provider As IFormatProvider) As String
```

If there is only one implementation found then use that implementation. 
EG  ```Public Property Get TimeOfDay() As ITimeSpan ```
Expect output  ```Public Property Get TimeOfDay() As DotNetLib.TimeSpan ```
Require to search known types may require searching dependent external type libraries?.



**Utilize [RubberDuck Com Management](https://github.com/rubberduck-vba/Rubberduck](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.VBEEditor/ComManagement))**
Items of interest

- [ComReflection](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.Parsing/ComReflection)
    - [ComLibraryProvider](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComLibraryProvider.cs)
        - public ITypeLib LoadTypeLibrary(string libraryPath)
    - [is-there-a-way-to-view-com-entries-by-traversing-a-tlb-file-in-net](https://stackoverflow.com/questions/43875454/is-there-a-way-to-view-com-entries-by-traversing-a-tlb-file-in-net) 

- [TypeLibs](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.VBEEditor/ComManagement/TypeLibs)

- [TypeLibWrapper.cs](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.VBEEditor/ComManagement/TypeLibs/TypeLibWrapper.cs)
    - constructor  internal TypeInfoWrapper(ComTypes.ITypeInfo rawTypeInfo)

- [A dumb container of ITypeInfos.](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.VBEEditor/ComManagement/TypeLibs/Utility/SimpleCustomTypeLibrary.cs)

- [LibraryReferencedDeclarationsCollector](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/LibraryReferencedDeclarationsCollector.cs)
  - IReadOnlyCollection<Declaration> CollectedDeclarations(ReferenceInfo reference)


