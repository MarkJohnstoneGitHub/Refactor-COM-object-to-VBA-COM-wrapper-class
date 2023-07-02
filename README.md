# Refactor-COM-object-to-VBA-COM-wrapper-class

Aim: To refactor a COM object to extract a VBA COM wrapper class.  This is in regard for the requirement to construct VBA COM wrappers for classes at [DotNetLib]( https://github.com/MarkJohnstoneGitHub/DotNetLib).  As there are "numerous" properties and methods for each COM object to wrap aim to automate.

Suggestion for RubberDuck feature [Adding refactoring of COM objects](https://github.com/rubberduck-vba/Rubberduck/discussions/6111)


From the type library info for the required class obtain the class template to extract to a VBA COM wrapper classs.

- Utilizing [RubberDuck](https://github.com/rubberduck-vba/Rubberduck) COM typelib wrappers

For testing using [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/DotNetLib/blob/main/bin/Release/DotNetLib.tlb) for the Com object DateTime 
Sample output [DateTime.cls](https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class/blob/main/C%23/ComRefactorConsoleRDComponents/Output/DateTime.cls)


**Development:**

Outline for implemention of refactoring a Com Object to implement a VBA class COM wrapper. 

1) Locate Com type library required and load using ComLibraryProvider class.
2) Find the Com object required by name [ComCoClass.cs](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComCoClass.cs) object.
3) Selecting various options for implementing :
- Intially the default implementation will be template of the ComInterface object required, using early binding, including descriptions and Rubberduck annotations for a predeclared class.
- Selecting various options would require a GUI etc. 
- Select which methods to implement.
- Ordering/grouping of methods/properties.
     - Ie. group as constructor  functions, properties,  methods in alpabetical order or same as Com object template.
     - Issue distinguishing between constructors and other similar factory functions that don't wish  to group as constructor overloads.
     - In Com objects created I'm placing these in the order as constructors, properties, methods.
           - Therefore possible could obtain a list  of constructor functions before the first property found.   
           - This might not be the case for all Com objects obtained to wrap. 
           - Might require an additional property for indentified constructor factory methods, which might require manual inspection to alter.
           - Note similar factory functions not to confuse with constructor factory functions.
     - Include class and method descriptions
     - Add Rubberduck annotations required. eg classs and  method description, hidden items, default, enumeration, predeclared class etc.
     - Option to implement as predeclared class.
     - Early or late binding implementation of wrapping the Com object in VBA.
- May require option to create static helper class for field properties returning constant objects. eg. DateTime.MaxValue
  
If allowing for various options create a copy of the Com object ComInterface according to selections. 
i.e. The list of methods is copied according to required methods required and ordering. 
Update ComInterface.IsPreDeclared property?
To allow for sorting/grouping of methods maybe have to expose some lists eg the methods list?

4) Parse default interface for the ComCoClass object obtained to build the VBA Com wrapper class.
- Add Class Header hidden attributes [VBA Attributes](https://vbaplanet.com/attributes.php#:~:text=VBA%20code%20modules%20contain%20attributes,module%20in%20a%20text%20editor.)
  - Add Class description if required.
  - Add Rubberduck module description annotation  if required. '@ModuleDescription("")
  - Add Rubberduck PredeclaredId annotation if required. '@PredeclaredId
  - Add in comments? i.e. Generated by etc?
  - Potential issues if class name already exists in project or reserved names?
- Declarations
   - Add Option Explicit
   - Add private variable for Com object being wrapped in the type libary.
   - EG. Private mDateTime As DotNetLib.DateTime
- Add VBA Constructors and Destructors
  - Private Sub Class_Initialize()
  - Private Sub Class_Terminate()
- Add VBA Internal helper properties to access the Com object wrapped.
- Add Class members (constructors/factory methods, properties, methods)
  - Create in the order of the ComInterface template.
  - Add Rubberduck Description, DefaultMember, Enumerator annotation if required.
  - Add Method signature. Eg Public Function CreateFromDate(ByVal year as long, ByVal month as Long, ByVal day as Long) As DateTime
     - Parameters all one line or split? Maybe make as a selection option? Initially all one line. Potientially issue for extremely long method signatures.
     - Issue the ComInterface properties returning IDateTime how to determine if require an interface or object?
  - Add hidden method attributes required eg. method description, default property, NewEnum etc.
  - Add Com object reference being wrapped.
  - Add Method End
     - Eg.  End Function, End Property, End Sub
5) Write VBA class output to a file with extension ".cls" for the output path obtained. eg. DateTime.cls
6) May also require creating a static helper class for constant field values/objects. eg  DateTimeStatic.cls
7) To implement inherited interfaces, so far the DateTime example only has a default interface.

Will require to investigate VBA wrappers of objects eg a Collection wrapper to check correct implementation.

Any custom error handling required to be done manually and/or extending the VBA Com wrapper class as required.

**Implemented July 2nd 20223**
- Obtain a type library by path.
- Obtain the ComCoClass required by name
- Created the VBA Com wrapper class for all properties and members, including Rubberduck annotations and attributes.
- Added internal helper properties to access the wrapped Com object.
- To implement:
     - Covert if required parameters and return types to implemented object eg ITimeSpan
     - Would require finding it's implementating object i.e. From known types or maybe require searching an external typelib from GUID?
     - Static helper class if required? Would require option to select which members required for static fields.
     - Eg.Public Property Get MaxValue() As DateTime should be implemented in a static DateTime helper class.
     - Inherited interfaces i.e. require Implements section and generate private VBA members including references to the COM object being wrapped.
 
- Issues
  - Member names using reserved VBA words. Eg. Date see: DotNetLib.DateTime.Date method
  - Paramaters for object being wrapped displayed as interface of the object. eg eg ITimeSpan

Overall preforming resonable well with some outstanding issues regarding parameters and return types required to convert the interface to the object required.
This issue may require some restructing to search dependent external type libraries. Also issue member names using reserved words.

i.e. For the DotNetLib.tlb example IFormatProvider is referenced from \Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb type library.
In this example is the interface is required as there are multiple implementations. 

```
Public Function ToString4(ByVal format As String, ByRef provider As IFormatProvider) As String
```

If there is only one implementation found then use that implementation. 
EG  ```Public Property Get TimeOfDay() As ITimeSpan ```
Expect output  ```Public Property Get TimeOfDay() As TimeSpan ```
Require to search known types may require searching dependent external type libraries? Currently an outstanding issue, quick fix perform manualy.


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

- [ComReflection](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.Parsing/ComReflection)
- [ComLibraryProvider](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComLibraryProvider.cs)
        - public ITypeLib LoadTypeLibrary(string libraryPath)
  - [is-there-a-way-to-view-com-entries-by-traversing-a-tlb-file-in-net](https://stackoverflow.com/questions/43875454/is-there-a-way-to-view-com-entries-by-traversing-a-tlb-file-in-net) 


