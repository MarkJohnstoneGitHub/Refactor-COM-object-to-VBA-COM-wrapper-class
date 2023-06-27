using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Claims;
using System.Windows.Controls;
using System.Xml.Linq;

namespace ComRefactorConsole.ComRefactor.Notes
{
    internal class MyNotes
    {
    }
}



//Class Description
//Documentation.DocString

//Methods see Members
//Get list of members which is not restricted
//IsRestricted = false

//Documentation.DocString

//Method Signature
//AsTypeName -> _codeMethod -> MemberDeclartion
//"Public Function CreateFromTicks As IDateTime"

//RD Annotation "'Description("" & Documentation.DocString & ")"
//Method Signiture "Public" & " " & ("Sub" or "Function") &" " & MemberName "(" &
//GetMethodParameterSigniture separated by commas & ")"

//If Function "AS " & ReturnType

//Or list parameters on separate lines?


//Next line

//"("
//Insert VBA Attributes if required
//i.e. for description.
//default item
//enumeration
//	If object "Set " & MemberName & " = " & comObjectWrapperName & "." & MemberName & "(" & GetParameters & ")"
  	
//")"


//Also Properties
//Get
//Set
//Let
//Public Property  Get XYZ


// A bit overwhelmed following Rubberduck model for implementing a class from an interface

// Aim is converting an objects type information to another.
// E.G From Com object to VBA class/interface or potientially any other language.
// or VBA object to another VBA object i.e. VBA interface to class etc.

// An ICodeBuilder is the mapping from a source object informaation to target object implementation.
// Could potientially map property/methods to different names? 


//ICodeBuilder
//Create(source ITypeInfo, target ITypeInfo)


//where source could be a VBA ITypeInfo or Com ITypeInfo
// i.e.  The Com object information maps from another language. 


//Thefore the Com ojbect information obtained from a ComTypeLib requires to be compatible with an VBA object type information
// I.e Both implement a common interface. i.e. call it  ITypeInfo or IType
// Is  ComProject is a type of ComTypeLib? i.e. could be for a VBA Project or Com type library or C# project?
// Simlar to the VBA Extensibility adding code to a project can you do the same for other languages eg C# project?



//Rubberduck.Parsing.ComReflection.LibraryReferencedDeclarationsCollector : ReferencedDeclarationsCollectorBase



//Rubberduck.Parsing.ComReflection.IComLibraryProvider
//namespace Rubberduck.Parsing.ComReflection
//{
//    public interface IComLibraryProvider
//    {
//        ITypeLib LoadTypeLibrary(string libraryPath);
//        IComDocumentation GetComDocumentation(ITypeLib typelib);
//        ReferenceInfo GetReferenceInfo(ITypeLib typelib, string name, string path);
//    }
//}


//namespace Rubberduck.Parsing.ComReflection
//{
//    public interface IDeclarationsFromComProjectLoader
//    {
//        IReadOnlyCollection<Declaration> LoadDeclarations(ComProject type, string projectId = null);
//    }
//}


// Rubberduck.Refactorings.RefactoringBase
//namespace Rubberduck.Refactorings
//{
//    public abstract class RefactoringBase : IRefactoring
//    {
//        protected readonly ISelectionProvider SelectionProvider;

//        protected RefactoringBase(ISelectionProvider selectionProvider)
//        {
//            SelectionProvider = selectionProvider;
//        }

//        public virtual void Refactor()
//        {
//            var activeSelection = SelectionProvider.ActiveSelection();
//            if (!activeSelection.HasValue)
//            {
//                throw new NoActiveSelectionException();
//            }

//            Refactor(activeSelection.Value);
//        }

//        public virtual void Refactor(QualifiedSelection targetSelection)
//        {
//            var target = FindTargetDeclaration(targetSelection);

//            if (target == null)
//            {
//                throw new NoDeclarationForSelectionException(targetSelection);
//            }

//            Refactor(target);
//        }

//        protected abstract Declaration FindTargetDeclaration(QualifiedSelection targetSelection);
//        public abstract void Refactor(Declaration target);
//    }
//}

// Rubberduck.VBEditor.QualifiedModuleName
///// <summary>
///// Creates a QualifiedModuleName for a library reference.
///// Do not use this overload for referenced user projects.
///// </summary>
//public QualifiedModuleName(ReferenceInfo reference)
//        :this(reference.Name,
//            reference.FullPath,
//            reference.Name)
//        { }

///// <summary>
///// Gets the standard projectId for a library reference.
///// Do not use this overload for referenced user projects.
///// </summary>
//public static string GetProjectId(ReferenceInfo reference)
//{
//    return new QualifiedModuleName(reference).ProjectId;
//}




// Rubberduck.Refactorings.IRefactoring
//void Refactor();
//void Refactor(QualifiedSelection target);
//void Refactor(Declaration target);



//namespace Rubberduck.VBEditor.Utility
//{
//    public interface ISelectionProvider
//    {
//        /// <summary>
//        /// Gets the QualifiedModuleName for the component that is currently selected in the Project Explorer.
//        /// </summary>
//        QualifiedModuleName ProjectExplorerSelection();
//        QualifiedSelection? ActiveSelection();
//        ICollection<QualifiedModuleName> OpenModules();
//        Selection? Selection(QualifiedModuleName module);
//    }
//}


// Current investigation of Rubber components to refactorr a ComInterface to extract a VBA Com Wrapper
//Require to find how to incorporate the ComInterface object acquired


// Issue > Rubberduck.Parsing.ComReflection.DeclarationsFromComProjectLoader
//         ->LoadDeclarations(ComProject type, string projectId = null)

//using Rubberduck.Parsing.Symbols;

//namespace Rubberduck.Refactorings.ImplementInterface
//{
//    public class ImplementInterfaceModel : IRefactoringModel
//    {
//        public ClassModuleDeclaration TargetInterface { get; }
//        public ClassModuleDeclaration TargetClass { get; }

//        public ImplementInterfaceModel(ClassModuleDeclaration targetInterface, ClassModuleDeclaration targetClass)
//        {
//            TargetInterface = targetInterface;
//            TargetClass = targetClass;
//        }
//    }
//}


//Possible issue with custom version of ComProject
// As ProjecctDeclaration constructor parameter
// If ComProject is for a Com Type Library, would obtain list of declarations
// Possible need to create a Declaration where type is a ComProject,  projectId = ComTypeLib.Name eg DotNetLib ??

//  Rubberduck.VBEditor.QualifiedModuleName
//  QualifiedModuleName(string projectName, string projectPath, string componentName, string projectId = null)



//IDeclarationsFromComProjectLoader
//namespace Rubberduck.Parsing.ComReflection
//{
//    public interface IDeclarationsFromComProjectLoader
//    {
//        IReadOnlyCollection<Declaration> LoadDeclarations(ComProject type, string projectId = null);
//    }
//}


//Rubberduck.Parsing.ComReflection.DeclarationsFromComProjectLoader
// public IReadOnlyCollection<Declaration> LoadDeclarations(ComProject type, string projectId = null)
// private static ICollection<Declaration> GetDeclarationsForModule(IComType module, QualifiedModuleName moduleName, ProjectDeclaration project)
// private static ICollection<Declaration> GetDeclarationsForModule(IComType module, QualifiedModuleName moduleName, ProjectDeclaration project)


//public IReadOnlyCollection<Declaration> LoadDeclarations(ComProject type, string projectId = null)
//{
//    var declarations = new List<Declaration>();

//    var projectName = new QualifiedModuleName(type.Name, type.Path, type.Name, projectId);
//    var project = new ProjectDeclaration(type, projectName);
//    declarations.Add(project);

//    foreach (var alias in type.Aliases.Select(item => new AliasDeclaration(item, project, projectName)))
//    {
//        declarations.Add(alias);
//    }

//    foreach (var module in type.Members)
//    {
//        var moduleIdentifier = module.Type == DeclarationType.Enumeration || module.Type == DeclarationType.UserDefinedType
//            ? $"_{module.Name}"
//            : module.Name;
//        var moduleName = new QualifiedModuleName(type.Name, type.Path, moduleIdentifier);

//        var moduleDeclarations = GetDeclarationsForModule(module, moduleName, project);
//        declarations.AddRange(moduleDeclarations);
//    }

//    return declarations;
//}








//Rubberduck.Refactorings.ImplementInterface.ImplementInterfaceRefactoring
//public override void Refactor(QualifiedSelection target)
//{
//    var targetInterface = _declarationFinderProvider.DeclarationFinder.FindInterface(target);

//    if (targetInterface == null)
//    {
//        throw new NoImplementsStatementSelectedException(target);
//    }

//    var targetModule = _declarationFinderProvider.DeclarationFinder
//        .ModuleDeclaration(target.QualifiedName);

//    if (!ImplementingModuleTypes.Contains(targetModule.DeclarationType))
//    {
//        throw new InvalidDeclarationTypeException(targetModule);
//    }

//    var targetClass = targetModule as ClassModuleDeclaration;

//    if (targetClass == null)
//    {
//        //This really should never happen. If it happens the declaration type enum value
//        //and the type of the declaration are inconsistent.
//        throw new InvalidTargetDeclarationException(targetModule);
//    }

//    var model = Model(targetInterface, targetClass);
//    _refactoringAction.Refactor(model);
//}




// Rubberduck.Parsing.Symbols.ClassModuleDeclaration
//
//public ClassModuleDeclaration(ComInterface @interface, Declaration parent, QualifiedModuleName module,
//    Attributes Attributes)
//            : this(
//                module.QualifyMemberName(@interface.Name),
//                parent,
//                @interface.Name,
//                false,
//                new List<IParseTreeAnnotation>(),
//                Attributes)
//        { }


//Rubberduck.VBEditor.QualifiedModuleName

/// <summary>
/// Creates a QualifiedModuleName for a library reference.
/// Do not use this overload for referenced user projects.
/// </summary>
//public QualifiedModuleName(ReferenceInfo reference)
//        :this(reference.Name,
//            reference.FullPath,
//            reference.Name)
//        { }



//  public class ClassModuleDeclaration : ModuleDeclaration
//  public abstract class ModuleDeclaration : Declaration
//  public class Declaration : IEquatable<Declaration>