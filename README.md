# Class diagram creator for C# code (Visio 2017)
Tool for marking classes/structs/interfaces/enums for being included in a class diagram, made with Visio 2017.

Writing every class, field, and method in a class diagram always take ages, even more so when you have a big project. And since the built in class diagram tool for Visual Studio is rather limited in functionality, I decided to create a tool, which could take a project, and create a class diagram for you, in Visio 2017.

The tool searches through all classes/structs/interfaces/enums in your assembly, which has been marked with specific attributes, and reads all information in them, and transfers it to a class diagram. If specified, you can also make it generate inheritance arrows, and/or relational arrows (though those are limited to what it can read from reflection, so they still require some manual labor).

Note: There is no implemented way of creating a layout automatically, so right now it is just making a X by Y grid of classes, X being specified by the starting method, and Y being total number of classes modulus X.

# Requirements
- Microsoft Visio 2017 (Maybe, just maybe, it works with other versions, but since I don't own them, I can't test it out)
- A project capable of running C# code and using C# attributes.
- A bit of patience (Visio commands are performed real time, so that also means that creating a big class diagram can take rather long)
- .NET 4.0 or later project.

# Usage
- First include the DLL generated by this project, in your project.

- Include the namespace "ClassDiagramMaker".

- Mark classes/structs with `[ClassDiagram]`, interfaces with `[ClassDiagramInterface]` and enums with `[ClassDiagramEnum]`.

- Somewhere in the code, where you can run a function, call the ClassDiagramCreator.MakeClassDiagram function.
  - This takes in a bitflag ArrowsToInclude, which dictates what arrows to include. Can be used as such: `ClassDiagramCreator.MakeClassDiagram(ArrowsToInclude.Inheritance | ArrowsToInclude.SuggestedRelationArrows);`.
  
- That is it! Once the method is called, if you have Visio installed, it should open a new document, and start generating the class diagram.

# Example
I tested this tool on an old school project, and the result can be seen in the `ClassDiagramCreatorExample.png` in the main folder of this project.
As you can see, the arrows/lines are quite confusing, so would probably suggest either only using inheritance arrows, or none at all.

# Future additions
There are a few things that I would like to add one day, if I get the chance.
- An actual layout algorithm, which can place classes in relation to what they inherit from, so more of the actual layout work is done for you.
- Default values and constants.
- Generic classes.
- Better layout for arrows.
- Faster creation (don't know if this is possible, due to the nature of calling commands in Visio).

