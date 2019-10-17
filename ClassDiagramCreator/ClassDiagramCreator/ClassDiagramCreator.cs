using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ClassDiagramMaker
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
    public class ClassDiagramAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Enum)]
    public class ClassDiagramEnumAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Interface)]
    public class ClassDiagramInterfaceAttribute : Attribute
    {
    }

    public enum ArrowsToInclude { None = 0, Inheritance = 1, SuggestedRelationArrows = 2 }

    [ClassDiagramEnum]
    public enum ConnectorArrows { Association = 12, Inheritance = 16, InterfaceInheritance = 99, EnumerationDependency = 100, Aggregation = 22, Composition = -1, Unknown = -2 }

    public static class ClassDiagramCreator
    {
        private static Application visApp;
        private static Dictionary<Type, Shape> AllShapeTypes = new Dictionary<Type, Shape>();
        private static Page visioPage;

        /// <summary>
        /// Opens Visio on the PC and creates a uml class diagram, based on what classes has the [UMLDiagram----] attributes.
        /// </summary>
        /// <param name="arrowsToInclude">Use as a flag if you want more than one (for all arrows, pass ArrowsToInclude.Inheritance |ArrowsToInclude.SuggestedRelationArrows) </param>
        public static void MakeClassDiagram(ArrowsToInclude arrowsToInclude = ArrowsToInclude.None, int NumberOfClassesPerRow = 10)
        {
            //int asdasd = 2;
            //var allClassesAndStructs = AppDomain.CurrentDomain
            //        .GetAssemblies()
            //        .SelectMany(x => x.GetTypes())
            //        .Where(x => x.IsClass || x.IsValueType)
            //        .Where(m => m.GetCustomAttributes(typeof(ClassDiagramAttribute), false).Length > 0)
            //        .ToArray();

            //var AllEnumTypes = AppDomain.CurrentDomain.GetAssemblies()
            //        .SelectMany(x => x.GetTypes())
            //        .Where(x => x.IsEnum)
            //        .Where(e => e.GetCustomAttributes(typeof(ClassDiagramEnumAttribute), false).Length > 0)
            //        .ToArray();

            //var AllInterfaces = AppDomain.CurrentDomain.GetAssemblies()
            //        .SelectMany(x => x.GetTypes())
            //        .Where(x => x.IsInterface)
            //        .Where(e => e.GetCustomAttributes(typeof(ClassDiagramInterfaceAttribute), false).Length > 0)
            //        .ToArray();

            //if (allClassesAndStructs.Length == 0 && AllEnumTypes.Length == 0 && AllInterfaces.Length == 0)
            //{
            //    return;
            //}



            List<Type> allClassesAndStructs = new List<Type>();
            List<Type> allEnumTypes = new List<Type>();
            List<Type> allInterfaces = new List<Type>();

            List<Type> AllTypes = AppDomain.CurrentDomain.GetAssemblies().SelectMany(x => x.GetTypes()).ToList();
            for (int i = 0; i < AllTypes.Count; i++)
            {
                if (AllTypes[i].GetCustomAttribute(typeof(ClassDiagramAttribute), false) != null)
                {
                    allClassesAndStructs.Add(AllTypes[i]);
                }
                else if (AllTypes[i].GetCustomAttribute(typeof(ClassDiagramEnumAttribute), false) != null)
                {
                    allEnumTypes.Add(AllTypes[i]);
                }
                else if (AllTypes[i].GetCustomAttribute(typeof(ClassDiagramInterfaceAttribute), false) != null)
                {
                    allInterfaces.Add(AllTypes[i]);

                }
            }

            if (allClassesAndStructs.Count == 0 && allEnumTypes.Count == 0 && allInterfaces.Count == 0)
            {
                return;
            }

            visApp = new Application();


            var doc = visApp.Documents.Add("");

            Documents visioDocs = visApp.Documents;
            //Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
            //    (short)VisOpenSaveArgs.visOpenDocked);
            Document umlStencils = null;
            try
            {
                umlStencils = visioDocs.OpenEx("USTRME_M.vssx",
                    (short)VisOpenSaveArgs.visOpenDocked);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
                Environment.Exit(0);
            }



            //This is like.. Super unperformant
            var listOfAllInheritanceClasses = GetInheritances(allClassesAndStructs);

            visioPage = visApp.ActivePage;

            Master visioClassMaster = umlStencils.Masters.get_ItemU(@"Class");

            #region classes and structs

            float biggestHeight = 0;
            float currentDepth = 12;
            float xPos = 0;
            for (int i = 0; i < allClassesAndStructs.Count; i++)
            {
                if (i % NumberOfClassesPerRow == 0)
                {
                    currentDepth -= (biggestHeight + 1);
                    biggestHeight = 0;
                    xPos = 0;
                }
                else
                {
                    xPos += 3;
                }
                Type curType = allClassesAndStructs[i];

                Shape visioClassShape = visioPage.Drop(visioClassMaster, xPos, currentDepth);
                visioClassShape.Text = curType.Name;
                //Gets all shape ID's within the newly created shape
                var ids = visioClassShape.ContainerProperties.GetMemberShapes((int)VisContainerFlags.visContainerFlagsDefault);

                int timesRunThrough = 0;

                foreach (int id in ids)
                {
                    //Get the specific shape from the page
                    Shape shape = visioPage.Shapes.ItemFromID[id];
                    //shape.Characters.Begin = shape.Characters.End;

                    #region fields

                    if (timesRunThrough == 0)
                    {
                        shape.Characters.Text = "";

                        BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly;

                        //For some god forsaken reason, we have to get a local reference, otherwise everything dies.
                        var shapeText = shape.Characters;

                        bool isFirst = true;
                        foreach (var field in curType.GetFields(bindFlags))
                        {
                            shapeText.Begin = shapeText.End;
                            if (!isFirst)
                            {
                                shapeText.Text = "\r\n";
                            }
                            else
                            {
                                isFirst = false;
                            }
                            if (field.IsPublic)
                            {
                                shapeText.Text += "+ " + CheckAndFixAutoPropName(field) + " : " + NicyfyFieldType(field.FieldType);
                            }
                            //Internal fields - https://docs.microsoft.com/en-us/dotnet/api/system.reflection.fieldinfo.isfamily?view=netframework-4.7.2
                            else if (field.IsAssembly)
                            {
                                shapeText.Text += "~ " + CheckAndFixAutoPropName(field) + " : " + NicyfyFieldType(field.FieldType);
                            }
                            //Protected fields - https://docs.microsoft.com/en-us/dotnet/api/system.reflection.fieldinfo.isfamily?view=netframework-4.7.2
                            else if (field.IsFamily)
                            {
                                shapeText.Text += "# " + CheckAndFixAutoPropName(field) + " : " + NicyfyFieldType(field.FieldType);
                            }
                            else if (!field.IsPublic)
                            {
                                shapeText.Text += "- " + CheckAndFixAutoPropName(field) + " : " + NicyfyFieldType(field.FieldType);
                            }
                            if (field.IsStatic)
                            {
                                shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visUnderLine);
                            }
                        }

                        //Force it back into the container, since linebreaks cause things to die
                        visioClassShape.ContainerProperties.InsertListMember(shape, 0);
                    }

                    #endregion fields

                    #region separator

                    if (timesRunThrough == 1)
                    {
                        timesRunThrough++;
                        continue;
                    }

                    #endregion separator

                    #region methods

                    if (timesRunThrough == 2)
                    {
                        shape.Characters.Text = "";

                        BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly;

                        //For some god forsaken reason, we have to get a local reference, otherwise everything dies.
                        var shapeText = shape.Characters;

                        bool isFirst = true;

                        //Properties
                        var properties = curType.GetProperties(bindFlags);
                        foreach (var property in properties)
                        {
                            for (int p = 0; p < 2; p++)
                            {
                                if (!property.CanRead && p == 0) continue;
                                if (!property.CanWrite && p == 1) continue;
                                MethodInfo propertyMethod;
                                string newName = "";
                                if (p == 0)
                                {
                                    propertyMethod = property.GetGetMethod(true);
                                    newName = propertyMethod.Name;
                                    newName = newName.Substring(4);
                                    newName = newName.Insert(0, "Get");
                                }
                                else
                                {
                                    propertyMethod = property.GetSetMethod(true);
                                    newName = propertyMethod.Name;
                                    newName = newName.Substring(4);
                                    newName = newName.Insert(0, "Set");
                                }
                                shapeText.Begin = shapeText.End;

                                if (!isFirst)
                                {
                                    shapeText.Text = "\r\n";
                                }
                                else
                                {
                                    isFirst = false;
                                }
                                if (propertyMethod.IsPublic)
                                {
                                    shapeText.Text += "+ " + newName + "(" + HandleParameterInfo(propertyMethod.GetParameters()) + ")" + " : " + NicyfyFieldType(propertyMethod.ReturnType);
                                }
                                //Internal methods - https://docs.microsoft.com/en-us/dotnet/api/system.reflection.fieldinfo.isfamily?view=netframework-4.7.2
                                else if (propertyMethod.IsAssembly)
                                {
                                    shapeText.Text += "~ " + newName + "(" + HandleParameterInfo(propertyMethod.GetParameters()) + ")" + " : " + NicyfyFieldType(propertyMethod.ReturnType);
                                }
                                //Protected methods - https://docs.microsoft.com/en-us/dotnet/api/system.reflection.fieldinfo.isfamily?view=netframework-4.7.2
                                else if (propertyMethod.IsFamily)
                                {
                                    shapeText.Text += "# " + newName + "(" + HandleParameterInfo(propertyMethod.GetParameters()) + ")" + " : " + NicyfyFieldType(propertyMethod.ReturnType);
                                }
                                else if (!propertyMethod.IsPublic)
                                {
                                    shapeText.Text += "- " + newName + "(" + HandleParameterInfo(propertyMethod.GetParameters()) + ")" + " : " + NicyfyFieldType(propertyMethod.ReturnType);
                                }
                                if (propertyMethod.IsStatic)
                                {
                                    shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visUnderLine);
                                }
                                else if (propertyMethod.IsAbstract)
                                {
                                    shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visItalic);
                                }
                                else
                                {
                                    shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visCaseNormal);

                                }
                            }
                        }
                        shapeText = shape.Characters;
                        //this also finds normal ctors, don't do that, TODO
                        //Constructors
                        var constructors = curType.GetConstructors(bindFlags);
                        foreach (var ctor in constructors)
                        {
                            shapeText.Begin = shapeText.End;
                            if (!isFirst)
                            {
                                shapeText.Text = "\r\n";
                            }
                            else
                            {
                                isFirst = false;
                            }
                            shapeText.Text += ctor.IsPublic ? "+ " : ctor.IsPrivate ? "- " : ctor.IsAssembly ? "~ " : ctor.IsFamily ? "# " : " ";
                            shapeText.Text += curType.Name + "(" + HandleParameterInfo(ctor.GetParameters()) + ")";
                            if (ctor.IsStatic)
                            {
                                shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visUnderLine);
                            }
                            else
                            {
                                shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visCaseNormal);
                            }
                        }
                        shapeText = shape.Characters;

                        //All methods (Apparently properties are tagged as "IsSpecialName", so discard those)
                        foreach (var method in curType.GetMethods(bindFlags).Where(m => !m.IsSpecialName))
                        {
                            shapeText.Begin = shapeText.End;
                            if (!isFirst)
                            {
                                shapeText.Text = "\r\n";
                            }
                            else
                            {
                                isFirst = false;
                            }
                            if (method.IsPublic)
                            {
                                shapeText.Text += "+ " + method.Name + "(" + HandleParameterInfo(method.GetParameters()) + ")" + " : " + NicyfyFieldType(method.ReturnType);
                            }
                            //Internal methods - https://docs.microsoft.com/en-us/dotnet/api/system.reflection.fieldinfo.isfamily?view=netframework-4.7.2
                            else if (method.IsAssembly)
                            {
                                shapeText.Text += "~ " + method.Name + "(" + HandleParameterInfo(method.GetParameters()) + ")" + " : " + NicyfyFieldType(method.ReturnType);
                            }
                            //Protected methods - https://docs.microsoft.com/en-us/dotnet/api/system.reflection.fieldinfo.isfamily?view=netframework-4.7.2
                            else if (method.IsFamily)
                            {
                                shapeText.Text += "# " + method.Name + "(" + HandleParameterInfo(method.GetParameters()) + ")" + " : " + NicyfyFieldType(method.ReturnType);
                            }
                            else if (!method.IsPublic)
                            {
                                shapeText.Text += "- " + method.Name + "(" + HandleParameterInfo(method.GetParameters()) + ")" + " : " + NicyfyFieldType(method.ReturnType);
                            }

                            if (method.IsStatic)
                            {
                                shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visUnderLine);
                            }
                            else if (method.IsAbstract)
                            {
                                shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visItalic);
                            }
                            else
                            {
                                shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visCaseNormal);
                            }
                        }

                        //Force it back into the container, since linebreaks cause things to die
                        visioClassShape.ContainerProperties.InsertListMember(shape, 3);


                    }

                    #endregion methods

                    timesRunThrough++;
                }

                float height = (float)visioClassShape.Cells["Height"].Result[VisUnitCodes.visInches];
                if (height > biggestHeight)
                {
                    biggestHeight = height;
                }
                AllShapeTypes.Add(curType, visioClassShape);

                //If class is static
                if (curType.IsAbstract && curType.IsSealed)
                {
                    //https://docs.microsoft.com/en-us/office/vba/api/visio.characters.charprops
                    //Underline
                    visioClassShape.Characters.CharProps[2] = 4;
                }
                //If just abstract
                else if (curType.IsAbstract)
                {
                    //https://docs.microsoft.com/en-us/office/vba/api/visio.characters.charprops
                    //Italic
                    visioClassShape.Characters.CharProps[2] = 2;
                }
            }

            #endregion classes and structs

            visioPage.AutoSizeDrawing();

            #region Enums



            Master visioEnumMaster = umlStencils.Masters.get_ItemU(@"Enumeration");

            for (int i = 0; i < allEnumTypes.Count; i++)
            {
                Type curType = allEnumTypes[i];

                Shape visioEnumShape = visioPage.Drop(visioEnumMaster, (i * 3) + 5, 14);
                visioEnumShape.Text = curType.Name;

                //Gets all shape ID's within the newly created shape
                var ids = visioEnumShape.ContainerProperties.GetMemberShapes((int)VisContainerFlags.visContainerFlagsDefault);

                int timesRunThrough = 0;
                foreach (int id in ids)
                {
                    Shape shape = visioPage.Shapes.ItemFromID[id];
                    if (timesRunThrough == 0)
                    {
                        shape.Characters.Text = "";

                        //For some god forsaken reason, we have to get a local reference, otherwise everything dies.
                        var shapeText = shape.Characters;

                        var enums = Enum.GetValues(curType);
                        foreach (var enumVal in enums)
                        {
                            shape.Characters.Text += enumVal.ToString();
                            if ((int)enumVal != (int)enums.GetValue(enums.Length - 1))
                                shape.Characters.Text += "\r\n";
                        }

                        visioEnumShape.ContainerProperties.InsertListMember(shape, 0);
                    }
                    if (timesRunThrough > 0)
                    {
                        shape.Text = "";
                        shape.Delete();
                    }
                    timesRunThrough++;
                }
                AllShapeTypes.Add(curType, visioEnumShape);
            }

            #endregion Enums

            visioPage.AutoSizeDrawing();

            #region Interfaces



            Master visioInterfaceMaster = umlStencils.Masters.get_ItemU(@"Interface");

            for (int i = 0; i < allInterfaces.Count; i++)
            {
                Type curType = allInterfaces[i];

                Shape visioInterfaceShape = visioPage.Drop(visioInterfaceMaster, 5 + i * 3, 8);

                visioInterfaceShape.Text = curType.Name;

                var ids = visioInterfaceShape.ContainerProperties.GetMemberShapes((int)VisContainerFlags.visContainerFlagsDefault);

                int timesRunThrough = 0;

                foreach (int id in ids)
                {
                    //Get the specific shape from the page
                    Shape shape = visioPage.Shapes.ItemFromID[id];
                    //shape.Characters.Begin = shape.Characters.End;

                    #region fields

                    if (timesRunThrough == 0)
                    {
                        shape.Characters.Text = "";

                        BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static;

                        //For some god forsaken reason, we have to get a local reference, otherwise everything dies.
                        var shapeText = shape.Characters;

                        bool isFirst = true;
                        foreach (var field in curType.GetFields(bindFlags))
                        {
                            shapeText.Begin = shapeText.End;
                            if (!isFirst)
                            {
                                shapeText.Text = "\r\n";
                            }
                            else
                            {
                                isFirst = false;
                            }

                            shapeText.Text += "- " + CheckAndFixAutoPropName(field) + " : " + NicyfyFieldType(field.FieldType);
                        }

                        //Force it back into the container, since linebreaks cause things to die
                        visioInterfaceShape.ContainerProperties.InsertListMember(shape, 0);
                    }

                    #endregion fields

                    #region separator

                    if (timesRunThrough == 1)
                    {
                        timesRunThrough++;
                        continue;
                    }

                    #endregion separator

                    #region methods

                    if (timesRunThrough == 2)
                    {
                        shape.Characters.Text = "";

                        BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly;

                        //For some god forsaken reason, we have to get a local reference, otherwise everything dies.
                        var shapeText = shape.Characters;

                        bool isFirst = true;

                        //Properties
                        var properties = curType.GetProperties(bindFlags);
                        foreach (var property in properties)
                        {
                            for (int p = 0; p < 2; p++)
                            {
                                if (!property.CanRead && p == 0) continue;
                                if (!property.CanWrite && p == 1) continue;
                                MethodInfo method;
                                string newName = "";
                                if (p == 0)
                                {
                                    method = property.GetGetMethod(true);
                                    newName = method.Name;
                                    newName = newName.Substring(4);
                                    newName = newName.Insert(0, "Get");
                                }
                                else
                                {
                                    method = property.GetSetMethod(true);
                                    newName = method.Name;
                                    newName = newName.Substring(4);
                                    newName = newName.Insert(0, "Set");
                                }
                                shapeText.Begin = shapeText.End;

                                if (!isFirst)
                                {
                                    shapeText.Text = "\r\n";
                                }
                                else
                                {
                                    isFirst = false;
                                }

                                shapeText.Text += "+ " + newName + "(" + HandleParameterInfo(method.GetParameters()) + ")" + " : " + NicyfyFieldType(method.ReturnType);

                                if (method.IsAbstract)
                                {
                                    shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visItalic);
                                }
                            }
                        }

                        //All methods (Apparently properties are tagged as "IsSpecialName", so discard those)
                        foreach (var method in curType.GetMethods(bindFlags).Where(m => !m.IsSpecialName))
                        {
                            shapeText.Begin = shapeText.End;
                            if (!isFirst)
                            {
                                shapeText.Text = "\r\n";
                            }
                            else
                            {
                                isFirst = false;
                            }

                            shapeText.Text += "+ " + method.Name + "(" + HandleParameterInfo(method.GetParameters()) + ")" + " : " + NicyfyFieldType(method.ReturnType);

                            if (method.IsAbstract)
                            {
                                shapeText.set_CharProps((short)VisCellIndices.visCharacterStyle, (short)VisCellVals.visItalic);
                            }
                        }

                        //Force it back into the container, since linebreaks cause things to die
                        visioInterfaceShape.ContainerProperties.InsertListMember(shape, 3);
                    }

                    #endregion methods

                    timesRunThrough++;
                }
                AllShapeTypes.Add(curType, visioInterfaceShape);
            }

            #endregion Interfaces

            visioPage.AutoSizeDrawing();

            //const string BASIC_FLOWCHART_STENCIL = "BASFLO_M.VSSX";
            //const string DYNAMIC_CONNECTOR_MASTER = "Dynamic Connector";

            ////Drop a connector and delete it because Interfaces for SOME FREACKING REASON dies if no arrows are dropped
            //var stencil = visApp.Documents.OpenEx(
            //           BASIC_FLOWCHART_STENCIL,
            //           (short)Microsoft.Office.Interop.Visio.
            //               VisOpenSaveArgs.visOpenDocked);

            //// Get the dynamic connector master on the stencil by its
            //// universal name.
            //var masterInStencil = stencil.Masters.get_ItemU(
            //    DYNAMIC_CONNECTOR_MASTER);


            //double parentPosX = AllShapeTypes[typeof(TestOne)].Cells["PinX"].Result[VisUnitCodes.visMillimeters];
            //double parentPosY = AllShapeTypes[typeof(TestOne)].Cells["PinY"].Result[VisUnitCodes.visMillimeters];
            //double parentWidth = AllShapeTypes[typeof(TestOne)].Cells["Width"].Result[VisUnitCodes.visMillimeters];
            //double parentHeight = AllShapeTypes[typeof(TestOne)].Cells["Height"].Result[VisUnitCodes.visMillimeters];





            //Important
            if ((arrowsToInclude & ArrowsToInclude.Inheritance) != ArrowsToInclude.None)
            {
                var AllInterfaceInheritance = GetInterfaceInheritances(allClassesAndStructs, allInterfaces);
                RecursiveConnectAllInheritance(AllInterfaceInheritance);
                RecursiveConnectAllInheritance(listOfAllInheritanceClasses);
            }

            if ((arrowsToInclude & ArrowsToInclude.SuggestedRelationArrows) != ArrowsToInclude.None)
            {
                CreateConnectionsBetweenTypes(AllShapeTypes);
            }
            //CorrectInheritancePosition(listOfAllInheritanceClasses, AllShapeTypes);

            visioPage.AutoSizeDrawing();


        }


        private static void CreateConnectionsBetweenTypes(Dictionary<Type, Shape> allTypesEnumerationsInterfacesAndShapes)
        {
            foreach (var typeFrom in allTypesEnumerationsInterfacesAndShapes.Keys)
            {
                List<Type> typesToConnectTo = new List<Type>();
                foreach (var typeTo in allTypesEnumerationsInterfacesAndShapes.Keys)
                {
                    if (typeFrom == typeTo || typeTo.IsEnum || typeTo.IsInterface)
                    {
                        continue;
                    }
                    bool hasSavedRef = false;
                    //Fields
                    BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly;
                    var allFieldsInType = typeTo.GetFields(bindFlags);

                    foreach (var field in allFieldsInType)
                    {
                        //Lists
                        foreach (Type interfaceType in field.FieldType.GetInterfaces())
                        {
                            if (interfaceType.IsGenericType &&
                                interfaceType.GetGenericTypeDefinition()
                                == typeof(IList<>))
                            {
                                foreach (var genericArg in field.FieldType.GetGenericArguments())
                                {
                                    if (genericArg == typeFrom)
                                    {
                                        if (!typesToConnectTo.Contains(typeTo))
                                        {
                                            typesToConnectTo.Add(typeTo);
                                            hasSavedRef = true;
                                            break;
                                        }
                                    }
                                }

                                break;
                            }
                            if (hasSavedRef)
                            {
                                break;
                            }
                        }

                        //Arrays
                        if ((field.FieldType.BaseType != null && field.FieldType.BaseType.Name == "Array") || (field.FieldType.IsByRef && field.FieldType.Name.Contains("[]")))
                        {
                            Type arrayType = field.FieldType.GetElementType();
                            if (arrayType == typeFrom)
                            {
                                if (!typesToConnectTo.Contains(typeTo))
                                {
                                    hasSavedRef = true;
                                    typesToConnectTo.Add(typeTo);
                                    break;
                                }
                            }
                        }

                        if (field.FieldType == typeFrom)
                        {
                            if (!typesToConnectTo.Contains(typeTo))
                            {
                                hasSavedRef = true;
                                typesToConnectTo.Add(typeTo);
                                break;
                            }
                        }
                    }
                    //Methods
                    if (hasSavedRef)
                    {
                        continue;
                    }

                    var allMethodsInType = typeTo.GetMethods(bindFlags);
                    foreach (var method in allMethodsInType)
                    {
                        var parameters = method.GetParameters();

                        foreach (var parameter in parameters)
                        {
                            Type paramType = parameter.ParameterType;

                            //Lists
                            foreach (Type interfaceType in paramType.GetInterfaces())
                            {
                                if (interfaceType.IsGenericType &&
                                    interfaceType.GetGenericTypeDefinition()
                                    == typeof(IList<>))
                                {
                                    foreach (var genericArg in paramType.GetGenericArguments())
                                    {
                                        if (genericArg == typeFrom)
                                        {
                                            if (!typesToConnectTo.Contains(typeTo))
                                            {
                                                typesToConnectTo.Add(typeTo);
                                                hasSavedRef = true;
                                                break;
                                            }
                                        }
                                    }

                                    break;
                                }
                                if (hasSavedRef)
                                {
                                    break;
                                }
                            }


                            //Arrays
                            if ((paramType.BaseType != null && paramType.BaseType.Name == "Array") || (paramType.IsByRef && paramType.Name.Contains("[]")))
                            {
                                Type arrayType = paramType.IsByRef ? paramType.GetElementType().GetElementType() : paramType.GetElementType();
                                if (arrayType == typeFrom)
                                {
                                    if (!typesToConnectTo.Contains(typeTo))
                                    {
                                        hasSavedRef = true;
                                        typesToConnectTo.Add(typeTo);
                                        break;
                                    }
                                }
                            }

                            if (parameter.ParameterType == typeFrom)
                            {
                                if (!typesToConnectTo.Contains(typeTo))
                                {
                                    typesToConnectTo.Add(typeTo);
                                    break;
                                }
                            }
                        }
                        if (method.ReturnType == typeFrom)
                        {
                            if (!typesToConnectTo.Contains(typeTo))
                            {
                                typesToConnectTo.Add(typeTo);
                                break;
                            }
                        }
                    }
                    if (hasSavedRef)
                    {
                        continue;
                    }
                }

                foreach (var typeToConnectTo in typesToConnectTo)
                {
                    if (typeToConnectTo != typeFrom)
                    {
                        if (typeFrom.IsEnum)
                        {
                            ConnectWithDynamicGlueAndConnector(allTypesEnumerationsInterfacesAndShapes[typeToConnectTo], allTypesEnumerationsInterfacesAndShapes[typeFrom], ConnectorArrows.EnumerationDependency);
                        }
                        else
                        {
                            ConnectWithDynamicGlueAndConnector(allTypesEnumerationsInterfacesAndShapes[typeFrom], allTypesEnumerationsInterfacesAndShapes[typeToConnectTo], ConnectorArrows.Unknown);
                        }
                    }
                }
            }
        }

        private static List<MyVector2> CorrectInheritancePosition(List<TypeContainer> allClassesWithInheritance, Dictionary<Type, Shape> allTypesEnumerationsInterfacesAndShapes)
        {
            List<MyVector2> depthWidth = new List<MyVector2>();

            foreach (var tc in allClassesWithInheritance)
            {
                MyVector2 position;
                MyVector2 size;
                GetShapePosAndSize(allTypesEnumerationsInterfacesAndShapes[tc.ThisType], out position, out size);

                float xOffSetFromCenterOfParent = 0;
                if (tc.ChildTypes.Count % 2 == 0)
                {
                    xOffSetFromCenterOfParent = size.x / 2;
                }


                for (int i = 0; i < tc.ChildTypes.Count; i++)
                {
                    Shape childShape = allTypesEnumerationsInterfacesAndShapes[tc.ChildTypes[i].ThisType];
                    MyVector2 childPos;
                    MyVector2 childSize;
                    GetShapePosAndSize(childShape, out childPos, out childSize);

                    int multiplier = -1;
                    if (i != 1 && i % 2 == 0)
                    {
                        multiplier = 1;
                    }

                    MyVector2 neededXAndY = new MyVector2();
                    neededXAndY.x = position.x - childPos.x + xOffSetFromCenterOfParent + (multiplier * i * childSize.x);
                    neededXAndY.y = position.y + size.y + 1 - childPos.y;

                    MoveShapeAndChildren(childShape, visioPage, neededXAndY.x, neededXAndY.y * -1);


                    MyVector2 newchildPos;
                    MyVector2 newchildSize;
                    GetShapePosAndSize(childShape, out newchildPos, out newchildSize);
                    
                }

            }



            //foreach (var tc in allClassesWithInheritance)
            //{

            //    //Queue<Shape> shapesToVisit = new Queue<Shape>();
            //    //TypeContainer curCon = tc;
            //    //shapesToVisit.Enqueue(allTypesEnumerationsInterfacesAndShapes[tc.ThisType]);

            //    //RecursiveAddChildrenToQueue(tc, allTypesEnumerationsInterfacesAndShapes, ref shapesToVisit);
            //}

            return depthWidth;
        }

        //public static void RecursiveAddChildrenToQueue(TypeContainer tc, Dictionary<Type, Shape> allTypesEnumerationsInterfacesAndShapes, ref Queue<Shape> queue)
        //{
        //    foreach (var child in tc.ChildTypes)
        //    {
        //        queue.Enqueue(allTypesEnumerationsInterfacesAndShapes[child.ThisType]);
        //        RecursiveAddChildrenToQueue(child, allTypesEnumerationsInterfacesAndShapes, ref queue);
        //    }
        //}

        private static void RecursiveConnectAllInheritance(List<TypeContainer> typeContainers, TypeContainer parentType = null)
        {
            foreach (var tc in typeContainers)
            {
                if (parentType != null)
                {
                    ConnectWithDynamicGlueAndConnector(AllShapeTypes[tc.ThisType], AllShapeTypes[parentType.ThisType], parentType.ThisType.IsInterface ? ConnectorArrows.InterfaceInheritance : ConnectorArrows.Inheritance);
                }

                RecursiveConnectAllInheritance(tc.ChildTypes, tc);
            }
        }

        private static void GetShapePosAndSize(Shape shape, out MyVector2 position, out MyVector2 size)
        {
            position.x = (float)shape.Cells["PinX"].Result[VisUnitCodes.visInches];
            position.y = (float)shape.Cells["PinY"].Result[VisUnitCodes.visInches];
            size.x = (float)shape.Cells["Width"].Result[VisUnitCodes.visInches];
            size.y = (float)shape.Cells["Height"].Result[VisUnitCodes.visInches];

        }

        private static List<TypeContainer> GetInheritances(List<Type> allClassesAndStructs)
        {
            List<TypeContainer> listToReturn = new List<TypeContainer>();

            var baseClasses = allClassesAndStructs.Where(t => t.BaseType.BaseType == null || t.BaseType == typeof(ValueType)).ToList();
            foreach (var baseClass in baseClasses)
            {
                TypeContainer tc = new TypeContainer();
                tc.ThisType = baseClass;
                tc.Depth = 0;
                var temp = RecursiveGetChildClasses(allClassesAndStructs, tc);
                if (temp != null)
                {
                    tc.ChildTypes.AddRange(temp);
                }
                listToReturn.Add(tc);
            }

            return listToReturn;
        }

        private static List<TypeContainer> GetInterfaceInheritances(List<Type> allClassesAndStructs, List<Type> allInterfaces)
        {
            List<TypeContainer> listToReturn = new List<TypeContainer>();

            //var baseClasses = allClassesAndStructs.Where(t => t.BaseType.BaseType == null || t.BaseType == typeof(ValueType)).ToList();
            foreach (var baseInterface in allInterfaces)
            {
                TypeContainer tc = new TypeContainer();
                tc.ThisType = baseInterface;
                tc.Depth = 0;
                var temp = RecursiveGetChildClasses(allClassesAndStructs, tc);
                if (temp != null)
                {
                    tc.ChildTypes.AddRange(temp);
                }
                listToReturn.Add(tc);
            }

            return listToReturn;
        }

        private static List<TypeContainer> RecursiveGetChildClasses(List<Type> allClassesAndStructs, TypeContainer typeToCheck)
        {
            List<TypeContainer> returningTCList = new List<TypeContainer>();
            var typesThatInherit = allClassesAndStructs.Where(t => typeToCheck.ThisType.IsAssignableFrom(t) && t != typeToCheck.ThisType);
            if (typesThatInherit.Count() == 0)
            {
                return null;
            }
            foreach (var type in typesThatInherit)
            {
                if (typeToCheck.ThisType == type) continue;

                if (type.BaseType == typeToCheck.ThisType || type.GetInterfaces().Except(type.BaseType.GetInterfaces()).Contains(typeToCheck.ThisType))
                {
                    TypeContainer tc = new TypeContainer();
                    tc.ThisType = type;
                    tc.Depth = typeToCheck.Depth + 1;
                    returningTCList.Add(tc);
                    var temp = RecursiveGetChildClasses(allClassesAndStructs, tc);
                    if (temp != null)
                        tc.ChildTypes.AddRange(temp);
                }
            }
            return returningTCList;
        }

        private static string CheckAndFixAutoPropName(FieldInfo FI)
        {
            if (FI.Name.Contains("<"))
            {
                return FI.Name.Substring(1, FI.Name.IndexOf('>') - 1);
            }
            else
            {
                return FI.Name;
            }
        }

        private static string NicyfyFieldType(Type fieldType)
        {
            if (fieldType.Name.Contains("List"))
            {
                StringBuilder sb = new StringBuilder();
                if (fieldType.IsByRef)
                {
                    sb.Append(fieldType.Name.Remove(fieldType.Name.Length - 3, 3));

                }
                else
                {
                    sb.Append(fieldType.Name.Remove(fieldType.Name.Length - 2, 2));
                }
                sb.Append("<");

                foreach (var genericParam in fieldType.GetGenericArguments())
                {
                    sb.Append(NicyfyFieldType(genericParam));
                    sb.Append(", ");
                }
                if (sb[sb.Length - 2] == ',')
                {
                    sb.Remove(sb.Length - 2, 2);
                }
                sb.Append(">");
                return sb.ToString();
            }
            else if (fieldType.BaseType != null && fieldType.BaseType.Name == "Array")
            {
                StringBuilder sb = new StringBuilder();
                string typeWithoutBracket = fieldType.Name.Substring(0, fieldType.Name.IndexOf('['));
                string type = NicyfyStringType(typeWithoutBracket);
                return type += fieldType.Name.Substring(fieldType.Name.IndexOf('['));
            }
            else if (fieldType.IsByRef && fieldType.Name.Contains("[]"))
            {
                string type = NicyfyStringType(fieldType.Name.Substring(0, fieldType.Name.IndexOf("[]")));
                return type + "[]";

            }
            else if (fieldType.BaseType != null && (fieldType.BaseType != typeof(MulticastDelegate)))
            {
                string ft = fieldType.ToString();
                return NicyfyStringType(ft);
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                int indexOfApostophy = fieldType.Name.IndexOf('`');
                if (indexOfApostophy == -1)
                {
                    return fieldType.Name;
                }
                else
                {
                    sb.Append(fieldType.Name.Remove(fieldType.Name.Length - 2, 2));
                    sb.Append("<");
                    foreach (var genericParam in fieldType.GenericTypeArguments)
                    {
                        sb.Append(NicyfyFieldType(genericParam));
                        sb.Append(", ");
                    }
                    if (sb[sb.Length - 2] == ',')
                    {
                        sb.Remove(sb.Length - 2, 2);
                    }
                    sb.Append(">");
                }
                return sb.ToString();
            }
        }

        private static string NicyfyStringType(string typeAsString)
        {
            int lastIndexOfDot = typeAsString.LastIndexOf('.');
            string fieldTypeWithoutNameSpaces;
            if (lastIndexOfDot == -1)
            {
                fieldTypeWithoutNameSpaces = typeAsString;
            }
            else
            {
                fieldTypeWithoutNameSpaces = typeAsString.Substring(lastIndexOfDot + 1);
            }
            fieldTypeWithoutNameSpaces = fieldTypeWithoutNameSpaces.TrimEnd(']');
            //Probably more types that needs to be added here
            switch (fieldTypeWithoutNameSpaces)
            {
                case ("String"):
                    fieldTypeWithoutNameSpaces = "string";
                    break;

                case ("Int32"):
                    fieldTypeWithoutNameSpaces = "int";
                    break;

                case ("Single"):
                    fieldTypeWithoutNameSpaces = "float";
                    break;

                case ("Boolean"):
                    fieldTypeWithoutNameSpaces = "bool";
                    break;

                case ("Void"):
                    fieldTypeWithoutNameSpaces = "void";
                    break;
            }
            return fieldTypeWithoutNameSpaces;
        }

        private static string HandleParameterInfo(ParameterInfo[] paramInfos)
        {
            StringBuilder sb = new StringBuilder();
            bool first = true;
            foreach (var param in paramInfos)
            {
                if (first)
                {
                    first = false;
                    sb.Append("in " + param.Name + " : " + NicyfyFieldType(param.ParameterType));
                }
                else
                {
                    sb.Append(", in " + param.Name + " : " + NicyfyFieldType(param.ParameterType));
                }
            }

            return sb.ToString();
        }

        private static List<Shape> GetChildShapes(Shape parentShape, Page visioPage)
        {
            List<Shape> shapesToReturn = new List<Shape>();

            var childIDs = parentShape.ContainerProperties.GetMemberShapes((int)VisContainerFlags.visContainerFlagsDefault);

            foreach (int id in childIDs)
            {
                //Get the specific shape from the page
                Shape shape = visioPage.Shapes.ItemFromID[id];
                shapesToReturn.Add(shape);
            }
            return shapesToReturn;
        }

        private static void MoveShapeAndChildren(Shape masterShape, Page visioPage, float unitsToMoveX, float unitsToMoveY)
        {
            visApp.ActiveWindow.DeselectAll();

            var shapes = GetChildShapes(masterShape, visioPage);

            shapes.Insert(0, masterShape);
            foreach (Shape shape in shapes)
            {
                visApp.ActiveWindow.Select(shape, (short)2);
            }
            visApp.ActiveWindow.Selection.Move(unitsToMoveX, unitsToMoveY);
            visApp.ActiveWindow.DeselectAll();
        }

        /// <summary>This method accesses the Basic Flowchart Shapes stencil and
        /// the dynamic connector master on the stencil. It connects two 2-D
        /// shapes using the dynamic connector by gluing the connector to the
        /// PinX cells of the 2-D shapes to create dynamic (walking) glue.
        ///
        /// Note: To get dynamic glue, a dynamic connector must be used and
        /// connected to the PinX or PinY cell of the 2-D shape.
        /// For more information about dynamic glue, see the "Working with 1-D
        /// Shapes, Connectors, and Glue" section in the book, Developing
        /// Microsoft Visio Solutions.</summary>
        /// <param name="shapeFrom">Shape from which the dynamic connector
        /// begins</param>
        /// <param name="shapeTo">Shape at which the dynamic connector ends
        /// </param>
        private static void ConnectWithDynamicGlueAndConnector(
        Shape shapeFrom,
        Shape shapeTo,
        ConnectorArrows connectorType)
        {
            if (shapeFrom == null || shapeTo == null)
            {
                return;
            }

            const string BASIC_FLOWCHART_STENCIL = "BASFLO_M.VSSX";
            const string UML_STENCIL = "USTRME_M.vssx";
            const string DYNAMIC_CONNECTOR_MASTER = "Dynamic Connector";
            const string COMPOSITION_MASTER = "Composition";
            const string INTERFACEREALIZATION_MASTER = "Interface Realization";
            const string DEPENDENCY = "Dependency";
            const string MESSAGE_NOT_SAME_PAGE =
                "Both the shapes are not on the same page.";

            Application visioApplication;
            Document stencil;
            Master masterInStencil;
            Shape connector;
            Cell beginX;
            Cell endX;

            // Get the Application object from the shape.
            visioApplication = (Microsoft.Office.Interop.Visio.Application)
                shapeFrom.Application;

            try
            {
                // Verify that the shapes are on the same page.
                if (shapeFrom.ContainingPage != null && shapeTo.ContainingPage != null &&
                    shapeFrom.ContainingPage.Equals(shapeTo.ContainingPage))
                {
                    if (connectorType == ConnectorArrows.InterfaceInheritance)
                    {
                        // Access the Basic Flowchart Shapes stencil from the
                        // Documents collection of the application.
                        stencil = visioApplication.Documents.OpenEx(
                            UML_STENCIL,
                            (short)Microsoft.Office.Interop.Visio.
                                VisOpenSaveArgs.visOpenDocked);

                        // Get the dynamic connector master on the stencil by its
                        // universal name.
                        masterInStencil = stencil.Masters.get_ItemU(
                            INTERFACEREALIZATION_MASTER);
                    }
                    else if (connectorType == ConnectorArrows.EnumerationDependency)
                    {
                        // Access the Basic Flowchart Shapes stencil from the
                        // Documents collection of the application.
                        stencil = visioApplication.Documents.OpenEx(
                            UML_STENCIL,
                            (short)VisOpenSaveArgs.visOpenDocked);

                        // Get the dynamic connector master on the stencil by its
                        // universal name.
                        masterInStencil = stencil.Masters.get_ItemU(
                            DEPENDENCY);
                    }
                    else if (connectorType != ConnectorArrows.Composition)
                    {
                        // Access the Basic Flowchart Shapes stencil from the
                        // Documents collection of the application.
                        stencil = visioApplication.Documents.OpenEx(
                            BASIC_FLOWCHART_STENCIL,
                            (short)Microsoft.Office.Interop.Visio.
                                VisOpenSaveArgs.visOpenDocked);

                        // Get the dynamic connector master on the stencil by its
                        // universal name.
                        masterInStencil = stencil.Masters.get_ItemU(
                            DYNAMIC_CONNECTOR_MASTER);
                    }
                    else
                    {
                        // Access the Basic Flowchart Shapes stencil from the
                        // Documents collection of the application.
                        stencil = visioApplication.Documents.OpenEx(
                            UML_STENCIL,
                            (short)Microsoft.Office.Interop.Visio.
                                VisOpenSaveArgs.visOpenDocked);

                        // Get the dynamic connector master on the stencil by its
                        // universal name.
                        masterInStencil = stencil.Masters.get_ItemU(
                            COMPOSITION_MASTER);
                    }

                    // Drop the dynamic connector on the active page.
                    connector = visioApplication.ActivePage.Drop(
                        masterInStencil, 0, 0);

                    // Connect the begin point of the dynamic connector to the
                    // PinX cell of the first 2-D shape.
                    beginX = connector.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowXForm1D,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.vis1DBeginX);

                    beginX.GlueTo(shapeFrom.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowXFormOut,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visXFormPinX));

                    // Connect the end point of the dynamic connector to the
                    // PinX cell of the second 2-D shape.
                    endX = connector.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowXForm1D,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.vis1DEndX);

                    if (connectorType == ConnectorArrows.Unknown)
                    {
                        connector.CellsU["EndArrow"].set_Result(VisUnitCodes.visNumber, 11);
                    }
                    else if (connectorType != ConnectorArrows.Composition && connectorType != ConnectorArrows.InterfaceInheritance && connectorType != ConnectorArrows.EnumerationDependency)
                        connector.CellsU["EndArrow"].set_Result(VisUnitCodes.visNumber, (int)connectorType);

                    endX.GlueTo(shapeTo.get_CellsSRC(
                        (short)Microsoft.Office.Interop.Visio.
                            VisSectionIndices.visSectionObject,
                        (short)Microsoft.Office.Interop.Visio.
                            VisRowIndices.visRowXFormOut,
                        (short)Microsoft.Office.Interop.Visio.
                            VisCellIndices.visXFormPinX));

                    visioApplication.ActiveWindow.DeselectAll();
                }
                else
                {
                    // Processing cannot continue because the shapes are not on
                    // the same page.
                    System.Diagnostics.Debug.WriteLine(MESSAGE_NOT_SAME_PAGE);
                }
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }
        }


        private class TypeContainer
        {
            public List<TypeContainer> ChildTypes = new List<TypeContainer>();
            public Type ThisType;
            public int Depth;
        }

        private struct MyVector2
        {
            public float x;
            public float y;
        }
    }

   
}