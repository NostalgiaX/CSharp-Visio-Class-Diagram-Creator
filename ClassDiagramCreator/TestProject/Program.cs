using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    class Program
    {
        static void Main(string[] args)
        {
            ClassDiagramCreator.MakeClassDiagram(ArrowsToInclude.Inheritance | ArrowsToInclude.SuggestedRelationArrows);

        }
    }
}
