using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{

    [ClassDiagram]
    static class SomeStaticClass
    {
        public static int something = 2;
        private static int somethingStatic = 2;
        static SomeStaticClass()
        {

        }
        private static void SomeVoidFunc()
        {

        }

    }
}
