using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    [ClassDiagram]
    abstract class AbstractTest
    {

        public int PublicAbstractInt = 2;
        public TestInterface SomeInterface;
        public AbstractTest SelfReferenceWorks;
        public static string StaticStringTest;
        public abstract float SomeAbstractFloatFunction(int one);
    }
}
