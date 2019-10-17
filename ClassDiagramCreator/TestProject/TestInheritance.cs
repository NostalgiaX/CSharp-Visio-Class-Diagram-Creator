using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    [ClassDiagram]
    class TestInheritance : AbstractTest, TestInterface
    {

        public TestOne someTestOne;

        public override float SomeAbstractFloatFunction(int one)
        {
            throw new NotImplementedException();
        }

        public static TestOne SomeTestOne
        {
            get
            {
                return SomeTestOne;
            }
            set
            {
                SomeTestOne = value;
            }
        }

        public int SomeOtherTestFunc(int i)
        {
            throw new NotImplementedException();
        }

        public void SomeTestFunc()
        {
            throw new NotImplementedException();
        }
    }
}
