using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    [ClassDiagramInterface]
    interface TestInterface
    {

        void SomeTestFunc();
        int SomeOtherTestFunc(int i);
    }
}
