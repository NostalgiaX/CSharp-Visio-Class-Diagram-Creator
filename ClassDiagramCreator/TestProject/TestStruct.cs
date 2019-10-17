using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    [ClassDiagram]
    struct TestStruct
    {
        public int MyProperty { get; set; }
        public int StructInt;
        public List<int> someList;
        public int[] someArray;
        public int[,] some2DArray;

        public event Action<TestOne, TestOne> OnTestOneSomethingTwo;
        public event Action OnTestOneSomething;


    }
}
