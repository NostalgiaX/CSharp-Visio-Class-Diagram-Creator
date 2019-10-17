using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    [ClassDiagram]
    internal class TestOne : TestInheritance, TestInterface
    {
      
        private int SetOnlyInt;

        private TestTwo[] ObjectArray;
        private string internalstring;
        protected int someprotecInt;

        public int Lol { set => this.SetOnlyInt = value; }
        internal string Internalstring { get => this.internalstring; set => this.internalstring = value; }

        public TestStruct someFunc()
        {
            return new TestStruct();
        }

        public static void SomeFunction() { }
        public void SomeNonStaticFunc() { }
        public TestOne(int one)
        {

        }
        public TestOne(int one, int two)
        {

        }

        public void TestRefList(ref int[] dafq) { }
    }

    [ClassDiagram]
    internal class TestTwo : TestInheritance
    {

        internal string internalstring;
        protected int someprotecInt;



    }
    [ClassDiagram]
    internal class TestThree : TestInheritance
    {
    
        internal string internalstring;
        protected int someprotecInt;




    }
    [ClassDiagram]
    internal class TestFour : TestThree
    {

        internal string internalstring;
        protected int someprotecInt;



    }
}
