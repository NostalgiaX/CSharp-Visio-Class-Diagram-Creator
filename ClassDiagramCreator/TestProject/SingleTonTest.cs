using ClassDiagramMaker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    [ClassDiagram]
    class SingleTonTest
    {
        private static SingleTonTest instance;
        private List<TestOne> listForProperty;
        private int intForProperty;

        public static SingleTonTest Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new SingleTonTest();
                }
                return instance;
            }
        }

        internal List<TestOne> Lololol { get => listForProperty; set => listForProperty = value; }
        public int Loloel { get => intForProperty; set => intForProperty = value; }

        public void SomeFunc()
        {

        }
        public static void SomeStaticFunc()
        {

        }

        public void SomeNonStaticFunc()
        {

        }
    }
}
