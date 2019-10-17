using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject
{
    [ClassDiagramMaker.ClassDiagram]
    class GenericClass<T>
    {
        private T obj;

        public GenericClass(T something)
        {
            this.obj = something;
        }
        public T GetT()
        {
            return obj;
        }
    }
}
