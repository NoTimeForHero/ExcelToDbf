using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Jint;

namespace ExcelToDbf.Core.Services.Scripts.Context
{
    internal class AbstractContext
    {
        protected Engine engine;

        public AbstractContext(Engine engine)
        {
            this.engine = engine;
        }
    }
}
