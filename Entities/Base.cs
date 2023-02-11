using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LearnWinInterop.Entities
{
    public partial class LaboratoryEntities
    {
        private static LaboratoryEntities Context;
        public static LaboratoryEntities GetContext()
        {
            if(Context == null)
                Context = new LaboratoryEntities();
            return Context;
        }
    }
}
