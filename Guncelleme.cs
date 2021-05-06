using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sınav_Planlama_Cetveli
{
    class Guncelleme
    {
        public bool status;
        public string statusComment;
        public string version;
        public string path;

        public Guncelleme()
        {
            status = true;
            statusComment = "";
            version = "";
            path = "";
        }
    }
}
