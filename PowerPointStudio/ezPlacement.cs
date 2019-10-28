using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointStudio
{
    public class ezPlacement
    {
        public ezOnSlide onSlide;

        [JsonConstructor]
        public ezPlacement()
        {

        }
    }
}
