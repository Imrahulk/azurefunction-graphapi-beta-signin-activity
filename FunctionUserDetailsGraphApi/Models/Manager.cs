using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace FunctionUserDetailsGraphApi.Models
{
    public class Manager
    {
        [JsonProperty("@odata.type")]
        public string OdataType { get; set; }
        public string givenName { get; set; }
        public string mail { get; set; }
    }
}
