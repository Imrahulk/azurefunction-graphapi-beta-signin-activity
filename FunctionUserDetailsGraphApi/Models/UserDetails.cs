using System.Collections.Generic;
using Newtonsoft.Json;

namespace FunctionUserDetailsGraphApi.Models
{
    internal class UserDetails
    {
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }

        [JsonProperty("@odata.nextLink")]
        public string OdataNextLink { get; set; }
        public List<Value> value { get; set; }
    }
}
