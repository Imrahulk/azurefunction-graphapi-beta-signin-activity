using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunctionUserDetailsGraphApi.Models
{
    public class SignInReport
    {
        public string DisplayName { get; set; }
        public string Id { get; set; }
        public DateTime LastSignInDateTime { get; set; }
        public string ManagerName { get; set; }
        public string ManagerEmail { get; set; }
    }
}
