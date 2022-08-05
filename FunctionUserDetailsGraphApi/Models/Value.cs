
namespace FunctionUserDetailsGraphApi.Models
{
    public class Value
    {
        public string displayName { get; set; }
        public string id { get; set; }
        public SignInActivity signInActivity { get; set; }
        public Manager manager { get; set; }
    }
}
