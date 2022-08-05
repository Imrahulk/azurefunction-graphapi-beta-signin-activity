using System;
using System.Globalization;

namespace FunctionUserDetailsGraphApi.Models
{
    public class Constants
    {
        public static string ClientId => "provide graph api clientId";
        public static string AzureActiveDirectoryInstance => "https://login.microsoftonline.com/{0}";
        public static string TenantId => "provide graph api tenant id";
        public static string Resource => "https://graph.microsoft.com";
        public static string ClientSecret => "provide graph api secret";
        public static string TenantName => "GraphApiBuilder";
        public static string Authority => String.Format(CultureInfo.InvariantCulture, AzureActiveDirectoryInstance, TenantId);
        public static string BetaBaseUrl => "https://graph.microsoft.com/beta/users?$select=displayName,signInActivity&$expand=manager($select=givenName,mail)";
        public static DateTime LastLoginDate => DateTime.Today.AddDays(-45);

        public static string ContainerName => "provide your blob container id";
        public static string ConnectionString => "provide your blob connection string";
        public static string BlobName => "provide your blob name with extension";

    }
}
