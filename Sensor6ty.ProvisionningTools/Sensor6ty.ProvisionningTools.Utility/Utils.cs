using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Sensor6ty.ProvisionningTools.Utility
{
    public class Utils
    {
        #region Id Netexware
        public static string siteCollectionUrl = "https://netexware.sharepoint.com";
        public static string webSiteUrl = "https://netexware.sharepoint.com/PnpTemplate%20Sensor6ty";
        public static string siteUrlAdmin = "https://netexware.sharepoint.com/sites/bambey/_layouts/15/viewlsts.aspx?view=14";
        public static string targetWebUrl = "https://netexware.sharepoint.com/sites/bambey";
        public static string userName = "admin@netexware.onmicrosoft.com";
        public static string password = "bbsemou2010#";
        #endregion

        /// <summary>
        /// Connect the client to the tenant AAD
        /// </summary>
        /// <returns>ClientContext</returns>
        public static ClientContext GetAuthentificateClient()
        {
            //Namespace: It belongs to Microsoft.SharePoint.Client
            ClientContext context = new ClientContext(webSiteUrl);
            ClientContext contextAdmin = new ClientContext(siteUrlAdmin);
            Utils obj = new Utils();
            try
            {
                context = obj.ConnectToSharePointOnline(context, userName, password);
            }
            catch (Exception ex)
            {
                string msg = ex.Message.ToString();

            }
            SecureString pwd = new SecureString();
            foreach (char c in password.ToCharArray()) pwd.AppendChar(c);

            return context;
        }

        /// <summary>
        /// Retrieve the enter values from the keyboard
        /// </summary>
        /// <param name="label"></param>
        /// <param name="isPassword"></param>
        /// <param name="defaultForeground"></param>
        /// <returns></returns>
        public static string GetInput(string label, bool isPassword, ConsoleColor defaultForeground)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("{0} : ", label);
            Console.ForegroundColor = defaultForeground;

            string value = "";

            for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
            {
                if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    if (value.Length > 0)
                    {
                        value = value.Remove(value.Length - 1);
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    }
                }
                else if (keyInfo.Key != ConsoleKey.Enter)
                {
                    if (isPassword)
                    {
                        Console.Write("*");
                    }
                    else
                    {
                        Console.Write(keyInfo.KeyChar);
                    }
                    value += keyInfo.KeyChar;
                }
            }
            Console.WriteLine("");

            return value;
        }


        /// <summary>
        /// Connect to sharepoint online
        /// </summary>
        /// <param name="siteCollUrl">Site url</param>
        /// <param name="userName">user name</param>
        /// <param name="password">password</param>
        public ClientContext ConnectToSharePointOnline(ClientContext ctx, string userName, string password)
        {
            // Namespace: It belongs to System.Security
            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);

            // Namespace: It belongs to Microsoft.SharePoint.Client
            ctx.Credentials = new SharePointOnlineCredentials(userName, secureString);

            // Namespace: It belongs to Microsoft.SharePoint.Client
            Site mySite = ctx.Site;

            ctx.Load(mySite);
            ctx.ExecuteQuery();

            Console.WriteLine(mySite.Url.ToString());
            return ctx;
        }
    }
}
