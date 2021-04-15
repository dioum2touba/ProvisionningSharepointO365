using System;
using System.Linq;
using System.Security;
using System.Threading;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using Sensor6ty.ProvisionningTools.Utility;

namespace Sensor6ty.ProvisionningTools.SPSite
{
    public class SPSiteTools
    {
        #region Connect AAD & Manage sharepoint website
        /// <summary>
        /// Export the template from the sharepoint site
        /// </summary>
        /// <param name="defaultForeground">Default foreground</param>
        /// <param name="webUrl">web site url</param>
        /// <param name="userName">user name</param>
        /// <param name="pwd">password</param>
        /// <returns></returns>
        public static ProvisioningTemplate GetProvisioningTemplate(ConsoleColor defaultForeground, string webUrl, string userName, SecureString pwd)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is:" + ctx.Web.Title);
                Console.ForegroundColor = defaultForeground;

                ProvisioningTemplateCreationInformation ptci
                    = new ProvisioningTemplateCreationInformation(ctx.Web);

                // Create FileSystemConnector to store a temporary copy of the template
                ptci.FileConnector = new FileSystemConnector("C:\\Users\\DIOUM2TOUBA\\pnpprovisioningdemo", "C:\\Users\\DIOUM2TOUBA\\pnpprovisioningdemo");
                ptci.PersistComposedLookFiles = true;
                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    // Only to output progress for console UI
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Execute actual extraction of the template
                ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);


                // We can serialize this template to save and reuse it
                // Optional step
                XMLTemplateProvider provider =
                    new XMLFileSystemTemplateProvider(@"C:\Users\DIOUM2TOUBA\pnpprovisioningdemo", "");
                provider.SaveAs(template, "PnPProvisioningDemo.xml");

                return template;
            }
        }

        /// <summary>
        /// Apply a template on Sharepoint site
        /// </summary>
        /// <param name="targetWebUrl">Url Web Site template</param>
        /// <param name="userName">User Name</param>
        /// <param name="pwd">Password</param>
        /// <param name="template">The template from Sharepoint</param>
        public static void ApplyProvisioningTemplate(string targetWebUrl, string userName, SecureString pwd, ProvisioningTemplate template)
        {
            using (var ctx = new ClientContext(targetWebUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                Web web = ctx.Web;

                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Associate file connector for assets
                FileSystemConnector connector = new FileSystemConnector(@"C:\Users\DIOUM2TOUBA\pnpprovisioningdemo", "");
                template.Connector = connector;

                // Because the template is actual object, we can modify this using code as needed
                template.Lists.Add(new ListInstance()
                {
                    Title = "PnP Sample Contacts",
                    Url = "lists/PnPContacts",
                    TemplateType = (Int32)ListTemplateType.Contacts,
                    EnableAttachments = true
                });

                web.ApplyProvisioningTemplate(template, ptai);
            }
        }

        /// <summary>
        /// Get properties from Sharepoint website
        /// </summary>
        /// <param name="context">Client context</param>
        public static void GetPropertiesSharepointSite(ClientContext context)
        {

            // The SharePoint web at the URL.
            Web web = context.Web;

            // We want to retrieve the web's properties.
            context.Load(web);

            // Execute the query to the server.
            context.ExecuteQuery();

            // Now, the web's properties are available and we could display
            // web properties, such as title.
            Console.WriteLine("Get properties from Sharepoint website");
            Console.WriteLine(web.Title);
            Console.WriteLine(web.Description);
            Console.WriteLine("------------------------------------- Fin Get properties ----------------------------------");
        }

        /// <summary>
        /// Set properties from Sharepoint website
        /// </summary>
        /// <param name="context">The client context object</param>
        public static void SetPropertiesSharepointSite(ClientContext context)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;
            string title = Utils.GetInput("Give the new title", false, defaultForeground);
            string description = Utils.GetInput("Give the new description", false, defaultForeground);
            // The SharePoint web at the URL.
            Web web = context.Web;

            web.Title = title;
            web.Description = description;

            // Note that the web.Update() doesn't trigger a request to the server.
            // Requests are only sent to the server from the client library when
            // the ExecuteQuery() method is called.
            web.Update();

            // Execute the query to server.
            context.ExecuteQuery();
            Console.WriteLine("------------------------------------- End Set properties ----------------------------------");
        }

        /// <summary>
        /// Create a new SharePoint website
        /// </summary>
        /// <param name="context"></param>
        public static void CreateNewSharePointWebsite(ClientContext context)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;
            string title = Utils.GetInput("Give the new title for the new", false, defaultForeground);
            string description = Utils.GetInput("Give the new description for the new", false, defaultForeground);

            WebCreationInformation creation = new WebCreationInformation();
            creation.Url = title;
            creation.Title = description;
            Web newWeb = context.Web.Webs.Add(creation);

            // Retrieve the new web information.
            context.Load(newWeb, w => w.Title);
            context.ExecuteQuery();

            Console.WriteLine(newWeb.Title);
            Console.WriteLine("------------------------------------- Site successfull created ----------------------------------");
        }
        #endregion
    }
}
