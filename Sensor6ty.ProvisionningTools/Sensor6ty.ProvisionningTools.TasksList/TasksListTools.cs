using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Sensor6ty.ProvisionningTools.Utility;

namespace Sensor6ty.ProvisionningTools.TasksList
{
    // Les exemples suivants indiquent comment utiliser le modèle objet client .NET Framework pour effectuer des tâches en lien avec la liste
    public class TasksListTools
    {
        #region SharePoint List Tasks
        /// <summary>
        /// Retrieve all SharePoint lists on a website
        /// </summary>
        /// <param name="context"></param>
        public static void RetrieveAllSharePointLists(ClientContext context)
        {
            // The SharePoint web at the URL.
            Web web = context.Web;

            // Retrieve all lists from the server.
            // For each list, retrieve Title and Id.
            context.Load(web.Lists,
                lists => lists.Include(list => list.Title,
                    list => list.Id));

            // Execute query.
            context.ExecuteQuery();

            // Enumerate the web.Lists.
            web.Lists.ToList().ForEach((elt) => Console.WriteLine("Id: {0} \nTitle: {1} \n\n", elt.Id, elt.Title));
            Console.WriteLine(
                "--------------------------------- End of sharepoint list --------------------------------------------------");
        }

        /// <summary>
        /// Create and update a SharePoint list
        /// </summary>
        /// <param name="context"></param>
        public static void CreateListOnSharePoint(ClientContext context)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;
            string title = Utils.GetInput("Give the new title for the list", false, defaultForeground);
            string description = Utils.GetInput("Give the new description for the list", false, defaultForeground);

            // The SharePoint web at the URL.
            Web web = context.Web;

            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = title;
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;
            List list = web.Lists.Add(creationInfo);
            list.Description = description;

            list.Update();
            context.ExecuteQuery();

            RetrieveAllSharePointLists(context);
            Console.WriteLine(
                "--------------------------------- List created successfull on sharepoint website --------------------------------------------------");
        }

        /// <summary>
        /// Delete a SharePoint list
        /// </summary>
        /// <param name="context"></param>
        public static void DeleteListFromSharepointList(ClientContext context)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;
            string title = Utils.GetInput("Give the title list to deleted", false, defaultForeground);

            // The Sharepoint web at the url.
            Web web = context.Web;

            List list = web.Lists.GetByTitle(title);
            list.DeleteObject();

            context.ExecuteQuery();
            Console.WriteLine(
                $"--------------------------------- List deleted successfull from {title} sharepoint website --------------------------------------------------");
        }

        /// <summary>
        /// Retrieve items from a SharePoint list
        /// </summary>
        /// <param name="context"></param>
        public static void GetTasksFromSharepointList(ClientContext context)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;
            string title = Utils.GetInput("Give the title sharepoint list", false, defaultForeground);

            // Assume the web has a list named {Title}
            List titleList = context.Web.Lists.GetByTitle(title);

            // This creates a CalmQuery that has a RowLimit of 100, and also specifies Scope="RecursiveAll"
            // So that it grabs all list items, regardless of the folder they are in.
            CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
            ListItemCollection items = titleList.GetItems(query);

            // Retrieve all items in the ListItemCollection from List.GetItems(Query).
            context.Load(items);
            context.ExecuteQuery();
            foreach (ListItem listItem in items)
            {
                // We have all the list item data. For example, Title.
                if(listItem.FieldValues["Key"].Equals("ToDo") || listItem.FieldValues["Key"].Equals("Title") || listItem.FieldValues["Key"].Equals("Description"))
                    Console.WriteLine($"{listItem.FieldValues["Key"]}: {listItem.FieldValues["Value"]}");
            }
        }

        /// <summary>
        /// Create item on sharepoint list
        /// </summary>
        /// <param name="context"></param>
        public static void CreateTaskOnSharepointList(ClientContext context)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;
            string title = Utils.GetInput("Give the title of the list sharepoint", false, defaultForeground);
            string titleItem = Utils.GetInput("Give the title task to create on the list", false, defaultForeground);
            //string bodyItems = Utils.GetInput("Give the body task to create on the list", false, defaultForeground);
            string descriptionItems = Utils.GetInput("Give the description task to create on the list", false, defaultForeground);
            string toDoItems = Utils.GetInput("Give the status task to create on the list", false, defaultForeground);

            // Assume that the web has a list named "Announcements".
            List titleList = context.Web.Lists.GetByTitle(title);

            // We are just creating a regular list item, so we don't need to
            // set any properties. If we wanted to create a new folder, for
            // example, we would have to set properties such as
            // UnderlyingObjectType to FileSystemObjectType.Folder.
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem newItem = titleList.AddItem(itemCreateInfo);
            newItem["Title"] = titleItem;
            //newItem["Body"] = bodyItems;
            newItem["Description"] = descriptionItems;
            newItem["ToDo"] = toDoItems;
            newItem.Update();

            context.ExecuteQuery();
        }
        #endregion
    }
}
