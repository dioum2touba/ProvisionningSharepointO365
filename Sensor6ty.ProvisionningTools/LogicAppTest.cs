using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sensor6ty.ProvisionningTools.TasksList;
using Sensor6ty.ProvisionningTools.Utility;
using Sensor6ty.ProvisionningTools.Website;

namespace Sensor6ty.ProvisionningTools
{
    class LogicAppTest
    {
        static void Main(string[] args)
        {
            var context = Utils.GetAuthentificateClient();

            // Create sharepoint website
            // WebsiteTools.CreateNewSharePointWebsite(context);

            // Create list from sharepoint website
            // TasksListTools.CreateListOnSharePoint(context);

            // Retrieve lists from sharepoint website
            // TasksListTools.RetrieveAllSharePointLists(context);

            // Create task on sharepoint list
            // TasksListTools.CreateTaskOnSharepointList(context);

            // Retrieve task from sharepoint list
            // TasksListTools.GetTasksFromSharepointList(context);

            Console.WriteLine("We're done. Press Enter to continue.");
            Console.ReadLine();
        }
    }
}
