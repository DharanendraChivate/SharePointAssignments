//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace SharepointPractice
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//        }
//    }
//}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            string userName = "dharanendra.sheetal@acuvate.com";
            Console.WriteLine("Enter your password.");

            SecureString password = GetPassword();

            // ClienContext - Get the context for the SharePoint Online Site  
            // SharePoint site URL - https://c986.sharepoint.com  
            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/SharePointDemo"))
            {
                // SharePoint Online Credentials  
                clientContext.Credentials = new SharePointOnlineCredentials(userName, password);
                // Get the SharePoint web  
                Web web = clientContext.Web;
                // Load the Web properties  
                clientContext.Load(web, w => w.Title, w => w.Description, w => w.Url);
                // Execute the query to the server.  
                clientContext.ExecuteQuery();
                // Web properties - Display the Title and URL for the web
                //   ClientRuntimeContext crt = new ClientRuntimeContext("https://acuvatehyd.sharepoint.com/teams/SharePointDemo");

                Console.WriteLine("Title: " + web.Title + "\n URL: " + web.Url + "\n Description: " + web.Description);

                web.Description = "DecriptionChanged";

                /**********************************************************/
                clientContext.Load(web);
                clientContext.ExecuteQuery();
                Console.WriteLine("Title: " + web.Title + "\n URL: " + web.Url + "\n Description: " + web.Description);

                /**********Employee List Display Data***********/
                  DisplayEmployeeListItems(clientContext);

                /********Create Subsite*********************/
                   CreateNewSubsite(clientContext);

                /************Display User group for a site************/
                DisplayUserGroup(clientContext);

                /****************Retrive all lists in a site***************/
                 DisplayAllListsInSite(clientContext);

                /****************Create list***********************/
                CreateList(clientContext);

                /******************Delete list***************/
                DeleteList(clientContext);

                /**************************Create Folder***************************/
                CreateFolder(clientContext);

                Console.ReadLine();
            }
        }

        /**********Employee List Display Data***********/
        public static void DisplayEmployeeListItems(ClientContext clientContext)
        {
            List emplist = clientContext.Web.Lists.GetByTitle("Employees");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><RowLimit></RowLimit></View>";

            ListItemCollection empcoll = emplist.GetItems(camlQuery);

            clientContext.Load(
                empcoll,

                items => items.Take(5).Include(
                    item => item["FirstName"],
                    item => item["Company"],
                    item => item["Department"]
                    )
                );
            clientContext.ExecuteQuery();
            foreach (ListItem employee in empcoll)
            {
                Console.WriteLine("\n First Name: {0} \n Company: {1}\n Department: {2}\n-----------------------\n", employee["FirstName"], employee["Company"], employee["Department".ToString()]);
            }

        }

        /********Create Subsite*********************/
        public static void CreateNewSubsite(ClientContext clientcntx)
        {

            //   clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
            WebCreationInformation crete = new WebCreationInformation();
            Console.WriteLine("Enter Site Name");


            crete.Url = Console.ReadLine().Trim().Replace(" ", "");
            Console.WriteLine("Enter the title for share point site");

            crete.Title = Console.ReadLine();
            clientcntx.Web.Webs.Add(crete);

            clientcntx.Load(clientcntx.Web, w => w.Title);
            try
            {
                clientcntx.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error : " + e);
                throw e;
            }
            Console.WriteLine("New site" + crete.Title);
            Console.ReadKey();

        }

        /*********Display User Group*****************/
        public static void DisplayUserGroup(ClientContext clientcntx)
        {
            UserCollection userlist = clientcntx.Web.SiteUsers;

            clientcntx.Load(userlist);
            clientcntx.ExecuteQuery();
            Console.WriteLine("------------------User List----------------------");
            foreach (User u in userlist)
            {
                Console.WriteLine("Name {0} \t email {1} \t Site Admin", u.LoginName, u.Email, u.IsSiteAdmin.ToString());
            }

        }

        /****************Retrive all lists in a site***************/
        public static void DisplayAllListsInSite(ClientContext clientcntx)
        {
            ListCollection allList = clientcntx.Web.Lists;

            clientcntx.Load(allList);
            clientcntx.ExecuteQuery();

            Console.WriteLine("------------------List Collection----------------------");
            foreach (List l in allList)
            {
                Console.WriteLine("List Name: {0} ", l.Title);
            }
        }

        /*****************Create List***************************/
        public static void CreateList(ClientContext clientcntx)
        {
            Console.WriteLine("-------------------------Create List-------------------------");
            ListCreationInformation listInfo = new ListCreationInformation();
            Console.WriteLine("Enter title");
            listInfo.Title = Console.ReadLine();
            Console.WriteLine("Enter url");
            listInfo.Url = Console.ReadLine().Trim().Replace(" ", "");
            Console.WriteLine("Enter description");
            listInfo.Description = Console.ReadLine();

            listInfo.TemplateType = (int)ListTemplateType.GenericList;

            clientcntx.Load(clientcntx.Web.Lists.Add(listInfo));

            try
            {
                clientcntx.ExecuteQuery();
                Console.WriteLine("Name of the list: " + clientcntx.Web.Lists.Add(listInfo).Title);
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e);
            }
        }

        /**********************List Deletion***************************/
        public static void DeleteList(ClientContext clientcntx)
        {
            Console.WriteLine("-------------------------Delete List-------------------------");

            Console.WriteLine("Enter List Ttle");
            List l = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());

            l.DeleteObject();
            try
            {
                clientcntx.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine("exc :" + e);
            }
        }

        /**************************Create Folder***************************/
        public static void CreateFolder(ClientContext clientCntx)
        {
            Console.WriteLine("--------------------Create Folder-------------------");

            Console.WriteLine("Enter Folder Name to be created");

            var list = clientCntx.Web.Lists.GetByTitle("Documents");
            var folder = list.RootFolder;
            clientCntx.Load(folder);

            try
            {
                clientCntx.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exc :" + e);
            }
            folder = folder.Folders.Add(Console.ReadLine());
            Console.Read();

            try
            {
                clientCntx.ExecuteQuery();
                Console.WriteLine("folder created");
            }
            catch (Exception e)
            {
                Console.WriteLine("Exc :" + e);
            }
        }

        public static void AddColumnToSpecificList(ClientContext clientCntx)
        {
            Console.WriteLine("-----------------Add Column to Specific List---------------- ");
            Console.WriteLine("Enter Column Name");
            List l = clientCntx.Web.Lists.GetByTitle(Console.ReadLine());
            clientCntx.Load(l);

            
        
        }

        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString  
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}


