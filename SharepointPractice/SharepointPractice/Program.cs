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
using System.IO;
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
                  // DisplayEmployeeListItems(clientContext);

                /********Create Subsite*********************/
                //   CreateNewSubsite(clientContext);

                /************Display User group for a site************/
                //   DisplayUserGroup(clientContext);

                /****************Retrive all lists in a site***************/
                //   DisplayAllListsInSite(clientContext);

                /****************Create list***********************/
                //   CreateList(clientContext);

                /******************Delete list***************/
                //   DeleteList(clientContext);

                /**************************Create Folder***************************/
                //   CreateFolder(clientContext);

                /******************Add Column to a specific list*****************/
                //AddColumnToSpecificList(clientContext);

                /******************Delete Column from a specific list*****************/
                //DeleteColumnFromSpecificList(clientContext);

                /********************Upload File In Folder**************/
                //UploadFileInFolder(clientContext);

                /******************Create Folder in Specific Document Library List******************/
                // CreateFolderInDocLibList(clientContext);

                /***************Delete User from Group********************/
                DeleteUserFromGroup(clientContext);

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

            /****************Add Items in list***************/
            //ListItemCreationInformation lici = new ListItemCreationInformation();
            //ListItem addItem = emplist.AddItem(lici);
            //addItem["Title"] = "Final Ra";
            //addItem["Department"] = "1";

            //try
            //{
            //    addItem.Update();
            //    clientContext.ExecuteQuery();
            //    Console.WriteLine("Insert success");
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine("Exce: "+e);
            //}

            /****************Edit Item in list***************/
            ListItem updateItem = emplist.GetItemById(7);
            updateItem["Title"] = "Updated Title1";
            updateItem["FirstName"] = "First Update";
            updateItem["Company"] = "Update Company";
            updateItem["Email%5fAddress"] = "Updated Email";
            updateItem["Department"] = "3";

            try
            {
                updateItem.Update();
                clientContext.ExecuteQuery();
                Console.WriteLine("Update success");
            }
            catch (Exception e)
            {
                Console.WriteLine("Exce: " + e);
            }

            /****************Delete Items in list***************/



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

        /********************Upload File In Folder**************/
        public static void UploadFileInFolder(ClientContext clientContext)
        {
            Console.WriteLine("---------------Uploading file in Folder--------------");
            var newfile = @"D:/DharanendraCode/SharepointPractice/SharepointPractice/errorlist.txt";

            FileCreationInformation file = new FileCreationInformation();
            file.Content = System.IO.File.ReadAllBytes(newfile);
            file.Overwrite = true;
            file.Url = Path.Combine("DemoLibrary/spDeletef/", Path.GetFileName(newfile));

            List l = clientContext.Web.Lists.GetByTitle("DemoLibrary");
            var f = l.RootFolder.Files.Add(file);

            try
            {
                clientContext.Load(f);
                clientContext.ExecuteQuery();
                Console.WriteLine("Uploaded Successfully");

                var files = l.RootFolder.Folders.GetByUrl("spDeletef/").Files.GetByUrl("errorlist.txt");

                //clientContext.Load(files);

                try
                {
                    files.DeleteObject();
                    clientContext.ExecuteQuery();
                    Console.WriteLine("Delete file suc");
                    var folder = l.RootFolder.Folders.GetByUrl("spDeletef/");

                    try
                    {
                        folder.DeleteObject();
                        clientContext.ExecuteQuery();
                        Console.WriteLine("Delete fol suc");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Exception : " + e);
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exc: " + e);
                }

                //files.DeleteObject();
                //clientContext.ExecuteQuery();

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e);
            }


        }

        /******************Create Folder in Specific Document Library List******************/
        public static void CreateFolderInDocLibList(ClientContext clientcontext)
        {
            List l = clientcontext.Web.Lists.GetByTitle("DemoLibrary");
            clientcontext.Load(l);
            clientcontext.ExecuteQuery();

            l.RootFolder.AddSubFolder("myfantfold");// Folder RootFolder;
            try
            {
                clientcontext.ExecuteQuery();
                Console.WriteLine("Folder created successfully");
            }
            catch (Exception e)
            {
                Console.WriteLine("exc : " + e);
            }
        }

        /******************Add Column to a specific list*****************/
        public static void AddColumnToSpecificList(ClientContext clientCntx)
        {
            Console.WriteLine("-----------------Add Column to Specific List---------------- ");
            //Console.WriteLine("Enter Column Name");
            //List l = clientCntx.Web.Lists.GetByTitle(Console.ReadLine());
            //clientCntx.Load(l);
            //clientCntx.ExecuteQuery();
            //string s = "Nationality," + FieldType.Text + "," + true;
            //l.Fields.Add();

            ////AddFieldOptions f = new AddFieldOptions();

            ////clientCntx.Web.Lists.GetByTitle(Console.ReadLine()).Fields.Add();

            //List list = clientCntx.Web.Lists.GetByTitle(Console.ReadLine());
            //clientCntx.Load(list);

            //Field field = clientCntx.Web.Lists.GetByTitle("").Fields.Add()
            //list.Fields.Add() AddFieldAsXml(@"<field Name='Sex' DisplayName='Gender' type='text' Required='FALSE'><Default>[male]</Default></Field>", true, AddFieldOptions.DefaultValue);
            //list.Update();
            //context.ExecuteQuery();
            Console.WriteLine("Enter List Name");
            List list = clientCntx.Web.Lists.GetByTitle(Console.ReadLine());

            Field field = list.Fields.AddFieldAsXml(@"<Field Name='Nationality' DisplayName='Nationality' Key='Nationality' Type='Text' Required='FALSE'/>", true, AddFieldOptions.DefaultValue);

            field.Update();
            clientCntx.ExecuteQuery();
        }

        /******************Delete Column from a specific list*****************/
        public static void DeleteColumnFromSpecificList(ClientContext clientCntx)
        {
            Console.WriteLine("-----------------Delete Column from a Specific List---------------- ");
            Console.WriteLine("Enter List Name");
            List l = clientCntx.Web.Lists.GetByTitle(Console.ReadLine());
            //clientCntx.Load(l);
            //clientCntx.ExecuteQuery();
            Console.WriteLine("Enter Field Title");
            Field f = l.Fields.GetByTitle(Console.ReadLine());
            f.DeleteObject();
            clientCntx.ExecuteQuery();
        }

        /**********************Add User in a group***********************/
        public static void DeleteUserFromGroup(ClientContext clientCntx)
        {
           
            UserCollection uc = clientCntx.Web.SiteGroups.GetByName("InsertUsers").Users;
            clientCntx.Load(uc);
            clientCntx.ExecuteQuery();

            /*****************To Delete User From Specific Group**************/
       /*     foreach (User ur in uc)
            {
                if(ur.Email == "venu.kalam@acuvate.com")
                {
                    try
                    {
                        User u1 = uc.GetByEmail(ur.Email);
                        uc.Remove(u1);
                        clientCntx.ExecuteQuery();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Exe  :"+e);
                    }
                }
            }
            */

           Console.WriteLine("-----------------Add user in group---------------- ");
            UserCreationInformation userCreationInformation = new UserCreationInformation();
            //User u = clientCntx.Web.user
            userCreationInformation.Email = "venu.kalam@acuvate.com";

            userCreationInformation.LoginName = "venu.kalam@acuvate.com";
            try
            {
                User user1 = clientCntx.Web.SiteGroups.GetByName("InsertUsers").Users.Add(userCreationInformation);
                user1.Update();
                clientCntx.ExecuteQuery();
                Console.WriteLine("Venu added");
            }
            catch (Exception e)
            {
                Console.WriteLine("eer "+e);
            }
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


