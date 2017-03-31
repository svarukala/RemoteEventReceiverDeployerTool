using System;
using System.Security;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Configuration;

namespace DeployRER
{
    class Program
    {
        private static string userName = "admin@mod362200.onmicrosoft.com";
        private static SecureString password;
        private static string receiverUrl;

        public static SecureString GetPassword()
        {
            var pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            return pwd;
        }

        static void Main(string[] args)
        {
            #region secure
            userName = ConfigurationManager.AppSettings["userid"] ?? null;
            if(userName.Length == 0)
            {
                Console.WriteLine("Provide user id in the app.config file.");
                return;
            }
            string passString = ConfigurationManager.AppSettings["password"] ?? null;
            password = new SecureString();
            if (passString.Length == 0)
            {
                Console.WriteLine("Please enter the password");
                password = GetPassword();
            }
            else
            {
                foreach (char c in passString)
                {
                    password.AppendChar(c);
                }
            }
            #endregion

            receiverUrl= ConfigurationManager.AppSettings["receiverurl"] ?? null;
            if (receiverUrl.Length == 0)
            {
                Console.WriteLine("Provide the receiver url in the app.config file.");
                return;
            }

            string pubsite = ConfigurationManager.AppSettings["siteurl"] ?? null;
            if(pubsite.Length == 0)
            {
                Console.WriteLine("Provide the site url in the app.config file.");
                return;
            }
            string listName = ConfigurationManager.AppSettings["listname"] ?? null;
            if (listName.Length == 0)
            {
                Console.WriteLine("Provide the list name in the app.config file.");
                return;
            }
            string receiverName = ConfigurationManager.AppSettings["receivername"] ?? null;
            if (receiverName.Length == 0)
            {
                Console.WriteLine("Provide the receiver name in the app.config file.");
                return;
            }
            var eventType = EventReceiverType.ItemDeleting;
            try
            {
                string et = ConfigurationManager.AppSettings["eventtype"];
                if (et.Length == 0) throw new Exception();
                eventType = (EventReceiverType)Enum.Parse(typeof(EventReceiverType), et);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Provide a valid event type in app.config file");
                return;
            }

            Console.WriteLine("User Id: {0}\nSite url: {1}\nList name: {2}\nReceiver name: {3}\nEvent type: {4}\nReceiver url: {5}\n", userName, pubsite, listName, receiverName, eventType.ToString(), receiverUrl);


            bool exit = false;
            while (!exit)
            {
                Console.WriteLine("\n***************************");
                Console.WriteLine("Select one of the options:");
                Console.WriteLine("1. List all event receivers for a list/library");
                Console.WriteLine("2. Attach event receiver to the list");
                Console.WriteLine("3. Remove event receiver from the list");
                Console.WriteLine("4. Attach event receiver at Site Collection level");
                Console.WriteLine("5. Remove event receiver at Site Collection level");
                Console.WriteLine("6. Attach event receiver at Web level");
                Console.WriteLine("7. Remove event receiver at Web level");
                Console.WriteLine("8. List all event receivers at Site Collection level");
                Console.WriteLine("9. List all event receivers at Web level");
                Console.WriteLine("10. Exit");
                Console.WriteLine("***************************");

                string select = Console.ReadLine();


                switch (select)
                {
                    case "1":
                        GetReceivers(pubsite, listName);
                        break;
                    case "2":
                        AttachReceiver(pubsite, listName, receiverName, eventType);
                        break;
                    case "3":
                        RemoveReceiver(pubsite, receiverName, listName);
                        break;
                    case "4":
                        AttachReceiver(pubsite, receiverName, eventType, true);
                        break;
                    case "5":
                        RemoveReceiver(pubsite, receiverName, true);
                        break;
                    case "6":
                        AttachReceiver(pubsite, receiverName, eventType, false);
                        break;
                    case "7":
                        RemoveReceiver(pubsite, receiverName, false);
                        break;
                    case "8":
                        GetReceivers(pubsite, true);
                        break;
                    case "9":
                        GetReceivers(pubsite, false);
                        break;
                    case "10":
                        exit = true;
                        break;
                }
            }
            Console.WriteLine("Done. Press any key to exit.");
            Console.ReadLine();
        }

        private static void RemoveReceiver(string spUrl, string receiverName, bool isSiteCollectionLevel)
        {
            try
            {
                using (var ctx = new ClientContext(spUrl))
                {
                    ctx.Credentials = new SharePointOnlineCredentials(userName, password);

                    if (isSiteCollectionLevel)
                    {
                        ctx.Load(ctx.Web);
                        ctx.ExecuteQuery();
                        ReceiverHelper.RemoveEventReceiver(ctx, ctx.Site, receiverName);
                    }
                    else
                    {
                        ctx.Load(ctx.Web);
                        ctx.ExecuteQuery();
                        ReceiverHelper.RemoveEventReceiver(ctx, ctx.Web, receiverName);
                    }
                    Console.WriteLine("Removed the event receiver!");
                }
            }
            catch (Exception ex1)
            {
                var msg = ex1.ToString();
           }
        }
        private static void RemoveReceiver(string spUrl, string receiverName, string listName)
        {
            using (var ctx = new ClientContext(spUrl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                ctx.Load(ctx.Web);
                ctx.ExecuteQuery();

                try
                {

                    List _list = ctx.Web.Lists.GetByTitle(listName);
                    ReceiverHelper.RemoveEventReceiver(ctx, _list, receiverName);
                    Console.WriteLine("Removed the event receiver!");
                }
                catch (Exception ex1)
                {
                    var msg = ex1.ToString();
                }

                //try
                //{
                //    ReceiverHelper.RemoveEventReceiver(ctx, ctx.Web, receiverName);//.AddEventReceiver(ctx, _list, _rec);
                //}
                //catch (Exception ex1)
                //{
                //    var msg = ex1.ToString();
                //}

            }

        }
        private static void AttachReceiver(string spUrl, string listName, string rerName, EventReceiverType eventType)
        {
            using (var ctx = new ClientContext(spUrl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                ctx.Load(ctx.Web, w => w.Title);
                ctx.ExecuteQuery();

                try 
                {

                    List _list = ctx.Web.Lists.GetByTitle(listName);
                    EventReceiverDefinitionCreationInformation _rec = ReceiverHelper.CreateEventReciever(rerName, eventType, receiverUrl);
                    ReceiverHelper.AddEventReceiver(ctx, _list, _rec);
                    Console.WriteLine("Attached the event receiver!");
                }
                catch(Exception ex1)
                {
                    var msg = ex1.ToString();
                }

            }

        }
        private static void AttachReceiver(string spUrl, string rerName, EventReceiverType eventType)
        {
               Uri uri = new Uri(spUrl);
            var targetRealm = TokenHelper.GetRealmFromTargetUrl(uri);
            var targetPrincipalName = TokenHelper.SharePointPrincipal;
            var targetHost = new Uri(spUrl).Authority;
            var accessToken = TokenHelper.GetAppOnlyAccessToken(targetPrincipalName, targetHost, targetRealm).AccessToken;

            using (ClientContext ctx = TokenHelper.GetClientContextWithAccessToken(spUrl, accessToken))
            {

                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();

                try
                {
                    EventReceiverDefinitionCreationInformation _rec = ReceiverHelper.CreateEventReciever(rerName, eventType, receiverUrl);
                    ReceiverHelper.AddEventReceiver(ctx, web, _rec);
                    Console.WriteLine("Attached the event receiver!");
                }
                catch (Exception ex1)
                {
                    var msg = ex1.ToString();
                    Console.WriteLine(msg);
                }
            }
        }

        private static void AttachReceiver(string spUrl, string rerName, EventReceiverType eventType, bool isSiteCollectionLevel)
        {
            EventReceiverDefinitionCreationInformation _rec = ReceiverHelper.CreateEventReciever(rerName, eventType, receiverUrl);

            using (var ctx = new ClientContext(spUrl))
            {
                ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                try
                {                    
                    if(isSiteCollectionLevel)
                    {
                        Site site = ctx.Site;
                        ctx.Load(site);
                        ctx.ExecuteQuery();
                        ReceiverHelper.AddEventReceiver(ctx, site, _rec);
                    }
                    else
                    {
                        Web web = ctx.Web;
                        ctx.Load(web);
                        ctx.ExecuteQuery();
                        ReceiverHelper.AddEventReceiver(ctx, web, _rec);
                    }
                    Console.WriteLine("Attached the event receiver!");
                }
                catch (Exception ex1)
                {
                    var msg = ex1.ToString();
                    Console.WriteLine(msg);
                }
            }
        }
        private static void GetReceivers(string spUrl, string listName)
        {
            //string spUrl = "https://varuk.sharepoint.com/sites/dev"; //"https://varuk.sharepoint.com/sites/pubsite/en-us";
            //string webUrl = "https://rerhost.azurewebsites.net";
            //string userName = "srini@varuk.onmicrosoft.com";
            //Console.WriteLine("Enter your password.");
            //SecureString password = GetPasswordFromConsoleInput();

            using (var context = new ClientContext(spUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                ListReceivers(context,  listName);
                //RegisterReceiver(context, EventReceiverType.ItemAdding, EventReceiverSynchronization.Synchronous, listName, webUrl);
            }

        }
        private static void GetReceivers(string spUrl, bool isSiteCollectionLevel)
        {
            
            //Console.WriteLine("Enter your password.");
            //SecureString password = "Corp123!1";//GetPasswordFromConsoleInput();
            Console.WriteLine("Remote event receivers for site:" + spUrl);
            using (var context = new ClientContext(spUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                ListReceivers(context, isSiteCollectionLevel);
            }

        }
        private static SecureString GetPasswordFromConsoleInput()
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
        private static void ListReceivers(ClientContext clientContext, string listName)
        {
            List targetList = clientContext.Web.Lists.GetByTitle(listName);
            clientContext.Load(targetList);

            EventReceiverDefinitionCollection ec = targetList.EventReceivers;
            clientContext.Load(ec);
            clientContext.ExecuteQuery();

            //Get rid of old rer registration in the case that we are re-deploying
            for (int i = 0; i < ec.Count; i++)
            {
                Console.WriteLine("{0} - {1} - {2}", ec[i].ReceiverName, ec[i].EventType.ToString(), ec[i].ReceiverUrl);
            }
        }
        private static void ListReceivers(ClientContext clientContext, bool isSiteCollectionLevel)
        {
            EventReceiverDefinitionCollection ec;
            if (isSiteCollectionLevel)
            {
                Site site = clientContext.Site;
                clientContext.Load(site);
                ec = site.EventReceivers;
                //clientContext.Load(ec);
                //clientContext.ExecuteQuery();
            }
            else
            {
                Web web = clientContext.Web;
                clientContext.Load(web);
                ec = web.EventReceivers;
                //clientContext.Load(ec);
                //clientContext.ExecuteQuery();
            }

            clientContext.Load(ec);
            clientContext.ExecuteQuery();

            //Get rid of old rer registration in the case that we are re-deploying
            for (int i = 0; i < ec.Count; i++)
            {
                Console.WriteLine("{0} - {1} - {2}", ec[i].ReceiverName, ec[i].EventType.ToString(), ec[i].ReceiverUrl);
            }
        }
        private static void GetPropBagValues(string spUrl)
        {

            //Console.WriteLine("Enter your password.");
            //SecureString password = "Corp123!1";//GetPasswordFromConsoleInput();
            Console.WriteLine("Remote event receivers for site:" + spUrl);
            using (var context = new ClientContext(spUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                Web web = context.Web;
                context.Load(web, w => w.AllProperties);
                context.ExecuteQuery();
                if (web.AllProperties.FieldValues.ContainsKey("IsOnHold"))
                {
                    Console.WriteLine("IsOnHold: {0}", ((string)web.AllProperties["IsOnHold"] == "1") ? "TRUE" : "FALSE");
                }
                else
                    Console.WriteLine("No prop bag entry");
            }

        }
        private static void DeleteSite(string spUrl)
        {
            string adminUrl = "https://varuk-admin.sharepoint.com";
            Uri uri = new Uri(adminUrl);
            var targetRealm = TokenHelper.GetRealmFromTargetUrl(uri);
            var targetPrincipalName = TokenHelper.SharePointPrincipal;
            var targetHost = new Uri(adminUrl).Authority;
            var accessToken = TokenHelper.GetAppOnlyAccessToken(targetPrincipalName, targetHost, targetRealm).AccessToken;

            using (ClientContext tenantContext = TokenHelper.GetClientContextWithAccessToken(adminUrl, accessToken))
            {                
                var tenant = new Tenant(tenantContext);
                SpoOperation spoOperation = tenant.RemoveSite(spUrl);
                tenantContext.Load(spoOperation);
                //tenantContext.RequestTimeout = 2;
                tenantContext.ExecuteQuery();
            }
        }

        private static void DeleteWeb(string spUrl)
        {
            Uri uri = new Uri(spUrl);
            var targetRealm = TokenHelper.GetRealmFromTargetUrl(uri);
            var targetPrincipalName = TokenHelper.SharePointPrincipal;
            var targetHost = new Uri(spUrl).Authority;
            var accessToken = TokenHelper.GetAppOnlyAccessToken(targetPrincipalName, targetHost, targetRealm).AccessToken;

            using (ClientContext ctx = TokenHelper.GetClientContextWithAccessToken(spUrl, accessToken))
            {
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                web.DeleteObject();
                ctx.ExecuteQuery();
                Console.WriteLine("Web deleted");
            }
        }

    }
}
