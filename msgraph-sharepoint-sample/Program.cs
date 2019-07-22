using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace msgraph_sharepoint_sample
{
    public class Program
    {
        private readonly static string _groupId = "30ecd9bc-6e8b-4280-adbc-18dccc0815a6";
        private readonly static string _siteId = "titusgicheru.sharepoint.com";
        private static string _listId = null;

        static void Main(string[] args)
        {
            /* 
             * Display all Authors list items
             */
            GetAuthors();
            Console.ReadLine();

            /*
             * Add Author ListItem
            */
            AddAuthor();
            GetAuthors();
            Console.ReadLine();

            /*
             * Update Author ListItem
            */
            UpdateAuthor();
            GetAuthors();
            Console.ReadLine();
        }


        private async static Task<ISiteListsCollectionPage> GetList()
        {
            ISiteListsCollectionPage lists = await Sites.GetSiteLists(_groupId, _siteId);
            return lists;
        }

        private async static Task<IListItemsCollectionPage> GetListItems(string listId)
        {
            IListItemsCollectionPage listItems = await Sites.GetSiteListItems(_groupId, _siteId, listId);

            return listItems;
        }

        private async static void GetAuthors()
        {
            ISiteListsCollectionPage lists = await GetList();

            /* We will show existing site list 
             * and filter the Authors List for use 
             * in this sample  
             */
            var list = lists.Where(l => l.DisplayName.Contains("Authors")).FirstOrDefault();

            //Show existing Site List in the Current Site
            Console.WriteLine($"Display all {list.DisplayName}");
            Console.WriteLine("***************************");

            //assign the global listId for use in other methods 
            _listId = list.Id;

            //Getting listItems using msgraph
            IListItemsCollectionPage listItems = await GetListItems(list.Id);

            foreach (var item in listItems)
            {
                IDictionary<string, object> columns = item.Fields.AdditionalData;

                if (columns != null)
                {
                    string[] filterColumns = new string[] { "Title", "LastName" };
                    foreach (var col in columns)
                    {
                        if (filterColumns.Contains(col.Key))
                        {
                            Console.WriteLine($"Id: {item.Id} Field: {col.Key} Value: {col.Value}");
                        }
                    }
                }
            }
        }

        private async static void AddAuthor()
        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Add New Author");

            Console.WriteLine("Enter FirstName");
            string firstName = Console.ReadLine();
            data.Add("Title", firstName);

            Console.WriteLine("Enter LastName");
            string lastName = Console.ReadLine();
            data.Add("LastName", lastName);

            bool result = await Sites.CreateListItem(_groupId, _siteId, _listId, data);

            if (result)
                Console.WriteLine("Item Created");
            else
                Console.WriteLine("Item Not Created");
        }


        private async static void UpdateAuthor()
        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Update an Author");

            Console.WriteLine("Enter Id");
            string authorId = Console.ReadLine();

            Console.WriteLine("Enter FirstName");
            string firstName = Console.ReadLine();
            data.Add("Title", firstName);

            Console.WriteLine("Enter LastName");
            string lastName = Console.ReadLine();
            data.Add("LastName", lastName);

            bool result = await Sites.UpdateListItem(_groupId, _siteId, _listId, authorId, data);

            if (result)
                Console.WriteLine("Item Updated");
            else
                Console.WriteLine("Item Not Update");
        }
    }
}
