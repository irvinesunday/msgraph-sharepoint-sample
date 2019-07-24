using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace msgraph_sharepoint_sample
{
    public class Program
    {
        private readonly static string _groupId = "30ecd9bc-6e8b-4280-adbc-18dccc0815a6";
        private readonly static string _siteId = "m365b267815.sharepoint.com,6e1261a1-6d03-432a-95c0-e1c7705aef5f,f43d258c-ece0-476a-a1c0-018d359817d5";
        private static string _listId = null;
        private static ISiteListsCollectionPage lists;
        private static List<OfficeBook> officeBooks = new List<OfficeBook>();

        static async Task Main(string[] args)
        {
            lists = await GetList();
            MenuSelection();

        }

        #region Menu
        private static void MenuSelection()
        {
            string menuSelection;
            do
            {
                menuSelection = drawMenu();

                switch (menuSelection)
                {
                    case "1":
                        /* 
                         * Display all Books list items
                         */
                        GetBooks();
                        break;
                    case "2":
                        /*
                         * Add Book ListItem
                        */
                        AddBook();
                        break;
                    case "3":
                        /*
                         * Update Book ListItem
                        */
                        UpdateBook();
                        break;
                    case "4":
                        //delete Book ListItem
                        DeleteBook();
                        break;
                    case "5":
                        //Exit Menu
                        break;
                    default:
                        Console.WriteLine("Invalid Selection");
                        break;
                }

                Console.ReadLine();
            } while (menuSelection != "5");
        }

        private static string drawMenu()
        {
            Console.Clear();
            Console.WriteLine("*****************************");
            Console.WriteLine("\tBooks Menu");
            Console.WriteLine("*****************************");

            Console.WriteLine("1.\tShow Books");
            Console.WriteLine("2.\tAdd New Book");
            Console.WriteLine("3.\tUpdate Book");
            Console.WriteLine("4.\tDelete Book");
            Console.WriteLine("5.\tExit");
            Console.WriteLine("*****************************\n");

            Console.WriteLine("Please select the menu option");
            string selection = Console.ReadLine();

            return selection;
        }
        #endregion

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

        private async static void GetBooks()
        {

            /* We will show existing site list 
             * and filter the Books List for use 
             * in this sample  
             */
            var list = lists.Where(b => b.DisplayName.Contains("Books")).FirstOrDefault();

            //Show existing Site List in the Current Site
            Console.WriteLine($"Display all {list.DisplayName}");
            Console.WriteLine("***************************");

            //assign the global listId for use in other methods 
            _listId = list.Id;

            //Getting listItems using msgraph
            IListItemsCollectionPage listItems = await GetListItems(list.Id);

            foreach (var item in listItems)
            {
                IDictionary<string, object> booksList = item.Fields.AdditionalData;

                var jsonString = JsonConvert.SerializeObject(booksList);
                var officeBook = JsonConvert.DeserializeObject<OfficeBook>(jsonString);
                officeBook.SharePointItemId = item.Id;

                officeBooks.Add(officeBook);
            }

            foreach (var book in officeBooks)
            {
                Console.WriteLine(book.Title + " : " + book.BookId);
            }
        }

        private async static void AddBook()
        {
            var list = lists.Where(b => b.DisplayName.Contains("Books")).FirstOrDefault();

            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Add New Book");

            Console.WriteLine("Enter Title");
            string title = Console.ReadLine();
            data.Add("Title", title);

            bool result = await Sites.CreateListItem(_groupId, _siteId, _listId, data);

            if (result)
                Console.WriteLine("Item Created");
            else
                Console.WriteLine("Item Not Created");
        }

        private async static void UpdateBook()
        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Update a Book Title");

            Console.WriteLine("Enter Id");
            string listItemId = Console.ReadLine();

            Console.WriteLine("Enter Title");
            string title = Console.ReadLine();
            data.Add("Title", title);

            var officeBookItems = await GetListItems(_listId); 

            bool result = await Sites.UpdateListItem(_groupId, _siteId, _listId, listItemId, data);

            if (result)
                Console.WriteLine("Item Updated");
            else
                Console.WriteLine("Item Not Update");
        }

        private async static void DeleteBook()
        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Delete a book record");

            Console.WriteLine("Enter Id");
            string listItemId = Console.ReadLine();

            bool result = await Sites.DeleteListItem(_groupId, _siteId, _listId, listItemId);

            if (result)
                Console.WriteLine("Item Deleted");
            else
                Console.WriteLine("Item Not Deleted");
        }
    }
}