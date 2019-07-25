using Microsoft.Graph;
using msgraph_sharepoint_sample.Models;
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
        private static List<Member> _members = new List<Member>();
        private static string sharePointItemId = null;
        private static GraphServiceClient _graphClient = null;
        private static User user = new User();
        static async Task Main(string[] args)
        {
            // Get an authenticated client
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();

            // Get the logged-in user
            user = _graphClient.Me.Request().GetAsync().Result;            

            // Get all the SharePoint Lists
            lists = await GetList();

            // Get the list of members from SharePoint
            await LoadMembers();

            // Add the logged-in user to the SharePoint Lists (if they don't exist)
            await AddLoggedInUser();

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
                        UpdateBook(sharePointItemId);
                        break;
                    case "4":
                        //delete Book ListItem
                        DeleteBook(sharePointItemId);
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
            Console.WriteLine("User: " + user.GivenName + " " + user.Surname);
            Console.WriteLine("Email: " + user.Mail + "\n\n");
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

        private async static Task AddLoggedInUser()
        {    
            if(user == null)
            {
                user = _graphClient.Me.Request().GetAsync().Result;                
            }
            // Test //
          //  var directory = _graphClient.Directory.Request().GetAsync().Result;

            

            // Check if the logged-in member is in list.
            var member = _members.FirstOrDefault(m => m.MemberId.Equals(user.Id));

            Member newMember = new Member();
            if (member == null) // Add new member
            {                
                newMember.MemberId = Guid.Parse(user.Id);
                newMember.FirstName = user.GivenName;
                newMember.LastName = user.Surname;
                newMember.Email = user.Mail;
                newMember.IsAdmin = "False";
            }
            await AddMember(newMember);
        }

        private async static Task AddMember(Member newMember)
        {
            IDictionary<string, object> memberDictionary = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Adding new member...");
            
            var jsonString = JsonConvert.SerializeObject(newMember);
            memberDictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.CreateListItem(_groupId, _siteId, _listId, memberDictionary);
            if (result)
            {
                Console.WriteLine("Member added");
                Console.WriteLine("***************************\n\n");

                await LoadMembers();                
            }
            else
            {
                Console.WriteLine("Failed to add new member");
                Console.WriteLine("***************************\n\n");
            }                
        }
        private async static Task LoadMembers()
        {
            // Clear list
            _members.Clear();

            var list = lists.Where(b => b.DisplayName.Contains("Members")).FirstOrDefault();

            //assign the global listId for use in other methods 
            _listId = list.Id;

            //Getting listItems using msgraph
            IListItemsCollectionPage listItems = await GetListItems(list.Id);

            foreach (var item in listItems)
            {
                IDictionary<string, object> memberList = item.Fields.AdditionalData;

                var jsonString = JsonConvert.SerializeObject(memberList);
                var member = JsonConvert.DeserializeObject<Member>(jsonString);
                member.SharePointItemId = item.Id;

                _members.Add(member);
            }

        }

        private async static Task LoadBooks()
        {
            // Clear list 
            officeBooks.Clear();

            var list = lists.Where(b => b.DisplayName.Contains("Books")).FirstOrDefault();

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
        }

        private async static void GetBooks()
        {

            /* We will show existing site list 
             * and filter the Books List for use 
             * in this sample  
             */

            // Load books first
            await LoadBooks();

            //Show existing Site List in the Current Site
            Console.WriteLine($"Display all Office Books");
            Console.WriteLine("***************************");                       

            foreach (var book in officeBooks)
            {
                Console.WriteLine("(" + book.SharePointItemId +") "+ book.Title + " : " + book.BookId);
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
            var officeBookItem = new OfficeBook();
            officeBookItem.Title = title;
            officeBookItem.BookId = Guid.NewGuid();

            var jsonString = JsonConvert.SerializeObject(officeBookItem);
            data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.CreateListItem(_groupId, _siteId, list.Id, data);
            if (result)
            {
                Console.WriteLine("Item Created");
                await LoadBooks();
            }                
            else
                Console.WriteLine("Item Not Created");
        }

        private async static void UpdateBook(string sharePointItemId)
        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Update Book");
            Console.WriteLine("***************************");
            Console.WriteLine("Enter ID");

            sharePointItemId = Console.ReadLine();
            string listItemId = sharePointItemId;

            var officeBookItem = officeBooks.Where(b => b.SharePointItemId.Contains(sharePointItemId)).FirstOrDefault();


            Console.WriteLine("Enter Title");
            string title = Console.ReadLine();

            officeBookItem.Title = title;

            var jsonString = JsonConvert.SerializeObject(officeBookItem);
            data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.UpdateListItem(_groupId, _siteId, _listId, listItemId, data);

            if (result)
            {
                Console.WriteLine("Item Updated");
                await LoadBooks();
            }                
            else
                Console.WriteLine("Item Not Update");
        }

        private async static void DeleteBook(string id)
        {
            Console.WriteLine("Enter ID");

            sharePointItemId = Console.ReadLine();

            string listItemId = sharePointItemId;

            var officeBookItem = officeBooks.Where(b => b.SharePointItemId.Contains(sharePointItemId)).FirstOrDefault();
            
            bool result = await Sites.DeleteListItem(_groupId, _siteId, _listId, listItemId);

            if (result)
            {
                Console.WriteLine("Item Deleted");
                await LoadBooks();
            }               
            else
                Console.WriteLine("Item Not Deleted");
        }
    }
}