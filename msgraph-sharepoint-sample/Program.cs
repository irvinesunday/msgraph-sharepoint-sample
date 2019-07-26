using Microsoft.Graph;
using msgraph_sharepoint_sample.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace msgraph_sharepoint_sample
{
    public class Program
    {
        private readonly static string _groupId = "30ecd9bc-6e8b-4280-adbc-18dccc0815a6";
        private readonly static string _siteId = "m365b267815.sharepoint.com,6e1261a1-6d03-432a-95c0-e1c7705aef5f,f43d258c-ece0-476a-a1c0-018d359817d5";
        private static string _listId = null;

        private static ISiteListsCollectionPage lists;
        private static List<OfficeBook> _officeBooks = new List<OfficeBook>();
        private static List<Member> _members = new List<Member>();
        private static List<OfficeItem> _officeItems = new List<OfficeItem>();
        private static string userItemId = null;
        private static List list;

        private static OfficeBook _officeBook = new OfficeBook();
        private static OfficeItem _officeItem = new OfficeItem();
        private static Member _member = new Member();
        
        private static GraphServiceClient _graphClient = null;
        private static User user = new User();

        private static StringBuilder _consoleMessage = new StringBuilder();
        

        static async Task Main(string[] args)
        {
            // Get an authenticated client
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();

            // Get the logged-in user
            user = _graphClient.Me.Request().GetAsync().Result;            

            // Get all the SharePoint Lists
            lists = await GetList();

            // Get the list of members from SharePoint
            await RetrieveListItemsFromSharePoint(_member);

            // Add the logged-in user to the SharePoint Lists (if they don't exist)
            await AddLoggedInUser();

            await RetrieveListItemsFromSharePoint(_officeItem);
            await RetrieveListItemsFromSharePoint(_officeBook);
            MenuSelection();
        }

        #region Menu
        private static void MenuSelection()
        {
            string menuSelection;
            string selectedList;
            do
            {
                menuSelection = drawMenu();

                switch (menuSelection)
                {
                    case "1":
                        /* 
                         * Display all Resources list items
                         */
                        Console.WriteLine("Select List");

                        Console.WriteLine("\t(a) Office Book");
                        Console.WriteLine("\t(b) Office Item");
                        selectedList = Console.ReadLine();

                        if (selectedList.ToLower() == "a" || selectedList.ToLower() == "(a)")
                        {
                            DisplayListItems(_officeBook);
                        }
                        else if (selectedList.ToLower() == "b" || selectedList.ToLower() == "(b)")
                        {
                            DisplayListItems(_officeItem);
                        }

                        break;
                    case "2":
                        /*
                         * Add Resource ListItem
                        */
                        Console.WriteLine("Select List");

                        Console.WriteLine("\t(a) Office Book");
                        Console.WriteLine("\t(b) Office Item");
                        selectedList = Console.ReadLine();

                        if (selectedList == "a" || selectedList.ToLower() == "(a)")
                        {
                            AddListItem(_officeBook);
                        }
                        else if (selectedList == "b" || selectedList.ToLower() == "(b)")
                        {
                            AddListItem(_officeItem);
                        }
                        break;
                    case "3":
                        /*
                         * Update Resource ListItem
                        */
                        Console.WriteLine("Select List");

                        Console.WriteLine("\t(a) Office Book");
                        Console.WriteLine("\t(b) Office Item");
                        selectedList = Console.ReadLine();

                        if (selectedList == "a" || selectedList.ToLower() == "(a)")
                        {
                            UpdateOfficeBook(_officeBook);
                        }
                        else if (selectedList == "b" || selectedList.ToLower() == "(b)")
                        {
                            UpdateOfficeItem(_officeItem);
                        }                        

                        break;
                    case "4":
                        /*
                         * Delete a Resource ListItem
                        */
                        Console.WriteLine("Select List");

                        Console.WriteLine("\t(a) Office Book");
                        Console.WriteLine("\t(b) Office Item");
                        selectedList = Console.ReadLine();

                        if (selectedList == "a" || selectedList.ToLower() == "(a)")
                        {
                            DeleteOfficeBook(_officeBook);
                        }
                        else if (selectedList == "b" || selectedList.ToLower() == "(b)")
                        {
                            DeleteOfficeItem(_officeItem);
                        }

                        break;
                    case "5":
                        // Show list of members

                        DisplayListItems(_member);
                        break;                    
                    default:
                        Console.WriteLine("Invalid Selection");
                        break;
                }
                Console.ReadLine();
            } while (menuSelection != "6");
        }

        private static string drawMenu()
        {
            Console.Clear();
            Console.WriteLine("User: " + user.GivenName + " " + user.Surname);
            Console.WriteLine("Email: " + user.Mail + "\n");
            Console.WriteLine(_consoleMessage);
            Console.WriteLine("*****************************");
            Console.WriteLine("\tResources Menu");
            Console.WriteLine("*****************************");

            Console.WriteLine("1.\tShow Resources");
            Console.WriteLine("2.\tAdd a New Resource");
            Console.WriteLine("3.\tUpdate a Resource");
            Console.WriteLine("4.\tDelete a Resource");
            Console.WriteLine("5.\tShow Members");
            Console.WriteLine("6.\tExit");
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

            // Check if the logged-in member is in list.
            var member = _members.FirstOrDefault(m => m.MemberId.Equals(Guid.Parse(user.Id)));

            Member newMember = new Member();
            if (member == null) // Member not in members list - add new member
            {                
                newMember.MemberId = Guid.Parse(user.Id);
                newMember.FirstName = user.GivenName;
                newMember.LastName = user.Surname;
                newMember.Email = user.Mail;
                newMember.IsAdmin = "False";

                await AddMember(newMember);
            }            
        }

        private async static Task AddMember(Member newMember)
        {
            IDictionary<string, object> memberDictionary = new Dictionary<string, object>();

            StringBuilder newMemberMessage = new StringBuilder();
            
            Console.WriteLine("***************************");
            Console.WriteLine("Adding new member...");
            
            var jsonString = JsonConvert.SerializeObject(newMember);
            memberDictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.CreateListItem(_groupId, _siteId, _listId, memberDictionary);
            if (result)
            {
                newMemberMessage.Append("\nMember added");
                newMemberMessage.Append("***************************\n\n");

                await RetrieveListItemsFromSharePoint(_member);             
            }
            else
            {
                newMemberMessage.Append("\nFailed to add new member");
                newMemberMessage.Append("***************************\n\n");
            }

            _consoleMessage.Append(newMemberMessage);
        }
        
        /// <summary>
        /// Retrieves list items from SharePoint
        /// </summary>
        /// <param name="obj">The object instance of the corresponding list to be retrieved</param>
        /// <returns></returns>
        private async static Task RetrieveListItemsFromSharePoint(object obj)
        {
            // Clear list 
            if (obj.GetType() == typeof(OfficeBook))
            {
                _officeBooks.Clear();

                list = lists.Where(b => b.DisplayName.Contains("Books")).FirstOrDefault();
            }
            else if (obj.GetType() == typeof(OfficeItem))
            {
                _officeItems.Clear();

                list = lists.Where(b => b.DisplayName.Contains("Items")).FirstOrDefault();               
            }
            else if(obj.GetType() == typeof(Member))
            {
                _members.Clear();

                list = lists.Where(b => b.DisplayName.Contains("Members")).FirstOrDefault();
            }

            //assign the global listId for use in other methods 
            _listId = list.Id;

            //Getting listItems using msgraph
            IListItemsCollectionPage listItems = await GetListItems(_listId);

            foreach (var item in listItems)
            {
                IDictionary<string, object> resourceList = item.Fields.AdditionalData;
                var jsonString = JsonConvert.SerializeObject(resourceList);

                if (obj.GetType() == typeof(OfficeBook))
                {
                    var officeResource = JsonConvert.DeserializeObject<OfficeBook>(jsonString);
                    officeResource.SharePointItemId = item.Id;
                    _officeBooks.Add(officeResource);
                }
                else if (obj.GetType() == typeof(OfficeItem))
                {
                    var officeResource = JsonConvert.DeserializeObject<OfficeItem>(jsonString);
                    officeResource.SharePointItemId = item.Id;
                    _officeItems.Add(officeResource);
                }
                else if(obj.GetType() == typeof(Member))
                {
                    var memberResource = JsonConvert.DeserializeObject<Member>(jsonString);
                    memberResource.SharePointItemId = item.Id;
                    _members.Add(memberResource);
                }
            }
        }

        private async static void DisplayListItems(object obj)
        {            
            // Retrieve the respective list items
            await RetrieveListItemsFromSharePoint(obj);                      

            if (obj.GetType() == typeof(OfficeBook))
            {
                Console.WriteLine($"Display all Office Books");
                Console.WriteLine("***************************");
                foreach (var book in _officeBooks)
                {                    
                    Console.WriteLine($"({book.SharePointItemId}) : {book.Title} : {book.BookId}");
                }
            }
            else if (obj.GetType() == typeof(OfficeItem))
            {
                Console.WriteLine($"Display all Office Items");
                Console.WriteLine("***************************");
                foreach (var officeItem in _officeItems)
                {                    
                    Console.WriteLine($"({officeItem.SharePointItemId}) : {officeItem.Title} : {officeItem.ItemId}");
                }
            }
            else if (obj.GetType() == typeof(Member))
            {
                Console.WriteLine($"Display all Members");
                Console.WriteLine("***************************");
                foreach (var member in _members)
                {                    
                    Console.WriteLine($"({member.SharePointItemId}) : {member.FirstName} {member.LastName} : {member.Email}");
                }
            }
        }

        private async static void AddListItem(object obj)
        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            string jsonString;
            if (obj.GetType() == typeof(OfficeBook))
            {
                list = lists.Where(b => b.DisplayName.Contains("Books")).FirstOrDefault();
                _listId = list.Id;

                Console.WriteLine("***************************");
                Console.WriteLine($"Add New {list.DisplayName}");

                Console.WriteLine("Enter Title");
                string title = Console.ReadLine();

                var officeBookItem = new OfficeBook();
                officeBookItem.Title = title;
                officeBookItem.BookId = Guid.NewGuid();

                jsonString = JsonConvert.SerializeObject(officeBookItem);
                data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            }
            else if (obj.GetType() == typeof(OfficeItem))
            {
                list = lists.Where(b => b.DisplayName.Contains("Items")).FirstOrDefault();
                _listId = list.Id;

                Console.WriteLine($"Add New {list.DisplayName}");

                Console.WriteLine("Enter Title");
                string title = Console.ReadLine();

                var officeItem = new OfficeItem();
                officeItem.Title = title;
                officeItem.ItemId = Guid.NewGuid();

                jsonString = JsonConvert.SerializeObject(officeItem);
                data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);
            }
            
            bool result = await Sites.CreateListItem(_groupId, _siteId, _listId, data);
            if (result)
            {
                Console.WriteLine("Item Created");
                await RetrieveListItemsFromSharePoint(obj);
            }
            else
            {
                Console.WriteLine("Item Not Created");
            }                
        }

        private async static void UpdateOfficeBook(OfficeBook officeBook)
        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Update Office Book Details");
            Console.WriteLine("***************************");
            Console.WriteLine("Enter ID");

            userItemId = Console.ReadLine();
            
            var officeBookItem = _officeBooks.Where(b => b.SharePointItemId.Equals(userItemId)).FirstOrDefault();

            if(officeBookItem == null)
            {
                Console.WriteLine($"Office Book with ID: {userItemId} doesn't exist.");
                return;
            }

            string listItemId = userItemId;

            Console.WriteLine("Enter Title");
            string title = Console.ReadLine();
            
            officeBookItem.Title = title;


            var jsonString = JsonConvert.SerializeObject(officeBookItem);
            data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.UpdateListItem(_groupId, _siteId, _listId, listItemId, data);
            if (result)
            {
                Console.WriteLine("Book Successfully Updated");
                await RetrieveListItemsFromSharePoint(officeBook);
            }
            else
                Console.WriteLine("Book Not Updated");
        }

        private async static void UpdateOfficeItem(OfficeItem officeItem)
        {
          //  await LoadResources(obj);
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Update Office Item Details");
            Console.WriteLine("***************************");

            Console.WriteLine("Enter ID");
            string userListItemId = Console.ReadLine();

            var item = _officeItems.Where(b => b.SharePointItemId.Equals(userListItemId)).FirstOrDefault();
            if (item == null)
            {
                Console.WriteLine($"Item with ID: {userListItemId} doesn't exist.");
                return;
            }

            string listItemId = userListItemId;

            Console.WriteLine("Enter Title");
            string title = Console.ReadLine();                  

            item.Title = title;

            var jsonString = JsonConvert.SerializeObject(item);
            data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.UpdateListItem(_groupId, _siteId, _listId, listItemId, data);
            if (result)
            {
                Console.WriteLine("Item Successfully Updated");
                await RetrieveListItemsFromSharePoint(officeItem);
            }
            else
                Console.WriteLine("Item Not Update");
        }

        //deletes office book in sharepoint Office Book List
        private async static void DeleteOfficeBook(OfficeBook officeBook)
        {
            Console.WriteLine("***************************");
            Console.WriteLine("Delete an Office Book");
            Console.WriteLine("***************************");
            Console.WriteLine("Enter ID");

            userItemId = Console.ReadLine();                        

            var officeBookItem = _officeBooks.Where(b => b.SharePointItemId.Contains(userItemId)).FirstOrDefault();

            if (officeBookItem == null)
            {
                Console.WriteLine($"Office Book with ID: {userItemId} doesn't exist.");
                return;
            }

            string listItemId = userItemId;

            bool result = await Sites.DeleteListItem(_groupId, _siteId, _listId, listItemId);

            if (result)
            {
                Console.WriteLine("Book Successfully Deleted");
                await RetrieveListItemsFromSharePoint(officeBook);
            }
            else
                Console.WriteLine("Book Not Deleted");
        }

        //deletes office item in sharepoint Office Item list
        private async static void DeleteOfficeItem(OfficeItem officeItem)
        {
            Console.WriteLine("***************************");
            Console.WriteLine("Delete an Office Item");
            Console.WriteLine("***************************");
            Console.WriteLine("Enter ID");

            userItemId = Console.ReadLine();                      

            var item = _officeItems.Where(b => b.SharePointItemId.Contains(userItemId)).FirstOrDefault();
            if (item == null)
            {
                Console.WriteLine($"Item with ID: {userItemId} doesn't exist.");
                return;
            }

            string listItemId = userItemId;

            bool result = await Sites.DeleteListItem(_groupId, _siteId, _listId, listItemId);

            if (result)
            {
                Console.WriteLine("Office Item Successfully Deleted");
                await RetrieveListItemsFromSharePoint(officeItem);
            }
            else
                Console.WriteLine("Office Item Not Deleted");
        }

        
    }
}