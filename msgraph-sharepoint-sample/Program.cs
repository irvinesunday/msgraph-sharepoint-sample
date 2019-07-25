﻿using Microsoft.Graph;
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
        private static List<OfficeItem> officeItems = new List<OfficeItem>();
        private static object officeBook = new OfficeBook();
        private static object officeItem = new OfficeItem();
        private static string sharePointItemId = null;
        private static GraphServiceClient _graphClient = null;
        private static User user = new User();
        private static List list;

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

            await LoadResources(officeItem);
            await LoadResources(officeBook);
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
                         * Display all Books list items
                         */
                        Console.WriteLine("Select List");

                        Console.WriteLine("\t\t(a) Office Book");
                        Console.WriteLine("\t\t(b) Office Item");
                        selectedList = Console.ReadLine();

                        if (selectedList == "a")
                        {
                            GetResources(officeBook);

                        }
                        else if (selectedList == "b")
                        {
                            GetResources(officeItem);
                        }

                        break;
                    case "2":
                        /*
                         * Add Book ListItem
                        */
                        Console.WriteLine("Select List");

                        Console.WriteLine("\t\t(a) Office Book");
                        Console.WriteLine("\t\t(b) Office Item");
                        selectedList = Console.ReadLine();

                        if (selectedList == "a")
                        {
                            AddResource(officeBook);

                        }
                        else if (selectedList == "b")
                        {
                            AddResource(officeItem);
                        }
                        break;
                    case "3":
                        /*
                         * Update Book ListItem
                        */

                        UpdateBook(sharePointItemId, officeBook);

                        break;
                    case "4":

                        UpdateItem(officeItem);


                        break;
                    case "5":
                        //delete Book ListItem

                        DeleteBook(sharePointItemId, officeBook);

                        break;
                    case "6":
                        DeleteItem(sharePointItemId, officeItem);

                        break;
                    case "7":
                        //Exit Menu
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
            Console.WriteLine("Email: " + user.Mail + "\n\n");
            Console.WriteLine("*****************************");
            Console.WriteLine("\tBooks Menu");
            Console.WriteLine("*****************************");

            Console.WriteLine("1.\tShow Resources");
            Console.WriteLine("2.\tAdd New Resource");
            Console.WriteLine("3.\tUpdate Book");
            Console.WriteLine("4.\tUpdate Item");
            Console.WriteLine("5.\tDelete Book");
            Console.WriteLine("6.\tDelete Item");
            Console.WriteLine("7.\tExit");
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

        private async static Task LoadResources(object obj)
        {
            // Clear list 
            if (obj.GetType() == typeof(OfficeBook))
            {
                officeBooks.Clear();

                list = lists.Where(b => b.DisplayName.Contains("Books")).FirstOrDefault();

                //assign the global listId for use in other methods 
                _listId = list.Id;
            }
            else if (obj.GetType() == typeof(OfficeItem))
            {
                officeItems.Clear();

                list = lists.Where(b => b.DisplayName.Contains("Items")).FirstOrDefault();

                //assign the global listId for use in other methods 
                _listId = list.Id;
            }

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
                    officeBooks.Add(officeResource);
                }
                else if (obj.GetType() == typeof(OfficeItem))
                {
                    var officeResource = JsonConvert.DeserializeObject<OfficeItem>(jsonString);
                    officeResource.SharePointItemId = item.Id;
                    officeItems.Add(officeResource);
                }
            }
        }

        private async static void GetResources(object obj)
        {

            /* We will show existing site list 
             * and filter the Books List for use 
             * in this sample  
             */

            // Load books first
            await LoadResources(obj);

            //Show existing Site List in the Current Site
            Console.WriteLine($"Display all Office Books");
            Console.WriteLine("***************************");

            if (obj.GetType() == typeof(OfficeBook))
            {
                foreach (var book in officeBooks)
                {
                    Console.WriteLine("(" + book.SharePointItemId + ") " + book.Title + " : " + book.BookId);
                }
            }
            else if (obj.GetType() == typeof(OfficeItem))
            {
                foreach (var officeItem in officeItems)
                {
                    Console.WriteLine("(" + officeItem.SharePointItemId + ") " + officeItem.Title + officeItem.ItemId);

                }
            }
        }

        private async static void AddResource(object obj)
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
                await LoadResources(obj);
            }
            else
                Console.WriteLine("Item Not Created");
        }

        private async static void UpdateBook(string sharePointItemId, object obj)

        {
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Update Book");
            Console.WriteLine("***************************");
            Console.WriteLine("Enter ID");

            sharePointItemId = Console.ReadLine();

            string listItemId = sharePointItemId;

            var officeBookItem = officeBooks.Where(b => b.SharePointItemId.Equals(sharePointItemId)).FirstOrDefault();

            Console.WriteLine("Enter Title");
            string title = Console.ReadLine();

            officeBookItem.Title = title;

            var jsonString = JsonConvert.SerializeObject(officeBookItem);
            data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.UpdateListItem(_groupId, _siteId, _listId, listItemId, data);
            if (result)
            {
                Console.WriteLine("Item Updated");
                await LoadResources(obj);
            }
            else
                Console.WriteLine("Item Not Update");
        }

        private async static void UpdateItem(object obj)
        {
          //  await LoadResources(obj);
            IDictionary<string, object> data = new Dictionary<string, object>();

            Console.WriteLine("***************************");
            Console.WriteLine("Update Item");
            Console.WriteLine("***************************");

            Console.WriteLine("Enter ID");
            string listItemId = Console.ReadLine();

            var officeItem = officeItems.Where(b => b.SharePointItemId.Equals(listItemId)).FirstOrDefault();

            Console.WriteLine("Enter Title");
            string title = Console.ReadLine();

            officeItem.Title = title;

            var jsonString = JsonConvert.SerializeObject(officeItem);
            data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);

            bool result = await Sites.UpdateListItem(_groupId, _siteId, _listId, listItemId, data);
            if (result)
            {
                Console.WriteLine("Item Updated");
                await LoadResources(obj);
            }
            else
                Console.WriteLine("Item Not Update");
        }

        //deletes office book in sharepoint Office Book List
        private async static void DeleteBook(string sharePointItemId, object obj)
        {
            Console.WriteLine("Enter ID");

            sharePointItemId = Console.ReadLine();

            string listItemId = sharePointItemId;

            var officeBookItem = officeBooks.Where(b => b.SharePointItemId.Contains(sharePointItemId)).FirstOrDefault();

            bool result = await Sites.DeleteListItem(_groupId, _siteId, _listId, listItemId);

            if (result)
            {
                Console.WriteLine("Item Deleted");
                await LoadResources(obj);
            }
            else
                Console.WriteLine("Item Not Deleted");
        }

        //deletes office item in sharepoint Office Item list
        private async static void DeleteItem(string sharePointItemId, object obj)
        {
            Console.WriteLine("Enter ID");

            sharePointItemId = Console.ReadLine();

            string listItemId = sharePointItemId;

            var officeItem = officeItems.Where(b => b.SharePointItemId.Contains(sharePointItemId)).FirstOrDefault();

            bool result = await Sites.DeleteListItem(_groupId, _siteId, _listId, listItemId);

            if (result)
            {
                Console.WriteLine("Item Deleted");
                await LoadResources(obj);
            }
            else
                Console.WriteLine("Item Not Deleted");
        }


        private async static void DeleteResource(string sharePointItemId, object obj)
        {
            Console.WriteLine("Enter ID");

            sharePointItemId = Console.ReadLine();

            string listItemId = sharePointItemId;

            bool result = await Sites.DeleteListItem(_groupId, _siteId, _listId, listItemId);

            if (result)
            {
                Console.WriteLine("Item Deleted");
                await LoadResources(obj);
            }
            else
                Console.WriteLine("Item Not Deleted");
        }
    }
}