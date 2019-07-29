using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace msgraph_sharepoint_sample
{
    public class Sites
    {
        private static GraphServiceClient _graphClient = null;

        public async static Task<ISiteListsCollectionPage> GetSiteLists(string siteId)
        {
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
            ISiteListsCollectionPage lists = await _graphClient
                            .Sites[siteId]
                            .Lists.Request().GetAsync();          
           
            return lists;
        }

        public async static Task<List> GetSiteList(string siteId, string listId)
        {
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
            List list = await _graphClient
                            .Sites[siteId]
                            .Lists[listId].Request().GetAsync();
            return list;
        }

        public async static Task<IListItemsCollectionPage> GetSiteListItems(string siteId, string listId)
        {
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
            IListItemsCollectionPage listItems = await _graphClient
                            .Sites[siteId]
                            .Lists[listId]
                            .Items
                            .Request().Expand("fields")
                            .GetAsync();
            return listItems;
        }

        public async static Task<bool> CreateListItem(string siteId, string listId, IDictionary<string, object> data)
        {
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();
            var listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = data,
                }
            };

            try
            {
                await _graphClient
                               .Sites[siteId]
                               .Lists[listId]
                               .Items
                               .Request()
                               .AddAsync(listItem);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        public async static Task<bool> UpdateListItem(string siteId, string listId, string itemId, IDictionary<string, object> data)
        {
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();

            var fieldValueSet = new FieldValueSet
            {
                AdditionalData = data,
            };

            try
            {
                await _graphClient
                                .Sites[siteId]
                                .Lists[listId]
                                .Items[itemId]
                                .Fields
                                .Request()
                                .UpdateAsync(fieldValueSet);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }


        public async static Task<bool> DeleteListItem(string siteId, string listId, string itemId)
        {
            _graphClient = GraphServiceClientProvider.GetAuthenticatedClient();

            try
            {
                await _graphClient
                                .Sites[siteId]
                                .Lists[listId]
                                .Items[itemId]
                                .Request()
                                .DeleteAsync();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
    }
}
