using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using SpResearchTracker.Utils;
using System.Collections.Generic;
using System.Linq;

namespace SpResearchTracker.Models
{
    public abstract class Repository
    {
        public static readonly string SiteUrl = ConfigurationManager.AppSettings["ida:SiteUrl"];
        public static readonly string ProjectsListName = ConfigurationManager.AppSettings["ProjectsListName"];
        public static readonly string ReferencesListName = ConfigurationManager.AppSettings["ReferencesListName"];
    }

    public abstract class RESTRepository : Repository
    {
        /// <summary>
        /// Implements common GET functionality
        /// </summary>
        /// <param name="requestUri">The REST endpoint</param>
        /// <param name="accessToken">The SharePoint access token</param>
        /// <returns>XElement with results of operation</returns>
        public async Task<HttpResponseMessage> Get(string requestUri, string accessToken, string eTag)
        {
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUri);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            if (eTag.Length > 0 && eTag != "*")
            {
                request.Headers.IfNoneMatch.Add(new EntityTagHeaderValue(eTag));
            }
            return await client.SendAsync(request);
        }

        public Task<HttpResponseMessage> Get(string requestUri, string accessToken)
        {
            return Get(requestUri, accessToken, string.Empty);
        }

        /// <summary>
        /// Implements common POST functionality
        /// </summary>
        /// <param name="requestUri">The REST endpoint</param>
        /// <param name="accessToken">The SharePoint access token</param>
        /// <param name="requestData">The POST data</param>
        /// <returns>XElement with results of operation</returns>
        public async Task<HttpResponseMessage> Post(string requestUri, string accessToken, StringContent requestData)
        {
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUri);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            requestData.Headers.ContentType = MediaTypeHeaderValue.Parse("application/atom+xml");
            request.Content = requestData;
            return await client.SendAsync(request);
        }

        /// <summary>
        /// Implements common PATCH functionality
        /// </summary>
        /// <param name="requestUri">The REST endpoint</param>
        /// <param name="accessToken">The SharePoint access token</param>
        /// <param name="eTag">The eTag of the item</param>
        /// <param name="requestData">The data to use during the update</param>
        /// <returns>XElement with results of operation</returns>
        public async Task<HttpResponseMessage> Patch(string requestUri, string accessToken, string eTag, StringContent requestData)
        {
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUri);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            requestData.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/atom+xml");
            if (eTag == "*")
            {
                request.Headers.Add("IF-MATCH", "*");
            }
            else
            {
                request.Headers.IfMatch.Add(new EntityTagHeaderValue(eTag));
            }
            request.Headers.Add("X-Http-Method", "PATCH");
            request.Content = requestData;
            return await client.SendAsync(request);
        }

        /// <summary>
        /// Implements common DELETE functionality
        /// </summary>
        /// <param name="requestUri">The REST endpoint</param>
        /// <param name="accessToken">The SharePoint access token</param>
        /// <param name="eTag">The eTag of the item</param>
        /// <returns>XElement with results of operation</returns>
        public async Task<HttpResponseMessage> Delete(string requestUri, string accessToken, string eTag)
        {
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, requestUri);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            if (eTag == "*")
            {
                request.Headers.Add("IF-MATCH", "*");
            }
            else
            {
                request.Headers.IfMatch.Add(new EntityTagHeaderValue(eTag));
            }

            return await client.SendAsync(request);
        }
    }

    public abstract class CSOMRepository : Repository
    {
        public async Task<ListItemCollection> GetListItemCollectionAsync(string accessToken, string listTitle)
        {
            ListItemCollection items = null;

            using (ClientContext ctx = GetClientContext(accessToken))
            {
                // Get the list
                List list = ctx.Web.Lists.GetByTitle(listTitle);

                // Get the items
                items = list.GetItems(CamlQuery.CreateAllItemsQuery());

                // Load 
                ctx.Load(items);

                // Execute query
                await ctx.ExecuteQueryAsync();
            }

            return items;
        }

        public async Task<ListItem> GetListItemAsync(string accessToken, string listTitle, int Id)
        {
            ListItem item = null;

            using (ClientContext ctx = GetClientContext(accessToken))
            {
                // Get the list
                List list = ctx.Web.Lists.GetByTitle(listTitle);

                // Get the item
                item = list.GetItemById(Id);

                // Load 
                ctx.Load(item);

                // Execute query
                await ctx.ExecuteQueryAsync();
            }

            return item;
        }

        public async Task<ListItem> CreateListItemAsync(string accessToken, string listTitle, IDictionary<string, object> fieldMappings)
        {
            ListItem item = null;

            using (ClientContext ctx = GetClientContext(accessToken))
            {
                // Get the list
                List list = ctx.Web.Lists.GetByTitle(listTitle);

                // Create listitem creation info
                ListItemCreationInformation listItemCreationInfo = new ListItemCreationInformation();
                item = list.AddItem(listItemCreationInfo);

                // Add the fields
                foreach (var fieldMapping in fieldMappings)
                {
                    item[fieldMapping.Key] = fieldMapping.Value;
                }
                
                // Update the item
                item.Update();

                // Load the item
                ctx.Load(item);

                // Execute query
                await ctx.ExecuteQueryAsync();                
            }

            return item;
        }

        public async Task<bool> UpdateListItemAsync(string accessToken, string listTitle, int Id, IDictionary<string, object> fieldMappings)
        {
            try
            {
                using (ClientContext ctx = GetClientContext(accessToken))
                {
                    // Get the list
                    List list = ctx.Web.Lists.GetByTitle(listTitle);

                    // Get the item
                    ListItem item = list.GetItemById(Id);

                    // Add the fields
                    foreach (var fieldMapping in fieldMappings)
                    {
                        item[fieldMapping.Key] = fieldMapping.Value;
                    }

                    // Update the item
                    item.Update();

                    // Execute query
                    await ctx.ExecuteQueryAsync();
                }

                return true;
            }
            catch (ServerException)
            {
                // do not let it go up
                return false;
            }
        }

        public async Task<bool> DeleteListItemAsync(string accessToken, string listTitle, int Id)
        {
            try
            {
                using (ClientContext ctx = GetClientContext(accessToken))
                {
                    // Get the list
                    List list = ctx.Web.Lists.GetByTitle(listTitle);

                    // Get the item
                    ListItem item = list.GetItemById(Id);

                    // Delete
                    item.DeleteObject();

                    // Execute query
                    await ctx.ExecuteQueryAsync();
                }

                return true;
            }
            catch (ServerException)
            {
                // do not let it go up
                return false;
            }
        }

        public async Task<bool> ListExistsAsync(string accessToken, string listTitle)
        {
            bool listExists = false;

            using (ClientContext ctx = GetClientContext(accessToken))
            {
                // Get the lists
                ListCollection listCollection = ctx.Web.Lists;

                // Load, only include those where the title is the same
                ctx.Load(listCollection, lists => lists.Include(l => l.Title).Where(l => l.Title == listTitle));
                await ctx.ExecuteQueryAsync();

                // If found => exists
                listExists = listCollection.Count > 0;
            }

            return listExists;
        }

        public async Task<bool> CreateListAsync(string accessToken, string listTitle, int templateType)
        {
            try
            {
                using (ClientContext ctx = GetClientContext(accessToken))
                {
                    // Create ListCreationInformation
                    ListCreationInformation listCreationInformation =
                        new ListCreationInformation()
                        {
                            Title = listTitle
                            ,
                            TemplateType = templateType
                        };

                    // Add
                    List list = ctx.Web.Lists.Add(listCreationInformation);
                    list.Update();

                    // Execue query
                    await ctx.ExecuteQueryAsync();
                }

                return true;
            }
            catch (ServerException)
            {
                // do not let it go up
                return false;
            }
        }

        public async Task<bool> AddFieldToListAsync(string accessToken, string listTitle, string fieldName, string fieldType)
        {
            try
            {
                using (ClientContext ctx = GetClientContext(accessToken))
                {
                    // Get the list
                    List list = ctx.Web.Lists.GetByTitle(listTitle);

                    // Add field
                    Field field = 
                        list.Fields.AddFieldAsXml(
                            string.Format("<Field DisplayName='{0}' Type='{1}' />", fieldName, fieldType),
                            true,
                            AddFieldOptions.DefaultValue);

                    // Update
                    field.Update();

                    // Execute
                    await ctx.ExecuteQueryAsync();
                }

                return true;
            }
            catch (ServerException)
            {
                // do not let it go up
                return false;
            }
        }

        protected ClientContext GetClientContext(string accessToken)
        {
            // Create clientcontext
            ClientContext clientContext = new ClientContext(SiteUrl);

            // Set the Authorization header
            clientContext.ExecutingWebRequest +=
                (sender, e) =>
                {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };

            return clientContext;
        }
    }
}