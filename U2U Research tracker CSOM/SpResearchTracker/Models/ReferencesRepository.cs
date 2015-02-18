using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SpResearchTracker.Utils;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;

namespace SpResearchTracker.Models
{
    public interface IReferencesRepository
    {
        Task<IEnumerable<Reference>> GetReferences(string accessToken);
        Task<Reference> GetReference(string accessToken, int Id, string eTag);
        Task<Reference> CreateReference(string accessToken, Reference reference);
        Task<bool> UpdateReference(string accessToken, Reference reference);
        Task<bool> DeleteReference(string accessToken, int Id, string eTag);
    }
    public class ReferencesRepository: RESTRepository, IReferencesRepository
    {
        public async Task<IEnumerable<Reference>> GetReferences(string accessToken)
        {
            List<Reference> references = new List<Reference>();

            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ReferencesListName)
                .Append("')/items?$select=ID,Title,URL,Comments,Project");

            HttpResponseMessage response = await Get(requestUri.ToString(), accessToken);
            string responseString = await response.Content.ReadAsStringAsync();
            XElement root = XElement.Parse(responseString);

            foreach (XElement entryElem in root.Elements().Where(e => e.Name.LocalName == "entry"))
            {
                references.Add(entryElem.ToReference());
            }

            return references.AsQueryable();
        }

        public async Task<Reference> GetReference(string accessToken, int Id, string eTag)
        {
            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ReferencesListName)
                .Append("')/getItemByStringId('")
                .Append(Id.ToString())
                .Append("')?$select=ID,Title,URL,Comments,Project");


            HttpResponseMessage response = await Get(requestUri.ToString(), accessToken, eTag);
            string responseString = await response.Content.ReadAsStringAsync();

            return XElement.Parse(responseString).ToReference();

        }

        public async Task<Reference> CreateReference(string accessToken, Reference reference)
        {
            StringBuilder requestUri = new StringBuilder()
                 .Append(SiteUrl)
                 .Append("/_api/web/lists/getbyTitle('")
                 .Append(ReferencesListName)
                 .Append("')/items");

            if (reference.Title == null || reference.Title.Length == 0)
            {
                reference.Title = await GetTitleFromLink(reference.Url);
            }

            XElement entry = reference.ToXElement();

            StringContent requestContent = new StringContent(entry.ToString());
            HttpResponseMessage response = await Post(requestUri.ToString(), accessToken, requestContent);
            string responseString = await response.Content.ReadAsStringAsync();

            return XElement.Parse(responseString).ToReference();

        }

        public async Task<bool> UpdateReference(string accessToken, Reference reference)
        {
            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ReferencesListName)
                .Append("')/getItemByStringId('")
                .Append(reference.Id.ToString())
                .Append("')");

            XElement entry = reference.ToXElement();

            StringContent requestContent = new StringContent(entry.ToString());
            HttpResponseMessage response = await Patch(requestUri.ToString(), accessToken, reference.__eTag, requestContent);
            return response.IsSuccessStatusCode;

        }

        public async Task<bool> DeleteReference(string accessToken, int Id, string eTag)
        {
            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ReferencesListName)
                .Append("')/getItemByStringId('")
                .Append(Id.ToString())
                .Append("')");


            HttpResponseMessage response = await Delete(requestUri.ToString(), accessToken, eTag);
            return response.IsSuccessStatusCode;

        }
        private async Task<string> GetTitleFromLink(string Url)
        {
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, Url);
            HttpResponseMessage response = await client.SendAsync(request);
            string responseString = await response.Content.ReadAsStringAsync();
            Match match = Regex.Match(responseString, @"<title>\s*(.+?)\s*</title>");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            else
            {
                return "Unknown Title";
            }

        }

    }

    public class ReferencesCSOMRepository : CSOMRepository, IReferencesRepository
    {
        public async Task<IEnumerable<Reference>> GetReferences(string accessToken)
        {
            // Get the listitems
            ListItemCollection items = await GetListItemCollectionAsync(accessToken, ReferencesListName);

            // Return
            return items.ToList().Select(item => item.ToReference());
        }

        public async Task<Reference> GetReference(string accessToken, int Id, string eTag)
        {
            // Get listitem
            ListItem item = await GetListItemAsync(accessToken, ReferencesListName, Id);

            // Return
            return item.ToReference();
        }

        public async Task<Reference> CreateReference(string accessToken, Reference reference)
        {
            // Create listitem
            ListItem item = await CreateListItemAsync(accessToken, ReferencesListName, reference.ToDictionary());

            // Return
            return item.ToReference();
        }

        public async Task<bool> UpdateReference(string accessToken, Reference reference)
        {
            bool result = await UpdateListItemAsync(accessToken, ReferencesListName, reference.Id, reference.ToDictionary());

            return result;
        }

        public async Task<bool> DeleteReference(string accessToken, int Id, string eTag)
        {
            bool result = await DeleteListItemAsync(accessToken, ReferencesListName, Id);

            return result;
        }
    }
}