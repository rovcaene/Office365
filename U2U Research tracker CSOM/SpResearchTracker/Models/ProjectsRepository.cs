﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using SpResearchTracker.Utils;
using Microsoft.SharePoint.Client;
using System.Configuration;

namespace SpResearchTracker.Models
{
    public interface IProjectsRepository
    {
        Task<IEnumerable<Project>> GetProjects(string accessToken);
        Task<Project> GetProject(string accessToken, int Id, string eTag);
        Task<Project> CreateProject(string accessToken, Project project);
        Task<bool> UpdateProject(string accessToken, Project project);
        Task<bool> DeleteProject(string accessToken, int Id, string eTag);
    }
    public class ProjectsRepository : RESTRepository, IProjectsRepository
    {
        public async Task<IEnumerable<Project>> GetProjects(string accessToken)
        {
            List<Project> projects = new List<Project>();
            
            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ProjectsListName)
                .Append("')/items?$select=ID,Title");

            HttpResponseMessage response = await Get(requestUri.ToString(), accessToken);
            string responseString = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                throw new Exception(responseString);
            }
            XElement root = XElement.Parse(responseString);
            
            foreach (XElement entryElem in root.Elements().Where(e => e.Name.LocalName == "entry"))
            {
                projects.Add(entryElem.ToProject());
            }

            return projects.AsQueryable();
        }

        public async Task<Project> GetProject(string accessToken, int Id, string eTag)
        {
            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ProjectsListName)
                .Append("')/getItemByStringId('")
                .Append(Id)
                .Append("')?$select=ID,Title");

            HttpResponseMessage response = await Get(requestUri.ToString(), accessToken, eTag);
            string responseString = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                throw new Exception(responseString);
            }

            return XElement.Parse(responseString).ToProject();

        }

        public async Task<Project> CreateProject(string accessToken, Project project)
        {
            StringBuilder requestUri = new StringBuilder()
                 .Append(SiteUrl)
                 .Append("/_api/web/lists/getbyTitle('")
                 .Append(ProjectsListName)
                 .Append("')/items");

            XElement entry = project.ToXElement();

            StringContent requestContent = new StringContent(entry.ToString());
            HttpResponseMessage response = await Post(requestUri.ToString(), accessToken, requestContent);
            string responseString = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                throw new Exception(responseString);
            }

            return XElement.Parse(responseString).ToProject();

        }

        public async Task<bool> UpdateProject(string accessToken, Project project)
        {
            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ProjectsListName)
                .Append("')/getItemByStringId('")
                .Append(project.Id)
                .Append("')");

            XElement entry = project.ToXElement();

            StringContent requestContent = new StringContent(entry.ToString());
            HttpResponseMessage response = await Patch(requestUri.ToString(), accessToken, project.__eTag, requestContent);
            return response.IsSuccessStatusCode;
        }

        public async Task<bool> DeleteProject(string accessToken, int Id, string eTag)
        {
            StringBuilder requestUri = new StringBuilder()
                .Append(SiteUrl)
                .Append("/_api/web/lists/getbyTitle('")
                .Append(ProjectsListName)
                .Append("')/getItemByStringId('")
                .Append(Id)
                .Append("')");

            HttpResponseMessage response = await Delete(requestUri.ToString(), accessToken, eTag);
            return response.IsSuccessStatusCode;

        }
    }

    public class ProjectsCSOMRepository : CSOMRepository, IProjectsRepository
    {
        public async Task<IEnumerable<Project>> GetProjects(string accessToken)
        {
            // Get the listitems
            ListItemCollection items = await GetListItemCollectionAsync(accessToken, ProjectsListName);

            // Return result
            return items.ToList().Select(item => item.ToProject());
        }

        public async Task<Project> GetProject(string accessToken, int Id, string eTag)
        {
            Project project = null;

            // Get the listitem
            ListItem item = await GetListItemAsync(accessToken, ProjectsListName, Id);

            // Set project
            project = item.ToProject();

            return project;
        }

        public async Task<Project> CreateProject(string accessToken, Project project)
        {
            Project result = null;

            // Create listitem
            ListItem item = await CreateListItemAsync(accessToken, ProjectsListName, project.ToDictionary());

            // Set project
            result = item.ToProject();

            return result;
        }

        public async Task<bool> UpdateProject(string accessToken, Project project)
        {
            bool result = await UpdateListItemAsync(accessToken, ProjectsListName, project.Id, project.ToDictionary());

            return result;
        }

        public async Task<bool> DeleteProject(string accessToken, int Id, string eTag)
        {
            bool result = await DeleteListItemAsync(accessToken, ProjectsListName, Id);

            return result;
        }
    }
}