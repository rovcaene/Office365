using Microsoft.SharePoint.Client;
using SpResearchTracker.Controllers;
using SpResearchTracker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Helpers;
using System.Xml.Linq;

namespace SpResearchTracker.Utils
{
    public static class ExtensionMethods
    {
        public static XNamespace atom = "http://www.w3.org/2005/Atom";
        public static XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";
        public static XNamespace m = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";

         public static Project ToProject(this XElement root)
        {
            string etag = root.Attribute(m + "etag").Value;
            string sptype = root.Elements().Where(e => e.Name.LocalName == "category").First().Attribute("term").Value;
            XElement properties = root.Descendants(m + "properties").First();
            int id = int.Parse(properties.Descendants(d + "ID").First().Value);
            string title = properties.Descendants(d + "Title").First().Value;
            return new Project(id, etag, title);
        }

         public static Project ToProject(this ListItem listItem)
         {
             int id = listItem.Id;
             string title = listItem.FieldValues["Title"] as string;

             DateTime modified = (DateTime)listItem.FieldValues["Modified"];
             string etag = modified.ToShortDateString() + " " + modified.ToLongTimeString(); 

             return new Project(id, etag, title);
         }

        public static Reference ToReference(this XElement root)
        {
            string etag = root.Attribute(m + "etag").Value;
            string sptype = root.Elements().Where(e => e.Name.LocalName == "category").First().Attribute("term").Value;
            XElement properties = root.Descendants(m + "properties").First();
            int id = int.Parse(properties.Descendants(d + "ID").First().Value);
            string notes = properties.Descendants(d + "Comments").First().Value;
            string project = properties.Descendants(d + "Project").First().Value;
            XElement URL = properties.Descendants(d + "URL").First();
            string url = URL.Descendants(d + "Url").First().Value;
            string title = URL.Descendants(d + "Description").First().Value;
            return new Reference(id, etag, title, url, notes, project);
        }

        public static Reference ToReference(this ListItem listItem)
        {
            int id = listItem.Id;
            string title = "";
            string url = "";

            FieldUrlValue fieldUrlValue = listItem.FieldValues["URL"] as FieldUrlValue;

            if (fieldUrlValue != null)
            {
                title = fieldUrlValue.Description;
                url = fieldUrlValue.Url;
            }

            string notes = listItem.FieldValues["Comments"] as string;
            string project = listItem.FieldValues["Project"] as string;

            DateTime modified = (DateTime)listItem.FieldValues["Modified"];
            string etag = modified.ToShortDateString() + " " + modified.ToLongTimeString(); 

            return new Reference(id, etag, title, url, notes, project);
        }

        public static String GetFieldValueString<T>(this ListItem listItem, string fieldName) where T : ClientValueObject
        {
            string result = null;

            // get the fieldvalue
            T fieldValue = listItem[fieldName] as T;

            if (fieldValue != null)
            {
                Type tType = typeof(T);

                if (tType == typeof(FieldUrlValue))
                {
                    result = (fieldValue as FieldUrlValue).Url;
                }
                else
                {
                    // fall back on ToString
                    result = fieldValue.ToString();
                }
            }

            return result;
        }

        public static Task ExecuteQueryAsync(this ClientContext clientContext)
        {
            return Task.Factory.StartNew(() =>
            {
                clientContext.ExecuteQuery();
            });
        }
    }
}