using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace U2U.Provisioning.SiteCollectionCreation.Utils
{
    public static class ExtensionMethods
    {
        public static Task ExecuteQueryAsync(this ClientContext clientContext)
        {
            return Task.Factory.StartNew(() =>
            {
                clientContext.ExecuteQuery();
            });
        }
    }
}