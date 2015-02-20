using System.Web.Mvc;

namespace U2U.Provisioning.SiteCollectionCreation
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
