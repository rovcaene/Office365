using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using U2U.Provisioning.SiteCollectionCreation.Repositories;
using model = U2U.Provisioning.SiteCollectionCreation.Models;

namespace U2U.Provisioning.SiteCollectionCreation.Controllers
{
    [Authorize]
    public class LanguageController : ApiController
    {
        private SharePointRepository _sharePointRepository; 

        public LanguageController()
        {
            _sharePointRepository = new SharePointRepository();
        }

        public async Task<IEnumerable<model.LanguageVM>> Get()
        {
            // Get the languages
            var availableLanguages = await _sharePointRepository.GetAvailableLanguagesAsync();

            // Return the available languages
            return availableLanguages;
        }
    }
}
