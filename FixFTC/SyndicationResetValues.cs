using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace FixFTC
{
    public class SyndicationResetValues
    {
        public List<string> ContentTypesToRemove;
        public List<string> ContentTypesToRenameAndHide;
        public List<string> FieldNamesToRemove;
        public List<string> SiteCollections;

        public bool UsePassword;
        public SecureString Password;
        public string UserName;

        public bool ProcessFields;
        public bool ProcessCTHide;
        public bool ProcessCTRemove;
        public bool ProcessRefreshCTFlag;

        public string ContentTypeHubURL;

    }
}
