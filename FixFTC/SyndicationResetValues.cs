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
        public bool ProcessRemoveFieldsFromLists;

        public string ContentTypeHubURL;
        public string FailReason;


        public bool ValidSettings()
        {
            bool valid = true;
            bool siteProcessing = ProcessFields || ProcessCTHide || ProcessCTRemove || ProcessRefreshCTFlag;
            if (siteProcessing && ProcessRemoveFieldsFromLists)
            {
                valid = false;
            }
            FailReason = "This tool does not support running the removal of fields from lists if one or more of the other processing properties is set to true.";
            return valid;
        }


    }
}
