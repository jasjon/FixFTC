using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace FixFTC
{
    public class ColumnRemoval
    {
        private ClientContext _ctx;
        private ContentTypeCollection _cts;
        private List<string> _CTsProcessedViaHierarchy = new List<string>();
        private Dictionary<string, Guid> _cthCols;
        
        private SyndicationResetValues _resetValues;
        private FieldCollection _siteFields;

        /// <summary>
        /// Returns a client context for connection to a provided site collection
        /// Optionally looks at the SyndicationResetValues to determine if a 
        /// non-integrated account should be used
        /// </summary>
        /// <param name="SiteCollectionURL"></param>
        /// <returns></returns>
        private ClientContext GetClientContext(string SiteCollectionURL)
        {
            ClientContext ctx = new ClientContext(SiteCollectionURL);
            if (_resetValues.UsePassword)
            {
                var securePassword = _resetValues.Password;
                System.Net.NetworkCredential _creds = new System.Net.NetworkCredential(_resetValues.UserName, securePassword);

                ctx.Credentials = _creds;
            }
            return ctx;
        }

        /// <summary>
        /// Removes all references to the list of fields.
        /// this deletes field links from content types.
        /// Only field links that refer to fields that have a different GUID
        /// to the equivalent one in the CTH are removed
        /// </summary>
        /// <param name="ResetValues"></param>
        /// <returns></returns>
        public string RemoveSCFields(SyndicationResetValues ResetValues)
        {
            string outputReport = String.Format("Report generated on: {0}{1}", DateTime.Now, Environment.NewLine);
            string stage = "Start";
            _resetValues = ResetValues;

            try
            {
                using (_ctx = GetClientContext(_resetValues.ContentTypeHubURL))
                {
                    stage = "Set CTH context";
                    GetContentHubColumnGUIDs();
                    stage = "Retrieved CTH values";
                    outputReport += "GUIDs read from the Hub" + Environment.NewLine;
                }

                stage = "Processing site collections";
                ProcessSiteCollections(ResetValues, ref outputReport);

                outputReport += "*************** Finished  ***************";
                stage = "Site collections processed";
            }
            catch (Exception ex)
            {
                outputReport += String.Format("{3}{3}ERROR: {0} {3}{1}{3}{3}At stage: {2}", ex.Message, ex.InnerException, stage, Environment.NewLine);
            }

            return outputReport;
        }

        /// <summary>
        /// Wrapper function to cover several activities (looping around all site collections)
        /// 1. Getting content types for the SC
        /// 2. removing field links to the fields
        /// 3. removing the fields
        /// 4. Hide and rename of content types
        /// 5. removal of content types
        /// 6. Reset the 'refresh CT publishing' flag
        /// </summary>
        /// <param name="ResetValues"></param>
        /// <param name="outputReport"></param>
        private void ProcessSiteCollections(SyndicationResetValues ResetValues, ref string outputReport)
        {
            Dictionary<string, string> jobOutput = new Dictionary<string, string>();
            foreach (string siteCollectionURL in _resetValues.SiteCollections)
            {
                ProgressReport scReport = new ProgressReport(siteCollectionURL);
                using (_ctx = GetClientContext(siteCollectionURL))
                {
                    RetreiveContentTypes();
                    scReport.AddEntry("Content types read from target site collection");

                    if (_resetValues.ProcessFields)
                    {
                        scReport.AddEntry("*** Removing field links");
                        jobOutput = RemoveFieldLinks();
                        AddJobOutputToReport(jobOutput, ref scReport);


                        scReport.AddEntry(String.Format("*** Removing Site Columns: {0} columns to process", _cthCols.Keys.Count));
                        jobOutput = RemoveSiteColumns();
                        AddJobOutputToReport(jobOutput, ref scReport);
                    }

                    if (_resetValues.ProcessCTHide)
                    {
                        scReport.AddEntry(String.Format("*** Rename and hiding {0} content types", ResetValues.ContentTypesToRenameAndHide.Count));
                        jobOutput = RenameAndHideContentTypes(ResetValues.ContentTypesToRenameAndHide);
                        AddJobOutputToReport(jobOutput, ref scReport);
                    }

                    if (_resetValues.ProcessCTRemove)
                    {
                        scReport.AddEntry(String.Format("*** Removing {0} content types", ResetValues.ContentTypesToRemove.Count));
                        jobOutput = RemoveContentTypes(ResetValues.ContentTypesToRemove);
                        AddJobOutputToReport(jobOutput, ref scReport);
                    }

                    if (_resetValues.ProcessRefreshCTFlag)
                    {
                        scReport.AddEntry("*** Setting the 'Refresh Content Types' Flag");
                        SetRefreshContentTypesFlag();
                    }

                }
                outputReport += scReport.Output() + Environment.NewLine;
                outputReport += "-----------------------------------------------------------" + Environment.NewLine + Environment.NewLine;
            }
        }

        /// <summary>
        /// sets the site collection property "metadatatimestamp" which
        /// will trigger a content type refresh
        /// </summary>
        private void SetRefreshContentTypesFlag()
        {
            var rootWebProperties = _ctx.Site.RootWeb.AllProperties;
            _ctx.Load(rootWebProperties);
            _ctx.ExecuteQuery();

            if (rootWebProperties["metadatatimestamp"].ToString() != String.Empty)
            {
                rootWebProperties["metadatatimestamp"] = String.Empty;
                _ctx.ExecuteQuery();
            }
        }
        /// <summary>
        /// Deletes content types (provided in the CTToProcess param)
        /// </summary>
        /// <param name="CTToProcess"></param>
        /// <returns></returns>
        private Dictionary<string, string> RemoveContentTypes(List<string> CTToProcess)
        {
            Dictionary<string, string> jobOutput = new Dictionary<string, string>();
            foreach (string ctName in CTToProcess)
            {
                
                foreach (var ct in _cts)
                {
                    if (ct.Name == ctName)
                    {   
                        try
                        {
                            ct.ReadOnly = false;
                            ct.DeleteObject();
                            _ctx.ExecuteQuery();
                            jobOutput.Add(ct.Name, "Removed content type");
                            break;
                        }
                        catch (Exception ex)
                        {
                            jobOutput.Add(ct.Name, String.Format("Failed to remove content type:\t{0}", ex.Message));
                        }
                    }
                }
            }
            

            return jobOutput;
        }

        /// <summary>
        /// Renames CTs to Retired_<old CT Name> and then sets their property to hidden
        /// Only processes CTs that have not already been renamed.
        /// </summary>
        /// <param name="CTToProcess"></param>
        /// <returns></returns>
        private Dictionary<string, string> RenameAndHideContentTypes(List<string> CTToProcess)
        {
            Dictionary<string, string> jobOutput = new Dictionary<string, string>();
            string jobText = String.Empty;
            foreach(string ctToHide in CTToProcess)
            {
                bool alreadyHidden = false;
                foreach (var ct in _cts)
                {
                    if (ct.Name == "Retired_" + ctToHide)
                    {
                        alreadyHidden = true;
                        jobText += String.Format("Content Type '{0}' has already been renamed{1}", ctToHide, Environment.NewLine);
                        break;
                    }
                }
                if (! alreadyHidden)
                {
                    foreach (var ct in _cts)
                    {
                        if (ct.Name == ctToHide)
                        {
                            jobText += String.Format("Renaming and hiding {0} type{1}", ct.Name, Environment.NewLine);
                            ct.ReadOnly = false;
                            ct.Name = "Retired_" + ct.Name;
                            ct.Hidden = true;
                            ct.Group = "Retired EY Content Types";
                          
                            ct.Update(false);
                            _ctx.ExecuteQuery();

                            ct.ReadOnly = true;
                            _ctx.ExecuteQuery();
                        }
                    }
                }
            }
            
            jobOutput.Add("RenameAndHideContentTypes", jobText);

            return jobOutput;
        }

        /// <summary>
        /// Helper method to add a dictionary object to the output report for 
        /// improved styling.
        /// </summary>
        /// <param name="JobOutput"></param>
        /// <param name="report"></param>
        private void AddJobOutputToReport(Dictionary<string, string> JobOutput, ref ProgressReport report)
        {
            foreach(string processColumn in JobOutput.Keys)
            {
                report.AddEntry(JobOutput[processColumn]);
            }
        }

        /// <summary>
        /// Retrieves the full set of content types and site columns from the 
        /// pre-configured content type hub
        /// </summary>
        private void GetContentHubColumnGUIDs()
        {
            var cthCols = _ctx.Site.RootWeb.Fields;
            _cthCols = new Dictionary<string, Guid>();
            _ctx.Load(cthCols);
            _ctx.ExecuteQuery();

            foreach (string internalName in _resetValues.FieldNamesToRemove)
            {
                foreach (Field sc in cthCols)
                {
                    if (sc.InternalName == internalName)
                    {
                        _cthCols.Add(internalName, sc.Id);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Removes columns from the site collection
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, string> RemoveSiteColumns()
        {
            Dictionary<string, string> progressReport = new Dictionary<string, string>();
            string columnDiscovery;
            List<Field> siteFields = _siteFields.ToList<Field>();

            foreach (string columnName in _cthCols.Keys)
            {
                columnDiscovery = String.Format("Site column '{0}' Not Found", columnName);
                foreach (Field sfld in siteFields)
                {
                    if (sfld.InternalName == columnName)
                    {
                        if (sfld.Id != _cthCols[columnName]) //Good the local definition is not the same as the Hub
                        {
                            sfld.DeleteObject();
                            _ctx.ExecuteQuery();
                            columnDiscovery = String.Format("Site column '{0}' was deleted", columnName);
                        }
                        else
                        {
                            columnDiscovery = String.Format("Site Column '{0}' shares CTH GUID - not processed", columnName);
                        }
                        
                    }
                }
                progressReport.Add(columnName, columnDiscovery);
            }
            return progressReport;
        }

        /// <summary>
        /// Removes field links from a content type, disassociating the field
        /// from the content type.
        /// This is required before the field can be deleted
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, string> RemoveFieldLinks()
        {
            Dictionary<string, string> progressReport = new Dictionary<string, string>();
            string columnDiscovery;
            bool columnFound = false;
            foreach (string columnName in _cthCols.Keys)
            {
                columnFound = false;
                columnDiscovery = String.Format("Processing {0}{1}", columnName, Environment.NewLine);
                var fieldToRemove = GetFieldToRemove(columnName, _cthCols[columnName]);
                if (fieldToRemove != null)
                {   
                    foreach (var contentType in _cts)
                    {
                        var fieldLinks = contentType.FieldLinks;
                        FieldLink linkToRemove = fieldLinks.FirstOrDefault(l => l.Id == fieldToRemove.Id);
                        if (linkToRemove != null)
                        {
                            //OpenCTHierarchy(contentType.Id.StringValue);
                            linkToRemove.DeleteObject();
                            contentType.Update(false);
                            _ctx.ExecuteQuery();
                            columnDiscovery += String.Format("\tField link '{0}' removed from content type '{1}'{2}", columnName, contentType.Name, Environment.NewLine);
                            columnFound = true;
                        }
                    }
                }
                if (columnFound == false)
                {
                    columnDiscovery += String.Format("\tField link '{0}' was not found in any of the Content Types{1}", columnName, Environment.NewLine);
                }

                progressReport.Add(columnName, columnDiscovery);
            }

            return progressReport;
        }

        /// <summary>
        /// Helper method to ensure that the field about to be processed is not
        /// identical to the one in the CTH as these do not need to be processed.
        /// </summary>
        /// <param name="FieldName"></param>
        /// <param name="ContentTypeHubGUID"></param>
        /// <returns></returns>
        private Field GetFieldToRemove(string FieldName, Guid ContentTypeHubGUID)
        {
            var field = _siteFields.FirstOrDefault(x => x.InternalName == FieldName & x.Id != ContentTypeHubGUID);
            return field;
        }

        
        /// <summary>
        /// Not Used
        /// </summary>
        /// <param name="contentTypeId"></param>
        private void OpenCTHierarchy(string contentTypeId)
        {
            bool updateNeeded = false;

            foreach (var ctToProcess in _cts.Where(x => x.Parent.Id.StringValue == contentTypeId))
            {
                _CTsProcessedViaHierarchy.Add(ctToProcess.Name);
                if (ctToProcess.Sealed == true)
                {
                    ctToProcess.Sealed = false;
                    updateNeeded = true;
                }

                if (ctToProcess.ReadOnly == true)
                {
                    ctToProcess.ReadOnly = false;
                    updateNeeded = true;
                }

                if (updateNeeded)
                {
                    ctToProcess.Update(true);
                    _ctx.ExecuteQuery();
                }

                //recurse ...

                OpenCTHierarchy(ctToProcess.Id.StringValue);
            }


            updateNeeded = false;
            var thisCT = _cts.First(x => x.Id.StringValue == contentTypeId);

            if (thisCT.Sealed == true) { thisCT.Sealed = false; updateNeeded = true; }
            if (thisCT.ReadOnly == true) { thisCT.ReadOnly = false; updateNeeded = true; }
            if (updateNeeded)
            {
                thisCT.Update(true);
                _ctx.ExecuteQuery();
            }
        }

        /// <summary>
        /// Gets the content types (explicitly declared) from the 
        /// site collection that is currently being processed
        /// </summary>
        private void RetreiveContentTypes()
        {
            _siteFields = _ctx.Site.RootWeb.Fields;
            _ctx.Load(_siteFields);
            _cts = _ctx.Site.RootWeb.ContentTypes;
            _ctx.Load(_cts);
            _ctx.ExecuteQuery();

            foreach (var contentType in _cts)
            {
                _ctx.Load(contentType,
                    ct => ct.Parent.Name,
                    ct => ct.Parent.Id,
                    ct => ct.Sealed,
                    ct => ct.Hidden,
                    ct => ct.Id, ct => ct.FieldLinks);
            }
            _ctx.ExecuteQuery();
           

        }

    }        
}
