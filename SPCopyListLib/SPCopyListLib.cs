using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Json;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SPCopyListLib
{
    public class SPCopyListLib
    {
        public string SourceSite { get; set; }
        public string TargetSite { get; set; }
        public ClientContext SourceContext { get; set; }
        public ClientContext TargetContext { get; set; }
        private const int MAX_ITEMS = 2000;

        public enum SPCopyListCopyOptions {
            AbortOnDuplicate, Rename, Overwrite, Merge
        }

        public SPCopyListCopyResult SPCopyList(string sourceSite, string targetSite, string listName)
        {
            return SPCopyList(sourceSite, targetSite, listName, SPCopyListCopyOptions.Merge);

        }

        public SPCopyListCopyResult SPCopyList(string sourceSite, string targetSite, string listName, SPCopyListCopyOptions options)
        {
            SPCopyListCopyResult result = new SPCopyListCopyResult();
            CheckConnected();

            Web web = SourceContext.Web;
            ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(web);
            JsonTemplateProvider jsonProvider = new JsonFileSystemTemplateProvider();
            var template = web.GetProvisioningTemplate();

            template.Lists.RemoveAll(l => l.Title != listName);

            Web targetWeb = TargetContext.Web;
            var ptai = new ProvisioningTemplateApplyingInformation();
            ptai.HandlersToProcess = Handlers.Lists;

            TargetContext.Load(targetWeb.Lists);
            TargetContext.ExecuteQueryRetry();

            bool listExists = ListExists(listName, targetWeb.Lists);

            //TODO - remove
            var debugList = template.Lists.Find(l => l.Title == listName);


            if (options == SPCopyListCopyOptions.AbortOnDuplicate)
            {
                if (listExists)
                {
                    result.Conflict = true;
                    result.ListCopied = false;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.Aborted;
                } else
                {
                    result.Conflict = false;
                    result.ListCopied = true;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.None;

                    targetWeb.ApplyProvisioningTemplate(template, ptai);
                }
            }
            else if (options == SPCopyListCopyOptions.Merge)
            {
                if (listExists)
                {
                    result.Conflict = true;
                    result.ListCopied = true;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.Merged;
                }
                else
                {
                    result.Conflict = false;
                    result.ListCopied = true;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.None;
                }

                targetWeb.ApplyProvisioningTemplate(template, ptai);

            }
            else if (options == SPCopyListCopyOptions.Overwrite)
            {
                if (listExists)
                {
                    result.Conflict = true;
                    result.ListCopied = true;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.Overwritten;

                    DeleteTargetList(targetWeb, listName);
                }
                else
                {
                    result.Conflict = false;
                    result.ListCopied = true;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.None;
                }

                targetWeb.ApplyProvisioningTemplate(template, ptai);

            }
            else if (options == SPCopyListCopyOptions.Rename)
            {
                if (listExists)
                {
                    result.Conflict = true;
                    result.ListCopied = true;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.Renamed;

                    var list = template.Lists.Find(l => l.Title == listName);
                    list.Title = GenerateUniqueListName(targetWeb.Lists, listName);
                    list.Url = AlignUrlWithTitle(list.Title, list.Url);
                    
                }
                else
                {
                    result.Conflict = false;
                    result.ListCopied = true;
                    result.ConflictAction = SPCopyListCopyResult.SPCopyListConflictAction.None;

                }

                targetWeb.ApplyProvisioningTemplate(template, ptai);
                

            }

            return result;
        }

        private string AlignUrlWithTitle(string title, string url)
        {
            string result = "";
            int slashIndex = url.LastIndexOf("/");
            if (slashIndex == -1)
            {
                result = title;
            }
            else
            {
                string firstPart = url.Substring(0, slashIndex);
                string secondPart = title;
                result = firstPart + "/" + secondPart;
            }

            return result;
        }

        private string GenerateUniqueListName(ListCollection lists, string listName)
        {
            string result = listName;
            if (!ListExists(listName, lists)) return listName;

            for (int i = 1; i <= (MAX_ITEMS + 1); i++)
            {
                result = listName + i.ToString();

                if (i == (MAX_ITEMS + 1)) throw new Exception("Max number of lists reached");

                if (!ListExists(result, lists))
                    break;
            }
            
            return result;
        }

        private void DeleteTargetList(Web targetWeb, string listName)
        {
            List oList = targetWeb.Lists.GetByTitle(listName);
            oList.DeleteObject();
            TargetContext.ExecuteQueryRetry();
        }

        private bool ListExists(string listName, ListCollection lists)
        {
            bool result = false;
            foreach (var item in lists)
            {
                if (item.Title == listName)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        private void CheckConnected()
        {
            if (SourceContext == null || TargetContext == null) throw new Exception("Source or Target context is not connected");
        }

        public void ConnectSource(string site, string user, string password)
        {
            SourceSite = site;
            SourceContext = new ClientContext(site);
            SecureString spwd = new SecureString();
            foreach (char c in password.ToCharArray()) spwd.AppendChar(c);
            SourceContext.Credentials = new SharePointOnlineCredentials(user, spwd);
            Web web = SourceContext.Web;
            SourceContext.ExecuteQueryRetry();
        }

        public void ConnectTarget(string site, string user, string password)
        {
            TargetSite = site;
            TargetContext = new ClientContext(site);
            SecureString spwd = new SecureString();
            foreach (char c in password.ToCharArray()) spwd.AppendChar(c);
            TargetContext.Credentials = new SharePointOnlineCredentials(user, spwd);
            Web web = TargetContext.Web;
            TargetContext.ExecuteQueryRetry();
        }

    }

    public class SPCopyListCopyResult
    {
        public SPCopyListCopyResult()
        {
            ConflictAction = SPCopyListConflictAction.None;
        }

        public enum SPCopyListConflictAction
        {
            None, Aborted, Renamed, Overwritten, Merged
        }

        public bool ListCopied { get; set; }
        public bool Conflict { get; set; }
        public SPCopyListConflictAction ConflictAction { get; set; }

    }
}
