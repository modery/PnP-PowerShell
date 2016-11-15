using System.Management.Automation;
using Microsoft.SharePoint.Client;
using web = Microsoft.SharePoint.Client.Web;
using SharePointPnP.PowerShell.CmdletHelpAttributes;
using SharePointPnP.PowerShell.Commands.Base.PipeBinds;
using System.Collections.Generic;

namespace SharePointPnP.PowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, "PnPRecycleBinItems")]
    [CmdletHelp("Returns all recycle bin items of the current web",
        Category = CmdletHelpCategory.Utilities,
        OutputType = typeof(RecycleBinItemCollection),
        OutputTypeLink = "https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.recyclebinitemcollection.aspx")]
    public class GetRecycleBinItems : SPOWebCmdlet
    {

        [Parameter(Mandatory = false)]
        public string pagingInfo;

        [Parameter(Mandatory = false, HelpMessage = "Number of rows to return. Defaults to 50")]
        public int rowLimit = 50;

        [Parameter(Mandatory = false, HelpMessage = "Sort in ascending order. Defaults to true")]
        public SwitchParameter isAscending = true;

        [Parameter(Mandatory = false, HelpMessage ="Property to be sorted on")]
        public RecycleBinOrderBy orderBy = RecycleBinOrderBy.Min;

        [Parameter(Mandatory = false)]
        public RecycleBinItemState itemState = RecycleBinItemState.FirstStageRecycleBin;

        [Parameter(Mandatory = false)]
        public bool showOnlyMyItems = false;


        protected override void ExecuteCmdlet()
        {
            var recQuery = new RecycleBinQueryInformation();

            if (!string.IsNullOrEmpty(pagingInfo))
            {
                recQuery.PagingInfo = pagingInfo;
            }
            recQuery.RowLimit = rowLimit;
            recQuery.IsAscending = isAscending;
            recQuery.OrderBy = orderBy;
            recQuery.ItemState = itemState;
            recQuery.ShowOnlyMyItems = showOnlyMyItems;
            var RecycleBinItemCollection = SelectedWeb.Context.LoadQuery(SelectedWeb.GetRecycleBinItemsByQueryInfo(recQuery));
            SelectedWeb.Context.ExecuteQueryRetry();
            WriteObject(RecycleBinItemCollection, true);
        }

    }
}
