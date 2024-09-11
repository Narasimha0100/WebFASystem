using Aspose.Cells;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security;

namespace AssetMonthlyUpdate
{
    class Program
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            logInfo("Started");
            UpdateFARequestItems();
            logInfo("Ended");
            //Console.ReadLine();
        }

        private static void UpdateFARequestItems()
        {
            logInfo("START UpdateFARequestItems");
            string _siteURL = Convert.ToString(ConfigurationManager.AppSettings["SiteURL"]);
            string _userName = Convert.ToString(ConfigurationManager.AppSettings["UserName"]);
            string _password = Convert.ToString(ConfigurationManager.AppSettings["Password"]);

            using (ClientContext clientContext = new ClientContext(_siteURL))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(_userName, GetSecureStringPassword(_password));
                clientContext.RequestTimeout = -1;
                logInfo("Initialized client context");
                try
                {
                    List<AssetHistoryExcelData> ahedListOriginal = ReadExcelFileAsObject();
                    List<AssetHistoryExcelData> ahedList = ahedListOriginal.Where(we => !String.IsNullOrEmpty(we.AssetNo_YGS)).Select(se => se).ToList();
                    logInfo("Total excel rows : " + ahedListOriginal.Count);
                    logInfo("Filtered excel rows (records without blank asset no) : " + ahedList.Count);
                    if (ahedList.Count > 0)
                    {
                        //var FA_M_Company = GetListItems(clientContext, "FA_M_Company");
                        var FA_M_CostCentre = GetListItems(clientContext, "FA_M_CostCentre");
                        //var FA_M_AssetClass = GetListItems(clientContext, "FA_M_AssetClass");
                        //var FA_M_Location = GetListItems(clientContext, "FA_M_Location");
                        //var FA_M_Currency = GetListItems(clientContext, "FA_M_Currency");
                        //var FA_M_Employee = GetListItems(clientContext, "FA_M_Employee");
                        //var FA_NewRequest_Items = GetListItems(clientContext, "FA_NewRequest_Items");
                        //var FA_NewRequest_List = GetListItems(clientContext, "FA_NewRequest_List");

                        int cnt = 0;
                        foreach (var excelRow in ahedList)
                        {
                            logInfo("----------------------------------------------------------");
                            logInfo("Processing Item : " + excelRow.AssetNo_YGS);
                            CamlQuery camlQuery = new CamlQuery();
                            camlQuery.ViewXml = @"<View>
				                <Query>
					                <Where>
						                <Eq>
							                <FieldRef Name='AssetNo_YGS' />
							                <Value Type='Text'>" + excelRow.AssetNo_YGS + @"</Value>
						                </Eq>
					                </Where>
				                </Query>
			                </View>";

                            ListItemCollection newRequestItems = clientContext.Web.Lists.GetByTitle("FA_NewRequest_Items").GetItems(camlQuery);
                            clientContext.Load(newRequestItems);
                            clientContext.ExecuteQuery();

                            if (newRequestItems != null)
                            {
                                logInfo("No of Items filtered Items : " + newRequestItems.Count);
                                if (newRequestItems.AreItemsAvailable && newRequestItems.Count > 0)
                                {
                                    var itemsList = newRequestItems.FirstOrDefault();
                                    int AssetRequestNo = 0;
                                    if (itemsList["AssetRequestNo"] != null)
                                    {
                                        var AssetRequestNoFieldLookup = itemsList["AssetRequestNo"] as FieldLookupValue;
                                        if (AssetRequestNoFieldLookup != null)
                                        {
                                            excelRow.AssetRequestNo = AssetRequestNoFieldLookup.LookupId.ToString();
                                            AssetRequestNo = AssetRequestNoFieldLookup.LookupId;
                                        }
                                    }
                                    logInfo("Updating the fields : {@value1}", excelRow);
                                    if (!String.IsNullOrEmpty(excelRow.FAInventory))
                                        itemsList["FAInventory"] = excelRow.FAInventory;
                                    if (!String.IsNullOrEmpty(excelRow.CurrentBookValue))
                                        itemsList["CurrentBookValue"] = excelRow.CurrentBookValue;
                                    if (!String.IsNullOrEmpty(excelRow.Costs))
                                        itemsList["Costs"] = excelRow.Costs;
                                    //if (!String.IsNullOrEmpty(excelRow.Quantity))
                                    //    itemsList["Quantity"] = excelRow.Quantity;
                                    if (!String.IsNullOrEmpty(excelRow.SerialNumber))
                                        itemsList["SerialNumber"] = excelRow.SerialNumber;
                                    if (!String.IsNullOrEmpty(excelRow.AssetDescription))
                                        itemsList["AssetDescription"] = excelRow.AssetDescription;
                                    if (!String.IsNullOrEmpty(excelRow.SAPPurchasedDate))
                                    {
                                        DateTime SAPPurchasedDate;
                                        if (DateTime.TryParse(excelRow.SAPPurchasedDate, out SAPPurchasedDate))
                                            itemsList["SAPAssetPurchaseDate"] = excelRow.SAPPurchasedDate;
                                    }
                                    if (!String.IsNullOrEmpty(excelRow.Deactivation_on))
                                    {
                                        DateTime Deactivation_on;
                                        if (DateTime.TryParse(excelRow.Deactivation_on, out Deactivation_on))
                                        {
                                            itemsList["ItemStatus"] = "Disposal Request Submitted";
                                            //itemsList["SAPAssetCreatedDate"] = excelRow.Deactivation_on;
                                            itemsList["SAPAssetCreatedDate"] = Convert.ToDateTime(itemsList["SAPAssetCreatedDate"], new CultureInfo("en-US"));
                                        }
                                    }
                                    itemsList.Update();
                                    clientContext.ExecuteQuery();
                                    logInfo("Updated Item...." );


                                    logInfo("Fetching main list item.....");

                                    CamlQuery camlQuery2 = new CamlQuery();
                                    camlQuery2.ViewXml = @"<View>
                                     <Query>
                                      <Where>
                                       <Eq>
                                        <FieldRef Name='ID' />
                                        <Value Type='Counter'>" + excelRow.AssetRequestNo + @"</Value>
                                       </Eq>
                                      </Where>
                                     </Query>
                                    </View>";

                                    var mainListItem = clientContext.Web.Lists
                                        .GetByTitle("FA_NewRequest_List")
                                        .GetItemById(AssetRequestNo);
                                    clientContext.Load(mainListItem);
                                    clientContext.ExecuteQuery();

                                    if (mainListItem != null)
                                    {
                                        logInfo("Updating mainList Item : " + mainListItem["AssetRequestNo"].ToString());
                                        if (!String.IsNullOrEmpty(excelRow.CostCenter))
                                        {
                                            var CostCenter = FA_M_CostCentre
                                                .Where(we => we.FieldValues["zskl"].ToString() == excelRow.CostCenter).Select(se => se)
                                                .FirstOrDefault();
                                            if (CostCenter != null)
                                            {
                                                mainListItem["CostCenter"] = CostCenter.FieldValues["ID"];
                                                mainListItem.Update();
                                                clientContext.ExecuteQuery();
                                                logInfo("Updated main list item.....");
                                            }
                                        }
                                    }

                                    if (++cnt % 10 == 0)
                                    {
                                        System.Threading.Thread.Sleep(5000);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message);
                    Console.WriteLine(ex.Message);
                }
            }
            logInfo("END UpdateFARequestItems");
        }        
                                                       
        private static List<ListItem> GetListItems(ClientContext context, String listname)
        {
            try
            {
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View>
                                        <Query>
                                            <OrderBy>
                                                <FieldRef Name='ID' Ascending='FALSE'/>
                                            </OrderBy>
                                        </Query>
                                        <RowLimit>5000</RowLimit>
                                    </View>";
                ListItemCollection configItems = context.Web.Lists.GetByTitle(listname).GetItems(camlQuery);
                context.Load(configItems);
                context.ExecuteQuery();

                return configItems.AsEnumerable().Select(x => x).ToList();
            }
            catch (Exception ex)
            {
                string strEx = ex.Message;
            }
            return null;

        }

        private static ListItem GetLatestListItem(ClientContext context, String listname)
        {
            try
            {
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View>
                                            <Query>
                                                <ViewFields>	                                
                                                    <FieldRef Name='ID'/>	                                
                                                    <FieldRef Name='Created' />	                            
                                                </ViewFields>
                                                <OrderBy>
                                                 <FieldRef Name='ID' Ascending ='False' />
                                                </OrderBy>
                                                <RowLimit>1</RowLimit>
                                            </Query>
                                        </View>";
                ListItemCollection configItems = context.Web.Lists.GetByTitle(listname).GetItems(camlQuery);
                context.Load(configItems);
                context.ExecuteQuery();

                return configItems.AsEnumerable().Select(x => x).FirstOrDefault();
            }
            catch (Exception ex)
            {
                string strEx = ex.Message;
            }
            return null;

        }

        public static FieldLookupValue GetLookupFieldFromValue(ClientContext context, string listName, string lookUpFiledName, string lookupValue)
        {
            FieldLookupValue lookUpValue;
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='" + lookUpFiledName + "'/><Value Type='Text'>" + lookupValue + "</Value></Eq></Where></Query></View>";

            ListItemCollection listItems = list.GetItems(query);
            context.Load(listItems);
            context.ExecuteQuery();
            if (listItems.AreItemsAvailable && listItems.Count > 0)
            {
                lookUpValue = new FieldLookupValue() { LookupId = listItems.First().Id };
                return lookUpValue;
            }
            return null;
        }

        private static FieldLookupValue GetLookupFiledFromID(ClientContext context, string listName, string lookUpFiledName, string lookUpID)
        {
            FieldLookupValue lookUpValue;
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View> 
                                <Query> 
                                <Where> 
                                <Eq>
	                                <FieldRef Name='ID' /> 
                                    <Value Type='Number'>" + lookUpID + @"</Value> 
                                </Eq> 
                                </Where> 
                                </Query> 
                            </View>";
            ListItemCollection listItems = list.GetItems(query);
            context.Load(listItems);
            context.ExecuteQuery();
            if (listItems.AreItemsAvailable && listItems.Count > 0)
            {
                lookUpValue = new FieldLookupValue() { LookupId = listItems.First().Id };
                return lookUpValue;
            }
            return null;
        }


        private static DataTable ReadExcelFile()
        {
            DataTable dt = new DataTable();
            string ShareFolderPath = Convert.ToString(ConfigurationManager.AppSettings["ShareFolderPath"]);
            Workbook wb = new Workbook(ShareFolderPath);
            dt = wb.Worksheets[0].Cells.ExportDataTable(0, 0, wb.Worksheets[0].Cells.MaxDataRow + 1, wb.Worksheets[0].Cells.MaxDataColumn + 1, true);
            return dt;
        }

        private static List<AssetHistoryExcelData> ReadExcelFileAsObject()
        {
            DataTable dt = new DataTable();
            string ShareFolderPath = Convert.ToString(ConfigurationManager.AppSettings["ShareFolderPath"]);
            Workbook wb = new Workbook(ShareFolderPath);
            dt = wb.Worksheets[0].Cells.ExportDataTable(0, 0, wb.Worksheets[0].Cells.MaxDataRow + 1, wb.Worksheets[0].Cells.MaxDataColumn + 1, new ExportTableOptions() { CheckMixedValueType = false, ExportColumnName = true, ExportAsString = true });

            List<AssetHistoryExcelData> ahedList = new List<AssetHistoryExcelData>();
            foreach (DataRow row in dt.Rows)
            {
                AssetHistoryExcelData ahed = new AssetHistoryExcelData();
                //if (row["Company Code"] != null)
                //    ahed.CompanyCode = Convert.ToString(row["Company Code"]);
                if (row["Cost Center"] != null)
                    ahed.CostCenter = Convert.ToString(row["Cost Center"]);
                if (row["Asset Class"] != null)
                    ahed.AssetClass = Convert.ToString(row["Asset Class"]);
                if (row["Location"] != null)
                    ahed.Location = Convert.ToString(row["Location"]);
                if (row["Currency"] != null)
                    ahed.Currency = Convert.ToString(row["Currency"]);
                //if (row["Asset Use By"] != null)
                //    ahed.AssetUseBy = Convert.ToString(row["Asset Use By"]);
                //if (row["Request By"] != null)
                //    ahed.RequestBy = Convert.ToString(row["Request By"]);
                //if (row["Requested Date"] != null)
                //    ahed.Requested_Date = Convert.ToString(row["Requested Date"]);

                if (row["Inventory note"] != null)
                    ahed.FAInventory = Convert.ToString(row["Inventory note"]);
                if (row["Serial number"] != null)
                    ahed.SerialNumber = Convert.ToString(row["Serial number"]);
                if (row["Asset"] != null)
                    ahed.AssetNo_YGS = Convert.ToString(row["Asset"]);
                if (row["Current APC"] != null)
                    ahed.Costs = Convert.ToString(row["Current APC"]);
                if (row["Asset description"] != null)
                    ahed.AssetDescription = Convert.ToString(row["Asset description"]);
                if (row["Curr.bk.val."] != null)
                    ahed.CurrentBookValue = Convert.ToString(row["Curr.bk.val."]);
                if (row["Quantity"] != null)
                    ahed.Quantity = Convert.ToString(row["Quantity"]);
                if (row["Capitalized on"] != null)
                    ahed.SAPPurchasedDate = Convert.ToString(row["Capitalized on"]); // SAP Purchased Date
                if (row["Deactivation on"] != null)
                    ahed.Deactivation_on = Convert.ToString(row["Deactivation on"]);
                //if (row["Requested Date"] != null)
                //    ahed.SAPDate = Convert.ToString(row["Requested Date"]); // SAP Purchased Date

                ahedList.Add(ahed);
            }
            return ahedList;
        }

        private static void logInfo(string msg, object obj=null) {
            if(obj==null)
                logger.Info(msg);
            else
                logger.Info(msg, obj);
        }

        private static SecureString GetSecureStringPassword(string sPassword)
        {
            SecureString passWord = new SecureString();
            foreach (char c in sPassword.ToCharArray()) passWord.AppendChar(c);
            return passWord;
        }
    }

    public class AssetsList
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string AssetRequestNo { get; set; }
        public string Company { get; set; }
        public string CostCenter { get; set; }
        public string AssetClass { get; set; }
        public string Location { get; set; }
        public string AssetUserName { get; set; }
        public string NatureOfPurchase { get; set; }
        public string AssetUseBy { get; set; }
        public string RequestedBy { get; set; }
        public string Currency { get; set; }
        public string Is3rdParty { get; set; }
        public string ReasonForPurchase { get; set; }
        public string CC { get; set; }
        public string DateRevised { get; set; }
        public string IsNewHire { get; set; }
        public string Receiver { get; set; }

        public string DivHead { get; set; }
        public string DeptHead { get; set; }
        public string ReplacementNo { get; set; }
        public string DateRequest { get; set; }
        public string ServiceStatus { get; set; }

        public string Status { get; set; }
        public string RevNo { get; set; }
        public string WorkflowStatus { get; set; }
        public string RPA_Status { get; set; }
        public string Sequence { get; set; }
        public string RunningNumber { get; set; }
        public string DivisionBudgetExceeded { get; set; }
        public string IsMDApprovalRequired { get; set; }
        public string RequestType { get; set; }
    }

    public class AssetItemsList
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string AssetRequestNo { get; set; }
        public string ReplacementNo { get; set; }
        public string FAInventory { get; set; }
        public string SerialNumber { get; set; }
        public string ItemID { get; set; }
        public string AssetNo_YGS { get; set; }
        public string ApprovedBudget { get; set; }
        public string QuotationAmount { get; set; }
        public string AssetDescription { get; set; }
        public string CurrentBookValue { get; set; }
        public string Quantity { get; set; }
        public string Costs { get; set; }
        public string Model { get; set; }
        public string ItemStatus { get; set; }
        public string SAPDate { get; set; }
        public string SAPAssetCreatedDate { get; set; }
        public string SAPAssetPurchaseDate { get; set; }
    }

    public class AssetHistoryExcelData
    {
        public string AssetRequestNo { get; set; }
        public string CompanyCode { get; set; }
        public string AssetClass { get; set; }
        public string AssetUseBy { get; set; }
        public string RequestBy { get; set; }
        public string Currency { get; set; }
        public string Location { get; set; }
        public string CostCenter { get; set; }
        public string Requested_Date { get; set; }

        public string AssetNo_YGS { get; set; }
        public string FAInventory { get; set; }
        public string SerialNumber { get; set; }
        public string AssetDescription { get; set; }
        public string SAPPurchasedDate { get; set; }
        public string SAPDate { get; set; }
        public string Quantity { get; set; }
        public string Costs { get; set; }
        public string ApprovedBudget { get; set; }
        public string QuotationAmount { get; set; }
        public string CurrentBookValue { get; set; }

        public string Dep_for_year { get; set; }
        public string Vendor { get; set; }
        public string Deactivation_on { get; set; }
        public string Useful_life { get; set; }
        public string Subnumber { get; set; }
        public string WBS_element { get; set; }
    }

}