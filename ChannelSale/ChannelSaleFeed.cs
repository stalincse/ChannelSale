#region "namespace"

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net;
using log4net.Config;
using MainStreet.SDK;
using MainStreet.BusinessFlow.SDK.Ws;
using MainStreet.BusinessFlow.SDK.Web;
using MainStreet.BusinessFlow.SDK;
using System.Data;
using System.Reflection;
using System.IO;
using System.Configuration;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;

#endregion

namespace ChannelSale
{
    public class ChannelSaleFeed
    {
        #region "Variables"

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private const string strBabyHavenSiteGuid = "8db36259-f93b-4276-a0c8-6574e70a5072";
        private const string strBabyHavenCatalogGuid = "812c6da0-f3b3-4399-9d37-9f1dec9010d0";
        private const string strAttributes = "Type,Color,Size,Manufacturer,FreeShipping,Product Weight,Product Length,Product Width,Product Height,Materials,Country Of Origin,Image URL,Features,Brand,SearsShipFromLocation,SearsPickUpNowEligible,SearsExpeditedShipping,SearsPremiumShipping,SearsGroundShipping,SearsPromotionalText";
        private const string strKitAssociatedItemAttributes = "Manufacturer,Product Weight,Product Length,Product Width,Product Height,Image URL,Features,Brand";
        private readonly int RECORDS_PER_PROCESS_BATCH = 50;
        private const string CHANNEL_DELIMITER = "\t";

        #endregion

        #region "Constructors"

        public ChannelSaleFeed()
        {
            RECORDS_PER_PROCESS_BATCH = Convert.ToInt32(ConfigurationManager.AppSettings["process"]);
            RECORDS_PER_PROCESS_BATCH = (RECORDS_PER_PROCESS_BATCH == 0 ? 50 : RECORDS_PER_PROCESS_BATCH);
            XmlConfigurator.Configure();
        }

        #endregion

        #region "Public Functions - ChannelDownloadFull"

        /// <summary>
        /// Get all the available items from BF and loop the ItemGetdetail Fill all the values into Datatable.
        /// (Night process)
        /// </summary>        
        public void ChannelDownloadFull()
        {
            #region "Variables"
            List<ChannelSale> lstChannel = null;
            ChannelSale objChannel = null;
            List<Guid> lstItemGuids = null;
            List<Guid> lstBatchItemGuids = null;
            ItemDetail itemDetail = null;
            int batchProcessPageNumber = 0;
            int iCount = 0;
            #endregion

            try
            {
                lstChannel = new List<ChannelSale>() { };
                lstItemGuids = GetAllItemGuids();
                if (null != lstItemGuids && 0 < lstItemGuids.Count)
                {
                    Console.WriteLine("Total items count: " + lstItemGuids.Count);
                    do
                    {
                        lstBatchItemGuids = lstItemGuids.Skip(batchProcessPageNumber * RECORDS_PER_PROCESS_BATCH).Take(RECORDS_PER_PROCESS_BATCH).ToList();
                        batchProcessPageNumber++;
                        if (null != lstBatchItemGuids && 0 < lstBatchItemGuids.Count)
                        {
                            Console.WriteLine(string.Format("*** Processing item batch {0} of {1}", batchProcessPageNumber, (Math.Round(0.00 + lstItemGuids.Count / RECORDS_PER_PROCESS_BATCH, 0)).ToString()));
                            itemDetail = GetItemsDetail(lstBatchItemGuids, false);
                            if (null != itemDetail && null != itemDetail.Items && 0 < itemDetail.Items.Count)
                            {
                                foreach (dsItemDetail.ItemsRow itemRow in itemDetail.Items)
                                {
                                    iCount++; Console.WriteLine(iCount + ". Processing item code:  " + (itemRow.Isitem_cdNull() ? "" : itemRow.item_cd));
                                    var varSiteGuid = itemDetail.Items.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemsRow>().ToList().Join(itemDetail.ItemSites.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemSitesRow>().ToList(), A => A.item_guid, B => B.item_guid, (A, B) => new { A.item_guid, B.site_guid }).SingleOrDefault(P => P.item_guid == itemRow.item_guid && P.site_guid == new Guid(strBabyHavenSiteGuid));
                                    if (null != varSiteGuid && 0 == varSiteGuid.site_guid.CompareTo(new Guid(strBabyHavenSiteGuid)))
                                    {
                                        objChannel = new ChannelSale();
                                        GenerateChannelFeed(ref objChannel, itemRow.item_guid, itemDetail);
                                        lstChannel.Add(objChannel);
                                    }
                                }
                            }
                        }
                    } while (lstBatchItemGuids.Count > 0);

                    if (null != lstChannel && 0 < lstChannel.Count)
                    {
                        //Add kit item details. If it has single association item
                        FillKitItemDetails(ref lstChannel);

                        //Add kit item(Child) details. If it has single association item
                        FillKitItemAssociatedChildDetails(ref lstChannel);

                        //Fill the attribute value intead of guid
                        FillAttributeValueFromGuid(ref lstChannel);

                        //Set the default value to channel. Set the Size of the fields.
                        NormalizeChannel(ref lstChannel);

                        //Export Channel feed file.
                        ExportExcel(lstChannel);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("\n\n ChannelDownloadFull: ", ex);
            }
        }
        #endregion

        #region "Private Functions"

        #region "GetAllItemGuids"

        /// <summary>
        /// Get item guid list from BF itemSync list.
        /// </summary>
        /// <returns></returns>
        private List<Guid> GetAllItemGuids()
        {
            ObjectGetSyncListRequest syncListReq = null;
            dsSyncList.SyncListDataTable dsSync = null;
            List<Guid> lstGuids = null;
            try
            {
                syncListReq = new ObjectGetSyncListRequest();
                lstGuids = new List<Guid>() { };
                //syncListReq.AddCriterion("item_guid", AdditionalColumnType.Database,
                //    ("15c724f8-8f7e-4941-8af5-06b2efaa0639,f0147c4d-eb42-490b-8905-028daaafa1f3,135d3a5b-d6dd-46c1-96aa-615515a6c252,8ea724ea-04a1-4b88-ab84-4e99cbb7d2f9").Split(','));
                //syncListReq.AddCriterion("item_guid", AdditionalColumnType.Database, "e033c93c-f0f5-4217-b039-969a6032d743", AdditionalCriterionCondition.Equal);

                syncListReq.MaxRows = Convert.ToInt32(ConfigurationManager.AppSettings["RowCount"]);
                syncListReq.AddCriterion("item_approved", AdditionalColumnType.Database, "1", AdditionalCriterionCondition.Equal);
                syncListReq.AddCriterion("item_available", AdditionalColumnType.Database, "1", AdditionalCriterionCondition.Equal);
                syncListReq.AddCriterion("item_track_inventory", AdditionalColumnType.Database, "1", AdditionalCriterionCondition.Equal);
                syncListReq.SyncGUID = Guid.NewGuid().ToString();
                dsSync = BusinessFlow.WebServices.Item.GetSyncList(syncListReq).SyncList;
                if (null != dsSync && 0 < dsSync.Count)
                    dsSync.AsEnumerable().AsQueryable().OfType<dsSyncList.SyncListRow>().ToList().Where(O => O.Isobject_guidNull() == false).ToList().ForEach(O => lstGuids.Add(O.object_guid));
                return lstGuids;

            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While getting the GetAllItemGuids: " + ex.Message);
                log.Error("GetAllItemGuids: ", ex);
                return null;
            }
        }
        #endregion

        #region "GetItemsDetail"

        /// <summary>
        /// BF Getitem detail call for the set of batch guids.
        /// </summary>
        /// <param name="lstBatchItemGuids"></param>
        /// <returns></returns>
        private ItemDetail GetItemsDetail(List<Guid> lstBatchItemGuids, bool isKitAssociatedItem)
        {
            ItemGetDataRequest itemGetDataReq = null;
            ItemDetail itemDetail = null;
            try
            {
                itemGetDataReq = new ItemGetDataRequest();
                lstBatchItemGuids.ForEach(O => itemGetDataReq.AddKey(O));
                itemDetail = BusinessFlow.WebServices.Item.GetDetail(itemGetDataReq);
                if (null != itemDetail && null != itemDetail.Items && 0 < itemDetail.Items.Count && !itemDetail.IsExtended)
                {
                    if (!isKitAssociatedItem)
                    {
                        foreach (Guid itemGuid in lstBatchItemGuids)
                        {
                            if (0 >= itemDetail.ItemSites.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemSitesRow>().
                                Count(O => O.item_guid == itemGuid && O.site_guid == new Guid(strBabyHavenSiteGuid)))
                            {
                                DataRow[] drItemRow = itemDetail.Items.Select("item_guid='" + itemGuid + "'", "");
                                if (null != drItemRow && 0 < drItemRow.Count())
                                    itemDetail.Items.Rows.Remove(drItemRow[0]);
                            }
                        }
                        if (null != itemDetail.Items && 0 < itemDetail.Items.Count)
                            itemDetail.ExtendAttributes(strAttributes.Split(','));
                    }
                    else
                        itemDetail.ExtendAttributes(strKitAssociatedItemAttributes.Split(','));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While getting the item detail for the list of items Guid" + ex.Message);
                log.Error("GetItemsDetail: ", ex);
            }
            return itemDetail;
        }
        #endregion


        #region "GenerateChannelFeed"

        /// <summary>
        /// Fill the all the channel fields values from BF Itemdetail. 
        /// Generate the necessary list items and call the functions to add the channel value.
        /// </summary>
        /// <param name="objChannel"></param>
        /// <param name="item_guid"></param>
        /// <param name="oItem"></param>
        private void GenerateChannelFeed(ref ChannelSale objChannel, Guid item_guid, ItemDetail oItem)
        {
            #region "Variables"
            dsItemDetail.ItemsRow itemRow = null;
            dsItemDetail.ItemSitesRow itemSitesRow = null;
            dsItemDetail.ItemVendorsRow itemVendorsRow = null;
            dsItemDetail.ItemCatalogsRow itemCatalogsRow = null;
            List<dsItemDetail.ItemImagesRow> lstItemImagesRow = null;
            List<dsItemDetail.ItemAssociationsRow> lstItemAssociationsRow = null;
            List<dsItemDetail.ItemCategoriesRow> lstItemCategoriesRow = null;
            #endregion

            try
            {
                if (null != objChannel && null != item_guid && null != oItem && 0 < oItem.Items.Count)
                {
                    itemRow = oItem.Items.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemsRow>().Where(O => O.item_guid == item_guid).FirstOrDefault();
                    itemSitesRow = oItem.ItemSites.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemSitesRow>().Where(O => O.site_guid == new Guid(strBabyHavenSiteGuid) && O.item_guid == item_guid).FirstOrDefault();
                    itemVendorsRow = oItem.ItemVendors.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemVendorsRow>().Where(O => O.item_guid == item_guid).FirstOrDefault();
                    itemCatalogsRow = oItem.ItemCatalogs.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemCatalogsRow>().Where(O => O.item_guid == item_guid && O.catalog_guid == new Guid(strBabyHavenCatalogGuid)).FirstOrDefault();
                    lstItemImagesRow = oItem.ItemImages.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemImagesRow>().Where(O => O.item_guid == item_guid).ToList();
                    lstItemAssociationsRow = oItem.ItemAssociations.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemAssociationsRow>().Where(O => O.parent_guid == item_guid).ToList();
                    lstItemCategoriesRow = oItem.ItemCategories.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemCategoriesRow>().Where(O => O.item_guid == item_guid).ToList();

                    if (null != itemRow)
                    {
                        AddItemdetailToChannel(ref objChannel, itemRow, false);
                        AddItemSitesToChannel(ref objChannel, itemSitesRow);
                        AddItemVendorToChannel(ref objChannel, itemVendorsRow);
                        AddItemCatalogToChannel(ref objChannel, itemCatalogsRow);
                        AddItemImageToChannel(ref objChannel, lstItemImagesRow);
                        objChannel.MerchantCategory = GetCategoryTree(lstItemCategoriesRow);
                        CheckKititemWithSingleAssociation(ref objChannel, itemRow, lstItemAssociationsRow);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While GenerateChannelFeed: " + ex.Message);
                log.Error("GenerateChannelFeed: ", ex);
            }
        }
        #endregion

        #region "FillKitItemDetails"

        /// <summary>
        /// Fill the data from associated item if channel dont have the value.
        /// </summary>
        /// <param name="lstChannel"></param>
        private void FillKitItemDetails(ref List<ChannelSale> lstChannel)
        {
            List<Guid> lstItemGuids = null;
            List<Guid> lstBatchItemGuids = null;
            ItemDetail oKitItem = null;
            int batchProcessPageNumber = 0;
            int iCount = 0;
            try
            {
                if (null != lstChannel && 0 < lstChannel.Count)
                {
                    lstItemGuids = lstChannel.Where(O => O.IsNeedKitItemData).Select(O => O.AssociatedMasterItemGuid).Distinct().ToList();
                    if (null != lstItemGuids && 0 < lstItemGuids.Count)
                    {
                        Console.WriteLine("---------------------------------------------------------");
                        Console.WriteLine("-------------Total kit associated items count: " + lstItemGuids.Count);
                        Console.WriteLine("---------------------------------------------------------");
                        do
                        {
                            lstBatchItemGuids = lstItemGuids.Skip(batchProcessPageNumber * RECORDS_PER_PROCESS_BATCH).Take(RECORDS_PER_PROCESS_BATCH).ToList();
                            batchProcessPageNumber++;
                            if (null != lstBatchItemGuids && 0 < lstBatchItemGuids.Count)
                            {
                                Console.WriteLine(string.Format("### Processing kit associated item batch {0} of {1}", batchProcessPageNumber, (Math.Round(0.00 + lstItemGuids.Count / RECORDS_PER_PROCESS_BATCH, 0)).ToString()));
                                oKitItem = GetItemsDetail(lstBatchItemGuids, true);
                                if (null != oKitItem && 0 < oKitItem.Items.Count)
                                {
                                    foreach (dsItemDetail.ItemsRow itemRow in oKitItem.Items)
                                    {
                                        iCount++; Console.WriteLine(iCount + ". Processing item code:  " + (itemRow.Isitem_cdNull() ? "" : itemRow.item_cd));
                                        List<ChannelSale> lstObjChannel = lstChannel.Where(O => O.AssociatedMasterItemGuid == itemRow.item_guid).ToList();
                                        if (null != lstObjChannel && 0 < lstObjChannel.Count)
                                        {
                                            for (int i = 0; i < lstObjChannel.Count; i++)
                                            {
                                                ChannelSale objChannel = lstObjChannel[i];
                                                AddItemdetailToChannel(ref objChannel, itemRow, true);
                                                if (null != oKitItem.ItemVendors && 0 < oKitItem.ItemVendors.Count && string.IsNullOrEmpty(Convert.ToString(objChannel.ManufacturerModel)))
                                                    AddItemVendorToChannel(ref objChannel, oKitItem.ItemVendors.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemVendorsRow>().Where(O => O.item_guid == itemRow.item_guid).FirstOrDefault());

                                                // Fill Item MerchantCategory value.
                                                if (string.IsNullOrEmpty(Convert.ToString(objChannel.MerchantCategory)) && null != oKitItem.ItemCategories && 0 < oKitItem.ItemCategories.Count)
                                                    objChannel.MerchantCategory = GetCategoryTree(oKitItem.ItemCategories.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemCategoriesRow>().Where(O => O.item_guid == itemRow.item_guid).ToList());

                                                // Fill kit associated child item`s Master item Guid.
                                                if (!itemRow.Ismaster_guidNull())
                                                    if (string.IsNullOrEmpty(Convert.ToString(objChannel.MerchantCategory)) || string.IsNullOrEmpty(Convert.ToString(objChannel.ItemTitle)) || string.IsNullOrEmpty(Convert.ToString(objChannel.ShortDescription)) || string.IsNullOrEmpty(Convert.ToString(objChannel.UPC)) || string.IsNullOrEmpty(Convert.ToString(objChannel.MSRP)) || string.IsNullOrEmpty(Convert.ToString(objChannel.Brand)) || string.IsNullOrEmpty(Convert.ToString(objChannel.SellingPrice)) || string.IsNullOrEmpty(Convert.ToString(objChannel.MainImageURL)) || string.IsNullOrEmpty(Convert.ToString(objChannel.Keyword)))
                                                        objChannel.ChildAssociatedMasterItemGuid = itemRow.master_guid;
                                            }
                                        }
                                    }
                                }
                            }
                        } while (lstBatchItemGuids.Count > 0);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While FillKitItemDetails: " + ex.Message);
                log.Error("\n FillKitItemDetails: ", ex);
            }
        }
        #endregion

        #region "FillKitItemAssociatedChildDetails"

        /// <summary>
        /// Fill the data from associated item(Child Item) if channel dont have the category value.
        /// </summary>
        /// <param name="lstChannel"></param>
        private void FillKitItemAssociatedChildDetails(ref List<ChannelSale> lstChannel)
        {
            List<Guid> lstItemGuids = null;
            List<Guid> lstBatchItemGuids = null;
            ItemDetail oKitItem = null;
            int batchProcessPageNumber = 0;
            int iCount = 0;
            try
            {
                if (null != lstChannel && 0 < lstChannel.Count)
                {
                    lstItemGuids = lstChannel.Where(O => O.ChildAssociatedMasterItemGuid.HasValue).Select(O => O.ChildAssociatedMasterItemGuid.Value).Distinct().ToList();
                    if (null != lstItemGuids && 0 < lstItemGuids.Count)
                    {
                        Console.WriteLine("---------------------------------------------------------");
                        Console.WriteLine("-------------Total kit associated items(Child) count: " + lstItemGuids.Count);
                        Console.WriteLine("---------------------------------------------------------");
                        do
                        {
                            lstBatchItemGuids = lstItemGuids.Skip(batchProcessPageNumber * RECORDS_PER_PROCESS_BATCH).Take(RECORDS_PER_PROCESS_BATCH).ToList();
                            batchProcessPageNumber++;
                            if (null != lstBatchItemGuids && 0 < lstBatchItemGuids.Count)
                            {
                                Console.WriteLine(String.Format("### Processing kit associated item(Child) batch {0} of {1}", batchProcessPageNumber, (Math.Round(0.00 + lstItemGuids.Count / RECORDS_PER_PROCESS_BATCH, 0)).ToString()));
                                oKitItem = GetItemsDetail(lstBatchItemGuids, true);
                                if (null != oKitItem && 0 < oKitItem.Items.Count)
                                {
                                    foreach (dsItemDetail.ItemsRow itemRow in oKitItem.Items)
                                    {
                                        iCount++; Console.WriteLine(iCount + ". Processing item code:  " + (itemRow.Isitem_cdNull() ? "" : itemRow.item_cd));
                                        List<ChannelSale> lstObjChannel = lstChannel.Where(O => O.ChildAssociatedMasterItemGuid == itemRow.item_guid).ToList();
                                        if (null != lstObjChannel && 0 < lstObjChannel.Count)
                                        {
                                            string strCatTree = "", strItemTitle = "", strShortDescription = "", strUPC = "", strMSRP = "", strBrand = "",
                                                strSellingPrice = "", strMainImageURL = "", strKeyword = "";

                                            strCatTree = GetCategoryTree(oKitItem.ItemCategories.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemCategoriesRow>().Where(O => O.item_guid == itemRow.item_guid).ToList());
                                            dsItemDetail.ItemsRow lstItemsRow = oKitItem.Items.AsEnumerable().AsQueryable().OfType<dsItemDetail.ItemsRow>().Where(O => O.item_guid == itemRow.item_guid).SingleOrDefault();
                                            if (null != lstItemsRow)
                                            {
                                                strItemTitle = lstItemsRow.Isitem_titleNull() ? "" : lstItemsRow.item_title;
                                                strShortDescription = lstItemsRow.Isitem_descriptionNull() ? "" : lstItemsRow.item_description;
                                                strUPC = lstItemsRow.Isitem_upcNull() ? "" : lstItemsRow.item_upc;
                                                strMSRP = lstItemsRow.Isitem_retail_priceNull() ? "" : lstItemsRow.item_retail_price.ToString();
                                                if (lstItemsRow.Table.Columns.Contains("Brand"))
                                                    strBrand = lstItemsRow["Brand"].ToString();
                                                if (lstItemsRow.Table.Columns.Contains("Image URL"))
                                                    strMainImageURL = lstItemsRow["Image URL"].ToString();
                                                strSellingPrice = lstItemsRow.Isitem_wholesale_priceNull() ? "" : lstItemsRow.item_wholesale_price.ToString();
                                                strKeyword = lstItemsRow.Isitem_description_keywordsNull() ? "" : lstItemsRow.item_description_keywords;
                                            }
                                            lstObjChannel.ForEach(O =>
                                                {
                                                    if (String.IsNullOrEmpty(O.MerchantCategory)) O.MerchantCategory = strCatTree;
                                                    if (String.IsNullOrEmpty(O.ItemTitle)) O.ItemTitle = strItemTitle;
                                                    if (String.IsNullOrEmpty(O.ShortDescription)) O.ShortDescription = strShortDescription;
                                                    if (String.IsNullOrEmpty(O.UPC)) O.UPC = strUPC;
                                                    if (String.IsNullOrEmpty(O.MSRP)) O.MSRP = strMSRP;
                                                    if (String.IsNullOrEmpty(O.Brand)) O.Brand = strBrand;
                                                    if (String.IsNullOrEmpty(O.SellingPrice)) O.SellingPrice = strSellingPrice;
                                                    if (String.IsNullOrEmpty(O.MainImageURL)) O.MainImageURL = strMainImageURL;
                                                    if (String.IsNullOrEmpty(O.Keyword)) O.Keyword = strKeyword;
                                                });
                                        }
                                    }
                                }
                            }
                        } while (lstBatchItemGuids.Count > 0);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While FillKitItemAssociatedChildDetails: " + ex.Message);
                log.Error("\n FillKitItemAssociatedChildDetails: ", ex);
            }
        }
        #endregion

        #region "NormalizeChannel"

        /// <summary>
        /// Set the default value to channel. Set the Size of the fields.
        /// </summary>
        /// <param name="objChannel"></param>
        private void NormalizeChannel(ref List<ChannelSale> lstChannel)
        {
            try
            {
                if (null != lstChannel && 0 < lstChannel.Count)
                {
                    lstChannel.ForEach(O =>
                    {
                        O.MPN = O.ManufacturerModel; O.Condition = "0"; O.SalesItem = "No"; O.IsTaxExempt = "No";
                        O.Catalog = ""; O.HandlingTime = ""; O.PromotionalText = "";
                        O.FreeShipping = string.IsNullOrEmpty(O.FreeShipping.Trim()) ? "No" : O.FreeShipping;
                        O.ShortDescription = (!string.IsNullOrEmpty(O.ShortDescription) && O.ShortDescription.Length > 1000) ? O.ShortDescription.Substring(0, 999) : O.ShortDescription;
                        O.ItemTitle = (!string.IsNullOrEmpty(O.ItemTitle) && O.ItemTitle.Length > 100) ? O.ItemTitle.Substring(0, 99) : O.ItemTitle;
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While NormalizeChannel: " + ex.Message);
                log.Error("NormalizeChannel: ", ex);
            }
        }
        #endregion

        #region "FillAttributeValueFromGuid"

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lstChannel"></param>
        private void FillAttributeValueFromGuid(ref List<ChannelSale> lstChannel)
        {
            Dictionary<Guid, string> dicAttributeValues = null;
            try
            {
                if (null != lstChannel && 0 < lstChannel.Count)
                {
                    Console.WriteLine("Fill Item Attribute values.");
                    dicAttributeValues = new Dictionary<Guid, string>() { };
                    LookupDataTable lkpAttributeValues = BusinessFlow.WebServices.LookupTables[LookupTables.AttributeValues];
                    if (null != lkpAttributeValues && 0 < lkpAttributeValues.Count)
                        lkpAttributeValues.AsEnumerable().ToList()
                            .ForEach(O => dicAttributeValues.Add(new Guid(O["attribute_value_guid"].ToString()), O["attribute_value_name"].ToString()));

                    if (null != dicAttributeValues && 0 < dicAttributeValues.Count)
                    {
                        string sd = dicAttributeValues.FirstOrDefault(O => O.Key == Guid.NewGuid()).Value;

                        lstChannel.ForEach(O =>
                        {
                            if (GlobalUtilities.IsGuid(O.Type)) O.Type = dicAttributeValues.FirstOrDefault(P => P.Key == new Guid(O.Type)).Value;
                            if (GlobalUtilities.IsGuid(O.Color)) O.Color = dicAttributeValues.FirstOrDefault(P => P.Key == new Guid(O.Color)).Value;
                            if (GlobalUtilities.IsGuid(O.Size)) O.Size = dicAttributeValues.FirstOrDefault(P => P.Key == new Guid(O.Size)).Value;
                            if (GlobalUtilities.IsGuid(O.Manufacturer)) O.Manufacturer = dicAttributeValues.FirstOrDefault(P => P.Key == new Guid(O.Manufacturer)).Value;
                            if (GlobalUtilities.IsGuid(O.FreeShipping)) O.FreeShipping = dicAttributeValues.FirstOrDefault(P => P.Key == new Guid(O.FreeShipping)).Value;
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("FillAttributeValueFromGuid: ", ex);
            }
        }
        #endregion

        #region "CheckKititemWithSingleAssociation"

        /// <summary>
        /// If channel item dont have the value and if it has only one associated master item then fill the master item value to channel. 
        /// </summary>
        /// <param name="objChannel"></param>
        /// <param name="itemRow"></param>
        /// <param name="lstItemAssociationsRow"></param>
        private void CheckKititemWithSingleAssociation(ref ChannelSale objChannel, dsItemDetail.ItemsRow itemRow, List<dsItemDetail.ItemAssociationsRow> lstItemAssociationsRow)
        {
            try
            {
                if (!itemRow.Isitem_type_idNull())
                {
                    objChannel.Master = (itemRow.item_type_id == 3 ? "1" : string.Empty);
                    if (!itemRow.Isitem_sub_type_idNull() && 1 == itemRow.item_sub_type_id && 1 == itemRow.item_type_id)
                    {
                        objChannel.ChildItem = "1";
                        if (string.IsNullOrEmpty(Convert.ToString(objChannel.FreeShipping.Trim())) && !itemRow.Isitem_wholesale_priceNull())
                            objChannel.FreeShipping = (itemRow.item_wholesale_price >= 100M ? "Yes" : "No");
                    }

                    if (1 == itemRow.item_type_id && !itemRow.Ismaster_guidNull() && !itemRow.Ismaster_item_cdNull())
                        objChannel.AssociatedMaster = itemRow.master_item_cd;
                }

                //Kit item
                if (!itemRow.Isitem_type_idNull() && !itemRow.Isitem_sub_type_idNull() && 1 == itemRow.item_type_id &&
                2 == itemRow.item_sub_type_id && null != lstItemAssociationsRow && 1 == lstItemAssociationsRow.Count && !lstItemAssociationsRow[0].Isitem_guidNull())
                {
                    objChannel.IsNeedKitItemData = true;
                    objChannel.AssociatedMasterItemGuid = lstItemAssociationsRow[0].item_guid;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While CheckKititemWithSingleAssociation: " + ex.Message);
                log.Error("CheckKititemWithSingleAssociation: ", ex);
            }
        }
        #endregion

        #region "AddItemImageToChannel"

        /// <summary>
        /// Add the BF item image details to channel feed.
        /// </summary>
        /// <param name="objChannel"></param>
        /// <param name="lstItemImagesRow"></param>
        private void AddItemImageToChannel(ref ChannelSale objChannel, List<dsItemDetail.ItemImagesRow> lstItemImagesRow)
        {
            try
            {
                if (null != lstItemImagesRow && 0 < lstItemImagesRow.Count)
                {
                    foreach (dsItemDetail.ItemImagesRow itemImagesRow in lstItemImagesRow)
                    {
                        if (!itemImagesRow.Isimage_guidNull() && !itemImagesRow.Isitem_image_type_idNull())
                        {
                            string strImageUrl = string.Format("http://cas07.businessflow.ms/Current/Media/item_image_sheet.aspx?domain=strollerbabies.com&item_guid={0}&image_guid={1}.jpg", itemImagesRow.item_guid, itemImagesRow.image_guid);
                            if (1 == itemImagesRow.item_image_type_id)
                                objChannel.AdditionalImageURL1 = strImageUrl;
                            else if (2 == itemImagesRow.item_image_type_id)
                                objChannel.AdditionalImageURL2 = strImageUrl;
                            else if (3 == itemImagesRow.item_image_type_id)
                                objChannel.AdditionalImageURL3 = strImageUrl;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While AddItemImageToChannel: " + ex.Message);
                log.Error("AddItemImageToChannel: ", ex);
            }
        }
        #endregion

        #region "AddItemCatalogToChannel"

        /// <summary>
        /// Add the BF item catalog details to channel feed.
        /// </summary>
        /// <param name="objChannel"></param>
        /// <param name="itemCatalogsRow"></param>
        private void AddItemCatalogToChannel(ref ChannelSale objChannel, dsItemDetail.ItemCatalogsRow itemCatalogsRow)
        {
            try
            {
                if (null != itemCatalogsRow)
                    objChannel.ProductURL = string.Format("http://www.babyhaven.com/product_details.aspx?item_guid={0}", itemCatalogsRow.item_guid);
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While AddItemCatalogToChannel: " + ex.Message);
                log.Error("AddItemCatalogToChannel: ", ex);
            }
        }
        #endregion

        #region "AddItemVendorToChannel"

        /// <summary>
        /// Add the BF item vendor details to channel feed.
        /// </summary>
        /// <param name="objChannel"></param>
        /// <param name="itemVendorsRow"></param>
        private void AddItemVendorToChannel(ref ChannelSale objChannel, dsItemDetail.ItemVendorsRow itemVendorsRow)
        {
            try
            {
                if (null != itemVendorsRow && !itemVendorsRow.Isvendor_item_cdNull())
                    objChannel.ManufacturerModel = (itemVendorsRow.vendor_item_cd.Length > 40 ? itemVendorsRow.vendor_item_cd.Substring(0, 39) : itemVendorsRow.vendor_item_cd);
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While AddItemVendorToChannel: " + ex.Message);
                log.Error("AddItemVendorToChannel: ", ex);
            }
        }
        #endregion

        #region "AddItemSitesToChannel"

        /// <summary>
        /// Get the item available and stock quantity from item site.
        /// </summary>
        /// <param name="objChannel"></param>
        /// <param name="itemSitesRow"></param>
        private void AddItemSitesToChannel(ref ChannelSale objChannel, dsItemDetail.ItemSitesRow itemSitesRow)
        {
            try
            {
                if (null != itemSitesRow && !itemSitesRow.Issite_quantity_on_handNull() && !itemSitesRow.Issite_quantity_on_holdNull())
                {
                    objChannel.AvailableInventory = Convert.ToString(itemSitesRow.site_quantity_on_hand - itemSitesRow.site_quantity_on_hold);
                    objChannel.StockQuantity = Convert.ToString(itemSitesRow.site_quantity_on_hand - itemSitesRow.site_quantity_on_hold);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While AddItemSitesToChannel: " + ex.Message);
                log.Error("AddItemSitesToChannel: ", ex);
            }
        }
        #endregion

        #region "AddItemdetailToChannel"

        /// <summary>
        /// Add the BF item details to channel feed.
        /// </summary>
        /// <param name="objChannel"></param>
        /// <param name="itemRow"></param>
        private void AddItemdetailToChannel(ref ChannelSale objChannel, dsItemDetail.ItemsRow itemRow, bool isKitAssociatedItem = false)
        {
            PropertyDescriptorCollection pdCollectionChannel = null;
            try
            {
                pdCollectionChannel = TypeDescriptor.GetProperties(objChannel);
                foreach (PropertyDescriptor pdChannel in pdCollectionChannel)
                {
                    PropertyInfo pInfo = null;
                    if (!string.IsNullOrEmpty(pdChannel.Description) && itemRow.Table.Columns.Contains(pdChannel.Description) && itemRow[pdChannel.Description] != DBNull.Value)
                    {
                        //if (!isKitAssociatedItem || !strRemoveKitAssociatedItem.Split(',').Contains(pdChannel.Name))
                        if (!isKitAssociatedItem || strKitAssociatedItemAttributes.Split(',').Contains(pdChannel.Name))
                        {
                            pInfo = objChannel.GetType().GetProperty(pdChannel.Name);
                            if (null != pInfo && string.IsNullOrEmpty(Convert.ToString(pInfo.GetValue(objChannel, null))))
                                pdChannel.SetValue(objChannel, Convert.ToString(itemRow[pdChannel.Description]));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("EXCEPTION. While AddItemdetailToChannel: " + ex.Message);
                log.Error("\n AddItemdetailToChannel: ", ex);
            }
        }
        #endregion

        #region "GetCategoryTree"

        /// <summary>
        /// Get the category tree from BF Category table.
        /// </summary>
        /// <param name="oItem"></param>
        /// <returns></returns>
        private string GetCategoryTree(List<dsItemDetail.ItemCategoriesRow> lstItemCategoriesRow)
        {
            string strCategoryTree = string.Empty;
            try
            {
                if (null != lstItemCategoriesRow && 0 < lstItemCategoriesRow.Count)
                {
                    var FilteredRow = lstItemCategoriesRow.FirstOrDefault(o => o["category_guid"].ToString() == "6ebfbe72-631e-4d05-bf3d-e878d9e4766b");
                    if (null != FilteredRow)
                        lstItemCategoriesRow.RemoveAll(o => o["category_level"].ToString() == FilteredRow["category_level"].ToString());
                    if (null != lstItemCategoriesRow && 0 < lstItemCategoriesRow.Count)
                    {
                        var category_level = lstItemCategoriesRow.GroupBy(o => o["category_level"]).OrderBy(o => o.Key).FirstOrDefault();
                        if (null != category_level)
                            lstItemCategoriesRow.Where(o => o["category_level"].ToString() == category_level.Key.ToString()).OrderByDescending(o => o["item_categories_seq_id"]).ToList()
                                .ForEach(o => strCategoryTree = (strCategoryTree == string.Empty ? o["category_name"].ToString() : strCategoryTree + " > " + o["category_name"].ToString()));
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("GetCategoryTree: ", ex);
            }
            return strCategoryTree;
        }
        #endregion


        #region "ExportExcel"
        /// <summary>
        /// Export the excel sheet from Channel table using StreamWritter
        /// </summary>
        /// <param name="dtChannelItem"></param>
        //private void ExportExcel(List<ChannelSale> lstChannel)
        //{
        //    string strFileName = string.Empty;
        //    StringBuilder strBldrChannel = null;
        //    try
        //    {
        //        strFileName = ConfigurationManager.AppSettings["FileName"];
        //        strBldrChannel = new StringBuilder();
        //        if (null != lstChannel && 0 < lstChannel.Count && !string.IsNullOrEmpty(strFileName))
        //        {
        //            Console.WriteLine("File export process started.");
        //            FileWritePreparation(strFileName);
        //            using (StreamWriter swChannelSale = new StreamWriter(strFileName, true))
        //            {
        //                strBldrChannel.Append(lstChannel[0].GetHeaderFieldsString());
        //                strBldrChannel.Append(System.Environment.NewLine);
        //                strBldrChannel.Append(string.Join(System.Environment.NewLine, lstChannel.Select(O => O.ToString()).ToArray()));

        //                swChannelSale.Write(strBldrChannel);
        //                swChannelSale.Close();
        //            }
        //            Console.WriteLine("File was created successfully.");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("ExportExcel: ", ex);
        //    }
        //}
        private void ExportExcel(List<ChannelSale> lstChannel)
        {
            string strFileName = string.Empty;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Int16 rowIndex = 2;
            Int16 HeaderIndex = 1;
            Int16 colIndex = 1;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range currentCellItemUpc = default(Excel.Range);
            Excel.Range currentCellMPN = default(Excel.Range);
            Excel.Range currentCellManufacturerModel = default(Excel.Range);
            List<string> lstFeedHeader = null;

            try
            {
                strFileName = ConfigurationManager.AppSettings["FileName"];
                if (null != lstChannel && 0 < lstChannel.Count && !string.IsNullOrEmpty(strFileName))
                {
                    Console.WriteLine("File export process started.");
                    lstFeedHeader = new List<string>() { };
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];
                    lstFeedHeader = lstChannel[0].GetHeaderFields();
                    foreach (string strHeader in lstFeedHeader)
                    {
                        xlWorkSheet.Cells[HeaderIndex, colIndex] = strHeader;
                        colIndex++;
                    }
                    foreach (ChannelSale objChannel in lstChannel)
                    {
                        xlWorkSheet.Cells[rowIndex, 1] = GetString(objChannel.SKU);
                        xlWorkSheet.Cells[rowIndex, 2] = GetString(objChannel.AvailableInventory);
                        xlWorkSheet.Cells[rowIndex, 3] = GetString(objChannel.Master);
                        xlWorkSheet.Cells[rowIndex, 4] = GetString(objChannel.ChildItem);
                        xlWorkSheet.Cells[rowIndex, 5] = GetString(objChannel.AssociatedMaster);
                        xlWorkSheet.Cells[rowIndex, 6] = GetString(objChannel.Type);
                        xlWorkSheet.Cells[rowIndex, 7] = GetString(objChannel.Color);
                        xlWorkSheet.Cells[rowIndex, 8] = GetString(objChannel.Size);
                        xlWorkSheet.Cells[rowIndex, 9] = GetString(objChannel.ItemTitle);
                        xlWorkSheet.Cells[rowIndex, 10] = GetString(objChannel.ShortDescription);
                        xlWorkSheet.Cells[rowIndex, 11] = GetString(objChannel.Manufacturer);
                        xlWorkSheet.Cells[rowIndex, 12] = GetString(objChannel.ManufacturerModel);
                        xlWorkSheet.Cells[rowIndex, 13] = GetString(objChannel.MPN);
                        xlWorkSheet.Cells[rowIndex, 14] = GetString(objChannel.MerchantCategory);
                        xlWorkSheet.Cells[rowIndex, 15] = GetString(objChannel.UPC);
                        xlWorkSheet.Cells[rowIndex, 16] = GetString(objChannel.Brand);
                        xlWorkSheet.Cells[rowIndex, 17] = GetString(objChannel.MSRP);
                        xlWorkSheet.Cells[rowIndex, 18] = GetString(objChannel.SellingPrice);
                        xlWorkSheet.Cells[rowIndex, 19] = GetString(objChannel.ProductURL);
                        xlWorkSheet.Cells[rowIndex, 20] = GetString(objChannel.MainImageURL);
                        xlWorkSheet.Cells[rowIndex, 21] = GetString(objChannel.StockQuantity);
                        xlWorkSheet.Cells[rowIndex, 22] = GetString(objChannel.Condition);
                        xlWorkSheet.Cells[rowIndex, 23] = GetString(objChannel.Keyword);
                        xlWorkSheet.Cells[rowIndex, 24] = GetString(objChannel.FreeShipping);
                        xlWorkSheet.Cells[rowIndex, 25] = GetString(objChannel.IsTaxExempt);
                        xlWorkSheet.Cells[rowIndex, 26] = GetString(objChannel.SalesItem);
                        xlWorkSheet.Cells[rowIndex, 27] = GetString(objChannel.ProductWeight);
                        xlWorkSheet.Cells[rowIndex, 28] = GetString(objChannel.ProductLength);
                        xlWorkSheet.Cells[rowIndex, 29] = GetString(objChannel.ProductWidth);
                        xlWorkSheet.Cells[rowIndex, 30] = GetString(objChannel.ProductHeight);
                        xlWorkSheet.Cells[rowIndex, 31] = GetString(objChannel.ProductCost);
                        xlWorkSheet.Cells[rowIndex, 32] = GetString(objChannel.PromotionalText);
                        xlWorkSheet.Cells[rowIndex, 33] = GetString(objChannel.LongDescription);
                        xlWorkSheet.Cells[rowIndex, 34] = GetString(objChannel.ShipWeight);
                        xlWorkSheet.Cells[rowIndex, 35] = GetString(objChannel.ShipLength);
                        xlWorkSheet.Cells[rowIndex, 36] = GetString(objChannel.ShipWidth);
                        xlWorkSheet.Cells[rowIndex, 37] = GetString(objChannel.ShipHeight);
                        xlWorkSheet.Cells[rowIndex, 38] = GetString(objChannel.Features);
                        xlWorkSheet.Cells[rowIndex, 39] = GetString(objChannel.AdditionalImageURL1);
                        xlWorkSheet.Cells[rowIndex, 40] = GetString(objChannel.AdditionalImageURL2);
                        xlWorkSheet.Cells[rowIndex, 41] = GetString(objChannel.AdditionalImageURL3);
                        xlWorkSheet.Cells[rowIndex, 42] = GetString(objChannel.Catalog);
                        xlWorkSheet.Cells[rowIndex, 43] = GetString(objChannel.HandlingTime);
                        xlWorkSheet.Cells[rowIndex, 44] = GetString(objChannel.SearsLocation);
                        xlWorkSheet.Cells[rowIndex, 45] = GetString(objChannel.SearsPickUpNow);
                        xlWorkSheet.Cells[rowIndex, 46] = GetString(objChannel.SearsExpeditedShipping);
                        xlWorkSheet.Cells[rowIndex, 47] = GetString(objChannel.SearsPremiumShipping);
                        xlWorkSheet.Cells[rowIndex, 48] = GetString(objChannel.SearsGroundShipping);
                        xlWorkSheet.Cells[rowIndex, 49] = GetString(objChannel.SearsPromotionalText);
                        xlWorkSheet.Cells[rowIndex, 50] = GetString(objChannel.Materials);
                        xlWorkSheet.Cells[rowIndex, 51] = GetString(objChannel.CountryofOrigin);

                        currentCellManufacturerModel = (Excel.Range)xlApp.ActiveCell[rowIndex, 12];
                        if (objChannel.ManufacturerModel.StartsWith("0"))
                        {
                            currentCellManufacturerModel.NumberFormat = "0#####################";
                            xlWorkSheet.Cells[rowIndex, 12] = currentCellManufacturerModel.Text;
                        }

                        currentCellMPN = (Excel.Range)xlApp.ActiveCell[rowIndex, 13];
                        if (objChannel.MPN.StartsWith("0"))
                        {
                            currentCellMPN.NumberFormat = "0#####################";
                            xlWorkSheet.Cells[rowIndex, 13] = currentCellMPN.Text;
                        }

                        currentCellItemUpc = (Excel.Range)xlApp.ActiveCell[rowIndex, 15];
                        if (objChannel.UPC.StartsWith("0"))
                        {
                            currentCellItemUpc.NumberFormat = "0#####################";
                            xlWorkSheet.Cells[rowIndex, 15] = currentCellItemUpc.Text;
                        }

                        rowIndex += 1;
                    }
                    FileWritePreparation(strFileName);
                    // Save the Excel workbok
                    xlWorkBook.SaveAs(strFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue,
                    misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    // Release the instances
                    ReleaseObject(xlApp);
                    ReleaseObject(xlWorkBook);
                    ReleaseObject(xlWorkSheet);
                    Console.WriteLine("File was created successfully.");

                    //Post the created file
                    //PostChannelFile(strFileName);
                }
            }
            catch (Exception ex)
            {
                log.Error("ExportExcel: ", ex);
            }
        }

        #endregion

        #region "PostChannelFile"

        /// <summary>
        /// Post the file to Babyhaven site using FTp.
        /// </summary>
        /// <param name="strFileName"></param>
        private void PostChannelFile(String strFileName)
        {
            #region Variables
            FtpWebRequest ftpRequest = null;
            FileInfo fileInfo = null;
            String csFTPServer = String.Empty, csFTPUser = String.Empty, csFTPPassword = String.Empty;
            #endregion
            try
            {
                fileInfo = new FileInfo(strFileName);
                csFTPServer = ConfigurationManager.AppSettings["CSFTPServer"];
                csFTPUser = ConfigurationManager.AppSettings["CSFTPUser"];
                csFTPPassword = ConfigurationManager.AppSettings["CSFTPPassword"];

                // Create FtpWebRequest object from the CJ FTP Server Uri provided
                Uri fileUri = new Uri("ftp://" + csFTPServer + "/" + fileInfo.Name);
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(fileUri);

                // Provide the WebPermission Credintials for FTP Access
                ftpRequest.Credentials = new NetworkCredential(csFTPUser, csFTPPassword);

                // By default KeepAlive is true, where the control connection is not closed
                // after a command is executed.
                ftpRequest.KeepAlive = false;

                // Specify the command to be executed.
                ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                // Specify the data transfer type.
                ftpRequest.UseBinary = false;
                ftpRequest.UsePassive = true;


                // Notify the server about the size of the uploaded file
                ftpRequest.ContentLength = fileInfo.Length;

                // The buffer size is set to 2kb since we want to transfer the data 2Kb at a time
                Int32 buffLength = 2048;
                Byte[] buffer = new Byte[buffLength];
                Int32 contentLength = 0;

                // Opens a file stream (System.IO.FileStream) to read the file to be uploaded
                FileStream fs = fileInfo.OpenRead();


                // Stream to which the file to be uploaded is written
                Stream ftpUploadStream = ftpRequest.GetRequestStream();


                // Read from the file stream 2kb at a time
                contentLength = fs.Read(buffer, 0, buffLength);

                // Till Stream content ends
                while (contentLength != 0)
                {
                    // Write Content from the file stream to the FTP Upload Stream
                    ftpUploadStream.Write(buffer, 0, contentLength);
                    contentLength = fs.Read(buffer, 0, buffLength);
                }
                // Close the file stream and the Request Stream
                ftpUploadStream.Close();
                fs.Close(); //fs.Dispose();
            }
            catch (Exception ex)
            {
                log.Error("FTP Error: PostChannelFile: ", ex);
            }
        }
        #endregion

        #region "FileWritePreparation"

        /// <summary>
        /// Check all file write location. Create the directory.
        /// </summary>
        /// <param name="strFileName"></param>
        public void FileWritePreparation(string strFileName)
        {
            string strDestinationFileName = string.Empty;
            try
            {
                Console.WriteLine("Check file write location.");
                GC.Collect();
                strDestinationFileName = Path.GetDirectoryName(strFileName) + "\\Processed\\Channelfeed-" + DateTime.Now.ToString("Mddyyyyhhmmssff") + ".xls";
                if (Directory.Exists(Path.GetDirectoryName(strFileName)))
                {
                    if (File.Exists(strFileName))
                    {
                        if (!Directory.Exists(Path.GetDirectoryName(strDestinationFileName)))
                            Directory.CreateDirectory(Path.GetDirectoryName(strDestinationFileName));
                        File.Move(strFileName, strDestinationFileName);
                        File.Delete(strFileName);
                    }
                }
                else
                    Directory.CreateDirectory(Path.GetDirectoryName(strFileName));
            }
            catch (Exception ex)
            {
                log.Error("FileWritePreparation: ", ex);
            }
        }
        #endregion

        #region"ReleaseObject"

        /// <summary>
        /// Release Object
        /// </summary>
        /// <param name="obj"></param>
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                log.Error("releaseObject", ex);
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion

        #region "GetString"

        /// <summary>
        /// Remove regex value and return the string
        /// </summary>
        /// <param name="strValue"></param>
        /// <returns></returns>
        private object GetString(string strValue)
        {
            try
            {
                return !(string.IsNullOrEmpty(Convert.ToString(strValue))) ? Convert.ToString(strValue).Replace('\n', ' ').Replace('\r', ' ').Replace('\t', ' ') : string.Empty;
            }
            catch (Exception ex)
            {
                log.Error("GetString: ", ex);
                return string.Empty;
            }
        }
        #endregion
        #endregion

    }
}