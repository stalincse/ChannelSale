#region "namespace"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using log4net;
using log4net.Config;
#endregion

namespace ChannelSale
{
    public class ChannelSale
    {
        #region "Variables"

        private const string CHANNEL_DELIMITER = "\t";
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #endregion

        #region "Properties"

        /// <summary>
        /// Gets or sets the item Guid
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_guid")] //BF having field name          
        public String item_guid { get; set; }

        private string _SKU = string.Empty;
        /// <summary>
        /// Gets or sets the SKU
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_cd")]
        [DisplayName("SKU")]
        public string SKU
        {
            get { return _SKU; }
            set { _SKU = value; }
        }

        private string _AvailableInventory = string.Empty;
        /// <summary>
        /// Gets or sets the Available Inventory
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>   
        [Description("Available Inventory")]
        [DisplayName("Available Inventory")]
        public string AvailableInventory
        {
            get { return _AvailableInventory; }
            set { _AvailableInventory = value; }
        }

        private string _Master = string.Empty;
        /// <summary>
        /// Gets or sets the Master
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Master")]
        [DisplayName("Master")]
        public string Master
        {
            get { return _Master; }
            set { _Master = value; }
        }

        private string _ChildItem = string.Empty;
        /// <summary>
        /// Gets or sets the Child Item
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Child Item")]
        [DisplayName("Child Item")]
        public string ChildItem
        {
            get { return _ChildItem; }
            set { _ChildItem = value; }
        }

        private string _AssociatedMaster = string.Empty;
        /// <summary>
        /// Gets or sets the Associated Master
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Associated Master")]
        [DisplayName("Associated Master")]
        public string AssociatedMaster
        {
            get { return _AssociatedMaster; }
            set { _AssociatedMaster = value; }
        }

        private string _ChildType = string.Empty;
        /// <summary>
        /// Gets or sets the Child Type
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Type")]
        [DisplayName("Child Type")]
        public string Type
        {
            get { return _ChildType; }
            set { _ChildType = value; }
        }

        private string _ChildColor = string.Empty;
        /// <summary>
        /// Gets or sets the Child Color
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Color")]
        [DisplayName("Child Color")]
        public string Color
        {
            get { return _ChildColor; }
            set { _ChildColor = value; }
        }

        private string _ChildSize = string.Empty;
        /// <summary>
        /// Gets or sets the Child Size
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Size")]
        [DisplayName("Child Size")]
        public string Size
        {
            get { return _ChildSize; }
            set { _ChildSize = value; }
        }

        private string _ItemTitle = string.Empty;
        /// <summary>
        /// Gets or sets the Item Title
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_title")]
        [DisplayName("Item Title")]
        public string ItemTitle
        {
            get { return _ItemTitle; }
            set { _ItemTitle = value; }
        }

        private string _ShortDescription = string.Empty;
        /// <summary>
        /// Gets or sets the Short Description
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_description")]
        [DisplayName("Short Description")]
        public string ShortDescription
        {
            get { return _ShortDescription; }
            set { _ShortDescription = value; }
        }

        private string _Manufacturer = string.Empty;
        /// <summary>
        /// Gets or sets the Manufacturer
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Manufacturer")]
        [DisplayName("Manufacturer")]
        public string Manufacturer
        {
            get { return _Manufacturer; }
            set { _Manufacturer = value; }
        }

        private string _ManufacturerModel = string.Empty;
        /// <summary>
        /// Gets or sets the ManufacturerModel
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("vendor_item_cd")]
        [DisplayName("Manufacturer Model")]
        public string ManufacturerModel
        {
            get { return _ManufacturerModel; }
            set { _ManufacturerModel = value; }
        }

        private string _MPN = string.Empty;
        /// <summary>
        /// Gets or sets the MPN
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("MPN")]
        [DisplayName("MPN")]
        public string MPN
        {
            get { return _MPN; }
            set { _MPN = value; }
        }

        private string _MerchantCategory = string.Empty;
        /// <summary>
        /// Gets or sets the MerchantCategory
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("MerchantCategory")]
        [DisplayName("MerchantCategory")]
        public string MerchantCategory
        {
            get { return _MerchantCategory; }
            set { _MerchantCategory = value; }
        }

        private string _UPC = string.Empty;
        /// <summary>
        /// Gets or sets the UPC
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_upc")]
        [DisplayName("UPC")]
        public string UPC
        {
            get { return _UPC; }
            set { _UPC = value; }
        }

        private string _Brand = string.Empty;
        /// <summary>
        /// Gets or sets the Brand
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Brand")]
        [DisplayName("Brand")]
        public string Brand
        {
            get { return _Brand; }
            set { _Brand = value; }
        }

        private string _MSRP = string.Empty;
        /// <summary>
        /// Gets or sets the MSRP
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_retail_price")]
        [DisplayName("MSRP")]
        public string MSRP
        {
            get { return _MSRP; }
            set { _MSRP = value; }
        }

        private string _SellingPrice = string.Empty;
        /// <summary>
        /// Gets or sets the Selling Price
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_wholesale_price")]
        [DisplayName("Selling Price")]
        public string SellingPrice
        {
            get { return _SellingPrice; }
            set { _SellingPrice = value; }
        }

        private string _ProductURL = string.Empty;
        /// <summary>
        /// Gets or sets the ProductURL
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Product URL")]
        [DisplayName("Product URL")]
        public string ProductURL
        {
            get { return _ProductURL; }
            set { _ProductURL = value; }
        }

        private string _MainImageURL = string.Empty;
        /// <summary>
        /// Gets or sets the Main Image URL
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Image URL")]
        [DisplayName("Main Image URL")]
        public string MainImageURL
        {
            get { return _MainImageURL; }
            set { _MainImageURL = value; }
        }

        private string _StockQuantity = string.Empty;
        /// <summary>
        /// Gets or sets the StockQuantity
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Stock Quantity")]
        [DisplayName("Stock Quantity")]
        public string StockQuantity
        {
            get { return _StockQuantity; }
            set { _StockQuantity = value; }
        }

        private string _Condition = string.Empty;
        /// <summary>
        /// Gets or sets the Condition
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Condition")]
        [DisplayName("Condition")]
        public string Condition
        {
            get { return _Condition; }
            set { _Condition = value; }
        }

        private string _Keyword = string.Empty;
        /// <summary>
        /// Gets or sets the Keyword
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_description_keywords")]
        [DisplayName("Keyword")]
        public string Keyword
        {
            get { return _Keyword; }
            set { _Keyword = value; }
        }

        private string _IsFreeShipping = string.Empty;
        /// <summary>
        /// Gets or sets the IsFreeShipping
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("FreeShipping")]
        [DisplayName("Is Free Shipping")]
        public string FreeShipping
        {
            get { return _IsFreeShipping; }
            set { _IsFreeShipping = value; }
        }

        private string _IsTaxExempt = string.Empty;
        /// <summary>
        /// Gets or sets the IsTaxExempt
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("IsTaxExempt")]
        [DisplayName("Is Tax Exempt")]
        public string IsTaxExempt
        {
            get { return _IsTaxExempt; }
            set { _IsTaxExempt = value; }
        }

        private string _SalesItem = string.Empty;
        /// <summary>
        /// Gets or sets the SalesItem
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Sales Item")]
        [DisplayName("Sales Item")]
        public string SalesItem
        {
            get { return _SalesItem; }
            set { _SalesItem = value; }
        }

        private string _ProductWeight = string.Empty;
        /// <summary>
        /// Gets or sets the Product Weight
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Product Weight")]
        [DisplayName("Product Weight")]
        public string ProductWeight
        {
            get { return _ProductWeight; }
            set { _ProductWeight = value; }
        }

        private string _ProductLength = string.Empty;
        /// <summary>
        /// Gets or sets the Product Length
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Product Length")]
        [DisplayName("Product Length")]
        public string ProductLength
        {
            get { return _ProductLength; }
            set { _ProductLength = value; }
        }

        private string _ProductWidth = string.Empty;
        /// <summary>
        /// Gets or sets the Product Width
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Product Width")]
        [DisplayName("Product Width")]
        public string ProductWidth
        {
            get { return _ProductWidth; }
            set { _ProductWidth = value; }
        }

        private string _ProductHeight = string.Empty;
        /// <summary>
        /// Gets or sets the Product Height
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Product Height")]
        [DisplayName("Product Height")]
        public string ProductHeight
        {
            get { return _ProductHeight; }
            set { _ProductHeight = value; }
        }

        private string _ProductCost = string.Empty;
        /// <summary>
        /// Gets or sets the Product Cost
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_purchase_price")]
        [DisplayName("Product Cost")]
        public string ProductCost
        {
            get { return _ProductCost; }
            set { _ProductCost = value; }
        }

        private string _PromotionalText = string.Empty;
        /// <summary>
        /// Gets or sets the PromotionalText
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("PromotionalText")]
        [DisplayName("Promotional Text")]
        public string PromotionalText
        {
            get { return _PromotionalText; }
            set { _PromotionalText = value; }
        }

        private string _LongDescription = string.Empty;
        /// <summary>
        /// Gets or sets the LongDescription
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_description")]
        [DisplayName("Long Description")]
        public string LongDescription
        {
            get { return _LongDescription; }
            set { _LongDescription = value; }
        }

        private string _ShipWeight = string.Empty;
        /// <summary>
        /// Gets or sets the Ship Weight
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_weight")]
        [DisplayName("Ship Weight")]
        public string ShipWeight
        {
            get { return _ShipWeight; }
            set { _ShipWeight = value; }
        }

        private string _ShipLength = string.Empty;
        /// <summary>
        /// Gets or sets the Ship Length
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_dimension_length")]
        [DisplayName("Ship Length")]
        public string ShipLength
        {
            get { return _ShipLength; }
            set { _ShipLength = value; }
        }

        private string _ShipWidth = string.Empty;
        /// <summary>
        /// Gets or sets the Ship Width
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_dimension_width")]
        [DisplayName("Ship Width")]
        public string ShipWidth
        {
            get { return _ShipWidth; }
            set { _ShipWidth = value; }
        }

        private string _ShipHeight = string.Empty;
        /// <summary>
        /// Gets or sets the Ship Width
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("item_dimension_height")]
        [DisplayName("Ship Height")]
        public string ShipHeight
        {
            get { return _ShipHeight; }
            set { _ShipHeight = value; }
        }

        private string _Features = string.Empty;
        /// <summary>
        /// Gets or sets the Features
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Features")]
        [DisplayName("Features")]
        public string Features
        {
            get { return _Features; }
            set { _Features = value; }
        }

        private string _AdditionalImageURL1 = string.Empty;
        /// <summary>
        /// Gets or sets the AdditionalImageURL1
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Additional ImageURL1")]
        [DisplayName("Additional ImageURL1")]
        public string AdditionalImageURL1
        {
            get { return _AdditionalImageURL1; }
            set { _AdditionalImageURL1 = value; }
        }

        private string _AdditionalImageURL2 = string.Empty;
        /// <summary>
        /// Gets or sets the AdditionalImageURL2
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Additional ImageURL2")]
        [DisplayName("Additional ImageURL2")]
        public string AdditionalImageURL2
        {
            get { return _AdditionalImageURL2; }
            set { _AdditionalImageURL2 = value; }
        }

        private string _AdditionalImageURL3 = string.Empty;
        /// <summary>
        /// Gets or sets the AdditionalImageURL3
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Additional ImageURL3")]
        [DisplayName("Additional ImageURL3")]
        public string AdditionalImageURL3
        {
            get { return _AdditionalImageURL3; }
            set { _AdditionalImageURL3 = value; }
        }

        private string _Catalog = string.Empty;
        /// <summary>
        /// Gets or sets the Catalog
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Catalog")]
        [DisplayName("Catalog")]
        public string Catalog
        {
            get { return _Catalog; }
            set { _Catalog = value; }
        }

        private string _HandlingTime = string.Empty;
        /// <summary>
        /// Gets or sets the Handling Time
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Handling Time")]
        [DisplayName("Handling Time")]
        public string HandlingTime
        {
            get { return _HandlingTime; }
            set { _HandlingTime = value; }
        }

        private string _SearsLocation = string.Empty;
        /// <summary>
        /// Gets or sets the Sears Location
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("SearsShipFromLocation")]
        [DisplayName("Sears Location ID")]
        public string SearsLocation
        {
            get { return _SearsLocation; }
            set { _SearsLocation = value; }
        }

        private string _SearsPickUpNow = string.Empty;
        /// <summary>
        /// Gets or sets the Sears Pick Up Now
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("SearsPickUpNowEligible")]
        [DisplayName("Sears Pick Up Now")]
        public string SearsPickUpNow
        {
            get { return _SearsPickUpNow; }
            set { _SearsPickUpNow = value; }
        }

        private string _SearsExpeditedShipping = string.Empty;
        /// <summary>
        /// Gets or sets the Sears Expedited Shipping
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("SearsExpeditedShipping")]
        [DisplayName("Sears Expedited Shipping")]
        public string SearsExpeditedShipping
        {
            get { return _SearsExpeditedShipping; }
            set { _SearsExpeditedShipping = value; }
        }

        private string _SearsPremiumShipping = string.Empty;
        /// <summary>
        /// Gets or sets the Sears Premium Shipping
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("SearsPremiumShipping")]
        [DisplayName("Sears Premium Shipping")]
        public string SearsPremiumShipping
        {
            get { return _SearsPremiumShipping; }
            set { _SearsPremiumShipping = value; }
        }

        private string _SearsGroundShipping = string.Empty;
        /// <summary>
        /// Gets or sets the Sears Ground Shipping
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("SearsGroundShipping")]
        [DisplayName("Sears Ground Shipping")]
        public string SearsGroundShipping
        {
            get { return _SearsGroundShipping; }
            set { _SearsGroundShipping = value; }
        }

        private string _SearsPromotionalText = string.Empty;
        /// <summary>
        /// Gets or sets the Sears Promotional Text
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("SearsPromotionalText")]
        [DisplayName("Sears Promotional Text")]
        public string SearsPromotionalText
        {
            get { return _SearsPromotionalText; }
            set { _SearsPromotionalText = value; }
        }

        private string _MaterialType = string.Empty;
        /// <summary>
        /// Gets or sets the Material Type
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Materials")]
        [DisplayName("Material Type")]
        public string Materials
        {
            get { return _MaterialType; }
            set { _MaterialType = value; }
        }

        private string _CountryofOrigin = string.Empty;
        /// <summary>
        /// Gets or sets the Sears Country of Origin
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        [Description("Country Of Origin")]
        [DisplayName("Country of Origin")]
        public string CountryofOrigin
        {
            get { return _CountryofOrigin; }
            set { _CountryofOrigin = value; }
        }

        private Boolean _IsNeedKitItemData = false;
        /// <summary>
        /// If item dont have the record and its kit item 
        /// and it associated with only one item the true.
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        //[Description("IsNeedKitItemData")]
        //[DisplayName("IsNeedKitItemData")]
        public Boolean IsNeedKitItemData
        {
            get { return _IsNeedKitItemData; }
            set { _IsNeedKitItemData = value; }
        }

        private Guid _AssociatedMasterItemGuid;
        /// <summary>
        /// If item dont have the record and its kit item and 
        /// it associated with only one item the map that associated item guid
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        //[Description("AssociatedMasterItemGuid")]
        //[DisplayName("AssociatedMasterItemGuid")]
        public Guid AssociatedMasterItemGuid
        {
            get { return _AssociatedMasterItemGuid; }
            set { _AssociatedMasterItemGuid = value; }
        }

        private Guid? _ChildAssociatedMasterItemGuid;
        /// <summary>
        /// If item dont have the record and its kit item and 
        /// it associated with only one item (Child item) the map that associated item guid
        /// </summary>
        /// <value>String</value>
        /// <returns>String</returns>
        /// <remarks></remarks>                
        public Guid? ChildAssociatedMasterItemGuid
        {
            get { return _ChildAssociatedMasterItemGuid; }
            set { _ChildAssociatedMasterItemGuid = value; }
        }

        #endregion

        #region "Constructors"

        public ChannelSale()
        {
            XmlConfigurator.Configure();
        }

        #endregion

        #region "ToString"

        /// <summary>
        /// Gets the Tab(\t) delimited representation of the object values
        /// </summary>
        /// <returns>String</returns>
        public override string ToString()
        {
            List<string> lstChannelFields = null;
            try
            {
                lstChannelFields = new List<string>();
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SKU) ? this.SKU : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.AvailableInventory) ? this.AvailableInventory : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Master) ? this.Master : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ChildItem) ? this.ChildItem : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.AssociatedMaster) ? this.AssociatedMaster : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Type) ? this.Type : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Color) ? this.Color : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Size) ? this.Size : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ItemTitle) ? this.ItemTitle : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ShortDescription) ? this.ShortDescription : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Manufacturer) ? this.Manufacturer : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ManufacturerModel) ? this.ManufacturerModel : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.MPN) ? this.MPN : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.MerchantCategory) ? this.MerchantCategory : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.UPC) ? this.UPC : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Brand) ? this.Brand : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.MSRP) ? this.MSRP : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SellingPrice) ? this.SellingPrice : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ProductURL) ? this.ProductURL : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.MainImageURL) ? this.MainImageURL : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.StockQuantity) ? this.StockQuantity : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Condition) ? this.Condition : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Keyword) ? this.Keyword : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.FreeShipping) ? this.FreeShipping : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.IsTaxExempt) ? this.IsTaxExempt : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SalesItem) ? this.SalesItem : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ProductWeight) ? this.ProductWeight : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ProductLength) ? this.ProductLength : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ProductWidth) ? this.ProductWidth : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ProductHeight) ? this.ProductHeight : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ProductCost) ? this.ProductCost : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.PromotionalText) ? this.PromotionalText : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.LongDescription) ? this.LongDescription : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ShipWeight) ? this.ShipWeight : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ShipLength) ? this.ShipLength : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ShipWidth) ? this.ShipWidth : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.ShipHeight) ? this.ShipHeight : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Features) ? this.Features : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.AdditionalImageURL1) ? this.AdditionalImageURL1 : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.AdditionalImageURL2) ? this.AdditionalImageURL2 : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.AdditionalImageURL3) ? this.AdditionalImageURL3 : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Catalog) ? this.Catalog : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.HandlingTime) ? this.HandlingTime : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SearsLocation) ? this.SearsLocation : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SearsPickUpNow) ? this.SearsPickUpNow : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SearsExpeditedShipping) ? this.SearsExpeditedShipping : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SearsPremiumShipping) ? this.SearsPremiumShipping : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SearsGroundShipping) ? this.SearsGroundShipping : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.SearsPromotionalText) ? this.SearsPromotionalText : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.Materials) ? this.Materials : string.Empty));
                lstChannelFields.Add((!string.IsNullOrEmpty(this.CountryofOrigin) ? this.CountryofOrigin : string.Empty));
            }
            catch (Exception ex)
            {
                log.Error("\n Critical ERROR:  ChannelSale.ToString: ", ex);
            }
            if (lstChannelFields != null)
                return string.Join(CHANNEL_DELIMITER, lstChannelFields.Select(chennelField => chennelField.Replace('\n', ' ').Replace('\r', ' ').Replace('\t', ' ')).ToArray());
            else
                return string.Empty;
        }
        #endregion

        #region "GetHeaderFields"

        /// <summary>
        /// Get the all channel fields as a string
        /// </summary>
        /// <returns>string</returns>
        public List<string> GetHeaderFields()
        {
            List<string> lstChannelHeader = null;
            PropertyDescriptorCollection pdCollectionChannel = null;
            try
            {
                lstChannelHeader = new List<string>() { };
                pdCollectionChannel = TypeDescriptor.GetProperties(typeof(ChannelSale));
                foreach (PropertyDescriptor pdChannel in pdCollectionChannel)
                {
                    if (!string.IsNullOrEmpty(pdChannel.Description) && pdChannel.Description != "item_guid")
                        lstChannelHeader.Add(pdChannel.DisplayName);
                }
                return lstChannelHeader;
            }
            catch (Exception ex)
            {
                log.Error("\n Critical ERROR: ChannelSale.GetHeaderFields: ", ex);
                return new List<string>() { };
            }
        }
        #endregion
    }
}