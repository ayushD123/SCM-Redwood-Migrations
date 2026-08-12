# ==================== PHASE 3 MODULAR ARCHITECTURE ====================
# page_configurations.py - Centralized page metadata and field mappings

from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass

@dataclass
class PageConfig:
    """Configuration for a specific Oracle Cloud page automation"""
    page_name: str  # Display name
    page_id: str  # Internal identifier
    navigation_path: List[Dict[str, str]]  # Steps to reach the page
    vb_search_term: str  # Search term in VB Studio
    vb_page_file: str  # Target metadata file in VB
    excel_template: str  # Expected Excel filename
    composite_mapping: Dict[Tuple[str, str], str]  # Field mapping
    requires_mfa: bool = False  # Whether page requires MFA
    description: str = ""

# ==================== COMPOSITE MAPPINGS ====================

REQUISITION_LINE_MAPPING = {
    ("ItemInformation.jsff.xml", "rate"): "purchaseRequisitionsLine.ConversionRate",
    ("ItemInformation.jsff.xml", "newSupplierCheckbox"): "purchaseRequisitionsLine.NewSupplierFlag",
    ("ItemInformation.jsff.xml", "it2"): "purchaseRequisitionsLine.SuggestedSupplier",
    ("ItemInformation.jsff.xml", "lineTypeId"): "purchaseRequisitionsLine.LineTypeId",
    ("ItemInformation.jsff.xml", "itemTypeId"): "purchaseRequisitionsLine.ItemTypeLookupCode",
    ("ItemInformation.jsff.xml", "itemNumberId"): "purchaseRequisitionsLine.ItemId",
    ("ItemInformation.jsff.xml", "infotext"): "",
    ("ItemInformation.jsff.xml", "soc1"): "purchaseRequisitionsLine.ItemRevision",
    ("ItemInformation.jsff.xml", "inputText3"): "purchaseRequisitionsLine.ItemDescription",
    ("ItemInformation.jsff.xml", "categoryNameId"): "purchaseRequisitionsLine.CategoryId",
    ("ItemInformation.jsff.xml", "quantity"): "purchaseRequisitionsLine.Quantity",
    ("ItemInformation.jsff.xml", "unitOfMeasurePrimaryId"): "purchaseRequisitionsLine.UOMCode",
    ("ItemInformation.jsff.xml", "ot2"): "purchaseRequisitionsLine.SecondaryUOMCode",
    ("ItemInformation.jsff.xml", "plam1"): "purchaseRequisitionsLine.SecondaryUOMCode",
    ("ItemInformation.jsff.xml", "secondaryQuantity"): "purchaseRequisitionsLine.SecondaryQuantity",
    ("ItemInformation.jsff.xml", "currUnitPriceItemRN"): "purchaseRequisitionsLine.UnitPrice",
    ("ItemInformation.jsff.xml", "plam3"): "purchaseRequisitionsLine.UnitPrice",
    ("ItemInformation.jsff.xml", "currencyAmountItemInfo"): "purchaseRequisitionsLine.CurrencyAmount",
    ("ItemInformation.jsff.xml", "currencyCodeId"): "purchaseRequisitionsLine.CurrencyCode",
    ("ItemInformation.jsff.xml", "rateType"): "purchaseRequisitionsLine.ConversionRateType",
    ("ItemInformation.jsff.xml", "rateDate"): "purchaseRequisitionsLine.ConversionRateDate",
    ("ItemInformation.jsff.xml", "negotiationRequiredCheckbox"): "purchaseRequisitionsLine.NegotiationRequiredFlag",
    ("ItemInformation.jsff.xml", "negotiatedCheckbox"): "purchaseRequisitionsLine.NegotiatedByPreparerFlag",
    ("ItemInformation.jsff.xml", "sourceType"): "purchaseRequisitionsLine.SourceTypeCode",
    ("ItemInformation.jsff.xml", "sourceOrgNameLabel"): "purchaseRequisitionsLine.SourceOrganizationId",
    ("ItemInformation.jsff.xml", "sourceSubinventory"): "purchaseRequisitionsLine.SourceSubinventory",
    ("ItemInformation.jsff.xml", "sourceDocumentType"): "purchaseRequisitionsLine.SourceTypeCode",
    ("ItemInformation.jsff.xml", "agreementNumber"): "purchaseRequisitionsLine.SourceAgreementHeaderId",
    ("ItemInformation.jsff.xml", "sourceDocLineId"): "purchaseRequisitionsLine.SourceAgreementLineId",
    ("ItemInformation.jsff.xml", "vendorNameId"): "purchaseRequisitionsLine.SupplierId",
    ("ItemInformation.jsff.xml", "vendorNameLinkId"): "purchaseRequisitionsLine.SupplierId",
    ("ItemInformation.jsff.xml", "readOnlySupplierSiteLabel"): "purchaseRequisitionsLine.SupplierSiteId",
    ("ItemInformation.jsff.xml", "suggestedVendorSiteId"): "purchaseRequisitionsLine.SuggestedSupplierSite",
    ("ItemInformation.jsff.xml", "suggestedVendorContactId"): "purchaseRequisitionsLine.SuggestedSupplierContact",
    ("ItemInformation.jsff.xml", "inputText14"): "purchaseRequisitionsLine.SuggestedSupplierSite",
    ("ItemInformation.jsff.xml", "inputText15"): "purchaseRequisitionsLine.SuggestedSupplierContact",
    ("ItemInformation.jsff.xml", "inputText16"): "purchaseRequisitionsLine.SuggestedSupplierContactPhone",
    ("ItemInformation.jsff.xml", "inputText18"): "purchaseRequisitionsLine.SuggestedSupplierContactFax",
    ("ItemInformation.jsff.xml", "inputText17"): "purchaseRequisitionsLine.SuggestedSupplierContactEmail",
    ("ItemInformation.jsff.xml", "inputText19"): "purchaseRequisitionsLine.SupplierItemNumber",
    ("ItemInformation.jsff.xml", "it1"): "purchaseRequisitionsLine.ManufacturerName",
    ("ItemInformation.jsff.xml", "it3"): "purchaseRequisitionsLine.ManufacturerPartNumber",
    ("DeliveryBilling.jsff.xml", "columnChargeAccount"): "purchaseReqLineDistribution.ChargeAccountInput",
    ("DeliveryBilling.jsff.xml", "column5"): "purchaseReqLineDistribution.Percentage",
    ("DeliveryBilling.jsff.xml", "column6"): "purchaseReqLineDistribution.Quantity",
    ("DeliveryBilling.jsff.xml", "columnDistributionAmount"): "purchaseReqLineDistribution.CurrencyAmount",
    ("DeliveryBilling.jsff.xml", "soc3"): "purchaseRequisitionsLine.SourceTypeCode",
    ("DeliveryBilling.jsff.xml", "plam5"): "purchaseRequisitionsLine.SourceAgreementHeaderId",
    ("DeliveryBilling.jsff.xml", "plam16"): "purchaseRequisitionsLine.SourceAgreementLineId",
    ("DeliveryBilling.jsff.xml", "sbc1"): "purchaseRequisitionsLine.NegotiationRequiredFlag",
    ("DeliveryBilling.jsff.xml", "sbc2"): "purchaseRequisitionsLine.NegotiatedByPreparerFlag",
    ("DeliveryBilling.jsff.xml", "plam6"): "purchaseRequisitionsLine.ManufacturerName",
    ("DeliveryBilling.jsff.xml", "plam4"): "purchaseRequisitionsLine.ManufacturerPartNumber",
    ("DeliveryBilling.jsff.xml", "plam9"): "purchaseRequisitionsLine.SourceOrganizationId",
    ("DeliveryBilling.jsff.xml", "selectOneChoice9"): "purchaseRequisitionsLine.UrgentFlag",
    ("DeliveryBilling.jsff.xml", "suggestedBuyerId"): "purchaseRequisitionsLine.suggestedBuyerField",
    ("SmartForm.jsff.xml", "plam14"): "purchaseRequisitionsLine.SupplierId",
    ("SmartForm.jsff.xml", "Supplier"): "purchaseRequisitionsLine.SupplierId",
    ("SmartForm.jsff.xml", "plam15"): "purchaseRequisitionsLine.SupplierSiteId",
    ("SmartForm.jsff.xml", "plam6"): "purchaseRequisitionsLine.Quantity",
    ("SmartForm.jsff.xml", "plam7"): "purchaseRequisitionsLine.UOMCode",
    ("SmartForm.jsff.xml", "plam8"): "purchaseRequisitionsLine.UnitPrice",
    ("SmartForm.jsff.xml", "plam18"): "purchaseRequisitionsLine.SupplierItemNumber"

}

# Add more mappings as needed
PURCHASE_ORDER_MAPPING = {
    # Example: Add your PO mappings here
    ("POHeader.jsff.xml", "poNumber"): "purchaseOrder.OrderNumber",
    # ... add more mappings
}

SUPPLIER_MANAGEMENT_MAPPING = {
    # Example: Add your supplier mappings here
    ("SupplierDetails.jsff.xml", "supplierName"): "supplier.Name",
    # ... add more mappings
}

# ==================== PAGE CONFIGURATIONS ====================

PAGE_CONFIGS: Dict[str, PageConfig] = {
    "requisition_line": PageConfig(
        page_name="Requisition Line - Delivery Billing",
        page_id="requisition_line",
        navigation_path=[
            {"action": "click_hamburger", "selector": "a#pt1\\:_UISmmLink"},
            {"action": "expand_section", "selector": "div#pt1\\:_UISnvr\\:0\\:nvgpgl2_groupNode_procurement"},
            {"action": "click_item", "selector": "a#pt1\\:_UISnvr\\:0\\:nv_itemNode_groupnode_procurement_NewPurchaseRequisitionsWorkArePwa"}
        ],
        vb_search_term="Delivery Billing - Edit",
        vb_page_file="createEditRequisitionLine",
        excel_template="createeditrequisitionline.xlsx",
        composite_mapping=REQUISITION_LINE_MAPPING,
        requires_mfa=False,
        description="Purchase Requisition Line - Delivery Billing and Edit page customization"
    ),
    
    # Example: Purchase Order page configuration
    "purchase_order": PageConfig(
        page_name="Purchase Order - Edit",
        page_id="purchase_order",
        navigation_path=[
            {"action": "click_hamburger", "selector": "a#pt1\\:_UISmmLink"},
            {"action": "expand_section", "selector": "div#pt1\\:_UISnvr\\:0\\:nvgpgl2_groupNode_procurement"},
            {"action": "click_item", "selector": "a#pt1\\:_UISnvr\\:0\\:nv_itemNode_purchase_orders"}
        ],
        vb_search_term="Purchase Order - Edit",
        vb_page_file="editPurchaseOrder",
        excel_template="purchaseorder.xlsx",
        composite_mapping=PURCHASE_ORDER_MAPPING,
        requires_mfa=False,
        description="Purchase Order Edit page customization"
    ),
    
    # Add more page configurations as needed
}

# ==================== HELPER FUNCTIONS ====================

def get_page_config(page_id: str) -> Optional[PageConfig]:
    """Get page configuration by ID"""
    return PAGE_CONFIGS.get(page_id)

def get_composite_mapping(page_id: str) -> Optional[Dict[Tuple[str, str], str]]:
    """Get composite mapping for a specific page"""
    config = get_page_config(page_id)
    return config.composite_mapping if config else None

def list_available_pages() -> List[Dict[str, str]]:
    """List all available page configurations"""
    return [
        {
            "page_id": page_id,
            "page_name": config.page_name,
            "description": config.description,
            "requires_mfa": config.requires_mfa
        }
        for page_id, config in PAGE_CONFIGS.items()
    ]

def validate_page_exists(page_id: str) -> bool:
    """Check if a page configuration exists"""
    return page_id in PAGE_CONFIGS