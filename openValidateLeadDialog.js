// URL of generative page :: /main.aspx?pagetype=control&controlName=MscrmControls.UxAgentControl&data={"refId":"5873dee8-e3bd-4293-a4ce-3fad8254f147"}


async function saveToLocalStorage(leadRecordData) {
    console.log("Lead Record Data :: ", leadRecordData)

    localStorage.removeItem("leadRecordData"); // clear the local storage

    const leadArray = leadRecordData.map(lead => ({
        legalEntity: lead["_msdyn_company_value@OData.Community.Display.V1.FormattedValue"] || "",
        contractingUnit: lead["_tcg_contractingunit_value@OData.Community.Display.V1.FormattedValue"] || "",
        marketSegment: lead["tcg_marketsegmentnew@OData.Community.Display.V1.FormattedValue"] || "",
        industry: lead["tcg_industrynew@OData.Community.Display.V1.FormattedValue"] || "",
        customerGroupId: lead["account1.msdyn_customergroupid@OData.Community.Display.V1.FormattedValue"] || "",
        accountType: lead["account1.tcg_accounttype@OData.Community.Display.V1.FormattedValue"] || "",
        accountName: lead["account1.name"] || "",
        accountNumber: lead["account1.accountnumber"] || ""
    }));

    localStorage.setItem("leadRecordData", JSON.stringify(leadArray));

}

async function getLeadRecord(leadId) {
    // Remove curly braces if present
    leadId = leadId.replace("{", "").replace("}", "");

    // Your FetchXML with a filter for the leadId
    var fetchXml = "?fetchXml=<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
        "<entity name='lead'>" +
        "<attribute name='fullname' />" +
        "<attribute name='companyname' />" +
        "<attribute name='leadid' />" +
        "<attribute name='lastname' />" +
        "<attribute name='firstname' />" +
        "<attribute name='tcg_marketsegmentnew' />" +
        "<attribute name='tcg_industrynew' />" +
        "<attribute name='tcg_contractingunit' />" +
        "<attribute name='msdyn_company' />" +
        "<attribute name='msdyncrm_leadid' />" +
        "<order attribute='fullname' descending='false' />" +
        "<filter type='and'>" +
        "<condition attribute='leadid' operator='eq' value='" + leadId + "' />" +
        "</filter>" +
        "<link-entity name='account' from='accountid' to='tcg_xistingaccount' visible='false' link-type='outer'>" +
        "<attribute name='accountnumber' />" +
        "<attribute name='name' />" +
        "<attribute name='msdyn_customergroupid' />" +
        "<attribute name='tcg_accounttype' />" +
        "</link-entity>" +
        "</entity>" +
        "</fetch>";


    // Call the Web API to retrieve the lead record
    var leadRecordData = await Xrm.WebApi.retrieveMultipleRecords("lead", fetchXml);
    // console.log("Lead Record Data:: ", leadRecordData);

    saveToLocalStorage(leadRecordData.entities);
}

async function fetchMetaData(PrimaryControl) {

    /* Fetching All Legal Entities  */
    var fetchXmlLegalEntities = "?fetchXml=<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
        "<entity name='tcg_companyaddress'>" +
        "<attribute name='tcg_name' />" +
        "<attribute name='tcg_owningcompany' />" +
        "<attribute name='tcg_companyaddressid' />" +
        "<order attribute='tcg_name' descending='false' />" +
        "<filter type='and'>" +
        "<condition attribute='statecode' operator='eq' value='0' />" +
        "</filter>" +
        "</entity>" +
        "</fetch>";
    var legalEntityMetaData = await Xrm.WebApi.retrieveMultipleRecords("tcg_companyaddress", fetchXmlLegalEntities);
    localStorage.removeItem("legalEntityMetaData");
    localStorage.setItem("legalEntityMetaData", JSON.stringify(legalEntityMetaData.entities));

    /* Fetching All Contracting unit */
    var fetchXmlContractingUnits = "?fetchXml=<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
        "<entity name='msdyn_organizationalunit'>" +
        "<attribute name='msdyn_name' />" +
        "<attribute name='tcg_company' />" +
        "<attribute name='msdyn_organizationalunitid' />" +
        "<order attribute='msdyn_name' descending='false' />" +
        "<filter type='and'>" +
        "<condition attribute='statecode' operator='eq' value='0' />" +
        "</filter>" +
        "</entity>" +
        "</fetch>";
    var contractingUnitMetaData = await Xrm.WebApi.retrieveMultipleRecords("msdyn_organizationalunit", fetchXmlContractingUnits);
    localStorage.removeItem("contractingUnitMetaData");
    localStorage.setItem("contractingUnitMetaData", JSON.stringify(contractingUnitMetaData.entities));

    // Fetching All Market Segment from lead
    var marketSegmentMetaData = await PrimaryControl.getAttribute("tcg_marketsegmentnew").getOptions();
    localStorage.removeItem("marketSegmentMetaData");
    localStorage.setItem("marketSegmentMetaData", JSON.stringify(marketSegmentMetaData));

    // Fetching All Industry from lead
    var IndustryMetaData = await PrimaryControl.getAttribute("tcg_industrynew").getOptions();
    localStorage.removeItem("IndustryMetaData");
    localStorage.setItem("IndustryMetaData", JSON.stringify(IndustryMetaData)); 

}

function openCustomDialog(PrimaryControl) {

    localStorage.removeItem("leadRecordData"); // clear the local storage

    // Get the Lead ID from the URL
    var leadId = Xrm.Page.data.entity.getId(); // Returns GUID with curly braces

    // Remove curly braces if needed
    leadId = leadId.replace("{", "").replace("}", "");
    console.log("Lead ID: " + leadId);

    getLeadRecord(leadId); // fetch current lead record data
    fetchMetaData(PrimaryControl);  // Fetch all meta data


    var webResourceName = "tcg_ValidateLeadDialogHTML"; // Use the unique name of your HTML web resource

    var pageInput = {
        pageType: "webresource",
        webresourceName: webResourceName, // Use your actual web resource name
        data: null
    };
    var navigationOptions = {
        target: 2,
        width: { value: 22, unit: "%" },
        height: { value: 64, unit: "%" },
        position: 1
    };
    Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        function (data) {
            // Handle submitted data here
            console.log("Dialog closed with data::", data);
        },
        function (error) {
            console.error(error);
        }
    );

    // var pageInput = {
    //     pageType: "control",
    //     controlName: "MscrmControls.UxAgentControl", // Use your actual control name
    //     data: {
    //         refId: "5873dee8-e3bd-4293-a4ce-3fad8254f147" // Use your actual ID
    //     }
    // };

    // var navigationOptions = {
    //     target: 2, // 2 = Dialog, 1 = New window, 0 = Same window
    //     width: { value: 600, unit: "px" },
    //     height: { value: 400, unit: "px" },
    //     position: 1 // Center
    // };

    // Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
    //     function (result) {
    //         // Handle result if needed
    //         console.log("Generative Pages is called successfully")
    //     },
    //     function (error) {
    //         console.error("Error in opening Generative page:: ", error);
    //     }
    // );
}