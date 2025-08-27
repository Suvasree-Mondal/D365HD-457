// URL of generative page :: /main.aspx?pagetype=control&controlName=MscrmControls.UxAgentControl&data={"refId":"5873dee8-e3bd-4293-a4ce-3fad8254f147"}


// it saves the lead record data (6 fields record) in local storage
async function saveToLocalStorage(leadRecordData) {
    console.log("Lead Record Data :: ", leadRecordData)

    localStorage.removeItem("leadRecordData"); // clear the local storage

    const leadArray = leadRecordData.map(lead => ({
        legalEntity: lead["_msdyn_company_value@OData.Community.Display.V1.FormattedValue"] || false,
        contractingUnit: lead["_tcg_contractingunit_value@OData.Community.Display.V1.FormattedValue"] || false,
        marketSegment: lead["tcg_marketsegmentnew@OData.Community.Display.V1.FormattedValue"] || false,
        industry: lead["tcg_industrynew@OData.Community.Display.V1.FormattedValue"] || false,
        existingAccount: lead.tcg_existingaccount,
        customerGroupId: lead["account1.msdyn_customergroupid@OData.Community.Display.V1.FormattedValue"] || lead["_tcg_customergroupidnewone_value@OData.Community.Display.V1.FormattedValue"],
        accountType: lead["account1.tcg_accounttype@OData.Community.Display.V1.FormattedValue"] || lead["tcg_accounttype@OData.Community.Display.V1.FormattedValue"],
        accountName: lead["account1.name"] || lead.companyname,
        accountNumber: lead["account1.accountnumber"] || false,
        // NAaccountType: lead["tcg_accounttype@OData.Community.Display.V1.FormattedValue"] || false,
        // NAcustomerGrpId: lead["_tcg_customergroupidnewone_value@OData.Community.Display.V1.FormattedValue"] || false,
        // NAaccountName: lead.companyname || false,
    }));

    localStorage.setItem("leadRecordData", JSON.stringify(leadArray));

}

// saves only meta data like legal entity, contracting unit, industry, market segment
function saveMetaDataToLocalStorage(key, data){
    localStorage.removeItem(key);
    localStorage.setItem(key, data);
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
        "<attribute name='tcg_existingaccount' />"+
        "<attribute name='tcg_customergroupidnewone' />"+
        "<attribute name='tcg_accounttype' />"+
        "<order attribute='fullname' descending='false' />" +
        "<filter type='and'>" +
        "<condition attribute='leadid' operator='eq' value='" + leadId + "' />" +
        "</filter>" +
        "<link-entity name='account' from='accountid' to='parentaccountid' visible='false' link-type='outer'>" +
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

    await saveToLocalStorage(leadRecordData.entities);
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
    saveMetaDataToLocalStorage("legalEntityMetaData", JSON.stringify(legalEntityMetaData.entities));

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
    saveMetaDataToLocalStorage("contractingUnitMetaData", JSON.stringify(contractingUnitMetaData.entities));

    // Fetching All Market Segment from lead
    var marketSegmentMetaData = await PrimaryControl.getAttribute("tcg_marketsegmentnew").getOptions();
    saveMetaDataToLocalStorage("marketSegmentMetaData", JSON.stringify(marketSegmentMetaData));

    // Fetching All Industry from lead
    var IndustryMetaData = await PrimaryControl.getAttribute("tcg_industrynew").getOptions();
    saveMetaDataToLocalStorage("IndustryMetaData", JSON.stringify(IndustryMetaData));

    // Fetching All Account Type options from lead
    var accountTypeMetaData = await PrimaryControl.getAttribute("tcg_accounttype").getOptions();
    saveMetaDataToLocalStorage("accountTypeMetaData", JSON.stringify(accountTypeMetaData));

    // var customerGroupIdMetaData = await 
    // saveMetaDataToLocalStorage("customerGroupIdMetaData", JSON.stringify(customerGroupIdMetaData));

}


async function applyPayloadToLead(formCtx, payload) {
    if (!payload) return;

    const withBraces = id => id ? "{" + (id + "").replace(/[{}]/g, "") + "}" : null;

    // Get lookup targets from CONTROLS (not attributes)
    const leCtrl = formCtx.getControl("msdyn_company") || formCtx.getAttribute("msdyn_company")?.controls?.get?.(0);
    const cuCtrl = formCtx.getControl("tcg_contractingunit") || formCtx.getAttribute("tcg_contractingunit")?.controls?.get?.(0);
    const leTarget = leCtrl?.getEntityTypes?.()[0];                 // e.g. "msdyn_company"
    const cuTarget = cuCtrl?.getEntityTypes?.()[0];                 // e.g. "msdyn_organizationalunit"

    console.log("leCtrl:", leCtrl, "cuCtrl:", cuCtrl);
    console.log("leTarget:", leTarget, "cuTarget:", cuTarget);

    // OptionSet metadata (already cached)
    const marketSegmentMeta = JSON.parse(localStorage.getItem("marketSegmentMetaData") || "[]");
    const industryMeta = JSON.parse(localStorage.getItem("IndustryMetaData") || "[]");

    async function resolveCdmCompanyByLabel(label) {
        if (!label) return null;
        const esc = label.replace(/'/g, "''");

        // Try by company code first (e.g., "LVUS"), then by name
        const tries = [
            `?$select=cdm_companyid,cdm_name,cdm_companycode&$top=1&$filter=${encodeURIComponent(`cdm_companycode eq '${esc}'`)}`,
            `?$select=cdm_companyid,cdm_name,cdm_companycode&$top=1&$filter=${encodeURIComponent(`cdm_name eq '${esc}'`)}`
        ];

        for (const q of tries) {
            try {
                const res = await Xrm.WebApi.retrieveMultipleRecords("cdm_company", q);
                if (res.entities?.length) return res.entities[0];
            } catch (e) { /* continue */ }
        }
        return null;
    }

    async function setLegalEntity(formCtx, label) {
        if (!label) return;
        const leAttr = formCtx.getAttribute("msdyn_company");
        if (!leAttr) return;

        // Resolve the control target (should be "cdm_company")
        const leCtrl = formCtx.getControl("msdyn_company") || leAttr.controls?.get?.(0);
        const leTarget = leCtrl?.getEntityTypes?.()[0]; // expect "cdm_company"
        if (leTarget !== "cdm_company") return; // adjust if your target differs

        const rec = await resolveCdmCompanyByLabel(label);
        if (!rec) return;

        const withBraces = id => "{" + (id + "").replace(/[{}]/g, "") + "}";

        leAttr.setValue([{
            id: withBraces(rec.cdm_companyid),
            name: rec.cdm_name || rec.cdm_companycode || label, // display text
            entityType: "cdm_company"
        }]);
        leAttr.fireOnChange?.();
    }

    // Helper: resolve a record ID by name for the given target entity
    async function resolveIdByName(entityLogicalName, label, nameFields) {
        if (!label) return null;
        const escaped = label.replace(/'/g, "''");
        for (const f of nameFields) {
            const query = `?$select=${entityLogicalName}id,${f}&$top=1&$filter=${encodeURIComponent(`${f} eq '${escaped}'`)}`;
            try {
                const res = await Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, query);
                if (res.entities?.length) return res.entities[0][`${entityLogicalName}id`];
            } catch (e) { /* try next field */ }
        }
        return null;
    }

    // Resolve and set LEGAL ENTITY lookup
    await setLegalEntity(formCtx, payload.legalEntity);

    // Resolve and set CONTRACTING UNIT lookup
    if (cuTarget && payload.contractingUnit) {
        const cuId = await resolveIdByName(cuTarget, payload.contractingUnit, ["msdyn_name", "name"]);
        console.log("cuId:", cuId);
        if (cuId) {
            formCtx.getAttribute("tcg_contractingunit")?.setValue([{ id: withBraces(cuId), name: payload.contractingUnit, entityType: cuTarget }]);
            formCtx.getAttribute("tcg_contractingunit")?.fireOnChange?.();
        }
    }

    // Set OPTION SETS
    const msOpt = marketSegmentMeta.find(x => (x.text || "").trim().toLowerCase() === (payload.marketSegment || "").trim().toLowerCase());
    const indOpt = industryMeta.find(x => (x.text || "").trim().toLowerCase() === (payload.industry || "").trim().toLowerCase());
    if (msOpt?.value !== undefined) formCtx.getAttribute("tcg_marketsegmentnew")?.setValue(parseInt(msOpt.value, 10));
    if (indOpt?.value !== undefined) formCtx.getAttribute("tcg_industrynew")?.setValue(parseInt(indOpt.value, 10));
    formCtx.getAttribute("tcg_marketsegmentnew")?.fireOnChange?.();
    formCtx.getAttribute("tcg_industrynew")?.fireOnChange?.();
    // formCtx.getAttribute("tcg_accounttype")?.fireOnChange?.();

    // Set ACCOUNT lookup (parentaccountid) if provided
    try {
        const accName = (payload.account || "").trim();
        const accIdRaw = (payload.accountId || "").replace(/[{}]/g, "");
        if (accName && accIdRaw) {
            const accountAttr = formCtx.getAttribute("parentaccountid");
            if (accountAttr) {
                accountAttr.setValue([{ id: withBraces(accIdRaw), name: accName, entityType: "account" }]);
                accountAttr.fireOnChange?.();
            }
        }
    } catch (e) { }

    // Save
    await formCtx.data.save();
}



async function openCustomDialog(PrimaryControl) {

    localStorage.removeItem("leadRecordData"); // clear the local storage

    // Get the Lead ID from the URL
    var leadId = Xrm.Page.data.entity.getId(); // Returns GUID with curly braces

    // Remove curly braces if needed
    leadId = leadId.replace("{", "").replace("}", "");
    console.log("Lead ID: " + leadId);

    await getLeadRecord(leadId); // fetch current lead record data
    await fetchMetaData(PrimaryControl);  // Fetch all meta data


    const returnKey = "validateLeadReturn:" + (Xrm.Utility.createGuid ? Xrm.Utility.createGuid() : Date.now().toString());
    var webResourceName = "tcg_ValidateLeadDialogHTML"; // Use the unique name of your HTML web resource

    var pageInput = {
        pageType: "webresource",
        webresourceName: webResourceName, // Use your actual web resource name
        data: JSON.stringify({ returnKey })
    };
    var navigationOptions = {
        target: 2,
        width: { value: 22, unit: "%" },
        height: { value: 70, unit: "%" },
        position: 1
    };


    Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        async function (result) {
            // Get payload (support fallback if you used the returnKey/sessionStorage pattern)
            let payload = (result && typeof result === "object") ? result.returnValue : result;
            if (!payload && returnKey) {
                try {
                    const raw = sessionStorage.getItem(returnKey);
                    if (raw) payload = JSON.parse(raw);
                } catch (e) { }
            }
            try {
                if (returnKey) {
                    sessionStorage.removeItem(returnKey);
                }
                console.log("Dialog closed with data::", payload);

            } catch (e) { }

            if (!payload) return;


            try {

                // Apply payload to Lead (your existing function that sets fields)
                await applyPayloadToLead(PrimaryControl, payload);

                await Xrm.Navigation.openAlertDialog(
                    { title: "Success", text: "Lead validated successfully." },
                    { height: 120, width: 320, position: 1 }
                );
            } catch (e) {
                await Xrm.Navigation.openErrorDialog({
                    message: "Failed to save lead. " + (e?.message || e || "")
                });
            }

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
