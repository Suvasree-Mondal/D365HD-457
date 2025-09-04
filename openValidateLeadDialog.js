
/* NORTH STAR IMPLEMENTATION 👇 
@licence Copyright (c) LabVantage Solution Pvt. Ltd. All rights reserved.  */


// it saves the lead record data (6 fields record) in local storage
async function saveToLocalStorage(leadRecordData) {
    console.log("Lead Record Data :: ", leadRecordData)

    localStorage.removeItem("leadRecordData"); // clear the local storage

    const leadArray = leadRecordData.map(lead => ({
        leadFullName: lead.fullname || false,
        legalEntity: lead["_msdyn_company_value@OData.Community.Display.V1.FormattedValue"] || false,
        contractingUnit: lead["_tcg_contractingunit_value@OData.Community.Display.V1.FormattedValue"] || false,
        marketSegment: lead["tcg_marketsegmentnew@OData.Community.Display.V1.FormattedValue"] || false,
        industry: lead["tcg_industrynew@OData.Community.Display.V1.FormattedValue"] || false,
        existingAccount: lead.tcg_existingaccount,
        customerGroupId: lead["account1.msdyn_customergroupid@OData.Community.Display.V1.FormattedValue"] || lead["_tcg_customergroupidnewone_value@OData.Community.Display.V1.FormattedValue"] || false,
        accountType: lead["account1.tcg_accounttype@OData.Community.Display.V1.FormattedValue"] || lead["tcg_accounttype@OData.Community.Display.V1.FormattedValue"] || false,
        accountName: lead["account1.name"] || lead.companyname || false,
        accountNumber: lead["account1.accountnumber"] || false,
        accountStatus: (lead["account1.statecode"] !== undefined && lead["account1.statecode"] !== null) ? (lead["account1.statecode"] == 0 ? "Active" : "Inactive") : false,
        // NAaccountType: lead["tcg_accounttype@OData.Community.Display.V1.FormattedValue"] || false,
        // NAcustomerGrpId: lead["_tcg_customergroupidnewone_value@OData.Community.Display.V1.FormattedValue"] || false,
        // NAaccountName: lead.companyname || false,
    }));

    localStorage.setItem("leadRecordData", JSON.stringify(leadArray));

}

// saves only meta data like legal entity, contracting unit, industry, market segment
function saveMetaDataToLocalStorage(key, data) {
    localStorage.removeItem(key);
    localStorage.setItem(key, data);
}

async function getLeadRecord(leadId) {
    // Remove curly braces if present
    // leadId = leadId.replace("{", "").replace("}", "");
    console.log("Fetching lead record data  of lead ID :: ", leadId)

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
        "<attribute name='tcg_existingaccount' />" +
        "<attribute name='tcg_customergroupidnewone' />" +
        "<attribute name='tcg_accounttype' />" +
        "<order attribute='fullname' descending='false' />" +
        "<filter type='and'>" +
        "<condition attribute='leadid' operator='eq' value='" + leadId + "' />" +
        "</filter>" +
        "<link-entity name='account' from='accountid' to='parentaccountid' visible='false' link-type='outer'>" +
        "<attribute name='accountnumber' />" +
        "<attribute name='name' />" +
        "<attribute name='msdyn_customergroupid' />" +
        "<attribute name='tcg_accounttype' />" +
        "<attribute name='statecode' />" +
        "<attribute name='tcg_accountstatusfo' />" +
        "</link-entity>" +
        "</entity>" +
        "</fetch>";


    // Call the Web API to retrieve the lead record
    var leadRecordData = await Xrm.WebApi.retrieveMultipleRecords("lead", fetchXml);
    // console.log("Lead Record Data:: ", leadRecordData);

    await saveToLocalStorage(leadRecordData.entities);
}

async function fetchMetaData(PrimaryControl, source) {

    /* Fetching All Legal Entities  */
    var fetchXmlLegalEntities = "?fetchXml=<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
        "<entity name='cdm_company'>" +
        "<attribute name='cdm_name' />" +
        "<attribute name='cdm_companycode' />" +
        "<attribute name='cdm_companyid' />" +
        "<order attribute='cdm_companycode' ascending='true' />" +
        "<filter type='and'>" +
        "<condition attribute='statecode' operator='eq' value='0' />" +
        "</filter>" +
        "</entity>" +
        "</fetch>";
    var legalEntityMetaData = await Xrm.WebApi.retrieveMultipleRecords("cdm_company", fetchXmlLegalEntities);
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

    /* 
    console.log("Just before getting Market Segment")
    // Fetching All Market Segment from lead
    var marketSegmentMetaData = await PrimaryControl.getAttribute("tcg_marketsegmentnew").getOptions();
    if(marketSegmentMetaData) saveMetaDataToLocalStorage("marketSegmentMetaData", JSON.stringify(marketSegmentMetaData));
    else{
        console.log("No Market Segment Meta Data found");
    }

    // Fetching All Industry from lead
    var IndustryMetaData = await PrimaryControl.getAttribute("tcg_industrynew").getOptions();
    if(IndustryMetaData) saveMetaDataToLocalStorage("IndustryMetaData", JSON.stringify(IndustryMetaData));
    else{
        console.log("No Industry Meta Data found");
    }

    // Fetching All Account Type options from lead
    var accountTypeMetaData = await PrimaryControl.getAttribute("tcg_accounttype").getOptions();
    if(accountTypeMetaData) saveMetaDataToLocalStorage("accountTypeMetaData", JSON.stringify(accountTypeMetaData));
    else{
        console.log("No Account Type Meta Data found");
    }
    
    */

    // Fetching All Market Segment from lead
    let marketSegmentMetaData;
    if (source === "form" && typeof PrimaryControl.getAttribute === "function") {
        marketSegmentMetaData = await PrimaryControl.getAttribute("tcg_marketsegmentnew").getOptions();
    } else {
        // Use Web API to get OptionSet metadata for grid context
        marketSegmentMetaData = await Xrm.Utility.getEntityMetadata("lead", ["tcg_marketsegmentnew"])
            .then(meta => meta.Attributes.get("tcg_marketsegmentnew").attributeDescriptor.OptionSet)
            .catch(() => []);
        // marketSegmentMetaData = await Xrm.Utility.getEntityMetadata("lead", ["tcg_marketsegmentnew"]) 
        //     .then(meta => {
        //         const attr = meta.Attributes.get("tcg_marketsegmentnew");
        //         const optSet = (attr && (attr.OptionSet || (attr.attributeDescriptor && attr.attributeDescriptor.OptionSet))) || {};
        //         const options = optSet.Options || [];
        //         return options.map(o => ({
        //             value: (o.Value !== undefined ? o.Value : o.value),
        //             text: ((o.Label && (o.Label.UserLocalizedLabel?.Label || (o.Label.LocalizedLabels && o.Label.LocalizedLabels[0]?.Label))) || o.text || o.label || "")
        //         })).filter(x => x.value !== undefined);
        //     }) 
            console.log("market Segment Meta Data JS :: ", marketSegmentMetaData)
    }
    if (marketSegmentMetaData) saveMetaDataToLocalStorage("marketSegmentMetaData", JSON.stringify(marketSegmentMetaData));
    else console.log("No Market Segment Meta Data found");

    // Fetching All Industry from lead
    let IndustryMetaData;
    if (source === "form" && typeof PrimaryControl.getAttribute === "function") {
        IndustryMetaData = await PrimaryControl.getAttribute("tcg_industrynew").getOptions();
    } else {
        IndustryMetaData = await Xrm.Utility.getEntityMetadata("lead", ["tcg_industrynew"])
            .then(meta => meta.Attributes.get("tcg_industrynew").attributeDescriptor.OptionSet)
            .catch(() => []);
        // IndustryMetaData = await Xrm.Utility.getEntityMetadata("lead", ["tcg_industrynew"]) 
        //     .then(meta => {
        //         const attr = meta.Attributes.get("tcg_industrynew");
        //         const optSet = (attr && (attr.OptionSet || (attr.attributeDescriptor && attr.attributeDescriptor.OptionSet))) || {};
        //         const options = optSet.Options || [];
        //         return options.map(o => ({
        //             value: (o.Value !== undefined ? o.Value : o.value),
        //             text: ((o.Label && (o.Label.UserLocalizedLabel?.Label || (o.Label.LocalizedLabels && o.Label.LocalizedLabels[0]?.Label))) || o.text || o.label || "")
        //         })).filter(x => x.value !== undefined);
        //     });
            console.log("Industry Meta Data JS :: ", IndustryMetaData)
    }
    if (IndustryMetaData) saveMetaDataToLocalStorage("IndustryMetaData", JSON.stringify(IndustryMetaData));
    else console.log("No Industry Meta Data found");

    // Fetching All Account Type options from lead
    let accountTypeMetaData;
    if (source === "form" && typeof PrimaryControl.getAttribute === "function") {
        accountTypeMetaData = await PrimaryControl.getAttribute("tcg_accounttype").getOptions();
    } else {
        accountTypeMetaData = await Xrm.Utility.getEntityMetadata("lead", ["tcg_accounttype"])
            .then(meta => meta.Attributes.get("tcg_accounttype").attributeDescriptor.OptionSet)
            .catch(() => []);
    }
    if (accountTypeMetaData) saveMetaDataToLocalStorage("accountTypeMetaData", JSON.stringify(accountTypeMetaData));
    else console.log("No Account Type Meta Data found");

    // if customer is new (existing account is no or false) Fetching All customer group Id lookup 
    var fetchXmlCustomerGrpId = "?fetchXml=<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
        "<entity name='msdyn_customergroup'>" +
        "<attribute name='msdyn_groupid' />" +
        "<attribute name='msdyn_company' />" +
        "<attribute name='msdyn_customergroupid' />" +
        "<attribute name='msdyn_description' />" +
        "<order attribute='msdyn_company' descending='false' />" +
        "<order attribute='msdyn_groupid' descending='false' />" +
        "<filter type='and'>" +
        "<condition attribute='statecode' operator='eq' value='0' />" +
        "<condition attribute='msdyn_company' operator='not-null' />" +
        "</filter>" +
        "</entity>" +
        "</fetch>";

    // for New account only
    const leadDataArr = JSON.parse(localStorage.getItem("leadRecordData"));
    if (leadDataArr && leadDataArr[0] && !leadDataArr[0].existingAccount) {
        var customerGroupIdMetaData = await Xrm.WebApi.retrieveMultipleRecords("msdyn_customergroup", fetchXmlCustomerGrpId);
        console.log("Customer Group id Meta Data:: ", customerGroupIdMetaData)
        saveMetaDataToLocalStorage("customerGroupIdMetaData", JSON.stringify(customerGroupIdMetaData));
    }

}

// Apply dialog payload when invoked from Form ribbon: upadate the record via form context API
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

    // Account Type (optionset) from payload
    try {
        const accountTypeMeta = JSON.parse(localStorage.getItem("accountTypeMetaData") || "[]");
        const atOpt = accountTypeMeta.find(x => (x.text || "").trim().toLowerCase() === (payload.accountType || "").trim().toLowerCase());
        if (atOpt?.value !== undefined) {
            formCtx.getAttribute("tcg_accounttype")?.setValue(parseInt(atOpt.value, 10));
            formCtx.getAttribute("tcg_accounttype")?.fireOnChange?.();
        }
    } catch (e) { }

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



    // Customer Group Id (lookup/label in dialog). Attempt to resolve to the lookup on lead if present
    try {
        const cgLabel = (payload.customerGroupId || "").trim();
        const cgAttrLogical = "tcg_customergroupidnewone"; // adjust if your schema differs
        const cgAttr = formCtx.getAttribute(cgAttrLogical);
        if (cgAttr && cgLabel) {
            // Resolve company from payload.legalEntity to disambiguate group by company
            let companyIdRaw = null;
            try {
                const comp = await resolveCdmCompanyByLabel(payload.legalEntity);
                companyIdRaw = comp && comp.cdm_companyid ? (comp.cdm_companyid + '').replace(/[{}]/g, '') : null;
            } catch (e) { companyIdRaw = null; }

            const esc = cgLabel.replace(/'/g, "''");
            const filters = [`msdyn_groupid eq '${esc}'`];
            if (companyIdRaw) filters.push(`_msdyn_company_value eq ${companyIdRaw}`);
            const filter = filters.join(' and ');
            const query = `?$select=msdyn_customergroupid,msdyn_groupid, _msdyn_company_value&$top=1&$filter=${encodeURIComponent(filter)}`;

            try {
                const res = await Xrm.WebApi.retrieveMultipleRecords("msdyn_customergroup", query);
                if (res.entities && res.entities.length) {
                    const id = res.entities[0].msdyn_customergroupid;
                    const withBracesLocal = v => "{" + (v + "").replace(/[{}]/g, "") + "}";
                    cgAttr.setValue([{ id: withBracesLocal(id), name: cgLabel, entityType: "msdyn_customergroup" }]);
                    cgAttr.fireOnChange?.();
                }
            } catch (e) { /* ignore */ }
        }
    } catch (e) { }


    // Save
    await formCtx.data.save();
}


// Apply dialog payload when invoked from Grid ribbon: update the record via Web API
async function applyPayloadToLeadGrid(leadId, payload) {
    if (!payload || !leadId) return;

    const stripBraces = v => (v + "").replace(/[{}]/g, "");
    const toGuid = v => stripBraces(v || "");

    // Resolve helpers mirroring form path logic
    async function resolveCdmCompanyByLabel(label) {
        if (!label) return null;
        const esc = label.replace(/'/g, "''");
        const tries = [
            `?$select=cdm_companyid,cdm_name,cdm_companycode&$top=1&$filter=${encodeURIComponent(`cdm_companycode eq '${esc}'`)}`,
            `?$select=cdm_companyid,cdm_name,cdm_companycode&$top=1&$filter=${encodeURIComponent(`cdm_name eq '${esc}'`)}`
        ];
        for (const q of tries) {
            try {
                const res = await Xrm.WebApi.retrieveMultipleRecords("cdm_company", q);
                if (res.entities?.length) return res.entities[0];
            } catch (e) { }
        }
        return null;
    }

    async function resolveIdByName(entityLogicalName, label, nameFields) {
        if (!label) return null;
        const escaped = label.replace(/'/g, "''");
        for (const f of nameFields) {
            const query = `?$select=${entityLogicalName}id,${f}&$top=1&$filter=${encodeURIComponent(`${f} eq '${escaped}'`)}`;
            try {
                const res = await Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, query);
                console.log("CU res :: ", res);
                if (res.entities?.length) return res.entities[0][`${entityLogicalName}id`];
            } catch (e) { }
        }
        return null;
    }

    // Resolve an optionset value by label using metadata
    async function resolveOptionValue(entityLogicalName, attributeLogicalName, label) {
        if (!label) return null;
        try {
            const meta = await Xrm.Utility.getEntityMetadata(entityLogicalName, [attributeLogicalName]);
            const attr = meta && meta.Attributes && meta.Attributes.get && meta.Attributes.get(attributeLogicalName);
            const optionSet = attr && (attr.OptionSet || (attr.attributeDescriptor && attr.attributeDescriptor.OptionSet));
            let options = optionSet && (optionSet.Options || optionSet);
            if (!options || !Array.isArray(options)) return null;
            const needle = (label || "").trim().toLowerCase();
            for (const opt of options) {
                const text = (opt.Label && (opt.Label.UserLocalizedLabel?.Label || (opt.Label.LocalizedLabels && opt.Label.LocalizedLabels[0]?.Label)))
                    || opt.text || opt.label || "";
                const value = (opt.Value !== undefined) ? opt.Value : opt.value;
                if (text && value !== undefined && (text + "").trim().toLowerCase() === needle) return parseInt(value, 10);
            }
        } catch (e) { }
        return null;
    }

    // Build update payload
    const update = {};

    // Retrieve lead metadata to validate lookup presence in this org/env
    let leadMeta = null;
    try {
        leadMeta = await Xrm.Utility.getEntityMetadata("lead", ["tcg_contractingunit", "msdyn_company", "tcg_customergroupidnewone"]);
        console.log("leadMeta :: ", leadMeta)
    } catch (e) { /* proceed defensively */ }
    const hasContractingUnitAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_contractingunit"));
    const hasCompanyAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("msdyn_company"));
    const hasCustomerGroupAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_customergroupidnewone"));

    console.log("hasContractingUnitAttr :: ", hasContractingUnitAttr)
    console.log("hasCompanyAttr :: ", hasCompanyAttr)
    console.log("hasCustomerGroupAttr :: ", hasCustomerGroupAttr)

    // OptionSets: prefer cached metadata, then fall back to live metadata
    try {
        const marketSegmentMeta = JSON.parse(localStorage.getItem("marketSegmentMetaData") || "[]");
        const industryMeta = JSON.parse(localStorage.getItem("IndustryMetaData") || "[]");
        const accountTypeMeta = JSON.parse(localStorage.getItem("accountTypeMetaData") || "[]");

        const msOpt = marketSegmentMeta.find(x => (x.Label || "").trim().toLowerCase() === (payload.marketSegment || "").trim().toLowerCase());
        if (msOpt?.Value !== undefined) update["tcg_marketsegmentnew"] = parseInt(msOpt.Value, 10);
        console.log("msOpt :: ", msOpt);
        console.log("update till msOpt:: ", update);

        const indOpt = industryMeta.find(x => (x.Label || "").trim().toLowerCase() === (payload.industry || "").trim().toLowerCase());
        if (indOpt?.Value !== undefined) update["tcg_industrynew"] = parseInt(indOpt.Value, 10);
        console.log("indOpt :: ", indOpt);
        console.log("update till indOpt:: ", update);

        const atOpt = accountTypeMeta.find(x => (x.Label || "").trim().toLowerCase() === (payload.accountType || "").trim().toLowerCase());
        if (atOpt?.Value !== undefined) update["tcg_accounttype"] = parseInt(atOpt.Value, 10);
        console.log("atOpt :: ", atOpt);
        console.log("update till atOpt:: ", update);

    } catch (e) { }

    // Fill any missing options via live metadata resolution
    if (update["tcg_marketsegmentnew"] === undefined) {
        const ms = await resolveOptionValue("lead", "tcg_marketsegmentnew", payload.marketSegment);
        if (ms !== null) update["tcg_marketsegmentnew"] = ms;
        console.log("ms :: ", ms)
    }
    if (update["tcg_industrynew"] === undefined) {
        const ind = await resolveOptionValue("lead", "tcg_industrynew", payload.industry);
        if (ind !== null) update["tcg_industrynew"] = ind;
        console.log("ind :: ", ind) 
    }
    if (update["tcg_accounttype"] === undefined) {
        const at = await resolveOptionValue("lead", "tcg_accounttype", payload.accountType);
        if (at !== null) update["tcg_accounttype"] = at;
        console.log("at :: ", at)
    }
    // Lookups
    // Legal Entity (msdyn_company -> cdm_company)
    if (payload.legalEntity && hasCompanyAttr) {
        const comp = await resolveCdmCompanyByLabel(payload.legalEntity);
        if (comp?.cdm_companyid) update["msdyn_company@odata.bind"] = `/cdm_companies(${toGuid(comp.cdm_companyid)})`;
    }
    console.log("Update Arr till legal entity :: ", update);

    // Contracting Unit (tcg_contractingunit -> msdyn_organizationalunit), only if attribute is a lookup
    try {
        const cuAttr = leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_contractingunit");
        console.log("cuAttr :: ", cuAttr);
        const isLookup = !!(cuAttr && (cuAttr._attributeTypeName.toLowerCase?.() === "lookup" || cuAttr.AttributeType === 6 || cuAttr.Targets || cuAttr.attributeDescriptor?.Targets));
        console.log("isLookup for CU :: ", isLookup);
        if (payload.contractingUnit && hasContractingUnitAttr && isLookup) {
            const cuId = await resolveIdByName("msdyn_organizationalunit", payload.contractingUnit, ["msdyn_name", "name"]);
            console.log("cuID :: ", cuId);
            if (cuId) {
                update["tcg_contractingunit@odata.bind"] = `/msdyn_organizationalunits(${payload.contractingUnit})`;
                console.log("update['tcg_contractingunit@odata.bind'] :: ", update["tcg_contractingunit@odata.bind"]);
                // update in database here for CU only
                try {
                    var cuObject = {}
                    cuObject["tcg_contractingunit@odata.bind"] = `/msdyn_organizationalunits(${payload.contractingUnit})`;
                    console.log("cuObject in try block :: ", cuObject);
                    await Xrm.WebApi.updateRecord("lead", toGuid(leadId), cuObject);
                    console.log("Successfully updated CU");
                } catch (e) {
                    console.error("Error updating CU: ", e);
                }
            }
            else console.log("Didn't updated the CU");
        }
        else console.log("CU is not valid");
    } catch (e) { }

    // Parent Account
    if ((payload.account || "").trim() && (payload.accountId || "").trim()) {
        update["parentaccountid@odata.bind"] = `/accounts(${toGuid(payload.accountId)})`;
        console.log("update in parent account :: ", update);  
    }

    // Customer Group (tcg_customergroupidnewone -> msdyn_customergroup)
    if ((payload.customerGroupId || "").trim() && hasCustomerGroupAttr) {
        let companyIdRaw = null;
        try {
            const comp = await resolveCdmCompanyByLabel(payload.legalEntity);
            companyIdRaw = comp && comp.cdm_companyid ? toGuid(comp.cdm_companyid) : null;
        } catch (e) { companyIdRaw = null; }

        const esc = payload.customerGroupId.trim().replace(/'/g, "''");
        const filters = [`msdyn_groupid eq '${esc}'`];
        if (companyIdRaw) filters.push(`_msdyn_company_value eq ${companyIdRaw}`);
        const filter = filters.join(' and ');
        const query = `?$select=msdyn_customergroupid&$top=1&$filter=${encodeURIComponent(filter)}`;
        try {
            const res = await Xrm.WebApi.retrieveMultipleRecords("msdyn_customergroup", query);
            if (res.entities && res.entities.length) {
                update["tcg_customergroupidnewone@odata.bind"] = `/msdyn_customergroups(${toGuid(res.entities[0].msdyn_customergroupid)})`;
                console.log("update in CGid :: ", update);
            }
        } catch (e) { }
    }

    // No changes resolved? Exit quietly
    if (Object.keys(update).length === 0) return;

    // Execute update with fallback for unknown nav properties
    try {
        console.log("update array :: ", update);
        await Xrm.WebApi.updateRecord("lead", toGuid(leadId), update);
    } catch (e) {
        const msg = (e && (e.message || e._message)) || "";
        // If server complains about tcg_contractingunit annotation, drop it and retry
        if (msg.toLowerCase().includes("tcg_contractingunit")) {
            try {
                console.log("Before delete :: ", update);
                delete update["tcg_contractingunit@odata.bind"];
                console.log("After delete :: ", update);
                await Xrm.WebApi.updateRecord("lead", toGuid(leadId), update);
            } catch (e2) {
                throw e2;
            }
        } else if (msg.toLowerCase().includes("tcg_customergroupidnewone")) {
            try {
                delete update["tcg_customergroupidnewone@odata.bind"];
                await Xrm.WebApi.updateRecord("lead", toGuid(leadId), update);
            } catch (e3) {
                throw e3;
            }
        } else {
            throw e;
        }
    }
}



// Apply dialog payload when invoked from Grid ribbon: update the record via Web API - latest
// async function applyPayloadToLeadGrid(leadId, payload) {
//     if (!payload || !leadId) return;

//     const stripBraces = v => (v + "").replace(/[{}]/g, "");
//     const toGuid = v => stripBraces(v || "");

//     // Resolve helpers mirroring form path logic
//     async function resolveCdmCompanyByLabel(label) {
//         if (!label) return null;
//         const esc = label.replace(/'/g, "''");
//         const tries = [
//             `?$select=cdm_companyid,cdm_name,cdm_companycode&$top=1&$filter=${encodeURIComponent(`cdm_companycode eq '${esc}'`)}`,
//             `?$select=cdm_companyid,cdm_name,cdm_companycode&$top=1&$filter=${encodeURIComponent(`cdm_name eq '${esc}'`)}`
//         ];
//         for (const q of tries) {
//             try {
//                 const res = await Xrm.WebApi.retrieveMultipleRecords("cdm_company", q);
//                 if (res.entities?.length) return res.entities[0];
//             } catch (e) { }
//         }
//         return null;
//     }

//     async function resolveIdByName(entityLogicalName, label, nameFields) {
//         if (!label) return null;
//         const escaped = label.replace(/'/g, "''");
//         for (const f of nameFields) {
//             const query = `?$select=${entityLogicalName}id,${f}&$top=1&$filter=${encodeURIComponent(`${f} eq '${escaped}'`)}`;
//             try {
//                 const res = await Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, query);
//                 if (res.entities?.length) return res.entities[0][`${entityLogicalName}id`];
//             } catch (e) { }
//         }
//         return null;
//     }

//     // Resolve an optionset value by label using metadata
//     async function resolveOptionValue(entityLogicalName, attributeLogicalName, label) {
//         if (!label) return null;
//         try {
//             const meta = await Xrm.Utility.getEntityMetadata(entityLogicalName, [attributeLogicalName]);
//             const attr = meta && meta.Attributes && meta.Attributes.get && meta.Attributes.get(attributeLogicalName);
//             const optionSet = attr && (attr.OptionSet || (attr.attributeDescriptor && attr.attributeDescriptor.OptionSet));
//             let options = optionSet && (optionSet.Options || optionSet);
//             if (!options || !Array.isArray(options)) return null;
//             const needle = (label || "").trim().toLowerCase();
//             for (const opt of options) {
//                 const text = (opt.Label && (opt.Label.UserLocalizedLabel?.Label || (opt.Label.LocalizedLabels && opt.Label.LocalizedLabels[0]?.Label)))
//                     || opt.text || opt.label || "";
//                 const value = (opt.Value !== undefined) ? opt.Value : opt.value;
//                 if (text && value !== undefined && (text + "").trim().toLowerCase() === needle) return parseInt(value, 10);
//             }
//         } catch (e) { }
//         return null;
//     }

//     // Build update payload
//     const update = {};

//     // Retrieve lead metadata to validate lookup presence in this org/env
//     let leadMeta = null;
//     try {
//         leadMeta = await Xrm.Utility.getEntityMetadata("lead", ["tcg_contractingunit", "msdyn_company", "tcg_customergroupidnewone"]);
//     } catch (e) { /* proceed defensively */ }
//     const hasContractingUnitAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_contractingunit"));
//     const hasCompanyAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("msdyn_company"));
//     const hasCustomerGroupAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_customergroupidnewone"));

//     // OptionSets: prefer cached metadata, then fall back to live metadata
//     try {
//         const marketSegmentMeta = JSON.parse(localStorage.getItem("marketSegmentMetaData") || "[]");
//         const industryMeta = JSON.parse(localStorage.getItem("IndustryMetaData") || "[]");
//         const accountTypeMeta = JSON.parse(localStorage.getItem("accountTypeMetaData") || "[]");

//         const msOpt = marketSegmentMeta.find(x => (x.Label || "").trim().toLowerCase() === (payload.marketSegment || "").trim().toLowerCase());
//         if (msOpt?.value !== undefined) update["tcg_marketsegmentnew"] = parseInt(msOpt.value, 10);

//         const indOpt = industryMeta.find(x => (x.Label || "").trim().toLowerCase() === (payload.industry || "").trim().toLowerCase());
//         if (indOpt?.value !== undefined) update["tcg_industrynew"] = parseInt(indOpt.value, 10);

//         const atOpt = accountTypeMeta.find(x => (x.Label || "").trim().toLowerCase() === (payload.accountType || "").trim().toLowerCase());
//         if (atOpt?.value !== undefined) update["tcg_accounttype"] = parseInt(atOpt.value, 10);
//     } catch (e) { }

//     // Fill any missing options via live metadata resolution
//     if (update["tcg_marketsegmentnew"] === undefined) {
//         const ms = await resolveOptionValue("lead", "tcg_marketsegmentnew", payload.marketSegment);
//         if (ms !== null) update["tcg_marketsegmentnew"] = ms;
//     }
//     if (update["tcg_industrynew"] === undefined) {
//         const ind = await resolveOptionValue("lead", "tcg_industrynew", payload.industry);
//         if (ind !== null) update["tcg_industrynew"] = ind;
//     }
//     if (update["tcg_accounttype"] === undefined) {
//         const at = await resolveOptionValue("lead", "tcg_accounttype", payload.accountType);
//         if (at !== null) update["tcg_accounttype"] = at;
//     }

//     // Lookups
//     // Legal Entity (msdyn_company -> cdm_company)
//     if (payload.legalEntity && hasCompanyAttr) {
//         const comp = await resolveCdmCompanyByLabel(payload.legalEntity);
//         if (comp?.cdm_companyid) update["msdyn_company@odata.bind"] = `/cdm_companies(${toGuid(comp.cdm_companyid)})`;
//     }

//     // Contracting Unit (tcg_contractingunit -> msdyn_organizationalunit), only if attribute is a lookup
//     try {
//         const cuAttr = leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_contractingunit");
//         const isLookup = !!(cuAttr && (cuAttr.AttributeTypeName?.Value?.toLowerCase?.() === "lookup" || cuAttr.AttributeType === 6 || cuAttr.Targets || cuAttr.attributeDescriptor?.Targets));
//         if (payload.contractingUnit && hasContractingUnitAttr && isLookup) {
//             const cuId = await resolveIdByName("msdyn_organizationalunit", payload.contractingUnit, ["msdyn_name", "name"]);
//             if (cuId) {
//                 const cuNav = (cuAttr && (cuAttr.SchemaName || cuAttr.attributeDescriptor?.SchemaName || cuAttr.LogicalName || cuAttr._logicalName)) || "tcg_contractingunit";
//                 update[`${cuNav}@odata.bind`] = `/msdyn_organizationalunits(${toGuid(cuId)})`;
//             }
//         }
//     } catch (e) { }

//     // Parent Account
//     if ((payload.account || "").trim() && (payload.accountId || "").trim()) {
//         update["parentaccountid@odata.bind"] = `/accounts(${toGuid(payload.accountId)})`;
//     }

//     // Customer Group (tcg_customergroupidnewone -> msdyn_customergroup)
//     if ((payload.customerGroupId || "").trim() && hasCustomerGroupAttr) {
//         let companyIdRaw = null;
//         try {
//             const comp = await resolveCdmCompanyByLabel(payload.legalEntity);
//             companyIdRaw = comp && comp.cdm_companyid ? toGuid(comp.cdm_companyid) : null;
//         } catch (e) { companyIdRaw = null; }

//         const esc = payload.customerGroupId.trim().replace(/'/g, "''");
//         const filters = [`msdyn_groupid eq '${esc}'`];
//         if (companyIdRaw) filters.push(`_msdyn_company_value eq ${companyIdRaw}`);
//         const filter = filters.join(' and ');
//         const query = `?$select=msdyn_customergroupid&$top=1&$filter=${encodeURIComponent(filter)}`;
//         try {
//             const res = await Xrm.WebApi.retrieveMultipleRecords("msdyn_customergroup", query);
//             if (res.entities && res.entities.length) {
//                 update["tcg_customergroupidnewone@odata.bind"] = `/msdyn_customergroups(${toGuid(res.entities[0].msdyn_customergroupid)})`;
//             }
//         } catch (e) { }
//     }

//     // No changes resolved? Exit quietly
//     if (Object.keys(update).length === 0) return;

//     // Execute update with fallback for unknown nav properties
//     try {
//         await Xrm.WebApi.updateRecord("lead", toGuid(leadId), update);
//     } catch (e) {
//         const msg = (e && (e.message || e._message)) || "";
//         // If server complains about tcg_contractingunit annotation, drop it and retry
//         if (msg.toLowerCase().includes("tcg_contractingunit")) {
//             try {
//                 delete update["tcg_contractingunit@odata.bind"];
//                 await Xrm.WebApi.updateRecord("lead", toGuid(leadId), update);
//             } catch (e2) {
//                 throw e2;
//             }
//         } else if (msg.toLowerCase().includes("tcg_customergroupidnewone")) {
//             try {
//                 delete update["tcg_customergroupidnewone@odata.bind"];
//                 await Xrm.WebApi.updateRecord("lead", toGuid(leadId), update);
//             } catch (e3) {
//                 throw e3;
//             }
//         } else {
//             throw e;
//         }
//     }
// }



// calling this function from Form level
function getRecordFromForm(PrimaryControl) {
    var formContext = PrimaryControl; // It's a form
    var leadId = Xrm.Page.data.entity.getId();
    var entityName = formContext.data.entity.getEntityName();

    console.log("Record Id:", leadId);
    console.log("Entity Name:", entityName);
    openCustomDialog(leadId, PrimaryControl, "form");
}


// calling this function from Grid level
function getRecordFromGrid(selectedItems, PrimaryControl) {
    console.log("Selected Items:", selectedItems);
    var leadId = selectedItems[0].Id;
    openCustomDialog(leadId, PrimaryControl, "grid");
}

// master function
async function openCustomDialog(leadId, PrimaryControl, source) {

    localStorage.removeItem("leadRecordData"); // clear the local storage

    // Get the Lead ID from the URL
    // var leadId = Xrm.Page.data.entity.getId(); // Returns GUID with curly braces

    // // Remove curly braces if needed
    if (source === "form") leadId = leadId.replace("{", "").replace("}", "");

    console.log("Lead ID: " + leadId);



    // spin loading
    Xrm.Utility.showProgressIndicator("Loading...");

    try {
        await getLeadRecord(leadId); // fetch current lead record data
        await fetchMetaData(PrimaryControl, source);  // Fetch all meta data

    } catch (error) {
        alert("Something went wrong, please try again: " + error.message);
        console.log("Error fetching data:", error);
        Xrm.Utility.closeProgressIndicator();
        // To - do
        // This should be return 
        // return;
    }

    Xrm.Utility.closeProgressIndicator();

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

                // Apply payload to Lead form (your existing function that sets fields)
                // await applyPayloadToLead(PrimaryControl, payload);

                // await Xrm.Navigation.openAlertDialog(
                //     { title: "Success", text: "Lead validated successfully." },
                //     { height: 120, width: 320, position: 1 }
                // );

                // // In this version 3, I will handle the integration of HTML form in Qualify button's code
                // Mscrm.LeadCommandActions.qualifyLeadQuick();

                if (source === "form") {
                    // Apply to current form
                    await applyPayloadToLead(PrimaryControl, payload);
                    await Xrm.Navigation.openAlertDialog(
                        { title: "Success", text: "Lead validated successfully." },
                        { height: 120, width: 500, position: 1 }
                    );
                    // Keep existing behavior for form ribbon
                    Mscrm.LeadCommandActions.qualifyLeadQuick();
                } else if (source === "grid") {
                    // Persist changes via Web API without needing form context
                    await applyPayloadToLeadGrid(leadId, payload);
                    await Xrm.Navigation.openAlertDialog(
                        { title: "Success", text: "Lead validated successfully." },
                        { height: 120, width: 500, position: 1 }
                    );
                    // Try to refresh the grid so the user sees updated values
                    try { if (PrimaryControl && typeof PrimaryControl.refresh === "function") await PrimaryControl.refresh(); } catch (e) { }
                    try { const g = PrimaryControl && PrimaryControl.getGrid && PrimaryControl.getGrid(); if (g && typeof g.refresh === "function") g.refresh(); } catch (e) { }
                    
                    // Do NOT trigger form-only qualify action here
                }

            } catch (e) {
                await Xrm.Navigation.openErrorDialog({
                    message: "Failed to save lead: " + (e?.message || e || "")
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
