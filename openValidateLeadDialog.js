// URL of generative page :: /main.aspx?pagetype=control&controlName=MscrmControls.UxAgentControl&data={"refId":"5873dee8-e3bd-4293-a4ce-3fad8254f147"}
// In this version 3, 
// 1. I will handle this JS by clicking the button from main grid page by selecting the rows -- working on it


var __assign = (this && this.__assign) || Object.assign || function (t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function () { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t;
    return { next: verb(0), "throw": verb(1), "return": verb(2) };
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var MarketingSales;
(function (MarketingSales) {
    var CustomerExperienceSurveyRequest = (function () {
        /**
         * @param teamName team name to submit feedback
         * @param surveyName survey name to submit feedback
         * @param callType to indicate which feedback api to use
         * @param eventName to pass event name
         * @param surveyBody feedback content
         */
        function CustomerExperienceSurveyRequest(teamName, surveyName, callType, eventName, surveyBody) {
            this.TeamName = teamName;
            this.SurveyName = surveyName;
            this.CallType = callType;
            this.EventName = eventName;
            this.SurveyBody = surveyBody;
        }
        /**
         *
         */
        CustomerExperienceSurveyRequest.prototype.getMetadata = function () {
            var metadata = {
                boundParameter: null,
                parameterTypes: {
                    TeamName: {
                        typeName: 'Edm.String',
                        structuralProperty: 1
                    },
                    SurveyName: {
                        typeName: 'Edm.String',
                        structuralProperty: 1
                    },
                    CallType: {
                        typeName: 'Edm.String',
                        structuralProperty: 1
                    },
                    EventName: {
                        typeName: 'Edm.String',
                        structuralProperty: 1
                    },
                    SurveyBody: {
                        typeName: 'Edm.String',
                        structuralProperty: 1
                    }
                },
                operationName: 'CustomerExperienceSurveysService',
                operationType: 0
            };
            return metadata;
        };
        return CustomerExperienceSurveyRequest;
    }());
    MarketingSales.CustomerExperienceSurveyRequest = CustomerExperienceSurveyRequest;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
var MarketingSales;
(function (MarketingSales) {
    var EntityNames = (function () {
        function EntityNames() {
        }
        return EntityNames;
    }());
    EntityNames.Account = "account";
    EntityNames.Appointment = "appointment";
    // ToDo: this should be in Marketing; avoiding that dependency for now, 
    // since Marketing does not use ClassicWeb / UCI type definitions yet; which 
    // would cause compile issues.
    EntityNames.BulkOperation = "bulkoperation";
    EntityNames.Campaign = "campaign";
    EntityNames.CampaignActivity = "campaignactivity";
    EntityNames.CampaignResponse = "campaignresponse";
    EntityNames.Contact = "contact";
    EntityNames.Email = "email";
    EntityNames.Fax = "fax";
    // ToDo: this should be in Marketing; avoiding that dependency for now, 
    // since Marketing does not use ClassicWeb / UCI type definitions yet; which 
    // would cause compile issues.
    EntityNames.Lead = "lead";
    EntityNames.Letter = "letter";
    EntityNames.Organization = "organization";
    EntityNames.PhoneCall = "phonecall";
    EntityNames.RecurringAppointmentMaster = "recurringappointmentmaster";
    EntityNames.SystemUser = "systemuser";
    EntityNames.Task = "task";
    EntityNames.TransactionCurrency = "transactioncurrency";
    EntityNames.UnresolvedAddress = "unresolvedaddress";
    EntityNames.Opportunity = "opportunity";
    EntityNames.AgentUsage = "msdyn_salesagentusage";
    MarketingSales.EntityNames = EntityNames;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
var MarketingSales;
(function (MarketingSales) {
    var MetadataDrivenDialogConstants = (function () {
        function MetadataDrivenDialogConstants() {
        }
        return MetadataDrivenDialogConstants;
    }());
    MetadataDrivenDialogConstants.CampaignLookup = "campaignassociatedview_id";
    MetadataDrivenDialogConstants.LogResponseId = "cbLogResponse_id";
    MetadataDrivenDialogConstants.QualifyLeadCommandName = "QualifyLead";
    MetadataDrivenDialogConstants.QualifyStatus = "qualifyStatus";
    MetadataDrivenDialogConstants.paramQualifyStatus = "param_qualifyStatus";
    MetadataDrivenDialogConstants.ParentAccountName = "parentAccountName";
    MetadataDrivenDialogConstants.ParentAccountId = "parentAccountId";
    MetadataDrivenDialogConstants.ParentAccountLookupId = "parentAccountLookup_id";
    MetadataDrivenDialogConstants.ParentContactName = "parentContactName";
    MetadataDrivenDialogConstants.ParentContactId = "parentContactId";
    MetadataDrivenDialogConstants.ParentContactLookupId = "parentContactLookup_id";
    MetadataDrivenDialogConstants.paramParentAccountId = "param_parentAccountId";
    MetadataDrivenDialogConstants.paramParentAccountName = "param_parentAccountName";
    MetadataDrivenDialogConstants.paramParentContactId = "param_parentContactId";
    MetadataDrivenDialogConstants.paramParentContactName = "param_parentContactName";
    MetadataDrivenDialogConstants.paramTransactionCurrencyId = "param_transactioncurrencyid";
    // Constants for lead qualifcation
    MetadataDrivenDialogConstants.paramIsLeadQualificationInProgress = "param_isLeadQualificationInProgress";
    MetadataDrivenDialogConstants.requestParamQualifyLeadV2 = "param_entityReferenceSetV2Request";
    MetadataDrivenDialogConstants.responseParamQualifyLeadV2 = "param_entityReferenceSetV2Response";
    MetadataDrivenDialogConstants.showQualifiedLeadCopilotSummary = "param_showCopilotSummary";
    MetadataDrivenDialogConstants.isCopilotSummaryEnabled = "param_isCopilotSummaryEnabled";
    MetadataDrivenDialogConstants.paramQualifyLead = "param_qualifyLeadParam";
    //Constants for convert to lead MDD from email entity
    MetadataDrivenDialogConstants.firstName = "firstName";
    MetadataDrivenDialogConstants.paramFirstName = "param_firstName";
    MetadataDrivenDialogConstants.lastName = "lastName";
    MetadataDrivenDialogConstants.paramLastName = "param_lastName";
    MetadataDrivenDialogConstants.company = "company";
    MetadataDrivenDialogConstants.paramCompany = "param_company";
    MetadataDrivenDialogConstants.email = "email";
    MetadataDrivenDialogConstants.paramEmail = "param_email";
    MetadataDrivenDialogConstants.paramEntityId = "param_entityId";
    MetadataDrivenDialogConstants.paramEntityTypeCode = "param_entityTypeCode";
    MetadataDrivenDialogConstants.paramSubject = "param_subject";
    MetadataDrivenDialogConstants.paramSaveActivty = "param_saveActivity";
    MetadataDrivenDialogConstants.SaveActivityControlId = "cbSaveActivity_id";
    MetadataDrivenDialogConstants.paramOpenNewRecord = "param_openNewRecord";
    MetadataDrivenDialogConstants.OpenNewRecordControlId = "cbOpenNew_id";
    MetadataDrivenDialogConstants.paramLeadId = "param_leadId";
    MetadataDrivenDialogConstants.paramProcessInstanceId = "param_processInstanceId";
    MetadataDrivenDialogConstants.paramLastButtonClicked = "param_lastButtonClicked";
    MetadataDrivenDialogConstants.paramOwnerId = "param_ownerId";
    MetadataDrivenDialogConstants.lastButtonClicked = "lastButtonClicked";
    MetadataDrivenDialogConstants.paramEmailAddress = "param_emailAddress";
    //Convert Lead Options Dialog
    MetadataDrivenDialogConstants.isGrid = "isGrid";
    MetadataDrivenDialogConstants.paramIsGrid = "param_isGrid";
    MetadataDrivenDialogConstants.paramRecordCount = "param_recordCount";
    MetadataDrivenDialogConstants.contactFlag = "contact_flag";
    MetadataDrivenDialogConstants.accountFlag = "account_flag";
    MetadataDrivenDialogConstants.opportunityFlag = "opportunity_flag";
    MetadataDrivenDialogConstants.qualifyLead = "qualify_lead";
    MetadataDrivenDialogConstants.description = "description";
    MetadataDrivenDialogConstants.notification = "notification";
    MetadataDrivenDialogConstants.information = "information";
    MetadataDrivenDialogConstants.error = "error";
    MetadataDrivenDialogConstants.marketingSalesModule = "MarketingSales";
    MetadataDrivenDialogConstants.duplicateAccount = "entity_record_account";
    MetadataDrivenDialogConstants.duplicateContact = "entity_record_contact";
    MetadataDrivenDialogConstants.duplicateContactId = "selected_Record_contact";
    MetadataDrivenDialogConstants.duplicateAccountId = "selected_Record_account";
    MetadataDrivenDialogConstants.april2021Update = "April2021Update";
    MarketingSales.MetadataDrivenDialogConstants = MetadataDrivenDialogConstants;
    var DialogName = (function () {
        function DialogName() {
        }
        return DialogName;
    }());
    DialogName.DupWarningDialog = "DupWarningDialog";
    DialogName.convertLeadDialog = "ConvertLeadDialog";
    DialogName.convertLeadDialogV2 = "ConvertLeadDialogV2";
    DialogName.DuplicateQualifyLeadDialog = "DuplicateQualifyLeadDialog";
    MarketingSales.DialogName = DialogName;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
var MarketingSales;
(function (MarketingSales) {
    var QualifyLeadV2Control = (function () {
        function QualifyLeadV2Control() {
        }
        return QualifyLeadV2Control;
    }());
    QualifyLeadV2Control.ControlId = "qualify_lead_mdd_v2_id";
    QualifyLeadV2Control.RequestParam = "qualify_lead_mdd_v2_id.fieldControl.EntityCollectionRequestParam";
    QualifyLeadV2Control.ShowCopilotSummary = "qualify_lead_mdd_v2_id.fieldControl.ShowCopilotSummary";
    QualifyLeadV2Control.FooterId = "ok_id";
    QualifyLeadV2Control.AccountOdataType = "Microsoft.Dynamics.CRM.account";
    QualifyLeadV2Control.ContactOdataType = "Microsoft.Dynamics.CRM.contact";
    QualifyLeadV2Control.OpportunityOdataType = "Microsoft.Dynamics.CRM.opportunity";
    QualifyLeadV2Control.TeamName = "salesinsight";
    QualifyLeadV2Control.SurveyName = "enhancedleadqualification";
    QualifyLeadV2Control.SurveyEventName = "leadqualificationevent";
    QualifyLeadV2Control.EnhancedLeadQualificationPlaceholder = "the new lead qualification experience";
    QualifyLeadV2Control.CESSurveyTimestampKey = 'LeadHandoverCESSurveyTimestamp';
    MarketingSales.QualifyLeadV2Control = QualifyLeadV2Control;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../../../TypeDefinitions/CRM/SIClientUtilityLogger/SIClientUtilityLogger.d.ts" />
var MarketingSales;
(function (MarketingSales) {
    var TelemetryReporter = (function () {
        function TelemetryReporter() {
        }
        TelemetryReporter.CreateTelemetryParametersDictionary = function (moduleName, componentName, functionName, eventType, eventDetails) {
            var parameters = {};
            parameters["moduleName"] = moduleName;
            parameters["componentName"] = componentName;
            parameters["functionName"] = functionName;
            parameters["eventType"] = eventType;
            parameters["eventDetails"] = eventDetails;
            return parameters;
        };
        TelemetryReporter.getDataParams = function (paramsArray) {
            var data = {};
            if (paramsArray) {
                paramsArray.forEach(function (param) {
                    data[param.name] = param.value;
                });
            }
            return data;
        };
        Object.defineProperty(TelemetryReporter, "SILogger", {
            get: function () {
                try {
                    TelemetryReporter.SITelemetryLogger = SIClientUtilityLogger && SIClientUtilityLogger.Telemetry ? SIClientUtilityLogger.Telemetry : null;
                }
                catch (ex) {
                    console.error(ex);
                }
                return TelemetryReporter.SITelemetryLogger;
            },
            enumerable: true,
            configurable: true
        });
        TelemetryReporter.ReportComponentSuccess = function (context, scenario, methodName, componentName, action, actionOn, message, paramsArray) {
            try {
                // push data to SITraceCIEvent
                TelemetryReporter.SILogger && TelemetryReporter.SILogger.ReportUserAction({
                    context: context,
                    methodName: methodName,
                    action: action,
                    actionOn: actionOn,
                    message: message,
                    area: TelemetryReporter.Area,
                    data: __assign({}, TelemetryReporter.getDataParams(paramsArray), { scenario: scenario })
                });
            }
            catch (ex) {
                TelemetryReporter.reportUCIError(componentName + "." + methodName, ex);
            }
        };
        TelemetryReporter.ReportInfo = function (context, methodName, componentName, action, actionOn, message, enableUCILogging, paramsArray) {
            if (enableUCILogging === void 0) { enableUCILogging = false; }
            try {
                // push data to SITraceBIEvent
                TelemetryReporter.SILogger && TelemetryReporter.SILogger.ReportInfo({
                    context: context,
                    methodName: methodName,
                    action: action,
                    actionOn: actionOn,
                    message: message,
                    area: TelemetryReporter.Area,
                    data: __assign({}, TelemetryReporter.getDataParams(paramsArray), { componentName: componentName })
                });
                // push data to uciMonitorSuccess
                if (enableUCILogging) {
                    Xrm.Reporting.reportSuccess(componentName + "." + methodName, paramsArray);
                }
            }
            catch (ex) {
                TelemetryReporter.reportUCIError(componentName + "." + methodName, ex);
            }
        };
        TelemetryReporter.ReportWarning = function (context, methodName, componentName, action, actionOn, message, paramsArray) {
            try {
                // push data to SITraceEvent
                TelemetryReporter.SILogger && TelemetryReporter.SILogger.ReportWarning({
                    context: context,
                    methodName: methodName,
                    action: action,
                    actionOn: actionOn,
                    message: message,
                    area: TelemetryReporter.Area,
                    data: __assign({}, TelemetryReporter.getDataParams(paramsArray), { componentName: componentName })
                });
            }
            catch (ex) {
                TelemetryReporter.reportUCIError(componentName + "." + methodName, ex);
            }
        };
        TelemetryReporter.ReportError = function (context, methodName, message, componentName, action, actionOn, paramsArray) {
            try {
                // push data to SITraceEvent
                TelemetryReporter.SILogger && TelemetryReporter.SILogger.ReportError({
                    context: context,
                    methodName: methodName,
                    message: message,
                    action: action,
                    actionOn: actionOn,
                    area: TelemetryReporter.Area,
                    data: __assign({}, TelemetryReporter.getDataParams(paramsArray), { componentName: componentName })
                });
            }
            catch (ex) {
                TelemetryReporter.reportUCIError(componentName + "." + methodName, ex);
            }
        };
        TelemetryReporter.reportUCIError = function (componentName, error) {
            try {
                Xrm.Reporting.reportFailure(componentName, Error(JSON.stringify(error)));
            }
            catch (ex) {
                console.error(ex);
            }
        };
        TelemetryReporter.LogExecutionTelemetryForWeb = function (eventName, parameters) {
            try {
                Xrm.Internal.addMetric(eventName, parameters);
            }
            catch (e) {
                console.error(e);
            }
        };
        return TelemetryReporter;
    }());
    TelemetryReporter.DialogName = "DuplicateQualifyLeadDialog";
    TelemetryReporter.Area = "Qualify Lead Duplicate Detection";
    MarketingSales.TelemetryReporter = TelemetryReporter;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../ClientCommon/EntityNames.ts" />
/// <reference path="../ClientCommon/MetadataDrivenDialogConstants.ts" />
/// <reference path="../ClientCommon/QualifyLeadV2Constants.ts" />
/// <reference path ="Reporter.ts" />
var MarketingSales;
(function (MarketingSales) {
    'use strict';
    var _this = this;
    var LeadUtil = (function () {
        function LeadUtil() {
            var _this = this;
            this.convertLeadOnOkClickHandler = function (context) {
                try {
                    var formContext = (context != null && context != undefined) ? context.getFormContext() : Xrm.Page;
                    ClientUtility.DialogUtil.setAttributeValue(ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked, ClientUtility.MetadataDrivenDialogConstants.DialogOkId);
                    formContext.ui.close();
                }
                catch (e) {
                    LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "convertLeadOnOkClickHandler", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
                }
            };
            this.convertLeadOnLoadHandler = function (context) {
                return __awaiter(_this, void 0, void 0, function () {
                    var formContext, isGrid, recordCount, description, header, description, header, notification, isQualifyLeadV2Enabled, bulkQualifyInfo, e_1;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                _a.trys.push([0, 4, , 5]);
                                formContext = (context != null && context != undefined) ? context.getFormContext() : Xrm.Page;
                                isGrid = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramIsGrid);
                                recordCount = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramRecordCount);
                                if (!!ClientUtility.DataUtil.isNullOrUndefined(isGrid)) return [3 /*break*/, 3];
                                if (!(!isGrid || recordCount == 1)) return [3 /*break*/, 1];
                                description = formContext.ui.controls.get(MarketingSales.MetadataDrivenDialogConstants.description);
                                if (description) {
                                    description.setLabel(MarketingSales.ResourceStringProvider.getResourceString("ConvertLead_Description"));
                                }
                                header = formContext.ui.controls.get(MarketingSales.MetadataDrivenDialogConstants.qualifyLead);
                                if (header) {
                                    header.setLabel(MarketingSales.ResourceStringProvider.getResourceString("ConvertLead_Header"));
                                }
                                return [3 /*break*/, 3];
                            case 1:
                                description = formContext.ui.controls.get(MarketingSales.MetadataDrivenDialogConstants.description);
                                if (description) {
                                    description.setLabel(MarketingSales.ResourceStringProvider.getResourceString("ConvertLeadGrid_Description"));
                                }
                                header = formContext.ui.controls.get(MarketingSales.MetadataDrivenDialogConstants.qualifyLead);
                                if (header) {
                                    header.setLabel(MarketingSales.ResourceStringProvider.getResourceString("ConvertLeadGrid_Header"));
                                }
                                notification = formContext.ui.controls.get(MarketingSales.MetadataDrivenDialogConstants.notification);
                                if (notification) {
                                    // label is picked directly from metadata, and shown only if this is grid.
                                    notification.setVisible(true);
                                }
                                return [4 /*yield*/, LeadUtil.isQualifyLeadV2Enabled()];
                            case 2:
                                isQualifyLeadV2Enabled = _a.sent();
                                if (isQualifyLeadV2Enabled) {
                                    bulkQualifyInfo = formContext.ui.controls.get("bulkQualifyInfo");
                                    if (bulkQualifyInfo) {
                                        bulkQualifyInfo.setVisible(true);
                                    }
                                }
                                _a.label = 3;
                            case 3: return [3 /*break*/, 5];
                            case 4:
                                e_1 = _a.sent();
                                LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "convertLeadOnLoadHandler", MarketingSales.MetadataDrivenDialogConstants.error, e_1.toString());
                                return [3 /*break*/, 5];
                            case 5: return [2 /*return*/];
                        }
                    });
                });
            };
        }
        LeadUtil.isQualifyLeadV2Enabled = function () {
            return __awaiter(this, void 0, void 0, function () {
                var isOctober2024FCBEnabled, isQualifyLeadV2FCSEnabled, isUserEligibleForQualifiedLeadV2, qualifyLeadAdminSettingV2Data, _a, isQualifyLeadV2AdminSettingEnabled;
                return __generator(this, function (_b) {
                    switch (_b.label) {
                        case 0:
                            if (ClientUtility.ClientUtil.isMobile() || ClientUtility.ClientUtil.isOutlook()) {
                                return [2 /*return*/, false];
                            }
                            isOctober2024FCBEnabled = LeadUtil.IsFcbEnabled(LeadUtil.FCB_October2024Update);
                            isQualifyLeadV2FCSEnabled = LeadUtil.getFeatureControlSetting(LeadUtil.FCS_QualifyLeadNamespace, LeadUtil.FCS_EnableQualifyLeadV2, false);
                            isUserEligibleForQualifiedLeadV2 = isQualifyLeadV2FCSEnabled && isOctober2024FCBEnabled;
                            if (!isUserEligibleForQualifiedLeadV2) return [3 /*break*/, 2];
                            return [4 /*yield*/, LeadUtil.getQualifyLeadAdminSettingData()];
                        case 1:
                            _a = _b.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            _a = null;
                            _b.label = 3;
                        case 3:
                            qualifyLeadAdminSettingV2Data = _a;
                            isQualifyLeadV2AdminSettingEnabled = qualifyLeadAdminSettingV2Data && qualifyLeadAdminSettingV2Data.IsQualifyLeadV2Enabled;
                            return [2 /*return*/, isQualifyLeadV2AdminSettingEnabled && isUserEligibleForQualifiedLeadV2];
                    }
                });
            });
        };
        LeadUtil.getFeatureControlSetting = function (namespace, settingKey, defaultValue) {
            var fcsValue = (Xrm.Utility.getGlobalContext()).getFeatureControlSetting(namespace, settingKey);
            if (fcsValue !== null && fcsValue !== undefined) {
                // Found the feature
                return fcsValue;
            }
            return defaultValue;
        };
        LeadUtil.convertLeadDialog = function (isGrid, recordCount) {
            try {
                // TODO :- Retreive admin setting value and add a check
                var dialogHeight = (recordCount > 1) ? 416 : 340;
                var options = {
                    height: dialogHeight,
                    width: 450,
                    position: 1 /* center */
                };
                var dialogParams = {};
                dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramIsGrid] = isGrid;
                dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramRecordCount] = recordCount;
                return Xrm.Navigation.openDialog(MarketingSales.DialogName.convertLeadDialog, options, dialogParams);
            }
            catch (e) {
                LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "convertLeadDialog", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
            }
        };
        ;
        LeadUtil.convertLeadDialogV2 = function (isGrid, leadIdParam, status, dialogParameters) {
            try {
                var options = {
                    height: 900,
                    width: 404,
                    position: 2 /* side */
                };
                var leadId = leadIdParam || ClientUtility.Guid.create(Xrm.Page.data.entity.getId());
                var dialogParams = {};
                if (dialogParameters && dialogParameters.parameters) {
                    var parentAccountId = dialogParameters.parameters.param_parentAccountId;
                    var parentContactId = dialogParameters.parameters.param_parentContactId;
                    var parentAccountName = dialogParameters.parameters.entity_record_account && dialogParameters.parameters.entity_record_account.name;
                    var contactFullName = "";
                    if (dialogParameters.parameters.entity_record_contact) {
                        contactFullName = ClientUtility.DataUtil.isNullOrUndefined(dialogParameters.parameters.entity_record_contact.firstname) ? dialogParameters.parameters.entity_record_contact.lastname :
                            (dialogParameters.parameters.entity_record_contact.firstname + " " + dialogParameters.parameters.entity_record_contact.lastname);
                    }
                    var parentContactName = dialogParameters.parameters.entity_record_contact && contactFullName;
                    dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId] = parentAccountId;
                    dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramParentContactId] = parentContactId;
                    dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountName] = parentAccountName;
                    dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramParentContactName] = parentContactName;
                    dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramQualifyLead] = dialogParameters.parameters.param_qualifyLeadParam;
                }
                dialogParams[MarketingSales.MetadataDrivenDialogConstants.paramIsGrid] = isGrid;
                Xrm.Navigation.openDialog(MarketingSales.DialogName.convertLeadDialogV2, options, dialogParams).then(function () {
                    Xrm.Page.data.refresh(true).then(function () {
                        Xrm.Page.ui.refreshRibbon();
                    });
                });
                return;
            }
            catch (e) {
                LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "convertLeadDialogV2", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
            }
        };
        ;
        LeadUtil.getIsUci = function () {
            try {
                if (Xrm && Xrm.Internal && Xrm.Internal.isUci && Xrm.Internal.isUci()) {
                    return true;
                }
                return false;
            }
            catch (e) {
                return false;
            }
        };
        ;
        LeadUtil.duplicateDetectionAdminSettings = function () {
            try {
                if (LeadUtil.isApril2021UpdateEnabled()) {
                    return true;
                }
                else {
                    var attributes = Xrm.Utility.getGlobalContext().organizationSettings.attributes;
                    var params = {};
                    var duplicateDialogoptions = null;
                    if (!ClientUtility.DataUtil.isNullOrUndefined(attributes)) {
                        // for UCI the attribute name is like attributes.qualifyleadadditionaloptions (in small case)
                        if (LeadUtil.getIsUci() && !ClientUtility.DataUtil.isNullOrUndefined(attributes.qualifyleadadditionaloptions) && attributes.qualifyleadadditionaloptions != "{}") {
                            if (attributes.qualifyleadadditionaloptions && attributes.qualifyleadadditionaloptions.length)
                                duplicateDialogoptions = JSON.parse(attributes.qualifyleadadditionaloptions);
                        }
                        else if (!ClientUtility.DataUtil.isNullOrUndefined(attributes.qualifyLeadAdditionalOptions) && attributes.qualifyLeadAdditionalOptions != "{}") {
                            if (attributes.qualifyLeadAdditionalOptions && attributes.qualifyLeadAdditionalOptions.length)
                                duplicateDialogoptions = JSON.parse(attributes.qualifyLeadAdditionalOptions);
                        }
                        else {
                            return false; // By default should return false.
                        }
                        if (duplicateDialogoptions) {
                            params["DuplicateDetectionMerge"] = !ClientUtility.DataUtil.isNullOrUndefined(duplicateDialogoptions["DuplicateDetectionMerge"]) ? duplicateDialogoptions["DuplicateDetectionMerge"].toLowerCase() === "true" : "false";
                            MarketingSales.TelemetryReporter.ReportInfo(null, "duplicateDetectionAdminSettings", MarketingSales.DialogName.DuplicateQualifyLeadDialog, MarketingSales.MetadataDrivenDialogConstants.information, null, "DuplicateDetectionMerge :" + params["DuplicateDetectionMerge"], true);
                            return params ? params["DuplicateDetectionMerge"] : false;
                        }
                        else
                            return false;
                    }
                }
                MarketingSales.TelemetryReporter.ReportInfo(null, "duplicateDetectionAdminSettings", MarketingSales.DialogName.DuplicateQualifyLeadDialog, MarketingSales.MetadataDrivenDialogConstants.information, null, "DuplicateDetectionMerge :" + params["DuplicateDetectionMerge"], true);
                return false; //By default should return false
            }
            catch (e) {
                MarketingSales.TelemetryReporter.ReportError(null, "duplicateDetectionAdminSettings", JSON.stringify(Error("Error fetching admin settings for duplicate and merge. : " + e)), MarketingSales.DialogName.DuplicateQualifyLeadDialog);
                return false; //By default should return false
            }
        };
        ;
        LeadUtil.isApril2021UpdateEnabled = function () {
            try {
                if (LeadUtil.getIsUci()) {
                    return Xrm.Internal.isFeatureEnabled(MarketingSales.MetadataDrivenDialogConstants.april2021Update);
                }
                else
                    return false;
            }
            catch (e) {
                LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "FCBAprilUpdateEnabled", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
                return false;
            }
        };
        /**
         * Checks if the current app is a single-session or multi-session app (e.g., CSW).
         * @returns Promise resolving to a boolean value.
         */
        LeadUtil.isCurrentAppMultiSession = function () {
            return __awaiter(this, void 0, void 0, function () {
                var globalContext, app, query, result, navigationType, error_1;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 3, , 4]);
                            if (Xrm.App.sessions) {
                                return [2 /*return*/, true];
                            }
                            globalContext = Xrm.Utility.getGlobalContext();
                            return [4 /*yield*/, globalContext.getCurrentAppProperties()];
                        case 1:
                            app = _a.sent();
                            query = "?$select=" + LeadUtil.AppModuleConstants.Fields.NavigationType.FieldName;
                            return [4 /*yield*/, Xrm.WebApi.retrieveRecord(LeadUtil.AppModuleConstants.AppModule, app.appId, query)];
                        case 2:
                            result = _a.sent();
                            navigationType = result.navigationtype;
                            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, this.isCurrentAppMultiSession.name, MarketingSales.MetadataDrivenDialogConstants.information, "App's session type retrieved successfully: " + navigationType);
                            return [2 /*return*/, navigationType === LeadUtil.AppModuleConstants.Fields.NavigationType.Type.MultiSession];
                        case 3:
                            error_1 = _a.sent();
                            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, this.isCurrentAppMultiSession.name, MarketingSales.MetadataDrivenDialogConstants.error, "Failed to get app's session type: " + error_1.message);
                            return [2 /*return*/, false];
                        case 4: return [2 /*return*/];
                    }
                });
            });
        };
        LeadUtil.createTabInFocusedSession = function (pageInput) {
            return __awaiter(this, void 0, void 0, function () {
                var focusedSession, tabInput, error_2;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            focusedSession = Xrm.App.sessions.getFocusedSession();
                            if (!focusedSession) {
                                throw new Error("No focused session found.");
                            }
                            tabInput = {
                                pageInput: pageInput,
                                options: { isFocused: true },
                            };
                            // Create a new tab in the focused session
                            return [4 /*yield*/, focusedSession.tabs.createTab(tabInput)];
                        case 1:
                            // Create a new tab in the focused session
                            _a.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            error_2 = _a.sent();
                            throw new Error("Failed to create a tab in the focused session: " + error_2.message);
                        case 3: return [2 /*return*/];
                    }
                });
            });
        };
        /**
         * Navigates to the qualified record in the appropriate session mode.
         * If the app is multi-session, it attempts to open the record in a new tab in currently focused session.
         * If navigation to a new tab fails, it falls back to opening in a new session.
         * In a single-session app, it directly navigates to the record.
         *
         * @param recordId - The ID of the record to navigate to.
         * @param recordType - The entity type of the record.
         */
        LeadUtil.navigateToQualifiedRecord = function (recordId, recordType) {
            return __awaiter(this, void 0, void 0, function () {
                var isMultiSession, pageInput_1, error_3, error_4;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 8, , 9]);
                            return [4 /*yield*/, LeadUtil.isCurrentAppMultiSession()];
                        case 1:
                            isMultiSession = _a.sent();
                            pageInput_1 = {
                                pageType: "entityrecord",
                                entityName: recordType,
                                entityId: recordId,
                            };
                            if (!isMultiSession) return [3 /*break*/, 6];
                            _a.label = 2;
                        case 2:
                            _a.trys.push([2, 4, , 5]);
                            return [4 /*yield*/, LeadUtil.createTabInFocusedSession(pageInput_1)];
                        case 3:
                            _a.sent();
                            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, this.navigateToQualifiedRecord.name, MarketingSales.MetadataDrivenDialogConstants.information, "Navigated to qualified opportunity in multi-session app");
                            return [3 /*break*/, 5];
                        case 4:
                            error_3 = _a.sent();
                            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, this.navigateToQualifiedRecord.name, MarketingSales.MetadataDrivenDialogConstants.error, "Navigation to qualified opportunity in multi-session app failed: " + error_3.message);
                            // Fallback: Navigate to the record in a new session
                            setTimeout(function () { return Xrm.Navigation.navigateTo(pageInput_1, { target: 1 }); }, 100);
                            return [3 /*break*/, 5];
                        case 5: return [3 /*break*/, 7];
                        case 6:
                            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, this.navigateToQualifiedRecord.name, MarketingSales.MetadataDrivenDialogConstants.information, "Navigated to qualified opportunity in single-session app");
                            // Workaround to avoid confirmation dialog upon navigation
                            setTimeout(function () { return Xrm.Navigation.navigateTo(pageInput_1, { target: 1 }); }, 100);
                            _a.label = 7;
                        case 7: return [3 /*break*/, 9];
                        case 8:
                            error_4 = _a.sent();
                            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, this.navigateToQualifiedRecord.name, MarketingSales.MetadataDrivenDialogConstants.error, "Unexpected error: " + error_4.message);
                            return [3 /*break*/, 9];
                        case 9: return [2 /*return*/];
                    }
                });
            });
        };
        return LeadUtil;
    }());
    LeadUtil.FCB_October2024Update = "October2024Update";
    LeadUtil.FCS_EnableQualifyLeadV2 = "EnableQualifyLeadV2";
    LeadUtil.FCS_SkipQualifyLeadV2Survey = "SkipQualifyLeadV2Survey";
    LeadUtil.FCS_QualifyLeadNamespace = "DynamicsCRMExtensions.QualifyLead";
    LeadUtil.qualifyLeadAdminSettingConstants = {
        qualifyLeadAdditionalOptions: "qualifyleadadditionaloptions",
        organization: "organization",
        isQualifyLeadV2Enabled: "IsQualifyLeadV2Enabled"
    };
    LeadUtil.AppModuleConstants = {
        AppModule: "appmodule",
        Fields: {
            NavigationType: {
                FieldName: "navigationtype", Type: { SingleSession: 0, MultiSession: 1 }
            }
        }
    };
    LeadUtil.getQualifyLeadAdminSettingData = function () {
        return __awaiter(_this, void 0, void 0, function () {
            var orgSettings, request, result, leadQualificationData, columnData;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        orgSettings = Xrm.Utility.getGlobalContext().organizationSettings;
                        request = "?$select=" + LeadUtil.qualifyLeadAdminSettingConstants.qualifyLeadAdditionalOptions;
                        return [4 /*yield*/, Xrm.WebApi.retrieveRecord(LeadUtil.qualifyLeadAdminSettingConstants.organization, orgSettings.organizationId, request)];
                    case 1:
                        result = _a.sent();
                        leadQualificationData = null;
                        columnData = JSON.parse(result[LeadUtil.qualifyLeadAdminSettingConstants.qualifyLeadAdditionalOptions]);
                        if (result[LeadUtil.qualifyLeadAdminSettingConstants.qualifyLeadAdditionalOptions].localeCompare("{}") != 0) {
                            leadQualificationData = {
                                IsQualifyLeadV2Enabled: columnData[LeadUtil.qualifyLeadAdminSettingConstants.isQualifyLeadV2Enabled]
                            };
                        }
                        return [2 /*return*/, leadQualificationData];
                }
            });
        });
    };
    LeadUtil.retrieveProcessInstanceId = function (pageData) {
        if (pageData === void 0) { pageData = Xrm.Page.data; }
        var processInstanceReference = null;
        if (pageData && pageData.process) {
            // ToDo: remove. this is a workaround for 'BPF does not get displayed after create and save' (work item 468994)
            var reference = pageData.process._entityRef;
            if (reference && reference.id && reference.id.guid === ClientUtility.Guid.Empty) {
                pageData.process._entityRef = {
                    etn: MarketingSales.EntityNames.Lead,
                    id: { guid: ClientUtility.Guid.create(pageData.entity.getId()) },
                    name: ""
                };
            }
            // end ToDo
            var instanceId = pageData.process.getInstanceId();
            if (!ClientUtility.DataUtil.isNullOrEmptyString(instanceId)) {
                processInstanceReference = { entityType: "businessprocessflowinstance", id: instanceId };
            }
        }
        return processInstanceReference;
    };
    LeadUtil.logTelemetryData = function (componentName, functionName, eventType, eventMessage, parameters) {
        if (eventMessage === void 0) { eventMessage = null; }
        if (parameters === void 0) { parameters = {}; }
        try {
            if (LeadUtil.getIsUci()) {
                var event_1 = new Array();
                event_1.push({ name: "functionName", value: functionName });
                for (var property in parameters) {
                    if (parameters[property] != null) {
                        event_1.push({ name: property, value: parameters[property] });
                    }
                }
                eventMessage && eventMessage != "" && event_1.push({ name: "eventMessage", value: eventMessage });
                if (eventType == MarketingSales.MetadataDrivenDialogConstants.error) {
                    Xrm.Reporting.reportFailure(MarketingSales.MetadataDrivenDialogConstants.marketingSalesModule + "." + componentName, event_1);
                }
                else {
                    Xrm.Reporting.reportSuccess(MarketingSales.MetadataDrivenDialogConstants.marketingSalesModule + "." + componentName, event_1);
                }
            }
            else {
                var event = {};
                if (eventMessage && eventMessage != "") {
                    parameters["eventMessage"] = eventMessage;
                }
                event["moduleName"] = MarketingSales.MetadataDrivenDialogConstants.marketingSalesModule;
                event["componentName"] = componentName;
                event["functionName"] = functionName;
                event["eventType"] = eventType;
                event["eventDetails"] = JSON.stringify(parameters);
                Xrm.Internal.addMetric("salesoperationevents", event);
            }
        }
        catch (e) {
        }
    };
    LeadUtil.qualifyLeadAdditionalOptionsAdminSettings = function () {
        try {
            var attributes = Xrm.Utility.getGlobalContext().organizationSettings.attributes;
            var params = {};
            var qualifyLeadAdditionalOptions = null;
            if (!ClientUtility.DataUtil.isNullOrUndefined(attributes)) {
                // for UCI the attribute name is like attributes.qualifyleadadditionaloptions (in small case)
                if (LeadUtil.getIsUci() && !ClientUtility.DataUtil.isNullOrUndefined(attributes.qualifyleadadditionaloptions) && attributes.qualifyleadadditionaloptions != "{}") {
                    qualifyLeadAdditionalOptions = JSON.parse(attributes.qualifyleadadditionaloptions);
                }
                else if (!ClientUtility.DataUtil.isNullOrUndefined(attributes.qualifyLeadAdditionalOptions) && attributes.qualifyLeadAdditionalOptions != "{}") {
                    qualifyLeadAdditionalOptions = JSON.parse(attributes.qualifyLeadAdditionalOptions);
                }
                else {
                    LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "qualifyLeadAdditionalOptionsAdminSettings", MarketingSales.MetadataDrivenDialogConstants.information, "attributes.qualifyleadadditionaloptions is empty object or null");
                    return true;
                }
                params["DefaultQualifyFlowState"] = !ClientUtility.DataUtil.isNullOrUndefined(qualifyLeadAdditionalOptions["DefaultQualifyFlow"]) ? qualifyLeadAdditionalOptions["DefaultQualifyFlow"].toLowerCase() === "true" : "false";
                LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "qualifyLeadAdditionalOptionsAdminSettings", MarketingSales.MetadataDrivenDialogConstants.information, "", params);
                return params["DefaultQualifyFlowState"];
            }
            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "qualifyLeadAdditionalOptionsAdminSettings", MarketingSales.MetadataDrivenDialogConstants.information, "Organization setting attributes or attribute qualifyleadadditionaloptions is undefined.");
            return true;
        }
        catch (e) {
            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "qualifyLeadAdditionalOptionsAdminSettings", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
            return true;
        }
    };
    LeadUtil.FCBOctoberUpdateEnabled = function () {
        try {
            if (LeadUtil.getIsUci()) {
                // UCI reads the FCB in different format.
                return Xrm.Internal.isFeatureEnabled('October2019Update');
            }
            return Xrm.Internal.isFeatureEnabled('FCB.October2019Update');
        }
        catch (e) {
            LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "FCBOctoberUpdateEnabled", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
            return false;
        }
    };
    LeadUtil.IsFcbEnabled = function (featureName) {
        try {
            if (LeadUtil.getIsUci()) {
                //Â UCIÂ readsÂ theÂ FCBÂ inÂ differentÂ format.
                return Xrm.Internal.isFeatureEnabled(featureName);
            }
            var fcbWeb = "FCB." + featureName;
            return Xrm.Internal.isFeatureEnabled(fcbWeb);
        }
        catch (e) {
            return false;
        }
    };
    MarketingSales.LeadUtil = LeadUtil;
})(MarketingSales || (MarketingSales = {}));
var getWindowObject = function () {
    return window.parent || window;
};
var setCESSurveyTimestamp = function (timestampkey) {
    if (!localStorage.getItem(timestampkey)) {
        localStorage.setItem(timestampkey, Date.now().toString());
    }
};
var checkCoolingPeriod = function (timestampkey) {
    var lastSubmittedTimestamp = localStorage.getItem(timestampkey);
    if (!lastSubmittedTimestamp)
        return true;
    var differenceInDays = (Date.now() - Number(lastSubmittedTimestamp)) / (1000 * 60 * 60 * 24);
    return differenceInDays > 90;
};
var getOrgInformation = function () {
    var globalContext = Xrm.Utility.getGlobalContext();
    var orgUrl = globalContext.getClientUrl();
    var crmRegion = getRegion();
    var crmGeo = GetGeo(new URL(orgUrl).hostname.split('.')[1]);
    var userLang = getLocale(globalContext.userSettings.languageId);
    var tenantId = globalContext.organizationSettings.organizationTenant;
    var userId = globalContext.userSettings.userId.replace(/[{}]/g, "");
    var envType = getEnvironmentType();
    return { orgUrl: orgUrl, crmRegion: crmRegion, crmGeo: crmGeo, userLang: userLang, tenantId: tenantId, userId: userId, envType: envType };
};
var getEnvironmentType = function () {
    var orgGeo = Xrm.Utility.getGlobalContext().organizationSettings.organizationGeo;
    return (orgGeo === "TIP" || orgGeo === "TST") ? "INT" : "PROD";
};
var getLocale = function (langCode) {
    var locales = {
        1025: 'ar', 1026: 'bg', 1027: 'ca', 1028: 'zh-Hant', 1029: 'cs',
        1030: 'da', 1031: 'de', 1034: 'es-ES', 1035: 'fi', 1036: 'fr-FR',
        1037: 'he', 1038: 'hu', 1040: 'it', 1041: 'ja', 1042: 'ko',
        1043: 'nl', 1044: 'nb-no', 1045: 'pl', 1046: 'pt-br', 1048: 'ro',
        1049: 'ru', 1050: 'hr', 1051: 'sk', 1053: 'sv', 1054: 'th',
        1055: 'tr', 1057: 'id', 1058: 'uk', 1060: 'sl', 1061: 'et',
        1062: 'lv', 1063: 'lt', 1066: 'vi', 1069: 'eu', 1081: 'hi',
        1086: 'ms', 1087: 'kk', 1110: 'gl', 2052: 'zh-Hans', 2068: 'nn-no',
        2070: 'pt-pt', 2110: 'ms', 9242: 'sr-latn-rs', 10266: 'sr-cyrl-rs'
    };
    return locales[langCode] || 'en-US';
};
var GetGeo = function (regionCode) {
    var geoMapping = {
        'crm': 'NAM', 'crm2': 'SAM', 'crm3': 'CAN', 'crm4': 'EUR',
        'crm5': 'APJ', 'crm6': 'OCE', 'crm7': 'JPN', 'crm8': 'IND',
        'crm10': 'TIP', 'crm11': 'GBR', 'crm12': 'FRA', 'crm14': 'ZAF',
        'crm15': 'UAE', 'crm17': 'CHE'
    };
    return geoMapping[regionCode] || regionCode;
};
var getRegion = function () {
    var orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
    if (orgUrl) {
        var hostname = new URL(orgUrl).hostname.split('.')[1];
        var europeHostnames = ["crm4", "crm11", "crm12", "crm16", "crm17", "crm19", "crm22", "crm23", "crm24"];
        if (europeHostnames.indexOf(hostname) !== -1) {
            return "EUROPE";
        }
    }
    return "WORLD";
};
/// <reference path="../../Lead/Contracts/CustomerExperienceSurveyRequest.ts" />
/// <reference path="../../Lead/LeadUtil.ts" />
/// <reference path="../QualifyLeadV2Constants.ts" />
/// <reference path="CESSurveyUtil.ts" />
var MarketingSales;
(function (MarketingSales) {
    var CESSurvey = (function () {
        function CESSurvey() {
        }
        CESSurvey.showSurvey = function (teamName, surveyName, surveyEventName) {
            var _this = this;
            var _a = getOrgInformation(), crmRegion = _a.crmRegion, userLang = _a.userLang, tenantId = _a.tenantId, userId = _a.userId, envType = _a.envType;
            var config = {
                teamName: teamName,
                surveyName: surveyName,
                userId: userId,
                tenantId: tenantId,
                locale: userLang,
                width: 520,
                uiType: "Modal",
                template: "SAT",
                environment: envType,
                region: crmRegion,
                callbackFunctions: {
                    submitFeedback: function (teamName, surveyName, userId, feedback) {
                        return __awaiter(_this, void 0, void 0, function () {
                            var feedbackRequest, response, error_5;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        feedbackRequest = new MarketingSales.CustomerExperienceSurveyRequest(teamName, surveyName, 'submitfeedback', surveyEventName, this.createRequestBody(feedback));
                                        _a.label = 1;
                                    case 1:
                                        _a.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, Xrm.WebApi.online.execute(feedbackRequest)];
                                    case 2:
                                        response = _a.sent();
                                        if (response && response.ok) {
                                            MarketingSales.LeadUtil.logTelemetryData("CESSurvey", "showSurvey", MarketingSales.MetadataDrivenDialogConstants.information, surveyName + " survey completed");
                                        }
                                        return [3 /*break*/, 4];
                                    case 3:
                                        error_5 = _a.sent();
                                        MarketingSales.LeadUtil.logTelemetryData("CESSurvey", "showSurvey", MarketingSales.MetadataDrivenDialogConstants.error, surveyName + ": " + error_5.message);
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        });
                    },
                },
                stringPlaceholders: {
                    appName: MarketingSales.QualifyLeadV2Control.EnhancedLeadQualificationPlaceholder
                },
                theme: {
                    accentBaseColor: '#742774',
                    layerCornerRadius: 0,
                    controlCornerRadius: 0
                }
            };
            var element = getWindowObject().document.getElementById('mySurvey');
            if (element) {
                element.config = config;
            }
            else {
                MarketingSales.LeadUtil.logTelemetryData("CESSuevey", "showSurvey", MarketingSales.MetadataDrivenDialogConstants.error, surveyName + "'s survey element not found");
            }
        };
        CESSurvey.createRequestBody = function (feedbackObj) {
            var _a = getOrgInformation(), orgUrl = _a.orgUrl, crmGeo = _a.crmGeo, userLang = _a.userLang, tenantId = _a.tenantId;
            var feedbackBody = {
                TenantId: tenantId,
                Culture: userLang,
                UrlReferrer: orgUrl,
                DeviceType: Xrm.Utility.getGlobalContext().client.getClient(),
                ProductContext: feedbackObj.ProductContext,
                IsDismissed: feedbackObj.IsDismissed,
                Geo: crmGeo,
                Feedbacks: feedbackObj.Feedbacks
            };
            return JSON.stringify(feedbackBody);
        };
        CESSurvey.loadCustomerExperienceSurveyScript = function () {
            return __awaiter(this, void 0, void 0, function () {
                return __generator(this, function (_a) {
                    return [2 /*return*/, new Promise(function (resolve, reject) {
                        var windowObject = getWindowObject();
                        var surveyComponentExists = windowObject.customElements.get('survey-sdk');
                        if (!surveyComponentExists) {
                            var surveyElement = document.createElement('survey-sdk');
                            surveyElement.id = 'mySurvey';
                            windowObject.document.body.appendChild(surveyElement);
                            var script = document.createElement('script');
                            script.type = 'text/javascript';
                            var url = Xrm.Utility.getGlobalContext().getWebResourceUrl('MarketingSales/External/survey.lib.umd.v1.0.10.min.js');
                            if (!url) {
                                reject(new Error('CustomerExperienceSurvey webresource not found'));
                                return;
                            }
                            script.src = url;
                            script.async = true;
                            script.onload = function () { return resolve(); };
                            script.onerror = function (error) { return reject(new Error("Failed to load script: " + JSON.stringify(error.message || error))); };
                            windowObject.document.head.appendChild(script);
                        }
                        else {
                            resolve();
                        }
                    })];
                });
            });
        };
        CESSurvey.showCESSurvey = function (teamName, surveyName, surveyEventName) {
            var _this = this;
            this.loadCustomerExperienceSurveyScript()
                .then(function () { return _this.showSurvey(teamName, surveyName, surveyEventName); })
                .catch(function (error) {
                    MarketingSales.LeadUtil.logTelemetryData("CESSurvey", "showCESSurvey", MarketingSales.MetadataDrivenDialogConstants.error, "Unable to load customer experience survey: " + error.message);
                });
        };
        return CESSurvey;
    }());
    MarketingSales.CESSurvey = CESSurvey;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="EntityNames.ts" />
var MarketingSales;
(function (MarketingSales) {
    var TransactionCurrency = (function () {
        function TransactionCurrency() {
        }
        /**
         * Gets the default currency for the user
         * @param systemUserId The user Id.
         * @param successCallback The success callback function.
         * @param errorCallback The error callback function.
         */
        TransactionCurrency.getDefaultTransactionCurrency = function (systemUserId, successCallback, errorCallback) {
            var currencyId = null;
            Xrm.WebApi.retrieveRecord(MarketingSales.EntityNames.SystemUser, systemUserId, ClientUtility.ODataUtil.getSelectOption(["transactioncurrencyid", "organizationid"])).then((function (response) {
                if (!ClientUtility.DataUtil.isNullOrUndefined(response)) {
                    var organizationId = response.organizationid;
                    currencyId = response.transactioncurrencyid;
                    if (ClientUtility.DataUtil.isNullOrUndefined(currencyId) && !ClientUtility.DataUtil.isNullOrUndefined(organizationId)) {
                        Xrm.WebApi.retrieveRecord(MarketingSales.EntityNames.Organization, organizationId, ClientUtility.ODataUtil.getSelectOption(["name", "_basecurrencyid_value"])).then((function (response) {
                            currencyId = response._basecurrencyid_value;
                            if (successCallback) {
                                successCallback(currencyId);
                            }
                        }), errorCallback);
                    }
                    else {
                        if (successCallback) {
                            successCallback(currencyId);
                        }
                    }
                }
            }), errorCallback);
        };
        ;
        /**
        * Gets the default currency for the current user.
        * Note: This is intended to work around issues with the retrieve record solution.
        * @param successCallback The success callback function.
        * @param errorCallback The error callback function.
        */
        TransactionCurrency.getUserDefaultCurrency = function (successCallback, failureCallback) {
            if (failureCallback === void 0) { failureCallback = function () { }; }
            Xrm.WebApi.online.execute(new ODataContract.RetrieveUserDefaultCurrencyRequest()).then(function (jsonResponse) {
                jsonResponse.json().then(function (response) {
                    successCallback(response);
                }, failureCallback);
            }, failureCallback);
        };
        /**
        * Gets the lookup value for the current user default currency, which also contains the currency name.
        * @param successCallback The success callback function.
        * @param errorCallback The error callback function.
        */
        TransactionCurrency.getUserDefaultCurrencyLookup = function (successCallback, failureCallback) {
            if (failureCallback === void 0) { failureCallback = function () { }; }
            var onGetSuccess = function (executeResponse) {
                if (!ClientUtility.DataUtil.isNullOrUndefined(executeResponse) && !ClientUtility.DataUtil.isNullOrUndefined(executeResponse.transactioncurrencyid)) {
                    Xrm.WebApi.retrieveRecord(MarketingSales.EntityNames.TransactionCurrency, executeResponse.transactioncurrencyid, ClientUtility.ODataUtil.getSelectOption(["currencyname"])).then((function (retrieveResponse) {
                        var defaultCurrencyLookup = {
                            id: retrieveResponse.transactioncurrencyid,
                            name: retrieveResponse.currencyname,
                            entityType: MarketingSales.EntityNames.TransactionCurrency
                        };
                        successCallback(defaultCurrencyLookup);
                    }), failureCallback);
                }
                else {
                    failureCallback();
                }
            };
            TransactionCurrency.getUserDefaultCurrency(onGetSuccess, failureCallback);
        };
        return TransactionCurrency;
    }());
    /**
     * Gets the currency from the given attribute or the default currency for the current user.
     * @param attributeName The currency attribute name.
     * @param successCallback The success callback function.
     * @param errorCallback The error callback function.
     */
    TransactionCurrency.getTransactionCurrency = function (attributeName, successCallback, errorCallback) {
        var transactionCurrencyIdAttribute = Xrm.Page.getAttribute(attributeName);
        var wasTransactionCurrencyRetrieved = false;
        if (!ClientUtility.DataUtil.isNullOrUndefined(transactionCurrencyIdAttribute)) {
            var transactionCurrencyIdAttributeValue = transactionCurrencyIdAttribute.getValue();
            wasTransactionCurrencyRetrieved = !ClientUtility.DataUtil.isNullOrUndefined(transactionCurrencyIdAttributeValue) && transactionCurrencyIdAttributeValue.length > 0;
            if (wasTransactionCurrencyRetrieved) {
                successCallback(transactionCurrencyIdAttributeValue[0].id);
            }
        }
        if (!wasTransactionCurrencyRetrieved) {
            var currencyId = null;
            if (!ClientUtility.DataUtil.isNullOrUndefined(Xrm.Utility.getGlobalContext()) && !ClientUtility.DataUtil.isNullOrUndefined(Xrm.Utility.getGlobalContext().userSettings)) {
                currencyId = Xrm.Utility.getGlobalContext().userSettings.transactionCurrencyId;
            }
            if (ClientUtility.DataUtil.isNullOrUndefined(currencyId) && !ClientUtility.DataUtil.isNullOrUndefined(Xrm.Utility.getGlobalContext()) && !ClientUtility.DataUtil.isNullOrUndefined(Xrm.Utility.getGlobalContext().organizationSettings)) {
                currencyId = Xrm.Utility.getGlobalContext().organizationSettings.baseCurrencyId;
            }
            if (!ClientUtility.DataUtil.isNullOrUndefined(currencyId)) {
                successCallback(currencyId);
            }
            else {
                var userId = Xrm.Page.context.getUserId();
                MarketingSales.TransactionCurrency.getDefaultTransactionCurrency(userId, successCallback, errorCallback);
            }
        }
    };
    MarketingSales.TransactionCurrency = TransactionCurrency;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="EntityNames.ts" />
/// <reference path="../ClientCommon/MetadataDrivenDialogConstants.ts" />
/// <reference path="../Lead/Reporter.ts" />
/// <reference path="../../../TypeDefinitions/CRM/SIClientUtilityLogger/SIClientUtilityLogger.d.ts" />
var MarketingSales;
(function (MarketingSales) {
    var SalesAgentUsageClient = (function () {
        function SalesAgentUsageClient() {
        }
        SalesAgentUsageClient.logSalesAgentUsage = function (agentName, eventJson) {
            return __awaiter(this, void 0, void 0, function () {
                var query, response, newGuid, recordId, error_6;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            if (!this.isEnableLoggingAgentUsageEnabled()) return [3 /*break*/, 8];
                            _a.label = 1;
                        case 1:
                            _a.trys.push([1, 7, , 8]);
                            query = "?$select=msdyn_salesagentusageid&$filter=msdyn_agentname eq '" + agentName + "'";
                            return [4 /*yield*/, Xrm.WebApi.retrieveMultipleRecords(MarketingSales.EntityNames.AgentUsage, query)];
                        case 2:
                            response = _a.sent();
                            newGuid = ClientUtility.Guid.create(MarketingSales.EntityNames.AgentUsage);
                            if (!(response.entities.length > 0)) return [3 /*break*/, 4];
                            recordId = response.entities[0]["msdyn_salesagentusageid"];
                            return [4 /*yield*/, this.updateSalesAgentUsage(recordId, new Date().toISOString(), newGuid, eventJson)];
                        case 3:
                            _a.sent();
                            return [3 /*break*/, 6];
                        case 4: return [4 /*yield*/, this.createSalesAgentUsage(agentName, new Date().toISOString(), newGuid, eventJson)];
                        case 5:
                            _a.sent();
                            _a.label = 6;
                        case 6:
                            Xrm.Reporting.reportSuccess(MarketingSales.MetadataDrivenDialogConstants.marketingSalesModule + "." + "logSalesAgentUsage");
                            return [3 /*break*/, 8];
                        case 7:
                            error_6 = _a.sent();
                            Xrm.Reporting.reportFailure(MarketingSales.MetadataDrivenDialogConstants.marketingSalesModule + "." + "logSalesAgentUsage", error_6);
                            return [3 /*break*/, 8];
                        case 8: return [2 /*return*/];
                    }
                });
            });
        };
        SalesAgentUsageClient.createSalesAgentUsage = function (agentName, eventDate, eventId, eventJson) {
            return __awaiter(this, void 0, void 0, function () {
                var newRecord, error_7;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            newRecord = {
                                msdyn_agentname: agentName,
                                msdyn_eventdate: eventDate,
                                msdyn_eventid: eventId,
                                msdyn_eventjson: eventJson
                            };
                            return [4 /*yield*/, Xrm.WebApi.createRecord(MarketingSales.EntityNames.AgentUsage, newRecord)];
                        case 1:
                            _a.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            error_7 = _a.sent();
                            Xrm.Reporting.reportFailure(MarketingSales.MetadataDrivenDialogConstants.marketingSalesModule + "." + "logSalesAgentUsage", error_7);
                            return [3 /*break*/, 3];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        };
        SalesAgentUsageClient.updateSalesAgentUsage = function (recordId, eventDate, eventId, eventJson) {
            return __awaiter(this, void 0, void 0, function () {
                var updatedFields, error_8;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            updatedFields = {
                                msdyn_eventdate: eventDate,
                                msdyn_eventid: eventId,
                                msdyn_eventjson: eventJson
                            };
                            return [4 /*yield*/, Xrm.WebApi.updateRecord(MarketingSales.EntityNames.AgentUsage, recordId, updatedFields)];
                        case 1:
                            _a.sent();
                            return [3 /*break*/, 3];
                        case 2:
                            error_8 = _a.sent();
                            Xrm.Reporting.reportFailure(MarketingSales.MetadataDrivenDialogConstants.marketingSalesModule + "." + "logSalesAgentUsage", error_8);
                            return [3 /*break*/, 3];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        };
        SalesAgentUsageClient.isEnableLoggingAgentUsageEnabled = function () {
            var isLoggingAgentUsageEnabled = MarketingSales.LeadUtil.getFeatureControlSetting(this.FCS_QualifyLeadNamespace, this.FCS_EnableLoggingAgentUsage, false);
            return isLoggingAgentUsageEnabled;
        };
        return SalesAgentUsageClient;
    }());
    SalesAgentUsageClient.FCS_EnableLoggingAgentUsage = "EnableLoggingAgentUsage";
    SalesAgentUsageClient.FCS_QualifyLeadNamespace = "DynamicsCRMExtensions.QualifyLead";
    MarketingSales.SalesAgentUsageClient = SalesAgentUsageClient;
})(MarketingSales || (MarketingSales = {}));
/// <reference path="../../../TypeDefinitions/CRM/ClientUtility.d.ts" />
/// <reference path="../../../TypeDefinitions/MarketingSales/Localization/ResourceStringProvider.d.ts" />
var MarketingSales;
(function (MarketingSales) {
    // ToDo: build from metadata
    var LeadStateCode;
    (function (LeadStateCode) {
        LeadStateCode[LeadStateCode["New"] = 0] = "New";
        LeadStateCode[LeadStateCode["Qualified"] = 1] = "Qualified";
        LeadStateCode[LeadStateCode["Disqualified"] = 2] = "Disqualified";
    })(LeadStateCode || (LeadStateCode = {}));
    var EntityStatusCode = (function () {
        function EntityStatusCode(text, value) {
            this.StatusCodeLabelText = text;
            this.StatusCodeLabelValue = value;
        }
        return EntityStatusCode;
    }());
    var EntityStateCode = (function () {
        function EntityStateCode(text, value) {
            this.EntityStatusCodes = [];
            this.StateCodeLabelText = text;
            this.StateCode = value;
        }
        return EntityStateCode;
    }());
    var PopulateDynamicMenu = (function () {
        function PopulateDynamicMenu() {
        }
        PopulateDynamicMenu.getData = function () {
            if (PopulateDynamicMenu._data) {
                return PopulateDynamicMenu._data;
            }
            PopulateDynamicMenu._data = PopulateDynamicMenu.retrieveStatusCodesForDynamicMenuGeneration();
            return PopulateDynamicMenu._data;
        };
        PopulateDynamicMenu.retrieveStatusCodesForDynamicMenuGeneration = function () {
            var deferred;
            if (ClientUtility.ClientUtil.isMobileOffline()) {
                deferred = Xrm.Utility.getEntityMetadata("lead", ["statecode", "statuscode"]).then(function (entityMetadata) {
                    return PopulateDynamicMenu.parseCustomStatusCodes(entityMetadata);
                }, (function (error) {
                    Xrm.Reporting.reportFailure("MarketingSales.PopulateDynamicMenu.retrieveStatusCodesForDynamicMenuGeneration in Offline", error, null);
                    PopulateDynamicMenu._data = null;
                    return Promise.reject(error);
                }));
            }
            else {
                var request = new ODataContract.RetrieveEntityStatusCodeRequest("lead", -1);
                deferred = Xrm.WebApi.online.execute(request)
                    .then(function (res) {
                        return res.json();
                    }).then(function (odataResponse) {
                        var menuContent = JSON.parse(odataResponse.Result);
                        return menuContent;
                    }, function (e) {
                        Xrm.Reporting.reportFailure("MarketingSales.PopulateDynamicMenu.retrieveStatusCodesForDynamicMenuGeneration", e, null);
                        PopulateDynamicMenu._data = null;
                        return Promise.reject(e);
                    });
            }
            return deferred;
        };
        PopulateDynamicMenu.parseCustomStatusCodes = function (entityMetadata) {
            var entityStateCodes = [];
            try {
                if (!ClientUtility.DataUtil.isNullOrUndefined(entityMetadata) && !ClientUtility.DataUtil.isNullOrUndefined(entityMetadata.Attributes)) {
                    var stateAttributeMetadata = entityMetadata.Attributes.get("statecode");
                    var statusAttributeMetadata = entityMetadata.Attributes.get("statuscode");
                    if (!ClientUtility.DataUtil.isNullOrUndefined(stateAttributeMetadata) && !ClientUtility.DataUtil.isNullOrUndefined(statusAttributeMetadata)) {
                        var stateValues = !ClientUtility.DataUtil.isNullOrUndefined(stateAttributeMetadata.attributeDescriptor) && !ClientUtility.DataUtil.isNullOrUndefined(stateAttributeMetadata.attributeDescriptor.OptionSet) ? stateAttributeMetadata.attributeDescriptor.OptionSet : [];
                        var statusValues = !ClientUtility.DataUtil.isNullOrUndefined(statusAttributeMetadata.attributeDescriptor) && !ClientUtility.DataUtil.isNullOrUndefined(statusAttributeMetadata.attributeDescriptor.OptionSet) ? statusAttributeMetadata.attributeDescriptor.OptionSet : [];
                        for (var i = 0; i < stateValues.length; i++) {
                            entityStateCodes.push(new EntityStateCode(stateValues[i].Label, stateValues[i].Value));
                        }
                        for (var i = 0; i < statusValues.length; i++) {
                            var stateCode = statusValues[i].State;
                            entityStateCodes[stateCode].EntityStatusCodes.push(new EntityStatusCode(statusValues[i].Label, statusValues[i].Value));
                        }
                    }
                }
            }
            catch (e) {
                Xrm.Reporting.reportFailure("MarketingSales.PopulateDynamicMenu.parseCustomStatusCodes in Offline", e, null);
            }
            return entityStateCodes;
        };
        return PopulateDynamicMenu;
    }());
    PopulateDynamicMenu.menuLabelIds = [
        'Lead_Disqualify_Lost_ToolTipDescription',
        'Lead_Disqualify_Lost_LabelText',
        'Lead_Disqualify_CannotContact_ToolTipDescription',
        'Lead_Disqualify_CannotContact_LabelText',
        'Lead_Disqualify_NoLongerInterested_ToolTipDescription',
        'Lead_Disqualify_NoLongerInterested_LabelText',
        'Lead_Disqualify_Canceled_ToolTipDescription',
        'Lead_Disqualify_Canceled_LabelText'
    ];
    ///used in telemetry logic if lead disqualify status are customized by the user.
    PopulateDynamicMenu.defaultleadDisqualifyStatusCodes = [
        4,
        5,
        6,
        7,
    ];
    PopulateDynamicMenu.generateMenuXml = function (menuContent, leadStateCode, isRibbon) {
        var dynamicMenuXml;
        if (!ClientUtility.DataUtil.isNullOrUndefined(menuContent) && menuContent.length > 0) {
            var leadStatusForSelectedStateCode = void 0;
            for (var i = 0; i < menuContent.length; i++) {
                if (menuContent[i].StateCode == leadStateCode) {
                    leadStatusForSelectedStateCode = menuContent[i];
                    break;
                }
            }
            if (!ClientUtility.DataUtil.isNullOrUndefined(leadStatusForSelectedStateCode)) {
                switch (leadStatusForSelectedStateCode.StateCode) {
                    //open
                    case 0:
                        //NOT IMPLEMENTED FOR OPEN STATUS
                        break;
                    //Qualified
                    case 1:
                        dynamicMenuXml = PopulateDynamicMenu.generateDynamicMenuForLeadStatus(leadStatusForSelectedStateCode, isRibbon, "QualifyLeadAs", "QualifyAs");
                        break;
                    //Disqualify
                    case 2:
                        dynamicMenuXml = PopulateDynamicMenu.generateDynamicMenuForLeadStatus(leadStatusForSelectedStateCode, isRibbon, "DisqualifyLeadAs", "DisqualifyAs");
                        break;
                }
            }
        }
        return dynamicMenuXml;
    };
    PopulateDynamicMenu.generateDynamicMenuForLeadStatus = function (statuscode, isRibbon, buttonId, command) {
        var buttons = [];
        var global = window.parent;
        var isCrmEncodeDecodeLibraryAvailable = !ClientUtility.DataUtil.isNullOrUndefined(global) && !ClientUtility.DataUtil.isNullOrUndefined(global.CrmEncodeDecode);
        var leadStateCodeMenuMetadata = statuscode.EntityStatusCodes;
        for (var i = 0; i < leadStateCodeMenuMetadata.length; i++) {
            var sequenceId = (i + 1) * 10;
            var nodeTextValue = leadStateCodeMenuMetadata[i].StatusCodeLabelText;
            if (isCrmEncodeDecodeLibraryAvailable) {
                nodeTextValue = global.CrmEncodeDecode.CrmXmlAttributeEncode(nodeTextValue);
            }
            //ribbon
            if (isRibbon) {
                buttons.push("<Button Id=\"lead|NoRelationship|Form|Mscrm.Form.lead.DM." + buttonId + "." + buttonId + ".Controls." + leadStateCodeMenuMetadata[i].StatusCodeLabelValue + "|" + leadStateCodeMenuMetadata[i].StatusCodeLabelValue + "\"");
                buttons.push(" Command=\"lead|NoRelationship|Global|Mscrm.Form.lead." + command + "\"");
                buttons.push(" Sequence=\"" + sequenceId + "\"");
                buttons.push(" ToolTipDescription=\"" + nodeTextValue + "\"");
                buttons.push(" Alt=\"" + nodeTextValue + "\"");
                buttons.push(" LabelText=\"" + nodeTextValue + "\"");
                buttons.push(" ModernImage=\" \" />");
            }
            else {
                buttons.push("<Button Id=\"lead | NoRelationship | Form | Mscrm.Grid.lead.DM." + buttonId + "." + buttonId + ".Controls." + leadStateCodeMenuMetadata[i].StatusCodeLabelValue + "|" + leadStateCodeMenuMetadata[i].StatusCodeLabelValue + "\"");
                buttons.push(" Command=\"lead|NoRelationship|Form|Mscrm.Grid.lead." + command + "\"");
                buttons.push(" Sequence=\"" + sequenceId + "\"");
                buttons.push(" ToolTipDescription=\"" + nodeTextValue + "\"");
                buttons.push(" Alt=\"" + nodeTextValue + "\"");
                buttons.push(" LabelText=\"" + nodeTextValue + "\"");
                buttons.push(" ModernImage=\" \" />");
            }
        }
        var menuXml;
        if (isRibbon) {
            menuXml = [
                "<Menu Id=\"lead|NoRelationship|Global|Mscrm.Form.lead.DM." + buttonId + "\">",
                "<MenuSection Id=\"lead|NoRelationship|Form|Mscrm.Form.lead.DM." + buttonId + "." + buttonId + "\" Sequence=\"10\" DisplayMode=\"Menu16\">",
                "<Controls Id=\"lead|NoRelationship|Form|Mscrm.Form.lead.DM." + buttonId + "." + buttonId + ".Controls\">"
            ];
        }
        else {
            menuXml = [
                "<Menu Id=\"lead|NoRelationship|Form|Mscrm.Grid.lead.DM." + buttonId + "\">",
                "<MenuSection Id=\"lead|NoRelationship|Form|Mscrm.Grid.lead.DM." + buttonId + "." + buttonId + "\" Sequence=\"10\" DisplayMode=\"Menu16\">",
                "<Controls Id=\"lead|NoRelationship|Form|Mscrm.Grid.lead.DM." + buttonId + "." + buttonId + ".Controls\">"
            ];
        }
        menuXml = menuXml.concat(buttons);
        var menuCloseXml = [
            '</Controls>',
            '</MenuSection>',
            '</Menu>'
        ];
        return menuXml.concat(menuCloseXml).join('');
    };
    PopulateDynamicMenu.getQualifyMenu = function (commandProperties) {
        var menu = PopulateDynamicMenu.getData().then(function (menuContent) {
            try {
                var menuXml = PopulateDynamicMenu.generateMenuXml(menuContent, LeadStateCode.Qualified, true);
                if (ClientUtility.DataUtil.isNullOrUndefined(menuXml))
                    menuXml = PopulateDynamicMenu.getDefaultFormQualifyMenuXML();
                return menuXml;
            }
            catch (e) {
                Xrm.Reporting.reportFailure("MarketingSales.PopulateDynamicMenu.getQualifyMenu", e, null);
                return PopulateDynamicMenu.getDefaultFormQualifyMenuXML();
            }
        }, function (e) {
            return PopulateDynamicMenu.getDefaultFormQualifyMenuXML();
        });
        commandProperties["PopulationXML"] = menu;
    };
    PopulateDynamicMenu.getDefaultFormQualifyMenuXML = function () {
        return [
            '<Menu Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.QualifyLeadAs">',
            '<MenuSection Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.QualifyLeadAs.QualifyLeadAs" Sequence="10" DisplayMode="Menu16">',
            '<Controls Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.QualifyLeadAs.QualifyLeadAs.Controls">',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.QualifyLeadAs.QualifyLeadAs.Controls.3|3"',
            ' Command="lead|NoRelationship|Form|Mscrm.Form.lead.QualifyAs"',
            ' Sequence="10"',
            ' ToolTipDescription="Qualified"',
            ' Alt="Qualified"',
            ' LabelText="Qualified"',
            ' ModernImage=" " />',
            '</Controls>',
            '</MenuSection>',
            '</Menu>'
        ].join('');
    };
    PopulateDynamicMenu.getQualifyMenuGrid = function (commandProperties) {
        var menu = PopulateDynamicMenu.getData().then(function (menuContent) {
            try {
                var menuXml = PopulateDynamicMenu.generateMenuXml(menuContent, LeadStateCode.Qualified, false);
                if (ClientUtility.DataUtil.isNullOrUndefined(menuXml))
                    menuXml = PopulateDynamicMenu.getDefaultGridQualifyMenuXML();
                return menuXml;
            }
            catch (e) {
                Xrm.Reporting.reportFailure("MarketingSales.PopulateDynamicMenu.getQualifyMenuGrid", e, null);
                return PopulateDynamicMenu.getDefaultGridQualifyMenuXML();
            }
        }, function (e) {
            return PopulateDynamicMenu.getDefaultGridQualifyMenuXML();
        });
        commandProperties["PopulationXML"] = menu;
    };
    PopulateDynamicMenu.getDefaultGridQualifyMenuXML = function () {
        return [
            '<Menu Id="lead|NoRelationship|Form|Mscrm.HomepageGrid.lead.DM.QualifyLeadAs">',
            '<MenuSection Id="lead|NoRelationship|Form|Mscrm.HomepageGrid.lead.DM.QualifyLeadAs.QualifyLeadAs" Sequence="10" DisplayMode="Menu16">',
            '<Controls Id="lead|NoRelationship|Form|Mscrm.HomepageGrid.lead.DM.QualifyLeadAs.QualifyLeadAs.Controls">',
            '<Button Id="lead|NoRelationship|Form|Mscrm.HomepageGrid.lead.DM.QualifyLeadAs.QualifyLeadAs.Controls.3|3"',
            ' Command="lead|NoRelationship|Form|Mscrm.Grid.lead.QualifyAs"',
            ' Sequence="10"',
            ' ToolTipDescription="Qualified"',
            ' Alt="Qualified"',
            ' LabelText="Qualified"',
            ' ModernImage=" " />',
            '</Controls>',
            '</MenuSection>',
            '</Menu>'
        ].join('');
    };
    PopulateDynamicMenu.getDisqualifyMenu = function (commandProperties) {
        var menu = PopulateDynamicMenu.getData().then(function (menuContent) {
            try {
                var menuXml = PopulateDynamicMenu.generateMenuXml(menuContent, LeadStateCode.Disqualified, true);
                if (ClientUtility.DataUtil.isNullOrUndefined(menuXml))
                    menuXml = PopulateDynamicMenu.getDefaultFormDisqualifyMenuXML();
                return menuXml;
            }
            catch (e) {
                Xrm.Reporting.reportFailure("MarketingSales.PopulateDynamicMenu.getDisqualifyMenu", e, null);
                return PopulateDynamicMenu.getDefaultFormDisqualifyMenuXML();
            }
        }, function (e) {
            return PopulateDynamicMenu.getDefaultFormDisqualifyMenuXML();
        });
        // ToDo: remove. this is a workaround for 'Dynamic menu ribbon rules are not evaulated correctly after create' (work item 508309)
        var pageId = Xrm.Page.ui.pageId;
        var applicationContext = !ClientUtility.DataUtil.isNullOrUndefined(pageId) ? Xrm.Page.applicationContext : null;
        var state = applicationContext && applicationContext.Router && applicationContext.Router._store ? applicationContext.Router._store.getState() : null;
        var ruleContext = state && state.commands && state.commands[pageId] ? state.commands[pageId].ruleContext : null;
        if (ruleContext && !ruleContext.EntityId) {
            ruleContext.EntityId = { guid: ClientUtility.Guid.create(Xrm.Page.data.entity.getId()) };
        }
        // end ToDo
        commandProperties["PopulationXML"] = menu;
    };
    PopulateDynamicMenu.getDefaultFormDisqualifyMenuXML = function () {
        var labels = PopulateDynamicMenu.getLabels(PopulateDynamicMenu.menuLabelIds);
        return [
            '<Menu Id="lead|NoRelationship|Global|Mscrm.Form.lead.DM.DisqualifyLeadAs">',
            '<MenuSection Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs" Sequence="10" DisplayMode="Menu16">',
            '<Controls Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls">',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.4|4"',
            ' Command="lead|NoRelationship|Global|Mscrm.Form.lead.DisqualifyAs"',
            ' Sequence="10"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_Lost_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_Lost_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_Lost_LabelText'] + "\"",
            ' ModernImage=" " />',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.5|5"',
            ' Command="lead|NoRelationship|Global|Mscrm.Form.lead.DisqualifyAs"',
            ' Sequence="20"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_CannotContact_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_CannotContact_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_CannotContact_LabelText'] + "\"",
            ' ModernImage=" " />',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.6|6"',
            ' Command="lead|NoRelationship|Global|Mscrm.Form.lead.DisqualifyAs"',
            ' Sequence="30"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_NoLongerInterested_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_NoLongerInterested_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_NoLongerInterested_LabelText'] + "\"",
            ' ModernImage=" " />',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Form.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.7|7"',
            ' Command="lead|NoRelationship|Global|Mscrm.Form.lead.DisqualifyAs"',
            ' Sequence="40"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_Canceled_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_Canceled_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_Canceled_LabelText'] + "\"",
            ' ModernImage=" " />',
            '</Controls>',
            '</MenuSection>',
            '</Menu>'
        ].join('');
    };
    PopulateDynamicMenu.getDisqualifyMenuGrid = function (commandProperties) {
        var menu = PopulateDynamicMenu.getData().then(function (menuContent) {
            try {
                var menuXml = PopulateDynamicMenu.generateMenuXml(menuContent, LeadStateCode.Disqualified, false);
                if (ClientUtility.DataUtil.isNullOrUndefined(menuXml))
                    menuXml = PopulateDynamicMenu.getDefaultGridDisqualifyMenuXML();
                return menuXml;
            }
            catch (e) {
                Xrm.Reporting.reportFailure("MarketingSales.PopulateDynamicMenu.getDisqualifyMenuGrid", e, null);
                return PopulateDynamicMenu.getDefaultGridDisqualifyMenuXML();
            }
        }, function (e) {
            return PopulateDynamicMenu.getDefaultGridDisqualifyMenuXML();
        });
        commandProperties["PopulationXML"] = menu;
    };
    PopulateDynamicMenu.getDefaultGridDisqualifyMenuXML = function () {
        var labels = PopulateDynamicMenu.getLabels(PopulateDynamicMenu.menuLabelIds);
        return [
            '<Menu Id="lead|NoRelationship|Form|Mscrm.Grid.lead.DM.DisqualifyLeadAs">',
            '<MenuSection Id="lead|NoRelationship|Form|Mscrm.Grid.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs" Sequence="10" DisplayMode="Menu16">',
            '<Controls Id="lead|NoRelationship|Form|Mscrm.Grid.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls">',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Grid.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.4|4"',
            ' Command="lead|NoRelationship|Form|Mscrm.Grid.lead.DisqualifyAs"',
            ' Sequence="10"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_Lost_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_Lost_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_Lost_LabelText'] + "\"",
            ' ModernImage=" " />',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Grid.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.5|5"',
            ' Command="lead|NoRelationship|Form|Mscrm.Grid.lead.DisqualifyAs"',
            ' Sequence="20"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_CannotContact_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_CannotContact_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_CannotContact_LabelText'] + "\"",
            ' ModernImage=" " />',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Grid.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.6|6"',
            ' Command="lead|NoRelationship|Form|Mscrm.Grid.lead.DisqualifyAs"',
            ' Sequence="30"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_NoLongerInterested_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_NoLongerInterested_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_NoLongerInterested_LabelText'] + "\"",
            ' ModernImage=" " />',
            '<Button Id="lead|NoRelationship|Form|Mscrm.Grid.lead.DM.DisqualifyLeadAs.DisqualifyLeadAs.Controls.7|7"',
            ' Command="lead|NoRelationship|Form|Mscrm.Grid.lead.DisqualifyAs"',
            ' Sequence="40"',
            " ToolTipDescription=\"" + labels['Lead_Disqualify_Canceled_ToolTipDescription'] + "\"",
            " Alt=\"" + labels['Lead_Disqualify_Canceled_LabelText'] + "\"",
            " LabelText=\"" + labels['Lead_Disqualify_Canceled_LabelText'] + "\"",
            ' ModernImage=" " />',
            '</Controls>',
            '</MenuSection>',
            '</Menu>'
        ].join('');
    };
    PopulateDynamicMenu.getLabels = function (labelIds) {
        var labels = {};
        if (labelIds && labelIds.length > 0) {
            for (var i = 0; i < labelIds.length; i++) {
                labels[labelIds[i]] = MarketingSales.ResourceStringProvider.getResourceString(labelIds[i]);
            }
        }
        return labels;
    };
    PopulateDynamicMenu.AppCommonNullCheck = function () {
        return ((typeof (AppCommon) !== 'undefined') &&
            (AppCommon != null && AppCommon != undefined) &&
            (AppCommon.TelemetryReporter != null && AppCommon.TelemetryReporter != undefined) &&
            (AppCommon.TelemetryReporter.Instance() != null && AppCommon.TelemetryReporter.Instance() != undefined));
    };
    MarketingSales.PopulateDynamicMenu = PopulateDynamicMenu;
})(MarketingSales || (MarketingSales = {}));
/// <reference path="../../../TypeDefinitions/CRM/ClientUtility.d.ts" />
var MarketingSales;
(function (MarketingSales) {
    var RetrieveEntityDefinitions = (function () {
        /**
         * constructor for ODataRetrieveEntityDefinitions
         * @param filter filter for the entity definition e.g.$filter=ObjectTypeCode eq 3
         * @param columns list of columns to be retrieved
         */
        function RetrieveEntityDefinitions(filter, columns) {
            this.filter = filter;
            this.columns = columns;
        }
        RetrieveEntityDefinitions.prototype.getMetadata = function () {
            return {
                boundParameter: undefined,
                parameterTypes: {},
                operationName: "EntityDefinitions",
                operationType: 2,
            };
        };
        return RetrieveEntityDefinitions;
    }());
    MarketingSales.RetrieveEntityDefinitions = RetrieveEntityDefinitions;
    var LeadErrorHandling = (function () {
        function LeadErrorHandling() {
        }
        return LeadErrorHandling;
    }());
    LeadErrorHandling.component = "LeadErrorHandling";
    LeadErrorHandling.actionFailedErrorDialog = function (errorResponse) {
        if (errorResponse && errorResponse.errorCode && LeadErrorHandling.ErrorCodesWithHelpLinkAvaliable(errorResponse.errorCode)) {
            errorResponse.errorRegion = "Sales";
        }
        if (MarketingSales.LeadUtil.getIsUci() && Xrm.Internal.isFeatureEnabled("EnhancedErrorDialog") && Xrm.Internal.isFeatureEnabled("October2020Update")) {
            LeadErrorHandling.addStepsInfo(errorResponse).then(function () {
                ClientUtility.ActionFailedHandler.actionFailedErrorDialog(errorResponse);
            });
        }
        else {
            ClientUtility.ActionFailedHandler.actionFailedErrorDialog(errorResponse);
        }
    };
    LeadErrorHandling.actionFailedCallback = function (errorResponse) {
        if (errorResponse && errorResponse.errorCode && LeadErrorHandling.ErrorCodesWithHelpLinkAvaliable(errorResponse.errorCode)) {
            errorResponse.errorRegion = "Sales";
        }
        ClientUtility.ActionFailedHandler.actionFailedCallback(errorResponse);
    };
    LeadErrorHandling.ErrorCodesWithHelpLinkAvaliable = function (errorCode) {
        //Todo: Add here all errocodes which are captured in our client code whose help link is avaialiable or need to be updated
        switch (errorCode) {
            case 2147746326: //0X80040216 -> UnExpected
            case 2148532226: //0X80100002 -> ActiveStageIsNotOnLeadEntity
            case 2147877914: //0x8006041A -> RequiredProcessStepIsNull
            case 2147746336: //0x80040220 -> PrivilegeDenied
            case 2147746405: //0x80040265 -> IsvAborted
            case 2147746340: //0x80040224 -> IsvUnexpected
            case 2147747097:
                return true;
            default:
                return false;
        }
    };
    LeadErrorHandling.addStepsInfo = function (errorResponse) {
        return new Promise(function (resolve, reject) {
            var columnList = ["LogicalName", "DisplayName"];
            var filter = "$filter=LogicalName eq 'lead' or LogicalName eq 'account' or LogicalName eq 'contact' " +
                "or LogicalName eq 'opportunity' or LogicalName eq 'competitor'";
            var fetchEntityReq = new RetrieveEntityDefinitions(filter, columnList);
            Xrm.WebApi.online.execute(fetchEntityReq).then(function (response) {
                response.json().then(function (jsonResponse) {
                    try {
                        var responseList = jsonResponse["value"];
                        var logicalNameMap = {};
                        for (var i = 0; i < responseList.length; i++) {
                            var logicalName = responseList[i]["LogicalName"];
                            var displayName = responseList[i]["DisplayName"]["UserLocalizedLabel"]["Label"];
                            logicalNameMap[logicalName] = displayName;
                        }
                        var parsedErrorResponse = errorResponse && errorResponse.raw && JSON.parse(errorResponse.raw);
                        var parsedAnnotations = parsedErrorResponse && parsedErrorResponse._errorFault && parsedErrorResponse._errorFault._annotations;
                        var parsedSteps = parsedAnnotations && parsedAnnotations["@Microsoft.PowerApps.CDS.ErrorDetails.StepErrorDetails"] && JSON.parse(parsedAnnotations["@Microsoft.PowerApps.CDS.ErrorDetails.StepErrorDetails"]);
                        var localizedSteps = [];
                        for (var i = 0; parsedSteps && i < parsedSteps.length; i++) {
                            var message = MarketingSales.ResourceStringProvider.getResourceString(parsedSteps[i]['id']);
                            var parsedMessage = LeadErrorHandling.parseMessage(message, logicalNameMap);
                            var status = parsedSteps[i]['status'];
                            localizedSteps.push({
                                message: parsedMessage,
                                status: status
                            });
                        }
                        errorResponse.steps = localizedSteps;
                    }
                    catch (error) {
                        Xrm.Reporting.reportFailure(LeadErrorHandling.component, error);
                    }
                    resolve();
                }, function (error) {
                    Xrm.Reporting.reportFailure(LeadErrorHandling.component, error);
                    resolve();
                });
            }, function (error) {
                Xrm.Reporting.reportFailure(LeadErrorHandling.component, error);
                resolve();
            });
        });
    };
    // replaces the placehodler text {lead} with mapping present in logicalNameMap
    LeadErrorHandling.parseMessage = function (rawMessage, logicalNameMap) {
        var matches = rawMessage.match(/\{(.*?)\}/);
        if (matches) {
            var submatch = matches[0];
            var name = matches[1];
            rawMessage = rawMessage.replace(submatch, logicalNameMap[name]);
        }
        return rawMessage;
    };
    MarketingSales.LeadErrorHandling = LeadErrorHandling;
})(MarketingSales || (MarketingSales = {}));
var MarketingSales;
(function (MarketingSales) {
    /* tslint:disable:crm-force-fields-private */
    var TargetFieldType;
    (function (TargetFieldType) {
        TargetFieldType[TargetFieldType["All"] = 0] = "All";
        TargetFieldType[TargetFieldType["ValidForCreate"] = 1] = "ValidForCreate";
        TargetFieldType[TargetFieldType["ValidForUpdate"] = 2] = "ValidForUpdate";
        TargetFieldType[TargetFieldType["ValidForRead"] = 3] = "ValidForRead";
    })(TargetFieldType = MarketingSales.TargetFieldType || (MarketingSales.TargetFieldType = {}));
    var GetAssociatedEntityRequest = (function () {
        function GetAssociatedEntityRequest(entityMoniker /*Microsoft.Dynamics.CRM.crmbaseentity*/, targetEntityName, targetFieldType) {
            this.EntityMoniker = entityMoniker;
            this.TargetEntityName = targetEntityName;
            this.TargetFieldType = targetFieldType;
        }
        GetAssociatedEntityRequest.prototype.getMetadata = function () {
            var metadata = {
                boundParameter: null,
                parameterTypes: {
                    "EntityMoniker": {
                        "typeName": "Microsoft.Dynamics.CRM.crmbaseentity",
                        "structuralProperty": 5,
                    },
                    "TargetEntityName": {
                        "typeName": "Edm.String",
                        "structuralProperty": 1,
                    },
                    "TargetFieldType": {
                        "typeName": "Microsoft.Dynamics.CRM.TargetFieldType",
                        "structuralProperty": 3,
                        "enumProperties": [,
                            {
                                "name": "All",
                                "value": 0,
                            },
                            {
                                "name": "ValidForCreate",
                                "value": 1,
                            },
                            {
                                "name": "ValidForUpdate",
                                "value": 2,
                            },
                            {
                                "name": "ValidForRead",
                                "value": 3,
                            },
                        ],
                    },
                },
                operationName: "InitializeFrom",
                operationType: 1,
            };
            return metadata;
        };
        return GetAssociatedEntityRequest;
    }());
    MarketingSales.GetAssociatedEntityRequest = GetAssociatedEntityRequest;
})(MarketingSales || (MarketingSales = {}));
/**
 * @license Copyright (c) Microsoft Corporation. All rights reserved.
 */
/**
 * IMPORTANT!
 * DO NOT MAKE CHANGES TO THIS FILE - THIS FILE IS AUTO-GENERATED FROM ODATA CSDL METADATA DOCUMENT
 * SEE https://msdn.microsoft.com/en-us/library/mt607990.aspx FOR MORE INFORMATION
 */
/**
 * IMPORTANT!
 * THIS FILE HAS BEEN MANUALLY MODIFIED TO ENABLE OPTIONAL PARAMETERS FOR THE REQUEST.
 * OPTIONAL PARAMETERS SHOULD BE IGNORED IN ORDER TO AVOID SERVER-SIDE DESERIALIZATION ISSUES.
 */
var MarketingSales;
(function (MarketingSales) {
    var GetEntityWiseDuplicatesCountRequest = (function () {
        function GetEntityWiseDuplicatesCountRequest(entity /*Microsoft.Dynamics.CRM.crmbaseentity*/) {
            this.Entity = entity;
        }
        GetEntityWiseDuplicatesCountRequest.prototype.getMetadata = function () {
            var metadata = {
                boundParameter: null,
                parameterTypes: {
                    Entity: {
                        "typeName": "Microsoft.Dynamics.CRM.crmbaseentity",
                        "structuralProperty": 5,
                    }
                },
                operationName: "GetEntityWiseDuplicatesCount",
                operationType: 1,
            };
            return metadata;
        };
        return GetEntityWiseDuplicatesCountRequest;
    }());
    MarketingSales.GetEntityWiseDuplicatesCountRequest = GetEntityWiseDuplicatesCountRequest;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../../../TypeDefinitions/CRM/ClientUtility.d.ts" />
/// <reference path="../../../TypeDefinitions/AppCommon/Telemetry/TelemetryLibrary.d.ts" />
/// <reference path="../../../TypeDefinitions/Sales/ClientCommon/Sales_ClientCommon.d.ts" />
/// <reference path="../../../TypeDefinitions/LeadManagement/Lead/Lead_main_system_library.d.ts" />
/// <reference path="../ClientCommon/CESSurvey/CESSurvey.ts" />
/// <reference path="../ClientCommon/CESSurvey/CESSurveyUtil.ts"/>
/// <reference path="../ClientCommon/MetadataDrivenDialogConstants.ts" />
/// <reference path="../ClientCommon/TransactionCurrency.ts" />
/// <reference path="../ClientCommon/SalesAgentUsageClient.ts" />
/// <reference path="LeadUtil.ts" />
/// <reference path="PopulateDynamicMenu.ts" />
/// <reference path="LeadErrorHandling.ts" />
/// <reference path="Contracts/GetAssociatedEntityRequest.ts" />
/// <reference path="Contracts/GetEntityWiseDuplicatesCountRequest.ts"/>
/// <reference path ="Reporter.ts" />
var MarketingSales;
(function (MarketingSales) {
    var LeadCommandActions = (function () {
        function LeadCommandActions() {
            var leadCommandsActions = new LeadCommandActionsImplementation();
            var global = window;
            var mscrm = global.Mscrm || {};
            mscrm.LeadCommandActions = mscrm.LeadCommandActions || {};
            mscrm.LeadCommandActions.qualifyLeadQuick = leadCommandsActions.qualifyLeadQuick;
            mscrm.LeadCommandActions.qualifyLead = leadCommandsActions.qualifyLead;
            mscrm.LeadCommandActions.qualifyLeadAs = leadCommandsActions.qualifyLeadAs;
            mscrm.LeadCommandActions.disqualifyLeadQuick = leadCommandsActions.disqualifyLeadQuick;
            mscrm.LeadCommandActions.disqualifyLeadAs = leadCommandsActions.disqualifyLeadAs;
            mscrm.LeadCommandActions.reactivateLead = leadCommandsActions.reactivateLead;
            mscrm.LeadCommandActions.getTransactionCurrencyId = leadCommandsActions.getTransactionCurrencyId;
            mscrm.LeadCommandActions.performActionAfterHandleLeadDuplication = leadCommandsActions.performActionAfterHandleLeadDuplication;
            mscrm.LeadCommandActions.dupWarningOnLoadHandler = leadCommandsActions.dupWarningOnLoadHandler;
            mscrm.LeadCommandActions.dupWarningOnOkClickHandler = leadCommandsActions.dupWarningOnOkClickHandler;
            mscrm.LeadCommandActions.closeDupWarningDialogCallback = leadCommandsActions.closeDupWarningDialogCallback;
            mscrm.LeadCommandActions.selectedRecordsChange = leadCommandsActions.selectedRecordsChange;
            mscrm.LeadCommandActions.convertLeadOnQualifyV2ClickHandler = leadCommandsActions.convertLeadOnQualifyV2ClickHandler;
        }
        LeadCommandActions.prototype.dialogClose = function (context) {
            var formContext = ClientUtility.DataUtil.isNullOrUndefined(context) ? Xrm.Page : context.getFormContext();
            ClientUtility.DialogUtil.setLastButtonClicked(ClientUtility.MetadataDrivenDialogConstants.DialogCancelId);
            formContext.ui.close();
        };
        return LeadCommandActions;
    }());
    MarketingSales.LeadCommandActions = LeadCommandActions;
    var LeadCommandActionsImplementation = (function () {
        function LeadCommandActionsImplementation() {
            var _this = this;
            // spec gives the defaults for flags as account=true, contact = true, opty = false on the Qualify Lead MDD.
            this.createContactFlag = true;
            this.createOpportunityFlag = false;
            this.createAccountFlag = true;
            // adminSetting is true implies fallback on the existing qualify lead rather than having new MDD.
            this.adminSettings = true;
            this.isGrid = false;
            this.fcbOctoberUpdate = false;
            this.duplicateDetectionAdminSettings = false;
            this.preventDefaultUCIErrorCode = 2200000005; // 0x83215605
            this.qualifyLeadQuick = function () {
                return __awaiter(_this, void 0, void 0, function () {
                    var _a;
                    return __generator(this, function (_b) {
                        switch (_b.label) {
                            case 0:
                                _a = this;
                                return [4 /*yield*/, MarketingSales.LeadUtil.isQualifyLeadV2Enabled()];
                            case 1:
                                _a.isQualifyLeadV2Enabled = _b.sent();
                                if (this.isQualifyLeadV2Enabled) {
                                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "qualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "FormQualifyButtonClickedV2");
                                    this.qualifyLeadV2();
                                }
                                else {
                                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "qualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "FormQualifyButtonClicked");
                                    this.qualifyLead();
                                }
                                return [2 /*return*/];
                        }
                    });
                });
            };
            this.qualifyLead = function () {
                _this.qualifyLeadAs(-1);
            };
            this.qualifyLeadV2 = function () {
                _this.qualifyLeadAsV2(-1);
            };
            this.qualifyLeadAs = function (parameter) {
                try {
                    _this.fcbOctoberUpdate = MarketingSales.LeadUtil.FCBOctoberUpdateEnabled();
                    _this.adminSettings = MarketingSales.LeadUtil.qualifyLeadAdditionalOptionsAdminSettings();
                    _this.duplicateDetectionAdminSettings = Xrm.Internal.isUci() && MarketingSales.LeadUtil.duplicateDetectionAdminSettings();
                    if (!_this.adminSettings && _this.fcbOctoberUpdate) {
                        MarketingSales.LeadUtil.convertLeadDialog(_this.isGrid, 1).then(function (dialogParams) {
                            if (!ClientUtility.DataUtil.isNullOrUndefined(dialogParams) &&
                                !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters) &&
                                dialogParams.parameters[ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked] === ClientUtility.MetadataDrivenDialogConstants.DialogOkId) {
                                _this.createAccountFlag = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.accountFlag] ? true : false;
                                _this.createContactFlag = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.contactFlag] ? true : false;
                                _this.createOpportunityFlag = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.opportunityFlag] ? true : false;
                                var parameters = {};
                                parameters["CreateContact"] = _this.createContactFlag;
                                parameters["CreateAccount"] = _this.createAccountFlag;
                                parameters["CreateOpportunity"] = _this.createOpportunityFlag;
                                parameters["isGrid"] = _this.isGrid;
                                MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "", parameters);
                                _this.qualifyLeadAsFormAction(parameter);
                            }
                        });
                    }
                    else {
                        _this.qualifyLeadAsFormAction(parameter);
                    }
                }
                catch (e) {
                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
                }
            };
            this.qualifyLeadAsV2 = function (parameter) {
                try {
                    _this.fcbOctoberUpdate = MarketingSales.LeadUtil.FCBOctoberUpdateEnabled();
                    _this.adminSettings = MarketingSales.LeadUtil.qualifyLeadAdditionalOptionsAdminSettings();
                    _this.duplicateDetectionAdminSettings = Xrm.Internal.isUci() && MarketingSales.LeadUtil.duplicateDetectionAdminSettings();
                    _this.qualifyLeadAsFormAction(parameter);
                }
                catch (e) {
                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
                }
            };
            this.qualifyLeadAsFormAction = function (parameter) {
                // parameter is the new status or 'commandproperties' from the ribbon
                if (Xrm.Page.data && !Xrm.Page.data.isValid()) {
                    Xrm.Utility.alertDialog(MarketingSales.ResourceStringProvider.getResourceString("RequiredFieldsNotFilledError"));
                    return;
                }
                var status = _this.getStatusFromParameter(parameter);
                if (ClientUtility.DataUtil.isNullOrUndefined(Xrm.Page.data.entity.getId())) {
                    return;
                }
                if (ClientUtility.ClientUtil.isMobileOffline()) {
                    if (Xrm.Page.data.entity.getIsDirty()) {
                        Xrm.Page.data.save({ saveMode: 16 /* Qualify */ }).then(function () {
                            _this.qualifyLeadAsHelper(status);
                        }, function (errorResponse) { _this.handleError(errorResponse); });
                    }
                    else {
                        _this.qualifyLeadAsHelper(status);
                    }
                }
                else {
                    Xrm.Page.data.save({ saveMode: 16 /* Qualify */ }).then(function () {
                        _this.qualifyLeadAsHelper(status);
                    }, function (errorResponse) { _this.handleError(errorResponse); });
                }
            };
            this.handleError = function (errorResponse) {
                // Absorb error if it due to event PreventDefault.
                if (!ClientUtility.DataUtil.isNullOrUndefined(errorResponse) && errorResponse.errorCode != _this.preventDefaultUCIErrorCode) {
                    MarketingSales.LeadErrorHandling.actionFailedErrorDialog(errorResponse);
                }
            };
            this.disqualifyLeadQuick = function () {
                _this.disqualifyLeadAs(-1);
            };
            this.disqualifyLeadAs = function (parameter) {
                // parameter is the new status or 'commandproperties' from the ribbon
                if (Xrm.Page.data && !Xrm.Page.data.isValid()) {
                    Xrm.Utility.alertDialog(MarketingSales.ResourceStringProvider.getResourceString("RequiredFieldsNotFilledError"));
                    return;
                }
                var status = _this.getStatusFromParameter(parameter);
                var state = LeadManagement.LeadState.Disqualified;
                var updateRecord = {
                    statecode: state,
                    statuscode: status
                };
                var progressIndicator = new ClientUtility.ProgressIndicator();
                progressIndicator.show();
                Xrm.Page.data.save({
                    saveMode: 15 /* Disqualify */
                })
                    .then(function () {
                        var leadId = Xrm.Page.data.entity.getId();
                        Xrm.WebApi.updateRecord(MarketingSales.EntityNames.Lead, leadId, updateRecord)
                            .then(function () {
                                Xrm.Page.data.refresh(false);
                                progressIndicator.hide();
                                if (Xrm.Internal.isUci()) {
                                    var toastLabel = MarketingSales.ResourceStringProvider.getResourceString("ToastMessage_Lead_Disqualify");
                                    _this.showLeadToastMessage(toastLabel);
                                }
                                // Trigger NSAT survey once lead is disqualified
                                MarketingSales.NSATSurvey.getNSATSurvey("lead", MarketingSales.NSATScenarios.DisqualifyLead);
                                // this.logStatusUpdate(leadId, "LeadCommandActions.disqualifyLeadAs", "Disqualify");
                            }, function (errorResponse) {
                                progressIndicator.hide();
                                _this.handleError(errorResponse);
                            });
                    }, function (saveErrorResponse) {
                        progressIndicator.hide();
                        ;
                        _this.handleError(saveErrorResponse);
                    });
            };
            this.reactivateLead = function () {
                var state = LeadManagement.LeadState.Open;
                var leadId = Xrm.Page.data.entity.getId();
                XrmCore.Commands.Common.setState(leadId, MarketingSales.EntityNames.Lead, state, null, null, null, null, 6 /* Reactivate */);
                // Trigger NSAT survey once lead is re-activated is complete
                MarketingSales.NSATSurvey.getNSATSurvey("lead", MarketingSales.NSATScenarios.ReactivateLead);
                // this.logStatusUpdate(leadId, "LeadCommandActions.reactivateLead", "Reactivate");
            };
            this.performActionAfterHandleLeadDuplication = function (returnValue) {
                var qualifyStatus = returnValue.qualifyStatus;
                var transactionCurrencyId = returnValue.transactionCurrencyId;
                var parentAccountId = returnValue.parentAccountId;
                var parentContactId = returnValue.parentContactId;
                var parentAccountName = returnValue.parentAccountName;
                var parentContactName = returnValue.parentContactName;
                _this.executeLeadDuplication(qualifyStatus, transactionCurrencyId, parentAccountId, parentContactId, parentAccountName, parentContactName);
            };
            this.executeLeadDuplication = function (qualifyStatus, transactionCurrencyId, parentAccountId, parentContactId, parentAccountName, parentContactName) {
                // check whether customer want to prevent opportunity duplicate creation on Ignore and save
                if (MarketingSales.LeadUtil.IsFcbEnabled("QualifyLeadDupDetectionFix")) {
                    var objectopportunity = {
                        parentEntity_Id: Xrm.Page.data.entity.getId(),
                        EntityName: "opportunity",
                    };
                    _this.getAssociatedEntityforDuplicate([objectopportunity]).then(function (response) {
                        if (response.length > 0) {
                            var opportunityRecord = response[0];
                            var request = new MarketingSales.GetEntityWiseDuplicatesCountRequest(opportunityRecord);
                            Xrm.WebApi.online.execute(request).then(function (response) {
                                response.json().then(function (res) {
                                    // if user clicked ignore and save for account/contact duplicates but there is duplicate of opportunity then prevent
                                    // that by throwing the alert 
                                    if (res && res.DuplicatesCount && res.DuplicatesCount.length > 0 && res.DuplicatesCount[0] > 0) {
                                        var alertStrings = {
                                            text: MarketingSales.ResourceStringProvider.getResourceString("LeadDuplicateRecordCreationErrorOpportunityDescription"),
                                            title: MarketingSales.ResourceStringProvider.getResourceString("LeadDuplicateRecordCreationErrorOpportunityTitle")
                                        };
                                        var alertOptions = { height: 120, width: 280 };
                                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                                    }
                                    else {
                                        _this.executeLeadDuplicationAction(qualifyStatus, transactionCurrencyId, parentAccountId, parentContactId, parentAccountName, parentContactName);
                                    }
                                });
                            }, function (errorResponse) {
                                _this.handleError(errorResponse);
                                return;
                            });
                        }
                    });
                }
                else {
                    _this.executeLeadDuplicationAction(qualifyStatus, transactionCurrencyId, parentAccountId, parentContactId, parentAccountName, parentContactName);
                }
            };
            this.executeLeadDuplicationAction = function (qualifyStatus, transactionCurrencyId, parentAccountId, parentContactId, parentAccountName, parentContactName) {
                _this.setLookupValue("parentaccountid", parentAccountName, parentAccountId, MarketingSales.EntityNames.Account, false);
                _this.setLookupValue("parentcontactid", parentContactName, parentContactId, MarketingSales.EntityNames.Contact, false);
                if (Xrm.Page.data.entity.getIsDirty()) {
                    Xrm.Page.data.save().then(function () {
                        _this.qualifyLeadSdkCall(qualifyStatus, transactionCurrencyId, true, parentAccountId, parentContactId);
                    }, function (errorResponse) { _this.handleError(errorResponse); });
                }
                else {
                    _this.qualifyLeadSdkCall(qualifyStatus, transactionCurrencyId, true, parentAccountId, parentContactId);
                }
            };
            this.getTransactionCurrencyId = function (succeedCallback, errorCallback) {
                MarketingSales.TransactionCurrency.getTransactionCurrency("transactioncurrencyid", succeedCallback, errorCallback);
            };
            this.dupWarningOnLoadHandler = function (context) {
                var parentEntityId = ClientUtility.DialogUtil.getAttributeValue(ClientUtility.MetadataDrivenDialogConstants.paramEntityId);
                var parentAccountName = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentAccountName);
                var parentAccountId = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId);
                _this.setLookupValue(MarketingSales.MetadataDrivenDialogConstants.ParentAccountLookupId, parentAccountName, parentAccountId, MarketingSales.EntityNames.Account, true);
                var parentContactName = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentContactName);
                var parentContactId = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentContactId);
                _this.setLookupValue(MarketingSales.MetadataDrivenDialogConstants.ParentContactLookupId, parentContactName, parentContactId, MarketingSales.EntityNames.Contact, true);
                var parentContactLookup = Xrm.Page.ui.controls.get(MarketingSales.MetadataDrivenDialogConstants.ParentContactLookupId);
                var emailAddress = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramEmailAddress);
                var lookUpFilterWithParent = function (context) {
                    var fetchXml = ClientUtility.StringUtil.format('<filter type="and"><condition attribute="emailaddress1" operator="eq" value="{0}"/></filter>', emailAddress);
                    parentContactLookup.addCustomFilter(fetchXml);
                };
                if (!ClientUtility.DataUtil.isNullOrUndefined(parentContactLookup)) {
                    if (!ClientUtility.DataUtil.isNullOrUndefined(emailAddress)) {
                        parentContactLookup.addPreSearch(lookUpFilterWithParent);
                    }
                }
                _this.duplicateDetectionAdminSettings = Xrm.Internal.isUci() && MarketingSales.LeadUtil.duplicateDetectionAdminSettings();
                if (context && _this.duplicateDetectionAdminSettings) {
                    var formContext = context.getFormContext();
                    var okButton = formContext && formContext.getControl("ok_id");
                    okButton.setDisabled(true);
                }
            };
            this.dupWarningOnOkClickHandler = function () {
                var parentEntityId = ClientUtility.DialogUtil.getAttributeValue(ClientUtility.MetadataDrivenDialogConstants.paramEntityId);
                ClientUtility.DialogUtil.setAttributeValue(ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked, ClientUtility.MetadataDrivenDialogConstants.DialogOkId);
                if (_this.duplicateDetectionAdminSettings) {
                    //Set the attribute value if contact is selected
                    var selectedContactAttribute = _this.getParameterFromContext(MarketingSales.MetadataDrivenDialogConstants.duplicateContactId);
                    if (!ClientUtility.DataUtil.isNullOrUndefined(selectedContactAttribute)) {
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentContactId, selectedContactAttribute);
                    }
                    //Set the attribute value if account is selected
                    var selectedAccountAttribute = _this.getParameterFromContext(MarketingSales.MetadataDrivenDialogConstants.duplicateAccountId);
                    if (!ClientUtility.DataUtil.isNullOrUndefined(selectedAccountAttribute)) {
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId, selectedAccountAttribute);
                    }
                }
                else {
                    //Set the attribute value if contact is selected
                    var contactLookupValue = _this.getLookupValue(MarketingSales.MetadataDrivenDialogConstants.ParentContactLookupId);
                    if (!ClientUtility.DataUtil.isNullOrUndefined(contactLookupValue)) {
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.ParentContactId, contactLookupValue.id);
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.ParentContactName, contactLookupValue.name);
                        // As above hiddencontrols are not supported in UCI, setting formparamters as well
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentContactId, contactLookupValue.id);
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentContactName, contactLookupValue.name);
                    }
                    //Set the attribute value if account is selected
                    var accountLookupValue = _this.getLookupValue(MarketingSales.MetadataDrivenDialogConstants.ParentAccountLookupId);
                    if (!ClientUtility.DataUtil.isNullOrUndefined(accountLookupValue)) {
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.ParentAccountId, accountLookupValue.id);
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.ParentAccountName, accountLookupValue.name);
                        // As above hiddencontrols are not supported in UCI, setting formparamters as well
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId, accountLookupValue.id);
                        ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramParentAccountName, accountLookupValue.name);
                    }
                }
                Xrm.Page.ui.close();
            };
            this.closeDupWarningDialogCallback = function (dialogParams) {
                if (_this.isQualifyLeadV2Enabled) {
                    var qualifyStatus = parseInt(dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.paramQualifyStatus].toString(), 10);
                    if (dialogParams.parameters[ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked] === ClientUtility.MetadataDrivenDialogConstants.DialogOkId) {
                        MarketingSales.LeadUtil.convertLeadDialogV2(_this.isGrid, null, qualifyStatus, dialogParams).then(function (dialogParameters) {
                            if (!ClientUtility.DataUtil.isNullOrUndefined(dialogParams) &&
                                !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters) &&
                                dialogParameters.parameters[ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked] === ClientUtility.MetadataDrivenDialogConstants.DialogOkId) {
                                MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "", dialogParameters);
                            }
                        });
                    }
                    return;
                }
                if (!ClientUtility.DataUtil.isNullOrUndefined(dialogParams) &&
                    !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters) &&
                    dialogParams.parameters[ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked] === ClientUtility.MetadataDrivenDialogConstants.DialogOkId) {
                    var qualifyStatus = parseInt(dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.paramQualifyStatus].toString(), 10);
                    // ToDo: undo workaround for bug 547179 when fixed
                    var transactionCurrency = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.paramTransactionCurrencyId];
                    var transactionCurrencyId = transactionCurrency ? (transactionCurrency.guid ? transactionCurrency.guid.toString() : transactionCurrency) : null;
                    if (_this.duplicateDetectionAdminSettings) {
                        var parentContactId = !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.paramParentContactId]) ?
                            dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.paramParentContactId] :
                            null;
                        var parentAccountId = !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId]) ?
                            dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId] :
                            null;
                        _this.executeLeadDuplication(qualifyStatus, transactionCurrencyId, parentAccountId, parentContactId, null, null);
                    }
                    else {
                        var parentAccount = !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.ParentAccountLookupId]) ?
                            dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.ParentAccountLookupId] :
                            null;
                        parentAccount = Array.isArray(parentAccount) ? parentAccount[0] : parentAccount;
                        var parentAccountId = parentAccount && parentAccount.id ? (parentAccount.id.guid ? parentAccount.id.guid.toString() : parentAccount.id) : null;
                        var parentContact = !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.ParentContactLookupId]) ?
                            dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.ParentContactLookupId] :
                            null;
                        parentContact = Array.isArray(parentContact) ? parentContact[0] : parentContact;
                        var parentContactId = parentContact && parentContact.id ? (parentContact.id.guid ? parentContact.id.guid.toString() : parentContact.id) : null;
                        _this.executeLeadDuplication(qualifyStatus, transactionCurrencyId, parentAccountId, parentContactId, null, null);
                    }
                }
            };
            this.showLeadToastMessage = function (toastLabel) {
                Xrm.Utility.getEntityMetadata(MarketingSales.EntityNames.Lead).then(function (result) {
                    var toastMessage = String.format(toastLabel, result.DisplayName);
                    Xrm.UI.addGlobalNotification(1 /* toast */, 1 /* success */, toastMessage, null, null, null);
                });
            };
            this.getOpportunityCustomer = function (leadQualificationV2Param, createAccount, createContact) {
                var opportunityCustomer = null;
                // Retrieve existing parent account and contact IDs
                var existingParentAccountId = null;
                var existingParentAccount = Xrm.Page.data.attributes.get(MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId);
                if (existingParentAccount && existingParentAccount.getValue()) {
                    existingParentAccountId = existingParentAccount.getValue();
                }
                var existingParentContactId = null;
                var existingParentContact = Xrm.Page.data.attributes.get(MarketingSales.MetadataDrivenDialogConstants.paramParentContactId);
                if (existingParentContact && existingParentContact.getValue()) {
                    existingParentContactId = existingParentContact.getValue();
                }
                var selectedAccountId = null;
                var selectedContactId = null;
                // Extract selected account and contact IDs from leadQualificationV2Param
                if (leadQualificationV2Param.length > 0) {
                    for (var _i = 0, leadQualificationV2Param_1 = leadQualificationV2Param; _i < leadQualificationV2Param_1.length; _i++) {
                        var v2param = leadQualificationV2Param_1[_i];
                        if (v2param.contactid) {
                            selectedContactId = v2param.contactid;
                        }
                        if (v2param.accountid) {
                            selectedAccountId = v2param.accountid;
                        }
                    }
                }
                var finalParentAccountId = !createAccount && (selectedAccountId || existingParentAccountId)
                    ? selectedAccountId || existingParentAccountId
                    : null;
                var finalParentContactId = !createContact && (selectedContactId || existingParentContactId)
                    ? selectedContactId || existingParentContactId
                    : null;
                // Determine the opportunity customer based on resolved parent account and contact
                if (finalParentAccountId) {
                    opportunityCustomer = { entityType: MarketingSales.EntityNames.Account, id: ClientUtility.Guid.create(finalParentAccountId) };
                }
                else if (finalParentContactId && !createAccount) {
                    opportunityCustomer = { entityType: MarketingSales.EntityNames.Contact, id: ClientUtility.Guid.create(finalParentContactId) };
                }
                return opportunityCustomer;
            };
            this.qualifyV2ClickHandler = function (formContext, controlContext, footerButton, entityCollectionRequestParamV2) {
                var start = performance.now();
                // Hide button
                footerButton.setVisible(false);
                // Set inProgress value to true
                ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramIsLeadQualificationInProgress, true);
                // Set focus back to V2 control
                controlContext.setFocus();
                var opportunityRecords = entityCollectionRequestParamV2.leadQualificationV2Param.filter(function (item) { return item["@odata.type"] === MarketingSales.QualifyLeadV2Control.OpportunityOdataType; });
                var accountRecords = entityCollectionRequestParamV2.leadQualificationV2Param.filter(function (item) { return item["@odata.type"] === MarketingSales.QualifyLeadV2Control.AccountOdataType; });
                var contactRecords = entityCollectionRequestParamV2.leadQualificationV2Param.filter(function (item) { return item["@odata.type"] === MarketingSales.QualifyLeadV2Control.ContactOdataType; });
                var createAccount = entityCollectionRequestParamV2.recordPreference.createAccount;
                var createContact = entityCollectionRequestParamV2.recordPreference.createContact;
                var createOpportunity = entityCollectionRequestParamV2.recordPreference.createOpportunity;
                var suppressDuplicateDetection = true;
                var leadQualificationV2Param = entityCollectionRequestParamV2.leadQualificationV2Param;
                var opportunityCustomer = _this.getOpportunityCustomer(leadQualificationV2Param, createAccount, createContact);
                var qualifyLeadParams = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramQualifyLead, formContext);
                var lead = qualifyLeadParams.leadReference;
                var qualifyStatus = qualifyLeadParams.qualifyStatus;
                var processInstanceReference = qualifyLeadParams.processInstanceReference;
                var transactionCurrencyReference = qualifyLeadParams.transactionCurrencyReference;
                var sourceCampaignReference = qualifyLeadParams.sourceCampaignReference;
                // Commenting optional paramter for now, would handle this for EA branch
                var request = new MarketingSales.QualifyLeadRequest(lead, createAccount, createContact, createOpportunity, transactionCurrencyReference || {}, opportunityCustomer || {}, sourceCampaignReference || {}, qualifyStatus, processInstanceReference, suppressDuplicateDetection, JSON.stringify(leadQualificationV2Param));
                Xrm.WebApi.online.execute(request).then(function (qualifyLeadResponse) {
                    qualifyLeadResponse.json().then(function (response) {
                        MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "qualifyV2ClickHandler", MarketingSales.MetadataDrivenDialogConstants.information, "leadQualificationParams", {
                            Account: {
                                DropDownSelection: LeadCommandActionsImplementation.getDropDownSelection(createAccount, accountRecords.length),
                                recordsCount: accountRecords.length
                            },
                            Contact: {
                                DropDownSelection: LeadCommandActionsImplementation.getDropDownSelection(createContact, contactRecords.length),
                                recordsCount: contactRecords.length
                            },
                            Opportunity: {
                                recordsCount: opportunityRecords.length
                            }
                        });
                        var shouldShowCopilotSummary = controlContext.getOutputs()[MarketingSales.QualifyLeadV2Control.ShowCopilotSummary].value;
                        var qualifiedOpportunityRecords = response.value.filter(function (item) { return item["@odata.type"].includes(MarketingSales.QualifyLeadV2Control.OpportunityOdataType); });
                        // If no records are created and summary generation is disabled, close leadQualification side pane and return.
                        // If record is qualified from gridView, close leadQualification side pane and return.
                        if ((response.value.length <= 0 || qualifiedOpportunityRecords.length <= 1) && !shouldShowCopilotSummary || qualifyLeadParams.isGrid) {
                            if (Xrm.Internal.isUci()) {
                                _this.showLeadToastMessage(MarketingSales.ResourceStringProvider.getResourceString("ToastMessage_Lead_Qualify"));
                            }
                            formContext.ui.close();
                            MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "qualifyV2ClickHandler", MarketingSales.MetadataDrivenDialogConstants.information, "SkipSummaryAndNavigateSuccess", {
                                qualifiedOpportunityRecordsCount: qualifiedOpportunityRecords.length
                            });
                            if (qualifiedOpportunityRecords.length == 1 && qualifiedOpportunityRecords[0].opportunityid) {
                                MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "qualifyV2ClickHandler", MarketingSales.MetadataDrivenDialogConstants.information, "SkipSummaryAndNavigateSuccess", {
                                    isOpportunityIdValid: true
                                });
                                MarketingSales.LeadUtil.navigateToQualifiedRecord(qualifiedOpportunityRecords[0].opportunityid, MarketingSales.EntityNames.Opportunity);
                                // this.logStatusUpdate(
                                // 	lead.id,
                                // 	"LeadCommandActions.qualifyV2ClickHandler",
                                // 	"Qualify",
                                // 	{ version: "v2" }
                                // );
                                var leadData = {
                                    leadid: lead.id,
                                };
                                var leadDataJsonString = JSON.stringify(leadData);
                                MarketingSales.SalesAgentUsageClient.logSalesAgentUsage("LeadQualifyEvent", leadDataJsonString);
                            }
                        }
                        else {
                            // Set response to control's output param
                            // Set inprogress to false, footer button's label to 'Finish' and make it visible
                            // Set focus back to control
                            ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.isCopilotSummaryEnabled, shouldShowCopilotSummary);
                            ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.responseParamQualifyLeadV2, { response: response });
                            ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramIsLeadQualificationInProgress, false);
                            footerButton.setLabel(MarketingSales.ResourceStringProvider.getResourceString("FinishLabel"));
                            footerButton.setVisible(true);
                            controlContext.setFocus();
                            LeadCommandActionsImplementation.isLeadQualified = true;
                            var end = performance.now();
                            MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "qualifyV2ClickHandler", MarketingSales.MetadataDrivenDialogConstants.information, "qualifyV2ClickHandlerSuccess", {
                                timeTaken: end - start
                            });
                        }
                    });
                }, function (errorResponse) {
                    var end = performance.now();
                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "qualifyV2ClickHandler", MarketingSales.MetadataDrivenDialogConstants.error, "qualifyV2ClickHandlerEndFailed" +
                        ': ' +
                        'Error code: ' +
                        errorResponse && errorResponse.code +
                        ' ,message: ' +
                        errorResponse && errorResponse.message, {
                        timeTaken: end - start
                    });
                    ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.responseParamQualifyLeadV2, { errorResponse: errorResponse });
                    ClientUtility.DialogUtil.setAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramIsLeadQualificationInProgress, false);
                    footerButton.setVisible(true);
                    controlContext.setFocus();
                    MarketingSales.LeadErrorHandling.actionFailedErrorDialog(errorResponse);
                });
            };
            this.convertLeadOnQualifyV2ClickHandler = function (context) {
                try {
                    var controlContext_1 = (context != null && context != undefined) ? context.getFormContext().getControl(MarketingSales.QualifyLeadV2Control.ControlId) : null;
                    if (ClientUtility.DataUtil.isNullOrUndefined(controlContext_1)) {
                        //TODO: log error, control not found
                        return;
                    }
                    var formContext_1 = context.getFormContext();
                    var footerButton_1 = formContext_1.getControl(MarketingSales.QualifyLeadV2Control.FooterId);
                    // If lead is not qualified, that means footer button's lable is 'Qualify'
                    if (!LeadCommandActionsImplementation.isLeadQualified) {
                        var entityCollectionRequestParamV2_1 = controlContext_1.getOutputs()[MarketingSales.QualifyLeadV2Control.RequestParam].value;
                        var hasOpportunity = entityCollectionRequestParamV2_1.leadQualificationV2Param.some(function (item) { return item["@odata.type"] === MarketingSales.QualifyLeadV2Control.OpportunityOdataType; });
                        // Show confirmation dialog if user is qualifying lead with no opportunity
                        if (!hasOpportunity) {
                            Xrm.Navigation.openConfirmDialog({
                                title: MarketingSales.ResourceStringProvider.getResourceString("QualifyConfirmationDialogTitle"),
                                text: MarketingSales.ResourceStringProvider.getResourceString("QualifyConfirmationDialogContext"),
                                confirmButtonLabel: MarketingSales.ResourceStringProvider.getResourceString("QualifyLabel"),
                                cancelButtonLabel: MarketingSales.ResourceStringProvider.getResourceString("CancelLabel")
                            }).then(function (result) {
                                if (result.confirmed) {
                                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, _this.convertLeadOnQualifyV2ClickHandler.name, MarketingSales.MetadataDrivenDialogConstants.information, "NoOptyDialogConfirmClicked");
                                    _this.qualifyV2ClickHandler(formContext_1, controlContext_1, footerButton_1, entityCollectionRequestParamV2_1);
                                }
                                else {
                                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, _this.convertLeadOnQualifyV2ClickHandler.name, MarketingSales.MetadataDrivenDialogConstants.information, "NoOptyDialogCancelClicked");
                                }
                            });
                        }
                        else {
                            _this.qualifyV2ClickHandler(formContext_1, controlContext_1, footerButton_1, entityCollectionRequestParamV2_1);
                        }
                    }
                    else {
                        // Get primary opportunyty from responseParam
                        var qualifiedLeadResponse = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.responseParamQualifyLeadV2);
                        var opportunityId = null;
                        if (qualifiedLeadResponse && qualifiedLeadResponse.response.value && qualifiedLeadResponse.response.value.length > 0) {
                            var i = 0;
                            while (i < qualifiedLeadResponse.response.value.length && !opportunityId) {
                                if (qualifiedLeadResponse.response.value[i].opportunityid) {
                                    opportunityId = qualifiedLeadResponse.response.value[i].opportunityid;
                                }
                                i++;
                            }
                        }
                        if (!ClientUtility.DataUtil.isNullOrUndefined(opportunityId)) {
                            if (Xrm.Internal.isUci()) {
                                _this.showLeadToastMessage(MarketingSales.ResourceStringProvider.getResourceString("ToastMessage_Lead_Qualify"));
                            }
                            formContext_1.ui.close();
                            MarketingSales.LeadUtil.navigateToQualifiedRecord(opportunityId, MarketingSales.EntityNames.Opportunity);
                        }
                        else {
                            if (Xrm.Internal.isUci()) {
                                _this.showLeadToastMessage(MarketingSales.ResourceStringProvider.getResourceString("ToastMessage_Lead_Qualify"));
                            }
                            formContext_1.ui.close();
                            Xrm.Page.ui.refreshRibbon();
                        }
                        var skipQualifyLeadV2Survey = MarketingSales.LeadUtil.getFeatureControlSetting(MarketingSales.LeadUtil.FCS_QualifyLeadNamespace, MarketingSales.LeadUtil.FCS_SkipQualifyLeadV2Survey, false);
                        MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, _this.convertLeadOnQualifyV2ClickHandler.name, MarketingSales.MetadataDrivenDialogConstants.information, "isV2SurveyBlocked: " + skipQualifyLeadV2Survey);
                        // Trigger enhanced lead qualification survey once lead qualification is complete
                        if (!skipQualifyLeadV2Survey && checkCoolingPeriod(MarketingSales.QualifyLeadV2Control.CESSurveyTimestampKey)) {
                            MarketingSales.CESSurvey.showCESSurvey(MarketingSales.QualifyLeadV2Control.TeamName, MarketingSales.QualifyLeadV2Control.SurveyName, MarketingSales.QualifyLeadV2Control.SurveyEventName);
                            setCESSurveyTimestamp(MarketingSales.QualifyLeadV2Control.CESSurveyTimestampKey);
                        }
                        var lead = ClientUtility.DialogUtil.getAttributeValue(MarketingSales.MetadataDrivenDialogConstants.paramQualifyLead, formContext_1).leadReference;
                        // this.logStatusUpdate(
                        // 	lead.id,
                        // 	"LeadCommandActions.convertLeadOnQualifyV2ClickHandler",
                        // 	"Qualify",
                        // 	{ version: "v2" }
                        // );
                        var leadData = {
                            leadid: lead.id,
                        };
                        var leadDataJsonString = JSON.stringify(leadData);
                        MarketingSales.SalesAgentUsageClient.logSalesAgentUsage("LeadQualifyEvent", leadDataJsonString);
                    }
                }
                catch (e) {
                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "convertLeadV2OnOkClickHandler", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
                }
            };
            this.qualifyLeadSdkCall = function (qualifyStatus, transactionCurrencyId, suppressDuplicateDetection, parentAccountId, parentContactId) {
                var _this = this;
                var parentAccountName = null;
                var parentContactName = null;
                var parentContactResolved = false;
                var parentAccountResolved = false;
                var companyName = ClientUtility.StringUtil.StringEmpty;
                var fullName = ClientUtility.StringUtil.StringEmpty;
                var createAccount = false;
                var createContact = true;
                var createOpportunity = true;
                var sourceCampaignId = this.getLookupId(this.getLookupValue("campaignid"));
                var opportunityCustomer = null;
                if (ClientUtility.DataUtil.isNullOrUndefined(parentAccountId)) {
                    var parentAccount = this.getLookupValue("parentaccountid");
                    parentAccountId = this.getLookupId(parentAccount);
                    if (!ClientUtility.DataUtil.isNullOrEmptyString(parentAccountId)) {
                        parentAccountResolved = true;
                        parentAccountName = parentAccount.name;
                    }
                }
                else {
                    parentAccountResolved = true;
                }
                if (ClientUtility.DataUtil.isNullOrUndefined(parentContactId)) {
                    var parentContact = this.getLookupValue("parentcontactid");
                    parentContactId = this.getLookupId(parentContact);
                    if (!ClientUtility.DataUtil.isNullOrEmptyString(parentContactId)) {
                        parentContactResolved = true;
                        parentContactName = parentContact.name;
                    }
                }
                else {
                    parentContactResolved = true;
                }
                var companyNameAttribute = Xrm.Page.getAttribute("companyname");
                if (!ClientUtility.DataUtil.isNullOrUndefined(companyNameAttribute)) {
                    var companyNameAttributeValue = companyNameAttribute.getValue();
                    if (!ClientUtility.DataUtil.isNullOrUndefined(companyNameAttributeValue)) {
                        companyName = companyNameAttributeValue;
                    }
                }
                var fullNameAttribute = Xrm.Page.getAttribute("fullname");
                if (!ClientUtility.DataUtil.isNullOrUndefined(fullNameAttribute)) {
                    var fullNameAttributeValue = fullNameAttribute.getValue();
                    if (!ClientUtility.DataUtil.isNullOrUndefined(fullNameAttributeValue)) {
                        fullName = fullNameAttributeValue;
                    }
                }
                if (!ClientUtility.DataUtil.isNullOrEmptyString(companyName) || (!ClientUtility.DataUtil.isNullOrUndefined(companyNameAttribute) && !companyNameAttribute.getUserPrivilege().canRead)) {
                    createAccount = true;
                }
                if (parentAccountResolved && parentContactResolved) {
                    createAccount = false;
                    createContact = false;
                    opportunityCustomer = { entityType: MarketingSales.EntityNames.Account, id: ClientUtility.Guid.create(parentAccountId) };
                }
                else {
                    if (parentAccountResolved) {
                        createAccount = false;
                        createContact = true;
                    }
                    else {
                        if (parentContactResolved) {
                            if (!ClientUtility.DataUtil.isNullOrEmptyString(companyName)) {
                                createAccount = true;
                                createContact = false;
                            }
                            else {
                                createAccount = false;
                                createContact = false;
                                opportunityCustomer = { entityType: MarketingSales.EntityNames.Contact, id: ClientUtility.Guid.create(parentContactId) };
                            }
                        }
                    }
                }
                var leadId = { entityType: MarketingSales.EntityNames.Lead, id: ClientUtility.Guid.create(Xrm.Page.data.entity.getId()) };
                var transactionCurrencyReference = ClientUtility.DataUtil.isNullOrWhiteSpace(transactionCurrencyId) ?
                    null :
                    { entityType: MarketingSales.EntityNames.TransactionCurrency, id: ClientUtility.Guid.create(transactionCurrencyId) };
                var sourceCampaignReference = ClientUtility.DataUtil.isNullOrWhiteSpace(sourceCampaignId) ?
                    null :
                    { entityType: MarketingSales.EntityNames.Campaign, id: ClientUtility.Guid.create(sourceCampaignId) };
                if (ClientUtility.ClientUtil.isMobileOffline()) {
                    transactionCurrencyReference = Xrm.Utility.getGlobalContext().userSettings.transactionCurrency;
                }
                var emailAddress;
                var emailAttribute = Xrm.Page.getAttribute("emailaddress1");
                if (!ClientUtility.DataUtil.isNullOrUndefined(emailAttribute)) {
                    emailAddress = emailAttribute.getValue();
                }
                var qualifyLeadFailedCallback = function (errorResponse) {
                    var actionFailed = true;
                    if (errorResponse) {
                        // 0x80040333 (2147746611) : Duplicate entity record error
                        if ((errorResponse.message && errorResponse.message.indexOf("ErrorCode: 0x80040333") >= 0) ||
                            errorResponse.errorCode === 2147746611 ||
                            errorResponse.errorCode === "0x80040333") {
                            actionFailed = false;
                            var isOpportunityEntityInResponse = false;
                            var isLeadEntityInResponse = false;
                            var isContactEntityInResponse = false;
                            var isAccountEntityInResponse = false;
                            var cloneDuplicateEntityFromResponse = _this.cloneDuplicateEntityFromErrorResponse(errorResponse);
                            var cloneDuplicateEntity = ClientUtility.DataUtil.isNullOrUndefined(cloneDuplicateEntityFromResponse) ?
                                ClientUtility.DataUtil.EmptyString : cloneDuplicateEntityFromResponse;
                            if (!ClientUtility.DataUtil.isNullOrUndefined(cloneDuplicateEntity.match("<LogicalName>opportunity</LogicalName>"))) {
                                isOpportunityEntityInResponse = cloneDuplicateEntity.match("<LogicalName>opportunity</LogicalName>").length > 0;
                            }
                            if (!ClientUtility.DataUtil.isNullOrUndefined(cloneDuplicateEntity.match("<LogicalName>lead</LogicalName>"))) {
                                isLeadEntityInResponse = cloneDuplicateEntity.match("<LogicalName>lead</LogicalName>").length > 0;
                            }
                            if (!ClientUtility.DataUtil.isNullOrUndefined(cloneDuplicateEntity.match("<LogicalName>contact</LogicalName>"))) {
                                isContactEntityInResponse = cloneDuplicateEntity.match("<LogicalName>contact</LogicalName>").length > 0;
                            }
                            if (!ClientUtility.DataUtil.isNullOrUndefined(cloneDuplicateEntity.match("<LogicalName>account</LogicalName>"))) {
                                isAccountEntityInResponse = cloneDuplicateEntity.match("<LogicalName>account</LogicalName>").length > 0;
                            }
                            // While qualifying the lead --> Duplicate contacts and accounts dialog will be shown when opportunity entity is not present in errorresponse
                            if (isAccountEntityInResponse || isContactEntityInResponse) {
                                _this.handleLeadDuplication(parentAccountId, parentContactId, parentAccountName, parentContactName, qualifyStatus, transactionCurrencyId, emailAddress);
                            }
                            else if (isOpportunityEntityInResponse || isLeadEntityInResponse) {
                                var entityNames = null;
                                if (isLeadEntityInResponse && isOpportunityEntityInResponse) {
                                    entityNames = MarketingSales.ResourceStringProvider.getResourceString("LeadDuplicateRecordCreationErrorEntityName") + " or " + MarketingSales.ResourceStringProvider.getResourceString("OpportunityDuplicateRecordCreationErrorEntityName");
                                }
                                else {
                                    entityNames = isOpportunityEntityInResponse ? MarketingSales.ResourceStringProvider.getResourceString("OpportunityDuplicateRecordCreationErrorEntityName") : MarketingSales.ResourceStringProvider.getResourceString("LeadDuplicateRecordCreationErrorEntityName");
                                }
                                errorResponse["message"] = entityNames + errorResponse["message"].substring(1);
                                MarketingSales.LeadErrorHandling.actionFailedCallback(errorResponse);
                            }
                            else {
                                MarketingSales.LeadErrorHandling.actionFailedCallback(errorResponse);
                            }
                        }
                        //	2147746326(Unexpected error) is a generic error code. UCI library replaces the outer error "message" in errorResponse with
                        //	error description corresponding to error code 2147746326. To throw a specific error, set errorResponse.message to a specific error message.
                        var processNoActiveStageError = String.format(MarketingSales.ResourceStringProvider.getResourceString("ErrorMessage_Lead_ReQualify_ProcessNoActiveStage"), MarketingSales.EntityNames.Lead);
                        if ((errorResponse.errorCode === 2147746326 || errorResponse.errorCode === -2147220970) && errorResponse.innerror && errorResponse.innerror.message &&
                            errorResponse.innerror.message.indexOf(processNoActiveStageError) >= 0) {
                            errorResponse.message = processNoActiveStageError;
                            errorResponse.errorRegion = "Sales";
                        }
                        else if (errorResponse.errorCode === 2148532226) {
                            errorResponse.errorRegion = "Sales";
                        }
                    }
                    if (actionFailed) {
                        MarketingSales.LeadErrorHandling.actionFailedErrorDialog(errorResponse);
                    }
                };
                if (!this.adminSettings && this.fcbOctoberUpdate) {
                    if (!ClientUtility.DataUtil.isNullOrUndefined(this.createContactFlag) && this.createContactFlag == false)
                        createContact = this.createContactFlag;
                    if (!ClientUtility.DataUtil.isNullOrUndefined(this.createAccountFlag) && this.createAccountFlag == false)
                        createAccount = this.createAccountFlag;
                    if (!ClientUtility.DataUtil.isNullOrUndefined(this.createOpportunityFlag) && this.createOpportunityFlag == false)
                        createOpportunity = this.createOpportunityFlag;
                }
                var request = new MarketingSales.QualifyLeadRequest(leadId, createAccount, createContact, createOpportunity, transactionCurrencyReference || {}, opportunityCustomer || {}, sourceCampaignReference || {}, qualifyStatus, MarketingSales.LeadUtil.retrieveProcessInstanceId(), suppressDuplicateDetection);
                if (ClientUtility.ClientUtil.isMobileOffline() && Xrm.Mobile.offline.isOfflineEnabled(Xrm.Page.data.entity.getEntityName())) {
                    var processStage = null;
                    var currentStage = null;
                    var nextStage = null;
                    try {
                        processStage = Xrm.Page.data.process.getActivePath();
                        currentStage = processStage.get(ClientUtility.MetadataDrivenDialogConstants.defaultIndex);
                        nextStage = processStage.get(ClientUtility.MetadataDrivenDialogConstants.firstIndex);
                    }
                    catch (ex) {
                    }
                    var nextStageId = ClientUtility.StringUtil.StringEmpty;
                    var nextTraversedPath = ClientUtility.StringUtil.StringEmpty;
                    var tempTraversedPath = ClientUtility.StringUtil.StringEmpty;
                    if (nextStage && nextStage.getEntityName().toLowerCase() !== Sales.EntityNames.Opportunity.toLowerCase()) {
                        tempTraversedPath = ClientUtility.MetadataDrivenDialogConstants.commaSeperator + nextStage.getId();
                        var length_1 = processStage.getLength();
                        for (var i = 2; i < length_1; i++) {
                            var processStageItem = processStage.get(i);
                            if (!ClientUtility.DataUtil.isNullOrUndefined(processStageItem)) {
                                tempTraversedPath = tempTraversedPath + ClientUtility.MetadataDrivenDialogConstants.commaSeperator + processStageItem.getId();
                                if (processStageItem.getEntityName().toLowerCase() === Sales.EntityNames.Opportunity.toLowerCase()) {
                                    nextStage = processStageItem;
                                    break;
                                }
                            }
                        }
                    }
                    if (!ClientUtility.DataUtil.isNullOrUndefined(currentStage) && !ClientUtility.DataUtil.isNullOrUndefined(nextStage)) {
                        var currentStageId = currentStage.getId();
                        nextStageId = nextStage.getId();
                        if (!ClientUtility.DataUtil.isNullOrEmptyString(tempTraversedPath)) {
                            nextTraversedPath = currentStageId + tempTraversedPath;
                        }
                        else {
                            nextTraversedPath = currentStageId + ClientUtility.MetadataDrivenDialogConstants.commaSeperator + nextStageId;
                        }
                    }
                    // as in outlook offline, creation of duplicate entries should be allowed even though duplicate detection is enabled.
                    request.SuppressDuplicateDetection = true;
                    var progressIndicator = new ClientUtility.ProgressIndicator();
                    progressIndicator.show();
                    Xrm.WebApi.offline.execute(request, { "EntityLogicalName": MarketingSales.EntityNames.Lead, "NextStageId": nextStageId, "TraversedPath": nextTraversedPath })
                        .then(progressIndicator.hideOnSuccess(function (qualifyLeadResponse) {
                            qualifyLeadResponse.json().then(function (response) {
                                var opportunityId = null;
                                if (response.createdEntities && response.createdEntities.length > 0) {
                                    var i = 0;
                                    while (i < response.createdEntities.length && !opportunityId) {
                                        if (response.createdEntities[i].LogicalName === Sales.EntityNames.Opportunity) {
                                            opportunityId = response.createdEntities[i].Id;
                                        }
                                        i++;
                                    }
                                }
                                if (!ClientUtility.DataUtil.isNullOrUndefined(opportunityId)) {
                                    Xrm.Page.data.refresh(true).then(function () {
                                        Xrm.Page.ui.refreshRibbon();
                                        Xrm.Utility.openEntityForm(Sales.EntityNames.Opportunity, opportunityId);
                                    }, function (errorResponse) { MarketingSales.LeadErrorHandling.actionFailedCallback(errorResponse); });
                                }
                            });
                        }), progressIndicator.hideOnError(qualifyLeadFailedCallback));
                }
                else {
                    if (this.isQualifyLeadV2Enabled) {
                        var qualifyLeadParams = {
                            leadReference: leadId,
                            createAccount: createAccount,
                            createContact: createContact,
                            createOpportunity: createOpportunity,
                            transactionCurrencyReference: transactionCurrencyReference || {},
                            opportunityCustomerReference: opportunityCustomer || {},
                            sourceCampaignReference: sourceCampaignReference || {},
                            qualifyStatus: qualifyStatus,
                            processInstanceReference: MarketingSales.LeadUtil.retrieveProcessInstanceId()
                        };
                        this.handleQualifyLeadV2Scenario(qualifyLeadParams, parentAccountId, parentContactId, emailAddress);
                        return;
                    }
                    var progressIndicator = new ClientUtility.ProgressIndicator();
                    progressIndicator.show();
                    var start_1 = performance.now();
                    Xrm.WebApi.online.execute(request)
                        .then(progressIndicator.hideOnSuccess(function (qualifyLeadResponse) {
                            qualifyLeadResponse.json().then(function (response) {
                                var opportunityId = null;
                                if (response && response.value && response.value.length > 0) {
                                    var i = 0;
                                    while (i < response.value.length && !opportunityId) {
                                        if (response.value[i].opportunityid) {
                                            opportunityId = response.value[i].opportunityid;
                                        }
                                        i++;
                                    }
                                }
                                var end = performance.now();
                                MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "qualifyLeadSdkCall", MarketingSales.MetadataDrivenDialogConstants.information, "qualifyLeadSdkSuccess", {
                                    timeTaken: end - start_1
                                });
                                if (!ClientUtility.DataUtil.isNullOrUndefined(opportunityId)) {
                                    if (Xrm.Internal.isUci() && _this.fcbOctoberUpdate) {
                                        _this.showLeadToastMessage(MarketingSales.ResourceStringProvider.getResourceString("ToastMessage_Lead_Qualify"));
                                    }
                                    Xrm.Utility.openEntityForm(Sales.EntityNames.Opportunity, opportunityId);
                                }
                                else {
                                    if (Xrm.Internal.isUci() && _this.fcbOctoberUpdate) {
                                        _this.showLeadToastMessage(MarketingSales.ResourceStringProvider.getResourceString("ToastMessage_Lead_Qualify"));
                                    }
                                    Xrm.Page.data.refresh(true).then(function () {
                                        Xrm.Page.ui.refreshRibbon();
                                    }, function (errorResponse) { MarketingSales.LeadErrorHandling.actionFailedCallback(errorResponse); });
                                }
                                // Trigger NSAT survey once lead qualification is complete
                                MarketingSales.NSATSurvey.getNSATSurvey("lead", MarketingSales.NSATScenarios.QualifyLead);
                                // this.logStatusUpdate(
                                // 	leadId.id,
                                // 	"LeadCommandActions.qualifyLeadSdkCall",
                                // 	"Qualify",
                                // 	{ version: "v1" }
                                // );
                                var leadData = {
                                    leadid: leadId.id,
                                };
                                var leadDataJsonString = JSON.stringify(leadData);
                                MarketingSales.SalesAgentUsageClient.logSalesAgentUsage("LeadQualifyEvent", leadDataJsonString);
                            });
                        }), progressIndicator.hideOnError(qualifyLeadFailedCallback));
                }
            };
            this.handleQualifyLeadV2Scenario = function (qualifyLeadParams, parentAccountId, parentContactId, emailAddress) {
                if (!qualifyLeadParams.createAccount && !qualifyLeadParams.createContact) {
                    _this.openQualifyLeadV2Dialog(parentAccountId, parentContactId, qualifyLeadParams);
                }
                else {
                    var progressIndicator_1 = new ClientUtility.ProgressIndicator();
                    progressIndicator_1.show();
                    var recordsToFetch = [];
                    var parentEntityId = Xrm.Page.data.entity.getId();
                    var options_1 = {
                        height: 600,
                        width: 800,
                        position: 1 /* center */
                    };
                    var dialogArguments_1 = {};
                    dialogArguments_1[ClientUtility.MetadataDrivenDialogConstants.EntityId] = parentEntityId;
                    // As above hiddencontrols are not supported in UCI, setting formparamters as well
                    dialogArguments_1[ClientUtility.MetadataDrivenDialogConstants.paramEntityId] = Xrm.Page.data.entity.getId();
                    dialogArguments_1[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId] = parentAccountId;
                    dialogArguments_1[MarketingSales.MetadataDrivenDialogConstants.paramParentContactId] = parentContactId;
                    dialogArguments_1[MarketingSales.MetadataDrivenDialogConstants.paramQualifyStatus] = qualifyLeadParams.qualifyStatus;
                    dialogArguments_1[MarketingSales.MetadataDrivenDialogConstants.paramEmailAddress] = emailAddress;
                    dialogArguments_1[MarketingSales.MetadataDrivenDialogConstants.paramQualifyLead] = qualifyLeadParams;
                    var objectAccount = {
                        parentEntity_Id: parentEntityId,
                        EntityName: MarketingSales.EntityNames.Account,
                    };
                    recordsToFetch.push(objectAccount);
                    var objectContact = {
                        parentEntity_Id: parentEntityId,
                        EntityName: MarketingSales.EntityNames.Contact,
                    };
                    recordsToFetch.push(objectContact);
                    _this.getAssociatedEntityforDuplicate(recordsToFetch).then(function (response) {
                        if (response.length > 1) {
                            _this.duplicateAccount = response[0];
                            _this.duplicateContact = response[1];
                            // These modifications are done to ensure 0 duplicates are returned as account/contact is already added through BPF.
                            if (!qualifyLeadParams.createAccount) {
                                _this.duplicateAccount.name = null;
                            }
                            if (!qualifyLeadParams.createContact) {
                                _this.duplicateContact.firstname = null;
                                _this.duplicateContact.lastname = null;
                            }
                            _this.handleDuplicateDetectionCount(response).then(progressIndicator_1.hideOnSuccess(function (response) {
                                var accountDuplicates = response.accountDuplicates;
                                var contactDuplicates = response.contactDuplicates;
                                // Open Duplicate detection dialog only when at least 1 duplicate is found.
                                if ((accountDuplicates &&
                                    accountDuplicates.DuplicatesCount &&
                                    accountDuplicates.DuplicatesCount.length > 0 &&
                                    accountDuplicates.DuplicatesCount > 0) ||
                                    (contactDuplicates &&
                                        contactDuplicates.DuplicatesCount &&
                                        contactDuplicates.DuplicatesCount.length > 0 &&
                                        contactDuplicates.DuplicatesCount > 0)) {
                                    dialogArguments_1[MarketingSales.MetadataDrivenDialogConstants.duplicateAccount] = _this.duplicateAccount;
                                    dialogArguments_1[MarketingSales.MetadataDrivenDialogConstants.duplicateContact] = _this.duplicateContact;
                                    Xrm.Navigation.openDialog(MarketingSales.DialogName.DuplicateQualifyLeadDialog, options_1, dialogArguments_1)
                                        .then(_this.closeDupWarningDialogCallback);
                                }
                                else {
                                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.DuplicateQualifyLeadDialog, "handleDuplicateDetectionCount", MarketingSales.MetadataDrivenDialogConstants.information, LeadCommandActionsImplementation.noDuplicatesFound);
                                    _this.openQualifyLeadV2Dialog(parentAccountId, parentContactId, qualifyLeadParams);
                                }
                            }), progressIndicator_1.hideOnError(function (error) {
                                MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.DuplicateQualifyLeadDialog, "handleDuplicateDetectionCount", MarketingSales.MetadataDrivenDialogConstants.error, error.toString());
                                _this.openQualifyLeadV2Dialog(parentAccountId, parentContactId, qualifyLeadParams);
                            }));
                        }
                        else {
                            progressIndicator_1.hide();
                            MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.DuplicateQualifyLeadDialog, "getAssociatedEntityforDuplicate", MarketingSales.MetadataDrivenDialogConstants.error, "Response length is: " + response.length.toString());
                            _this.openQualifyLeadV2Dialog(parentAccountId, parentContactId, qualifyLeadParams);
                        }
                    }, progressIndicator_1.hideOnError(function () {
                        MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.DuplicateQualifyLeadDialog, "getAssociatedEntityforDuplicate", MarketingSales.MetadataDrivenDialogConstants.error, LeadCommandActionsImplementation.failedToFetchRelatedEntities);
                        Xrm.Utility.alertDialog(MarketingSales.ResourceStringProvider.getResourceString("RelatedEntities_Fetch_Fail_Message"), function () {
                            _this.openQualifyLeadV2Dialog(parentAccountId, parentContactId, qualifyLeadParams);
                        });
                        return;
                    }));
                }
            };
            this.handleDuplicateDetectionCount = function (response) {
                return __awaiter(_this, void 0, void 0, function () {
                    var requestAccount, requestContact, accountDuplicates, contactDuplicates, hasAccountCallFailed, hasContactCallFailed, accountRequestPromise, contactRequestPromise;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                requestAccount = new MarketingSales.GetEntityWiseDuplicatesCountRequest(response[0]);
                                requestContact = new MarketingSales.GetEntityWiseDuplicatesCountRequest(response[1]);
                                accountDuplicates = null;
                                contactDuplicates = null;
                                hasAccountCallFailed = false;
                                hasContactCallFailed = false;
                                accountRequestPromise = Xrm.WebApi.online.execute(requestAccount)
                                    .then(function (response) { return response.json(); })
                                    .catch(function (error) {
                                        MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.DuplicateQualifyLeadDialog, "handleDuplicateDetectionCount", MarketingSales.MetadataDrivenDialogConstants.error, JSON.stringify({
                                            entity: "account",
                                            error: error.toString()
                                        }));
                                        hasAccountCallFailed = true;
                                        return null;
                                    });
                                contactRequestPromise = Xrm.WebApi.online.execute(requestContact)
                                    .then(function (response) { return response.json(); })
                                    .catch(function (error) {
                                        MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.DuplicateQualifyLeadDialog, "handleDuplicateDetectionCount", MarketingSales.MetadataDrivenDialogConstants.error, JSON.stringify({
                                            entity: "contact",
                                            error: error.toString()
                                        }));
                                        hasContactCallFailed = true;
                                        return null;
                                    });
                                return [4 /*yield*/, accountRequestPromise];
                            case 1:
                                accountDuplicates = _a.sent();
                                return [4 /*yield*/, contactRequestPromise];
                            case 2:
                                contactDuplicates = _a.sent();
                                if (hasAccountCallFailed && hasContactCallFailed) {
                                    return [2 /*return*/, Promise.reject("Account and Contact GetEntityWiseDuplicatesCountRequest calls failed")];
                                }
                                return [2 /*return*/, Promise.resolve({ accountDuplicates: accountDuplicates, contactDuplicates: contactDuplicates })];
                        }
                    });
                });
            };
            this.openQualifyLeadV2Dialog = function (parentAccountId, parentContactId, qualifyLeadParams) {
                var dialogArguments = {
                    parameters: (_a = {},
                        _a[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId] = parentAccountId,
                        _a[MarketingSales.MetadataDrivenDialogConstants.paramParentContactId] = parentContactId,
                        _a[MarketingSales.MetadataDrivenDialogConstants.paramQualifyLead] = qualifyLeadParams,
                        _a)
                };
                MarketingSales.LeadUtil.convertLeadDialogV2(_this.isGrid, null, qualifyLeadParams.qualifyStatus, dialogArguments).then(function (dialogParams) {
                    if (!ClientUtility.DataUtil.isNullOrUndefined(dialogParams) &&
                        !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters) &&
                        dialogParams.parameters[ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked] === ClientUtility.MetadataDrivenDialogConstants.DialogOkId) {
                        MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "OpenQualifyLeadV2Dialog", MarketingSales.MetadataDrivenDialogConstants.information, "", dialogParams);
                    }
                });
                var _a;
            };
            this.cloneDuplicateEntityFromErrorResponse = function (errorResponse) {
                var duplicateEntityAnnotation = null;
                if (!ClientUtility.DataUtil.isNullOrUndefined(errorResponse.raw)) {
                    var jsonObj = JSON.parse(errorResponse.raw);
                    if (!ClientUtility.DataUtil.isNullOrUndefined(jsonObj) &&
                        !ClientUtility.DataUtil.isNullOrUndefined(jsonObj['_errorFault']) &&
                        !ClientUtility.DataUtil.isNullOrUndefined(jsonObj['_errorFault']['_annotations']) &&
                        !ClientUtility.DataUtil.isNullOrUndefined(jsonObj['_errorFault']['_annotations']['@Microsoft.PowerApps.CDS.ErrorDetails.DuplicateEntity'])) {
                        duplicateEntityAnnotation = jsonObj['_errorFault']['_annotations']['@Microsoft.PowerApps.CDS.ErrorDetails.DuplicateEntity'];
                        var parser, xmlDoc;
                        parser = new DOMParser();
                        xmlDoc = parser.parseFromString(duplicateEntityAnnotation, "text/xml");
                        // Removing attributes tag from
                        // errorresponse-->[_annotations']['@Microsoft.PowerApps.CDS.ErrorDetails.DuplicateEntity']
                        // to filter the logicalname tags which has entity name
                        var sliceAttributestag = xmlDoc.querySelectorAll("Attributes");
                        var getEntityTag = Array.from(sliceAttributestag);
                        getEntityTag.forEach(function (x) { return x.remove(); });
                        var xmlSerializer = new XMLSerializer();
                        duplicateEntityAnnotation = xmlSerializer.serializeToString(xmlDoc);
                    }
                }
                return duplicateEntityAnnotation;
            };
            this.handleLeadDuplication = function (parentAccountId, parentContactId, parentAccountName, parentContactName, qualifyStatus, transactionCurrencyId, emailAddress) {
                if (_this.duplicateDetectionAdminSettings) {
                    var options_2 = {
                        height: 600,
                        width: 800,
                        position: 1 /* center */
                    };
                    var dialogArguments_2 = {};
                    var parentEntityId = Xrm.Page.data.entity.getId();
                    dialogArguments_2[ClientUtility.MetadataDrivenDialogConstants.EntityId] = parentEntityId;
                    // As above hiddencontrols are not supported in UCI, setting formparamters as well
                    dialogArguments_2[ClientUtility.MetadataDrivenDialogConstants.paramEntityId] = Xrm.Page.data.entity.getId();
                    dialogArguments_2[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId] = parentAccountId;
                    dialogArguments_2[MarketingSales.MetadataDrivenDialogConstants.paramParentContactId] = parentContactId;
                    dialogArguments_2[MarketingSales.MetadataDrivenDialogConstants.paramTransactionCurrencyId] = transactionCurrencyId;
                    dialogArguments_2[MarketingSales.MetadataDrivenDialogConstants.paramQualifyStatus] = qualifyStatus;
                    dialogArguments_2[MarketingSales.MetadataDrivenDialogConstants.paramEmailAddress] = emailAddress;
                    var recordsToFetch = [];
                    var objectAccount = {
                        parentEntity_Id: parentEntityId,
                        EntityName: MarketingSales.EntityNames.Account,
                    };
                    recordsToFetch.push(objectAccount);
                    var objectContact = {
                        parentEntity_Id: parentEntityId,
                        EntityName: MarketingSales.EntityNames.Contact,
                    };
                    recordsToFetch.push(objectContact);
                    //Get Account and contact entity from lead, where duplicates are triggered.
                    _this.getAssociatedEntityforDuplicate(recordsToFetch).then(function (response) {
                        if (response.length > 0) {
                            _this.duplicateAccount = response[0];
                            _this.duplicateContact = response[1];
                            dialogArguments_2[MarketingSales.MetadataDrivenDialogConstants.duplicateAccount] = _this.duplicateAccount;
                            dialogArguments_2[MarketingSales.MetadataDrivenDialogConstants.duplicateContact] = _this.duplicateContact;
                        }
                        Xrm.Navigation.openDialog(MarketingSales.DialogName.DuplicateQualifyLeadDialog, options_2, dialogArguments_2)
                            .then((_this.closeDupWarningDialogCallback));
                    }, function (error) {
                        Xrm.Utility.alertDialog(MarketingSales.ResourceStringProvider.getResourceString("RelatedEntities_Fetch_Fail_Message"), null);
                    });
                }
                else {
                    var options = {
                        height: 285,
                        width: 600,
                        position: 1 /* center */
                    };
                    var dialogArguments = {};
                    dialogArguments[ClientUtility.MetadataDrivenDialogConstants.EntityId] = Xrm.Page.data.entity.getId();
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.ParentAccountId] = parentAccountId;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.ParentAccountName] = parentAccountName;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.ParentContactId] = parentContactId;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.ParentContactName] = parentContactName;
                    // As above hiddencontrols are not supported in UCI, setting formparamters as well
                    dialogArguments[ClientUtility.MetadataDrivenDialogConstants.paramEntityId] = Xrm.Page.data.entity.getId();
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountId] = parentAccountId;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.paramParentAccountName] = parentAccountName;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.paramParentContactId] = parentContactId;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.paramParentContactName] = parentContactName;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.paramTransactionCurrencyId] = transactionCurrencyId;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.paramQualifyStatus] = qualifyStatus;
                    dialogArguments[MarketingSales.MetadataDrivenDialogConstants.paramEmailAddress] = emailAddress;
                    Xrm.Navigation.openDialog(MarketingSales.DialogName.DupWarningDialog, options, dialogArguments)
                        .then((_this.closeDupWarningDialogCallback));
                }
            };
            this.qualifyLeadAsHelper = function (qualifyStatus) {
                _this.getTransactionCurrencyId(function (transactionCurrencyId) {
                    _this.qualifyLeadSdkCall(qualifyStatus, transactionCurrencyId, false, null, null);
                }, function (errorResponse) { MarketingSales.LeadErrorHandling.actionFailedCallback(errorResponse); });
            };
            this.getLookupValue = function (lookupId) {
                // UCI dialogs don't have entity context
                var attribute = ClientUtility.DataUtil.isNullOrUndefined(Xrm.Page.data.entity) ? Xrm.Page.data.attributes.get(lookupId) : Xrm.Page.getAttribute(lookupId);
                if (!ClientUtility.DataUtil.isNullOrUndefined(attribute)) {
                    var attributeValue = attribute.getValue();
                    if (!ClientUtility.DataUtil.isNullOrUndefined(attributeValue)) {
                        return attributeValue[0];
                    }
                }
                return null;
            };
            this.getLookupId = function (lookup) {
                var id = null;
                if (!ClientUtility.DataUtil.isNullOrUndefined(lookup)) {
                    id = lookup.id;
                }
                return id;
            };
            this.setLookupValue = function (lookupId, lookupValueName, lookupValueId, entityLogicalName, isInitialLoad) {
                // UCI dialogs don't have entity context
                var attribute = ClientUtility.DataUtil.isNullOrUndefined(Xrm.Page.data.entity) ? Xrm.Page.data.attributes.get(lookupId) : Xrm.Page.getAttribute(lookupId);
                if (!ClientUtility.DataUtil.isNullOrUndefined(attribute) && !ClientUtility.DataUtil.isNullOrEmptyString(lookupValueId)) {
                    if (!ClientUtility.DataUtil.isNullOrUndefined(attribute.getValue())) {
                        if (attribute.getValue()[0].id === lookupValueId) {
                            return;
                        }
                    }
                    var lookupObject = { entityType: entityLogicalName, id: lookupValueId, name: lookupValueName };
                    var lookupObjects = [lookupObject];
                    attribute.setValue(lookupObjects);
                }
            };
            this.getStatusFromParameter = function (parameter) {
                var status = -1;
                if (typeof parameter === "object" && parameter.SourceControlId) {
                    var parts = parameter.SourceControlId.split('|');
                    status = parseInt(parts[parts.length - 1]) || -1;
                }
                else {
                    status = parseInt(parameter) || -1;
                }
                return status;
            };
            this.AppCommonNullCheck = function () {
                return ((typeof (AppCommon) !== 'undefined') &&
                    (AppCommon != null && AppCommon != undefined) &&
                    (AppCommon.TelemetryReporter != null && AppCommon.TelemetryReporter != undefined) &&
                    (AppCommon.TelemetryReporter.Instance() != null && AppCommon.TelemetryReporter.Instance() != undefined));
            };
            this.selectedRecordsChange = function (context) {
                var formContext = context.getFormContext();
                var okButton = formContext.getControl("ok_id");
                var ignore_SaveButton = formContext.getControl("ignore_save");
                var selectedContactId = formContext.data.attributes.get(MarketingSales.MetadataDrivenDialogConstants.duplicateContactId);
                var selectedAccountId = formContext.data.attributes.get(MarketingSales.MetadataDrivenDialogConstants.duplicateAccountId);
                var selectedAccountRecordsArray = [];
                if (selectedAccountId && selectedAccountId.getValue()) {
                    selectedAccountRecordsArray = JSON.parse(selectedAccountId.getValue());
                }
                var selectedContactRecordsArray = [];
                if (selectedContactId && selectedContactId.getValue()) {
                    selectedContactRecordsArray = JSON.parse(selectedContactId.getValue());
                }
                if (okButton) {
                    //If user selects more than one record for merge , then disable "Continue" button.
                    if ((Array.isArray(selectedAccountRecordsArray) && selectedAccountRecordsArray.length > 1) ||
                        (Array.isArray(selectedContactRecordsArray) && selectedContactRecordsArray.length > 1)) {
                        okButton.setDisabled(true);
                        ignore_SaveButton.setDisabled(true);
                    }
                    else if ((Array.isArray(selectedAccountRecordsArray) && selectedAccountRecordsArray.length == 1) ||
                        (Array.isArray(selectedContactRecordsArray) && selectedContactRecordsArray.length == 1)) {
                        okButton.setDisabled(false);
                        ignore_SaveButton.setDisabled(true);
                    }
                    else {
                        if ((selectedAccountRecordsArray && selectedAccountRecordsArray.length === 0) &&
                            (selectedContactRecordsArray && selectedContactRecordsArray.length === 0)) {
                            okButton.setDisabled(true);
                            ignore_SaveButton.setDisabled(false);
                        }
                    }
                }
            };
            // private logStatusUpdate(leadId: string, methodName: string, closeEventName: string, parameters?: any): void {
            // 	Xrm.WebApi.online.retrieveRecord(
            // 		EntityNames.Lead,
            // 		leadId,
            // 		"?$select=createdon"
            // 	)
            // 	.then((result: any) => {
            // 		let timeTakenSinceCreate = <any>new Date() - <any>new Date(result.createdon);
            // 		timeTakenSinceCreate /= 1000 * 60 * 60;
            // 		LeadUtil.logTelemetryData(
            // 			"LeadStatusUpdate",
            // 			methodName,
            // 			MetadataDrivenDialogConstants.information,
            // 			closeEventName,
            // 			{
            // 				leadId: leadId.replace(/[{}]/g, "").toLowerCase(),
            // 				timeTakenSinceCreate,
            // 				...parameters
            // 			}
            // 		);
            // 	});
            // }
        }
        //Get Account and contact entity from lead, where duplicates are triggered.
        LeadCommandActionsImplementation.prototype.getAssociatedEntityforDuplicate = function (records) {
            var _this = this;
            return new Promise(function (resolve, reject) {
                var getEntityrequestBatch = records.map(function (record) {
                    var leadEntity = { entityType: MarketingSales.EntityNames.Lead, id: record.parentEntity_Id };
                    return new MarketingSales.GetAssociatedEntityRequest(leadEntity, record.EntityName, MarketingSales.TargetFieldType.ValidForCreate);
                });
                Xrm.WebApi.online.executeMultiple(getEntityrequestBatch)
                    .then(function (getEntityResponse) {
                        var records = [];
                        getEntityResponse[0].json().then(function (response_account) {
                            //This fields are not used in fetching duplicate records , but casuses errors if values are null
                            //So remove this fields.
                            delete response_account["ownerid@odata.bind"];
                            delete response_account["originatingleadid@odata.bind"];
                            delete response_account["transactioncurrencyid@odata.bind"];
                            if (MarketingSales.LeadUtil.IsFcbEnabled("QualifyLeadDeleteLookupForDupDetection")) {
                                _this.deleteLookupAttributes(response_account, MarketingSales.EntityNames.Account);
                            }
                            records.push(response_account);
                            MarketingSales.TelemetryReporter.ReportInfo(null, "getAssociatedEntityforDuplicate", MarketingSales.DialogName.DuplicateQualifyLeadDialog, null, null, "Related entity account records fetched successfully", true);
                        });
                        getEntityResponse.length > 1 && getEntityResponse[1].json().then(function (response_contact) {
                            delete response_contact["ownerid@odata.bind"];
                            delete response_contact["originatingleadid@odata.bind"];
                            delete response_contact["transactioncurrencyid@odata.bind"];
                            if (MarketingSales.LeadUtil.IsFcbEnabled("QualifyLeadDeleteLookupForDupDetection")) {
                                _this.deleteLookupAttributes(response_contact, MarketingSales.EntityNames.Contact);
                            }
                            records.push(response_contact);
                            MarketingSales.TelemetryReporter.ReportInfo(null, "getAssociatedEntityforDuplicate", MarketingSales.DialogName.DuplicateQualifyLeadDialog, null, null, "Related entity contact records fetched successfully", true);
                        });
                        resolve(records);
                    }, function (error) {
                        reject(null);
                        MarketingSales.TelemetryReporter.ReportError(null, "getAssociatedEntityforDuplicate", MarketingSales.ResourceStringProvider.getResourceString("RelatedEntities_Fetch_Fail_Message"), MarketingSales.DialogName.DuplicateQualifyLeadDialog);
                    });
            });
        };
        /**
         * delete a lookup attribute, duplicate contact was not showing, it must show, removing a attribute.
         * @param {String} response List of attributes of entities
         * @param {String} EntityName name of the Entity like Account/Contact
         */
        LeadCommandActionsImplementation.prototype.deleteLookupAttributes = function (response, EntityName) {
            try {
                var _loop_1 = function (element) {
                    var attrElementName;
                    resultStart = element.startsWith('_');
                    if (resultStart) {
                        resultEnd = element.endsWith('_value');
                        if (resultEnd) {
                            attrElementName = element;
                            element = element.slice(1);
                            element = element.replace("_value", "");
                            Xrm.Utility.getEntityMetadata(EntityName, [element]).then(function (entityMetadata) {
                                if (entityMetadata.Attributes != null) {
                                    var attr = entityMetadata.Attributes.get(element);
                                    var attrDescriptor = attr && attr.attributeDescriptor;
                                    if (attrDescriptor != null && attrDescriptor.Type === Xrm.Constants.AttributeTypes.lookupType) {
                                        delete response[attrElementName];
                                    }
                                }
                            });
                        }
                    }
                };
                var resultStart, resultEnd;
                for (var element in response) {
                    _loop_1(element);
                }
            }
            catch (e) {
                MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "deleteAttribute", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
            }
        };
        LeadCommandActionsImplementation.prototype.getParameterFromContext = function (paramName) {
            var selectedRecordArray = [];
            var selectedRecordAttribute = Xrm.Page.data.attributes.get(paramName);
            if (selectedRecordAttribute && selectedRecordAttribute.getValue()) {
                selectedRecordArray = JSON.parse(selectedRecordAttribute.getValue());
            }
            if (Array.isArray(selectedRecordArray) && selectedRecordArray.length > 0) {
                return selectedRecordArray[0]; // If more than one record is selected by user first record from list will be set for merge.
            }
        };
        return LeadCommandActionsImplementation;
    }());
    LeadCommandActionsImplementation.isLeadQualified = false;
    LeadCommandActionsImplementation.noDuplicatesFound = "No duplicates found";
    LeadCommandActionsImplementation.failedToFetchRelatedEntities = "Failed to fetch related entities";
    LeadCommandActionsImplementation.getDropDownSelection = function (createEntity, recordsLength) {
        if (createEntity)
            return "CreateNew";
        return recordsLength > 0 ? "UseExisting" : "None";
    };
    MarketingSales.LeadCommandActionsImplementation = LeadCommandActionsImplementation;
})(MarketingSales || (MarketingSales = {}));
/**
 * @license Copyright (c) Microsoft Corporation. All rights reserved.
 */
/**
 * IMPORTANT!
 * DO NOT MAKE CHANGES TO THIS FILE - THIS FILE IS AUTO-GENERATED FROM ODATA CSDL METADATA DOCUMENT
 * SEE https://msdn.microsoft.com/en-us/library/mt607990.aspx FOR MORE INFORMATION
 */
/**
 * IMPORTANT!
 * THIS FILE HAS BEEN MANUALLY MODIFIED TO ENABLE OPTIONAL PARAMETERS FOR THE REQUEST.
 * OPTIONAL PARAMETERS SHOULD BE IGNORED IN ORDER TO AVOID SERVER-SIDE DESERIALIZATION ISSUES.
 */
var MarketingSales;
(function (MarketingSales) {
    /* tslint:disable:crm-force-fields-private */
    var QualifyLeadRequest = (function () {
        function QualifyLeadRequest(entity /*Microsoft.Dynamics.CRM.lead*/, createAccount, createContact, createOpportunity, opportunityCurrencyId /*Microsoft.Dynamics.CRM.transactioncurrency*/, opportunityCustomerId /*Microsoft.Dynamics.CRM.crmbaseentity*/, sourceCampaignId /*Microsoft.Dynamics.CRM.campaign*/, status, processInstanceId /*Microsoft.Dynamics.CRM.crmbaseentity*/, suppressDuplicateDetection, leadQualificationV2Param) {
            this.entity = entity;
            this.CreateAccount = createAccount;
            this.CreateContact = createContact;
            this.CreateOpportunity = createOpportunity;
            this.OpportunityCurrencyId = opportunityCurrencyId;
            this.OpportunityCustomerId = opportunityCustomerId;
            this.Status = status;
            if (ClientUtility.DataUtil.isNullOrUndefined(suppressDuplicateDetection)) {
                this.SuppressDuplicateDetection = false;
            }
            else {
                this.SuppressDuplicateDetection = suppressDuplicateDetection;
            }
            if (!ClientUtility.DataUtil.isNullOrUndefined(sourceCampaignId)) {
                this.SourceCampaignId = sourceCampaignId;
            }
            if (!ClientUtility.DataUtil.isNullOrUndefined(processInstanceId) && !ClientUtility.DataUtil.isNullOrUndefined(processInstanceId.id)) {
                this.ProcessInstanceId = processInstanceId;
            }
            if (ClientUtility.DataUtil.isNullOrUndefined(leadQualificationV2Param) || leadQualificationV2Param.length == 0) {
                this.LeadQualificationV2Param = null;
            }
            else {
                this.LeadQualificationV2Param = leadQualificationV2Param;
            }
        }
        QualifyLeadRequest.prototype.getMetadata = function () {
            var metadata = {
                boundParameter: "entity",
                parameterTypes: {
                    "entity": {
                        "typeName": "Microsoft.Dynamics.CRM.lead",
                        "structuralProperty": 5,
                    },
                    "CreateAccount": {
                        "typeName": "Edm.Boolean",
                        "structuralProperty": 1,
                    },
                    "CreateContact": {
                        "typeName": "Edm.Boolean",
                        "structuralProperty": 1,
                    },
                    "CreateOpportunity": {
                        "typeName": "Edm.Boolean",
                        "structuralProperty": 1,
                    },
                    "SuppressDuplicateDetection": {
                        "typeName": "Edm.Boolean",
                        "structuralProperty": 1,
                    },
                    "OpportunityCurrencyId": {
                        "typeName": "Microsoft.Dynamics.CRM.transactioncurrency",
                        "structuralProperty": 5,
                    },
                    "OpportunityCustomerId": {
                        "typeName": "Microsoft.Dynamics.CRM.crmbaseentity",
                        "structuralProperty": 5,
                    },
                    "SourceCampaignId": {
                        "typeName": "Microsoft.Dynamics.CRM.campaign",
                        "structuralProperty": 5,
                    },
                    "Status": {
                        "typeName": "Edm.Int32",
                        "structuralProperty": 1,
                    },
                    "ProcessInstanceId": {
                        "typeName": "Microsoft.Dynamics.CRM.crmbaseentity",
                        "structuralProperty": 5,
                    },
                    "LeadQualificationV2Param": {
                        "typeName": "Edm.String",
                        "structuralProperty": 1,
                    },
                },
                operationName: "QualifyLead",
                operationType: 0,
            };
            return metadata;
        };
        return QualifyLeadRequest;
    }());
    MarketingSales.QualifyLeadRequest = QualifyLeadRequest;
})(MarketingSales || (MarketingSales = {}));
var MarketingSales;
(function (MarketingSales) {
    var NSATScenarios;
    (function (NSATScenarios) {
        NSATScenarios[NSATScenarios["QualifyLead"] = 0] = "QualifyLead";
        NSATScenarios[NSATScenarios["DisqualifyLead"] = 1] = "DisqualifyLead";
        NSATScenarios[NSATScenarios["ReactivateLead"] = 2] = "ReactivateLead";
    })(NSATScenarios = MarketingSales.NSATScenarios || (MarketingSales.NSATScenarios = {}));
    ;
    var NSATSurvey = (function () {
        function NSATSurvey() {
        }
        //Invoke NSAT survey dialog
        NSATSurvey.getNSATSurvey = function (entityName, scenarioName) {
            try {
                Xrm.Internal.showSurvey(NSATScenarios[scenarioName]); // new NSAT survey API
                var eventParams = [
                    { name: NSATSurvey.EntityName, value: entityName },
                    { name: NSATSurvey.APIName, value: 'showSurvey' },
                    { name: NSATSurvey.APIScenario, value: NSATScenarios[scenarioName] }
                ];
                Xrm.Reporting.reportSuccess(NSATSurvey.ComponentName, eventParams);
            }
            catch (error) {
                Xrm.Reporting.reportFailure(NSATSurvey.ComponentName, error);
            }
        };
        return NSATSurvey;
    }());
    NSATSurvey.ComponentName = "NSATSurvey";
    NSATSurvey.EntityName = "EntityName";
    NSATSurvey.APIName = "API";
    NSATSurvey.APIScenario = "ApiScenario";
    MarketingSales.NSATSurvey = NSATSurvey;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../../../TypeDefinitions/CRM/ClientUtility.d.ts" />
/// <reference path="../../../TypeDefinitions/LeadManagement/Lead/Lead_main_system_library.d.ts" />
/// <reference path="Contracts/QualifyLeadRequest.ts" />
/// <reference path="LeadUtil.ts" />
/// <reference path="PopulateDynamicMenu.ts" />
/// <reference path ="LeadErrorHandling.ts" />
/// <reference path ="NSATSurvey.ts" />
var MarketingSales;
(function (MarketingSales) {
    var LeadGridCommandActions = (function () {
        function LeadGridCommandActions() {
            var leadGridCommandsActions = new LeadGridCommandActionsImplementation();
            var global = window;
            var mscrm = global.Mscrm || {};
            mscrm.LeadGridCommandActions = mscrm.LeadGridCommandActions || {};
            mscrm.LeadGridCommandActions.qualifyLeadQuick = leadGridCommandsActions.qualifyLeadQuick;
            mscrm.LeadGridCommandActions.qualifyLead = leadGridCommandsActions.qualifyLead;
            mscrm.LeadGridCommandActions.qualifyLeadAs = leadGridCommandsActions.qualifyLeadAs;
            mscrm.LeadGridCommandActions.disqualifyLeadQuick = leadGridCommandsActions.disqualifyLeadQuick;
            mscrm.LeadGridCommandActions.disqualifyLeadAs = leadGridCommandsActions.disqualifyLeadAs;
            mscrm.LeadGridCommandActions.reactivcateLead = leadGridCommandsActions.reactivateLead;
        }
        return LeadGridCommandActions;
    }());
    MarketingSales.LeadGridCommandActions = LeadGridCommandActions;
    var LeadGridCommandActionsImplementation = (function () {
        function LeadGridCommandActionsImplementation() {
            var _this = this;
            this.createContactFlag = true;
            this.createOpportunityFlag = false;
            this.createAccountFlag = true;
            this.adminSettings = true;
            this.isGrid = true;
            this.fcbOctoberUpdate = false;
            this.maximumConcurrentRequestsLimit = 52;
            this.qualifyLeadQuick = function (gridControl, records, entityTypeCode) {
                return __awaiter(_this, void 0, void 0, function () {
                    var _a;
                    return __generator(this, function (_b) {
                        switch (_b.label) {
                            case 0:
                                this.records = records;
                                this.entityTypeCode = entityTypeCode;
                                _a = this;
                                return [4 /*yield*/, MarketingSales.LeadUtil.isQualifyLeadV2Enabled()];
                            case 1:
                                _a.isQualifyLeadV2Enabled = _b.sent();
                                if (this.isQualifyLeadV2Enabled) {
                                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "qualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "GridQualifyButtonClicked", {
                                        recordsCount: records.length
                                    });
                                    this.qualifyLead(gridControl, records, entityTypeCode);
                                }
                                else {
                                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "qualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "GridQualifyButtonClicked");
                                    this.qualifyLead(gridControl, records, entityTypeCode);
                                }
                                return [2 /*return*/];
                        }
                    });
                });
            };
            this.qualifyLead = function (gridControl, records, entityTypeCode) {
                _this.qualifyLeadAs(gridControl, records, entityTypeCode, -1);
            };
            this.qualifyLeadV2 = function (gridControl, records, entityTypeCode) {
                _this.qualifyLeadAsV2(gridControl, records, entityTypeCode, -1, null);
            };
            this.qualifyLeadAs = function (gridControl, records, entityTypeCode, statusCode) {
                try {
                    _this.fcbOctoberUpdate = MarketingSales.LeadUtil.FCBOctoberUpdateEnabled();
                    _this.adminSettings = MarketingSales.LeadUtil.qualifyLeadAdditionalOptionsAdminSettings();
                    if ((!_this.adminSettings && _this.fcbOctoberUpdate) || (_this.isQualifyLeadV2Enabled && records.length >= 2)) {
                        MarketingSales.LeadUtil.convertLeadDialog(_this.isGrid, records.length).then(function (dialogParams) {
                            if (!ClientUtility.DataUtil.isNullOrUndefined(dialogParams) &&
                                !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters) &&
                                dialogParams.parameters[ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked] === ClientUtility.MetadataDrivenDialogConstants.DialogOkId) {
                                _this.createContactFlag = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.contactFlag] ? true : false;
                                _this.createAccountFlag = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.accountFlag] ? true : false;
                                _this.createOpportunityFlag = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.opportunityFlag] ? true : false;
                                var parameters = {};
                                parameters["CreateAccount"] = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.accountFlag];
                                parameters["CreateContact"] = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.contactFlag];
                                parameters["CreateOpportunity"] = dialogParams.parameters[MarketingSales.MetadataDrivenDialogConstants.opportunityFlag];
                                parameters["isGrid"] = _this.isGrid;
                                parameters["recordsCount"] = records.length;
                                MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "", parameters);
                                _this.qualifyLeadAsGridAction(gridControl, records, entityTypeCode, statusCode);
                            }
                        });
                    }
                    else {
                        _this.qualifyLeadAsGridAction(gridControl, records, entityTypeCode, statusCode);
                    }
                }
                catch (e) {
                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialog, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
                }
            };
            this.qualifyLeadAsV2 = function (gridControl, records, entityTypeCode, statusCode, dialogArguments) {
                try {
                    MarketingSales.LeadUtil.convertLeadDialogV2(_this.isGrid, records[0].Id, statusCode, dialogArguments).then(function (dialogParams) {
                        if (!ClientUtility.DataUtil.isNullOrUndefined(dialogParams) &&
                            !ClientUtility.DataUtil.isNullOrUndefined(dialogParams.parameters) &&
                            dialogParams.parameters[ClientUtility.MetadataDrivenDialogConstants.paramLastButtonClicked] === ClientUtility.MetadataDrivenDialogConstants.DialogOkId) {
                            var parameters = {};
                            MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.information, "", parameters);
                            _this.qualifyLeadAsGridAction(gridControl, records, entityTypeCode, statusCode);
                        }
                    });
                }
                catch (e) {
                    MarketingSales.LeadUtil.logTelemetryData(MarketingSales.DialogName.convertLeadDialogV2, "QualifyLeadQuick", MarketingSales.MetadataDrivenDialogConstants.error, e.toString());
                }
            };
            this.qualifyLeadAsGridAction = function (gridControl, records, entityTypeCode, statusCode) {
                if (ClientUtility.ClientUtil.isMobileOffline() &&
                    !Xrm.Mobile.offline.isOfflineEnabled(MarketingSales.EntityNames.Lead)) {
                    ClientUtility.DialogUtil.showMoCAOfflineError();
                    return;
                }
                if (records.length > _this.maximumConcurrentRequestsLimit) {
                    Xrm.Utility.alertDialog(MarketingSales.ResourceStringProvider.getResourceString("RecordsSelectedExceededLimitError"));
                    return;
                }
                var columnNames = null;
                // lookup column names has to be set differently for offline and online
                if (ClientUtility.ClientUtil.isMobileOffline()) {
                    columnNames = ClientUtility.ODataUtil.getSelectOption(["parentaccountid", "parentcontactid", "companyname", "campaignid", "transactioncurrencyid", "leadid", "createdon"]);
                }
                else {
                    columnNames = ClientUtility.ODataUtil.getSelectOption(["_parentaccountid_value", "_parentcontactid_value", "companyname", "_campaignid_value", "_transactioncurrencyid_value", "leadid", "createdon"]);
                }
                var recordCounter = 0;
                var progressIndicator = new ClientUtility.ProgressIndicator();
                progressIndicator.show();
                for (var i = 0; i < records.length; i++) {
                    var record = records[i];
                    Xrm.WebApi.retrieveRecord(MarketingSales.EntityNames.Lead, record.Id, columnNames)
                        .then(function (leadRecord) {
                            recordCounter++;
                            var transactionCurrencyId = leadRecord["_transactioncurrencyid_value"];
                            if (ClientUtility.ClientUtil.isMobileOffline() || ClientUtility.DataUtil.isNullOrUndefined(transactionCurrencyId)) {
                                transactionCurrencyId = Xrm.Utility.getGlobalContext().userSettings.transactionCurrencyId;
                                transactionCurrencyId = transactionCurrencyId || Xrm.Utility.getGlobalContext().organizationSettings.baseCurrencyId;
                            }
                            if (Xrm.Internal.isUci() && typeof statusCode === "object" && statusCode.SourceControlId) {
                                statusCode = _this.getStatusFromParameter(statusCode, recordCounter);
                            }
                            _this.qualifyLeadSdkCall(leadRecord, transactionCurrencyId, statusCode, recordCounter, gridControl, records.length, progressIndicator);
                        }, progressIndicator.hideOnError(_this.qualifyLeadFailedCallback));
                }
            };
            this.qualifyLeadSdkCall = function (leadRecord, transactionCurrencyId, statusCode, recordCounter, gridControl, numberOfRecords, progressIndicator) {
                if (ClientUtility.DataUtil.isNullOrUndefined(leadRecord)) {
                    throw "Response cannot be null";
                }
                var createAccount = false;
                var createContact = true;
                var createOpportunity = true;
                var leadReference = { entityType: MarketingSales.EntityNames.Lead, id: ClientUtility.DataUtil.toStringWithNullCheck(leadRecord["leadid"]) };
                var sourceCampaignId = leadRecord["_campaignid_value"];
                var account = leadRecord["_parentaccountid_value"];
                var contact = leadRecord["_parentcontactid_value"];
                var opportunityCustomer = null;
                var companyName = null;
                var request = null;
                var transactionCurrencyReference = ClientUtility.DataUtil.isNullOrWhiteSpace(transactionCurrencyId) ?
                    null :
                    { entityType: MarketingSales.EntityNames.TransactionCurrency, id: ClientUtility.Guid.create(transactionCurrencyId) };
                var sourceCampaignReference = ClientUtility.DataUtil.isNullOrWhiteSpace(sourceCampaignId) ?
                    null :
                    { entityType: MarketingSales.EntityNames.Campaign, id: ClientUtility.Guid.create(sourceCampaignId) };
                if (!ClientUtility.DataUtil.isNullOrUndefined(leadRecord["companyname"])) {
                    companyName = ClientUtility.DataUtil.toStringWithNullCheck(leadRecord["companyname"]);
                    createAccount = true;
                }
                if (!ClientUtility.DataUtil.isNullOrUndefined(account) && !ClientUtility.DataUtil.isNullOrUndefined(contact)) {
                    createAccount = false;
                    createContact = false;
                    opportunityCustomer = { entityType: MarketingSales.EntityNames.Account, id: ClientUtility.Guid.create(account) };
                }
                else {
                    if (!ClientUtility.DataUtil.isNullOrUndefined(account)) {
                        createAccount = false;
                        createContact = true;
                    }
                    else {
                        if (!ClientUtility.DataUtil.isNullOrUndefined(contact)) {
                            if (!ClientUtility.DataUtil.isNullOrUndefined(companyName)) {
                                createAccount = true;
                                createContact = false;
                            }
                            else {
                                createAccount = false;
                                createContact = false;
                                opportunityCustomer = { entityType: MarketingSales.EntityNames.Contact, id: ClientUtility.Guid.create(contact) };
                            }
                        }
                    }
                }
                if (!_this.adminSettings && _this.fcbOctoberUpdate) {
                    if (!ClientUtility.DataUtil.isNullOrUndefined(_this.createContactFlag) && _this.createContactFlag == false)
                        createContact = _this.createContactFlag;
                    if (!ClientUtility.DataUtil.isNullOrUndefined(_this.createAccountFlag) && _this.createAccountFlag == false)
                        createAccount = _this.createAccountFlag;
                    if (!ClientUtility.DataUtil.isNullOrUndefined(_this.createOpportunityFlag) && _this.createOpportunityFlag == false)
                        createOpportunity = _this.createOpportunityFlag;
                }
                request = new MarketingSales.QualifyLeadRequest(leadReference, createAccount, createContact, createOpportunity, transactionCurrencyReference || {}, opportunityCustomer || {}, sourceCampaignReference || {}, statusCode);
                if (ClientUtility.ClientUtil.isMobileOffline()) {
                    request.OpportunityCurrencyId = Xrm.Utility.getGlobalContext().userSettings.transactionCurrency;
                    var processStage = null;
                    var currentStage = null;
                    var nextStage = null;
                    if (!ClientUtility.DataUtil.isNullOrUndefined(Xrm.Page.data)) {
                        processStage = Xrm.Page.data.process.getActivePath();
                        currentStage = processStage.get(ClientUtility.MetadataDrivenDialogConstants.defaultIndex);
                        nextStage = processStage.get(ClientUtility.MetadataDrivenDialogConstants.firstIndex);
                    }
                    var nextStageId = "";
                    var nextTraversedPath = "";
                    if (!ClientUtility.DataUtil.isNullOrUndefined(currentStage) && !ClientUtility.DataUtil.isNullOrUndefined(nextStage)) {
                        var currentStageId = currentStage.getId();
                        nextStageId = nextStage.getId();
                        nextTraversedPath = currentStageId + ClientUtility.MetadataDrivenDialogConstants.commaSeperator + nextStageId;
                    }
                    // as in outlook offline, creation of duplicate entries should be allowed even though duplicate detection is enabled.
                    request.SuppressDuplicateDetection = true;
                    Xrm.WebApi.offline.execute(request, { "EntityLogicalName": MarketingSales.EntityNames.Lead, "NextStageId": nextStageId, "TraversedPath": nextTraversedPath }).then(function () {
                        _this.triggerGridRefreshOnLimitReached(gridControl, recordCounter, numberOfRecords, progressIndicator);
                    }, progressIndicator.hideOnError(_this.qualifyLeadFailedCallback));
                }
                else {
                    if (_this.isQualifyLeadV2Enabled && _this.records.length < 2) {
                        progressIndicator.hide();
                        var qualifyLeadParams = {
                            leadReference: leadReference,
                            createAccount: createAccount,
                            createContact: createContact,
                            createOpportunity: createOpportunity,
                            transactionCurrencyReference: transactionCurrencyReference || {},
                            opportunityCustomerReference: opportunityCustomer || {},
                            sourceCampaignReference: sourceCampaignReference || {},
                            qualifyStatus: statusCode,
                            processInstanceReference: MarketingSales.LeadUtil.retrieveProcessInstanceId(),
                            isGrid: true,
                        };
                        var dialogArguments = {
                            parameters: (_a = {},
                                _a[MarketingSales.MetadataDrivenDialogConstants.paramQualifyLead] = qualifyLeadParams,
                                _a[MarketingSales.MetadataDrivenDialogConstants.paramRecordCount] = _this.records.length,
                                _a)
                        };
                        _this.qualifyLeadAsV2(gridControl, _this.records, _this.entityTypeCode, -1, dialogArguments);
                    }
                    else {
                        // When leads are qualified from the main grid, customers are not warned for duplicate contacts/accounts.
                        request.SuppressDuplicateDetection = true;
                        Xrm.WebApi.online.execute(request).then(function () {
                            _this.triggerGridRefreshOnLimitReached(gridControl, recordCounter, numberOfRecords, progressIndicator);
                            // Trigger NSAT survey once lead qualification is complete
                            MarketingSales.NSATSurvey.getNSATSurvey("lead", MarketingSales.NSATScenarios.QualifyLead);
                            // this.logStatusUpdate(
                            // 	leadRecord["leadid"],
                            // 	leadRecord["createdon"],
                            // 	"LeadGridCommandActions.qualifyV2ClickHandler",
                            // 	"Qualify",
                            // 	{ version: "v1" }
                            // );
                            var leadData = {
                                leadid: leadRecord["leadid"],
                                createdon: leadRecord["createdon"]
                            };
                            var leadDataJsonString = JSON.stringify(leadData);
                            MarketingSales.SalesAgentUsageClient.logSalesAgentUsage("LeadQualifyEvent", leadDataJsonString);
                        }, progressIndicator.hideOnError(_this.qualifyLeadFailedCallback));
                    }
                }
                var _a;
            };
            this.disqualifyLeadQuick = function (gridControl, records, entityTypeCode) {
                _this.disqualifyLeadAs(gridControl, records, entityTypeCode, null, -1);
            };
            this.disqualifyLeadAs = function (gridControl, records, entityTypeCode, commandProperties, statusCode) {
                var status = _this.getStatusFromParameter(commandProperties, statusCode);
                _this.setStateAndStatus(gridControl, records, status, LeadManagement.LeadState.Disqualified);
            };
            this.reactivateLead = function (gridControl, records, entityTypeCode) {
                _this.setStateAndStatus(gridControl, records, LeadManagement.LeadStatus.New, LeadManagement.LeadState.Open);
            };
            this.setStateAndStatus = function (gridControl, records, status, state) {
                if (ClientUtility.ClientUtil.isMobileOffline() && !Xrm.Mobile.offline.isOfflineEnabled(MarketingSales.EntityNames.Lead)) {
                    ClientUtility.DialogUtil.showMoCAOfflineError();
                    return;
                }
                if (records.length > _this.maximumConcurrentRequestsLimit) {
                    Xrm.Utility.alertDialog(MarketingSales.ResourceStringProvider.getResourceString("RecordsSelectedExceededLimitError"));
                    return;
                }
                var recordCounter = 0;
                var updateRecord = {
                    statecode: state,
                    statuscode: status
                };
                var progressIndicator = new ClientUtility.ProgressIndicator();
                progressIndicator.show();
                var promises = [];
                var successfullyUpdatedRecords = [];
                var _loop_2 = function (i) {
                    promises.push(Xrm.WebApi.updateRecord(MarketingSales.EntityNames.Lead, records[i].Id, updateRecord).then(function () {
                        recordCounter++;
                        _this.triggerGridRefreshOnLimitReached(gridControl, recordCounter, records.length, progressIndicator);
                        successfullyUpdatedRecords.push(records[i]);
                    }, function (errorResponse) {
                        progressIndicator.hide();
                        MarketingSales.LeadErrorHandling.actionFailedErrorDialog(errorResponse);
                    }));
                };
                for (var i = 0; i < records.length; i++) {
                    _loop_2(i);
                }
                switch (state) {
                    case LeadManagement.LeadState.Disqualified:
                        // Trigger NSAT survey once lead is disqualified
                        MarketingSales.NSATSurvey.getNSATSurvey("lead", MarketingSales.NSATScenarios.DisqualifyLead);
                        break;
                    case LeadManagement.LeadState.Open:
                        // Trigger NSAT survey once lead is re-activated
                        MarketingSales.NSATSurvey.getNSATSurvey("lead", MarketingSales.NSATScenarios.ReactivateLead);
                        break;
                    default:
                        break;
                }
                Promise.all(promises);
                // .then(() => {
                // 	if (successfullyUpdatedRecords.length > 0) {
                // 		var options =
                // 			"?$filter=" +
                // 			successfullyUpdatedRecords
                // 				.map((record) => {
                // 					return "leadid eq '" + record.Id + "'";
                // 				})
                // 				.join(" or ") +
                // 			"&$select=createdon";
                // 		Xrm.WebApi.retrieveMultipleRecords(
                // 			EntityNames.Lead,
                // 			options
                // 		).then((leads: any) => {
                // 			leads.entities.forEach((lead: any) => {
                // 				this.logStatusUpdate(
                // 					lead.leadid,
                // 					lead.createdon,
                // 					"LeadGridCommandActions.setStateAndStatus",
                // 					state == LeadManagement.LeadState.Disqualified
                // 						? "Disqualify"
                // 						: "Reactivate"
                // 				);
                // 			});
                // 		});
                // 	}
                // });
            };
            this.triggerGridRefreshOnLimitReached = function (gridControl, recordCounter, numberOfRecords, progressIndicator) {
                if (recordCounter === numberOfRecords) {
                    progressIndicator.hide();
                    gridControl.refresh();
                }
            };
            this.getStatusFromParameter = function (commandProperties, statusParameter) {
                var status = -1;
                if (typeof commandProperties === "object" && commandProperties.SourceControlId) {
                    var parts = commandProperties.SourceControlId.split('|');
                    status = parseInt(parts[parts.length - 1]) || -1;
                }
                else {
                    if (!ClientUtility.DataUtil.isNullOrUndefined(statusParameter)) {
                        status = parseInt(statusParameter) || -1;
                    }
                    else {
                        status = parseInt(commandProperties) || -1;
                    }
                }
                return status;
            };
            this.qualifyLeadFailedCallback = function (errorResponse) {
                if (errorResponse) {
                    //	2147746326(Unexpected error) is a generic error code. UCI library replaces the outer error "message" in errorResponse with
                    //	error description corresponding to error code 2147746326. To throw a specific error, set errorResponse.message to a specific error message.
                    var processNoActiveStageError = String.format(MarketingSales.ResourceStringProvider.getResourceString("ErrorMessage_Lead_ReQualify_ProcessNoActiveStage"), MarketingSales.EntityNames.Lead);
                    if ((errorResponse.errorCode === 2147746326 || errorResponse.errorCode === -2147220970) && errorResponse.innerror && errorResponse.innerror.message &&
                        errorResponse.innerror.message.indexOf(processNoActiveStageError) >= 0) {
                        errorResponse.message = processNoActiveStageError;
                        errorResponse.errorRegion = "Sales";
                    }
                    else if (errorResponse.errorCode === 2148532226) {
                        errorResponse.errorRegion = "Sales";
                    }
                }
                MarketingSales.LeadErrorHandling.actionFailedErrorDialog(errorResponse);
            };
            // private logStatusUpdate(leadId: string, createdOn: string, methodName: string, closeEventName: string, parameters?: any): void {
            // 	let timeTakenSinceCreate = <any>new Date() - <any>new Date(createdOn);
            // 	timeTakenSinceCreate /= 1000 * 60 * 60;
            // 	LeadUtil.logTelemetryData(
            // 		"LeadStatusUpdate",
            // 		methodName,
            // 		MetadataDrivenDialogConstants.information,
            // 		closeEventName,
            // 		{
            // 			leadId: leadId.replace(/[{}]/g, "").toLowerCase(),
            // 			timeTakenSinceCreate,
            // 			...parameters
            // 		}
            // 	);
            // }
        }
        return LeadGridCommandActionsImplementation;
    }());
    MarketingSales.LeadGridCommandActionsImplementation = LeadGridCommandActionsImplementation;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../../ClientCommon/MetadataDrivenDialogConstants.ts" />
var MarketingSales;
(function (MarketingSales) {
    var LeadLibrary = (function () {
        function LeadLibrary() {
            this.formOnload = function () {
                if (Xrm.Internal.isUci() && Xrm.Page.data.process.getActiveProcess() != null) {
                    //populate the lead state & status codes
                    Xrm.Page.data.process.getActiveProcess().getStages().forEach(function (stage) {
                        var navigationBehavior = stage.getNavigationBehavior();
                        navigationBehavior.allowCreateNew = function (processStage) {
                            var activePath = Xrm.Page.data.process.getActivePath();
                            var activePathLength = activePath.getLength();
                            var nextProcessStage = null;
                            // Find next stage in the active path
                            for (var i = 0; i < (activePathLength - 1); i++) {
                                if ((activePath.get(i)).getId() === processStage.getId()) {
                                    nextProcessStage = activePath.get(i + 1);
                                    break;
                                }
                            }
                            // Case 1 - This is the last stage in the active path
                            if (nextProcessStage === null) {
                                return false;
                            }
                            // Case 2 - This stage is associated with lead and next stage with opportunity
                            if (processStage.getEntityName() === "lead" && nextProcessStage.getEntityName() === "opportunity") {
                                return false;
                            }
                            // Case 3 - This and next stages are associated with the same entity
                            if (processStage.getEntityName() === nextProcessStage.getEntityName()) {
                                return false;
                            }
                            // Case 4 - This is the not the last stage and next stage is associated with a different entity
                            return true;
                        };
                    });
                }
            };
        }
        return LeadLibrary;
    }());
    MarketingSales.LeadLibrary = LeadLibrary;
})(MarketingSales || (MarketingSales = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../../../TypeDefinitions/CRM/ClientUtility.d.ts" />
/// <reference path="LeadCommandActions.ts" />
/// <reference path="LeadGridCommandActions.ts" />
/// <reference path="UCI/LeadLibrary.ts" />
/// <reference path="LeadUtil.ts" />
var MarketingSales;
(function (MarketingSales) {
    var Lead = (function () {
        function Lead() {
        }
        return Lead;
    }());
    Lead.CommandActions = new MarketingSales.LeadCommandActions();
    Lead.GridCommandActions = new MarketingSales.LeadGridCommandActions();
    Lead.LeadUtil = new MarketingSales.LeadUtil();
    Lead.LeadLibrary = new MarketingSales.LeadLibrary();
    Lead.ctor = (function () {
        // These are needed on the window because of "general" command bar actions calling hard-coded methods with some conditions
    })();
    MarketingSales.Lead = Lead;
})(MarketingSales || (MarketingSales = {}));



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
                if (res.entities?.length) return res.entities[0][`${entityLogicalName}id`];
            } catch (e) { }
        }
        return null;
    }

    // Build update payload
    const update = {};

    // Retrieve lead metadata to validate lookup presence in this org/env
    let leadMeta = null;
    try {
        leadMeta = await Xrm.Utility.getEntityMetadata("lead", ["tcg_contractingunit", "msdyn_company", "tcg_customergroupidnewone"]);
    } catch (e) { /* proceed defensively */ }
    const hasContractingUnitAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_contractingunit"));
    const hasCompanyAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("msdyn_company"));
    const hasCustomerGroupAttr = !!(leadMeta && leadMeta.Attributes && leadMeta.Attributes.get && leadMeta.Attributes.get("tcg_customergroupidnewone"));

    // OptionSets: use cached metadata from preloaded localStorage
    try {
        const marketSegmentMeta = JSON.parse(localStorage.getItem("marketSegmentMetaData") || "[]");
        const industryMeta = JSON.parse(localStorage.getItem("IndustryMetaData") || "[]");
        const accountTypeMeta = JSON.parse(localStorage.getItem("accountTypeMetaData") || "[]");

        const msOpt = marketSegmentMeta.find(x => (x.text || "").trim().toLowerCase() === (payload.marketSegment || "").trim().toLowerCase());
        if (msOpt?.value !== undefined) update["tcg_marketsegmentnew"] = parseInt(msOpt.value, 10);

        const indOpt = industryMeta.find(x => (x.text || "").trim().toLowerCase() === (payload.industry || "").trim().toLowerCase());
        if (indOpt?.value !== undefined) update["tcg_industrynew"] = parseInt(indOpt.value, 10);

        const atOpt = accountTypeMeta.find(x => (x.text || "").trim().toLowerCase() === (payload.accountType || "").trim().toLowerCase());
        if (atOpt?.value !== undefined) update["tcg_accounttype"] = parseInt(atOpt.value, 10);
    } catch (e) { }

    // Lookups
    // Legal Entity (msdyn_company -> cdm_company)
    if (payload.legalEntity && hasCompanyAttr) {
        const comp = await resolveCdmCompanyByLabel(payload.legalEntity);
        if (comp?.cdm_companyid) update["msdyn_company@odata.bind"] = `/cdm_companies(${toGuid(comp.cdm_companyid)})`;
    }

    // Contracting Unit (skipped in grid path to avoid undeclared property error)
    // if (payload.contractingUnit && hasContractingUnitAttr) {
    //     const cuId = await resolveIdByName("msdyn_organizationalunit", payload.contractingUnit, ["msdyn_name", "name"]);
    //     if (cuId) update["tcg_contractingunit@odata.bind"] = `/msdyn_organizationalunits(${toGuid(cuId)})`;
    // }

    // Parent Account
    if ((payload.account || "").trim() && (payload.accountId || "").trim()) {
        update["parentaccountid@odata.bind"] = `/accounts(${toGuid(payload.accountId)})`;
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
            }
        } catch (e) { }
    }

    // No changes resolved? Exit quietly
    if (Object.keys(update).length === 0) return;

    // Execute update with fallback for unknown nav properties
    try {
        await Xrm.WebApi.updateRecord("lead", toGuid(leadId), update);
    } catch (e) {
        const msg = (e && (e.message || e._message)) || "";
        // If server complains about tcg_contractingunit annotation, drop it and retry
        if (msg.toLowerCase().includes("tcg_contractingunit")) {
            try {
                delete update["tcg_contractingunit@odata.bind"];
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
                if (source === "form") {
                    // Apply to current form
                    await applyPayloadToLead(PrimaryControl, payload);
                    await Xrm.Navigation.openAlertDialog(
                        { title: "Success", text: "Lead validated successfully." },
                        { height: 120, width: 320, position: 1 }
                    );
                    // Keep existing behavior for form ribbon
                    Mscrm.LeadCommandActions.qualifyLeadQuick();
                } else if (source === "grid") {
                    // Persist changes via Web API without needing form context
                    await applyPayloadToLeadGrid(leadId, payload);
                    await Xrm.Navigation.openAlertDialog(
                        { title: "Success", text: "Lead validated successfully." },
                        { height: 120, width: 320, position: 1 }
                    );
                    // Do NOT trigger form-only qualify action here
                }
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
