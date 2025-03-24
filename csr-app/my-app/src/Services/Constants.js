export const baseUrl = window.location.origin.includes('localhost:') ? 'http://localhost:5002' : window.location.origin;
// export const baseUrl = 'https://www.exdion-itest.studio:4000';
// export const baseUrl = 'https://www.exdion-itest.studio:4000';
// export const baseUrl ='https://www.exdiongrid.com:4000';
export const isAdLogin = false;
export const Disclaimer = 'Disclaimer: This checklist contains sensitive client information and is meant for consumption ONLY by you and your organization. Please do NOT forward or share this with any unauthorized person or people outside of your organization.';
export const Checklist = 'POLICY REVIEW CHECKLIST';
export const formCompare = 'FORM COMPARE';
export const initialData1 = {};
export const Formpolicydata = [
    "Do the Forms",
    "Are the forms",
    "Umbrella/Excess",
    "Inland Marine",
    "Employment Practices (EPL)",
    "Cargo/Transportation/Contingent/Ocean Cargo/Warehouse Legal Liability",
    "Property",
    "General Liability",
    "Crime / Kidnap/Ransom (K&R)",
    "Commercial Auto/Truckers/Contingent/Garage",
    "Cyber E&O / Directors & Officers (D & O) / Fiduciary / Pollution Liability / Professional (E&O)",
    "NULL"
];

export const confidenceScoreConfigStaticData = [
    {
        "Key": "EnableLockCell",
        "Value": "true",
        "CreatedOn": "2025-02-11T15:01:18.3226792+00:00"
    },
    {
        "Key": "MinLockCellScore",
        "Value": "95",
        "CreatedOn": "2025-02-11T15:01:18.3226792+00:00"
    },
    {
        "Key": "EnableConfidenceScore",
        "Value": "true",
        "CreatedOn": "2025-02-11T15:01:18.3226792+00:00"
    },
    {
        "Key": "EnableInsert",
        "Value": "false",
        "CreatedOn": "2025-02-11T15:01:18.3226792+00:00"
    },
    {
        "Key": "EnableAutoMatched",
        "Value": "false",
        "CreatedOn": "2025-02-11T15:01:18.3226792+00:00"
    }
];

export const excludedColumnlist = ["Id", "JobId", "Jobid", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp", "PolicyLob",];

export const staticExclusionData = [{
    "Id": "",
    "JobId": "",
    "FormName": "",
    "FormDescription": "",
    "Exclusion": "",
    "PageNumber": "",
    "CreatedOn": "",
    "UpdatedOn": "",
}];

export const staticExclusionCsrData = [{
    "Id": "",
    "JobId": "",
    "FormName": "",
    "FormDescription": "",
    "Exclusion": "",
    "PageNumber": "",
    "CreatedOn": "",
    "UpdatedOn": "",
    "ActionOnDiscrepancy": "",
    "RequestEndorsement": "",
    "Notes": "",
    "NotesFreeFill": ""
}];

export const updateData = [
    {
        "TemplateData": "[{\"CoverageSpecificationsMaster\":\"Attached Forms\",\"Checklist Questions\":\"CA2\",\"CurrentTermPolicyAttached\":\"Details not available in the document\",\"PriorTermPolicyAttached\":\"Details not available in the document\",\"Observation\":\"\",\"PolicyLob\":\"Attached Forms\",\"Page Number\":\"\"}]"
    }
]

export const policyData2_3 = [
    {
        "TemplateData": "[{\"CoverageSpecificationsMaster\":\"~~\",\"Checklist Questions\":\"~~\",\"CurrentTermPolicy\":\"~~Page #\",\"PriorTermPolicy\":\"~~Page #\",\"Observation\":\"~~\",\"PolicyLob\":\"Common Declaration\",\"Page Number\":\"~~\"}]"
    }
]

export const csrPolicyData2_3 = [
    {
        "TemplateData": "[{\"COVERAGE_SPECIFICATIONS_MASTER\":\"~~\",\"Checklist Questions\":\"~~\",\"Current Term Policy\":\"~~Page #\",\"Prior Term Policy\":\"~~Page #\",\"OBSERVATION\":\"~~\",\"POLICY LOB\":\"General Liability\",\"Page Number\":\"~~\"}]"
    }
]

export const policyDataForms = [
    {
        "TemplateData": "[{\"CoverageSpecificationsMaster\":\"~~\",\"Checklist Questions\":\"~~\",\"CurrentTermPolicyAttached\":\"~~Page #\",\"PriorTermPolicyAttached\":\"~~Page #\",\"Observation\":\"~~\",\"PolicyLob\":\"Do the Forms, Endorsements and Edition Dates match the source documents?\",\"Page Number\":\"~~\"}]"
    }
]

export const initialData = {
    "ChecklistData": {
        "XRayUrl": "",
        "QuestionID": null,
        "ID": 0,
        "JobID": "12345678",
        "DocID": "",
        "DocumentName": "Basic Document",
        "PageNum": null,
        "ChecklistQuestion": null,
        "Label": null,
        "ExtractedValue": null,
        "LabelClassID": null,
        "IsBaseDocument": null,
        "MatchScore": "0.0%",
        "RefID": null,
        "GroupID": null,
        "IsActive": false,
        "CreatedBy": null,
        "CreatedOn": "0001-01-01T00:00:00",
        "UpdatedBy": null,
        "UpdatedOn": "0001-01-01T00:00:00"
    },

    "tbl_CheckList": {
        "ID": "1",
        "JOBID": "12345678",
        "InsuredName": "JENNIFER LANE SPIELMANN",
        "InsCarrierName": "UNITED STATES LIABILITY INSURANCE COMPANY",
        "PortalJobID": "1142csr150072023045014",
        "PolicyNo": "DECLARATIONS NO IAE1550894J",
        "PolicyLOB": "pro",
        "PolicyTerm": "07/01/2023 - 07/01/2024",
        "FileName": "IAE023M0109_Applicant.pdf,IAE1550894J_Original.pdf,Jennifer Lane Spielmann Policy 22-23.pdf,MKT Binder.pdf",
        "PolicyLoaded": "2023-07-07",
        "PolicyType": "Marketed",
        "ChecklistQuestions": "RD1:Correct Retroactive Date, if applicable?",
        "Observation": "Current term content(s):07/01/2017 Proposal content(s):NO RECORDS",
        "CurrentTermValue": "07/01/2017",
        "ReferenceValue": "NO RECORDS",
        "PageNo": "Current term:17 Proposal:NO RECORDS",
        "MatchScore": "0.0%",
        "ActionTaken": null,
        "CSRAMComments": null,
        "CSRAMNotes": null,
        "CID": null,
        "UpdateBy": null,
        "UpdateDate": "0001-01-01T00:00:00",
        "INTDELETED": 0,
        "DeletedBy": null,
        "DeletedDate": "0001-01-01T00:00:00",
        "CreatedBy": null,
        "CreatedDate": null,
        "ExtractLOBS": null,
        "SamplingDetails": null,
        "ReferenceJobID": null
    }
}

export const comparisonRules = [
    { text: 'InsuredName', key: 'documentName' },
    { text: 'PolicyTerm', key: 'docId' },
    { text: 'PolicyLOB', key: 'checklistQuestion' },
    { text: 'PortalJobID', key: 'PortalJobID' },
    { text: 'InsCarrierName', key: 'PolicyLOB' },
    { text: 'ReferenceValue', key: 'createdOn' },
    { text: 'CreatedBy', key: 'createdBy' },
    { text: 'PolicyLoaded', key: 'extractedValue' },
    { text: 'Observation', key: 'isBaseDocument' },
    { text: 'JOBID', key: 'matchScore' },
];

export const baseData = {
    name: "Check list",
    styles: [
        {
            border: {
                bottom: ["medium", "#262626"],
                top: ["medium", "#262626"],
                left: ["medium", "#262626"],
                right: ["medium", "#262626"]
            },
            textwrap: true,
        }, //0
        {
            border: {
                bottom: ["thin", "#2d8aed"],
                top: ["medium", "#000"],
                left: ["medium", "#000"],
                right: ["medium", "#000"]
            },
            textwrap: true,
            font: { bold: true },
            bgcolor: "#2d8aed",
        }, //1
        {
            border: {
                left: ["medium", "#000"],
                top: ["thick", "#2d8aed"],
                right: ["medium", "#000"],
                bottom: ["medium", "#000"],
            },
            textwrap: true,
            font: { bold: true },
            align: "center", bgcolor: "#2d8aed"
        }, //2
        {
            border: {
                bottom: ["medium", "#000"],
                top: ["medium", "#000"],
                left: ["medium", "#000"],
                right: ["medium", "#000"]
            },
            textwrap: true,
            font: { bold: true },
            align: "center",
        }, //3
        {
            border: {
                bottom: ["medium", "#000"],
                top: ["medium", "#000"],
                left: ["thin", "#2d8aed"],
                right: ["thick", "#2d8aed"]
            },
            textwrap: true,
            font: { bold: true },
            align: "center",
            bgcolor: "#2d8aed ",
        }, //4
        {
            // align: "center",
            // bgcolor: "#ffffff",
            border: {
                bottom: ["medium", "#262626"],
                top: ["medium", "#262626"],
                left: ["medium", "#262626"],
                right: ["medium", "#262626"]
            },
            textwrap: true,
            color: "#008000",
            // underline: true,
        }, //5

        {
            textwrap: true,
            color: "#FFFFFF",
            // underline: true,
        }, //6
        {
            textwrap: true,
            color: "#FF0000",
            // underline: true,
        }, //7
    ],
    hyperlinks: [{
        text: 'Click here', // The display text for the hyperlink
        url: 'https://example.com', // The URL to link to
    }],
    //merges: ["B1:D1", "B26:L26"],
    rows: {
        "0": {
            cells: {
                "1": { merge: [0, 2], style: 3, text: "POLICY REVIEW CHECKLIST" },
                // "2": { style: 4 },
                // "3": { style: 4 },
                // "4": { style: 4 },
            },
            height: 30
        },
        "1": {
            cells: {
                "1": { text: "" },
                // "2": { style: 4 },
                // "3": { style: 4 },
                // "4": { style: 4 },
            },
            height: 30
        },
        // "2": {
        //     subColumns: {
        //         "1": { text:"" },
        //         "2": { style: 4 },
        //         "3": { style: 4 },
        //         "4": { style: 4 }
        //     },
        //     height: 30
        // }
    },
    cols: {
        "1": { width: 250 },
        "2": { width: 250 },
        "3": { width: 280 },
        "4": { width: 280 },
        "5": { width: 280 },
        "6": { width: 280 },
        "7": { width: 280 },
        "8": { width: 280 },
        "9": { width: 280 },
        "10": { width: 280 },
        "11": { width: 280 },
    },
    subColumns: [{}],
    validations: [],
    autofilter: {}
}

export const qacMasterSet = {
    name: "QAC not answered questions", // Worksheet name
    color: "", // Worksheet color
    config: {
        merge: {
        },
        // sheetcheck: "Formscompare",
        borderInfo: [],
        rowlen: {
        },
        columnlen: {
            "0": 830
        },
        "curentsheetView": "viewPage",//viewNormal, viewLayout, viewPage
        "sheetViewZoom": {
            "viewNormalZoomScale": 0.6,
            // "viewPageZoomScale": 1,
            "viewPageZoomScale": 0.6,
        },
    },
    // row: {
    //   len: 500, // This sets the default row length to 500
    // },
    //index: "0", // Worksheet index
    chart: [], // Chart configuration
    status: "0", // Activation status
    order: "0", // The order of the worksheet
    hide: 0, // Whether to hide
    column: 50, // Number of columns
    row: 50, // Number of rows
    celldata: [],
    // Forms_celldata: jobDataa,// Original cell data set
    // visibledatarow: [], // The position of all rows
    // visibledatacolumn: [], // The position of all columns
    ch_width: 2322, // The width of the worksheet area
    rh_height: 949, // The height of the worksheet area
    scrollLeft: 0,
    scrollTop: 0,
    luckysheet_select_save: [], // Selected area
    //luckysheet_conditionformat_save: {}, // Conditional format
    calcChain: [], // Formula chain
    isPivotTable: false, // Whether to pivot table
    pivotTable: {}, // Pivot table settings
    filter_select: null, // Filter range
    filter: null, // Filter configuration
    luckysheet_alternateformat_save: [], // Alternate colors
    luckysheet_alternateformat_save_modelCustom: [], // Customize alternate colors
    sheets: []
}