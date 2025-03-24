// Luckysheet2.jsx
import React, { useEffect, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Disclaimer, Checklist, updateData, formCompare, csrPolicyData2_3, baseUrl } from '../Services/Constants';
import { findTableForIndex, processAndUpdateToken, getText, findTblRowAllIndex, filterSelectedRowIndexForCopyPaste, CsrSaveHistoryApiCall, CsrPendingReport, autoupdate, brokerIdsGetData, qacTblRangeStructureFun } from '../Services/CommonFunctions';
import { DialogComponent, DiscrepancyOptionsDialogComponent, EndorsementDialogComponent, FilterCsrDialogComponent } from '../Services/dialogComponent';
import axios from "axios";
import { Icon } from '@fluentui/react';
import $ from 'jquery';
import { gradiationConverter } from '../Services/GradationDataConverter';
import { SimpleSnackbar } from '../Components/SnackBar';
import { UpdateJobPreviewStatus, UpdateJobSendPolicyInsured } from '../Services/PreviewChecklistDataService';
import { updateGridAuditLog } from '../Services/PreviewChecklistDataService';

export default function Luckysheet2(props) {
    const container = useRef();
    const sheetDataRowIndexRef = useRef({});
    const navigate = useNavigate();
    const [openDialog, setOpenDialog] = useState(false);
    const [msgText, setMsgText] = useState("");
    const luckysheet = window.luckysheet;
    const [yesDialog, setYesDialog] = useState(false);
    const [sheetState, setsheetState] = useState([]);
    const [loBResult, setLobResult] = useState([]);
    const [openFilterDialog, setOpenFilterDialog] = useState(false);
    const [filterSelectionData, setFilterSelectionData] = useState(null);
    const [csrPolicyData, setCsrPolicydata] = useState(props?.data);
    const [autoprogress, setautoprogress] = useState(false);
    const [jobId, setJobId] = useState(props?.selectedJob);
    const [issavessheet, setIssavessheet] = useState(false);
    const [dataForSavePolicy, setDataForSavePolicy] = useState([]); //AGN -- for update(previewchecklist)
    const [dataForSavePolicyPosition, setDataForSavePolicyPosition] = useState([]); //AGN -- for update(previewchecklist)
    const [dataForXRayMapping, setDataForXRayMapping] = useState([]); // for x-Ray Data(previewchecklist)
    const [formdataForXRayMapping, setFormDataForXRayMapping] = useState([]); // for x-Ray Data(formscompare)
    const [dataForSaveFormsCompare, setDataForSaveFormsCompare] = useState([]); //-- for update(formsCompare)
    const [dataForSaveFormsComparePosition, setDataForSaveFormsComparePosition] = useState([]); //-- for update(formsCompare)
    const [dataForSaveExclusion, setDataForSaveExclusion] = useState([]); //-- for update(exclusion)
    const [dataForSaveExclusionPosition, setDataForSaveExclusionPosition] = useState([]); //-- for update(exclusion)
    const [formsComparedata, setFormsComparedata] = useState(props.formCompareData);
    const [state, setState] = useState(props?.data);
    const [formstate, setFormState] = useState(props.formCompareData);
    const [exclusionstate, setExclusionState] = useState(props.exclusionRenderData);
    const [canRenderChecklist, setCanRenderChecklist] = useState(props?.data && Array.isArray(props?.data) && props?.data?.length > 0);
    const [canRenderFormData, setCanRenderFormData] = useState(props?.formCompareData && Array.isArray(props?.formCompareData) && props?.formCompareData?.length > 0);
    const [canRenderExclusionData, setCanRenderExclusionData] = useState(props?.exclusionRenderData && Array.isArray(props?.exclusionRenderData) && props.exclusionRenderData?.length > 0);
    const [canRenderGradation, setCanRenderGradation] = useState(props?.gradtionDataSet && Array.isArray(props?.gradtionDataSet) && props?.gradtionDataSet?.length > 0);
    const [canRenderQac, setCanRenderQac] = useState(props?.qacData);
    const [sheetDataRowIndex, setSheetDataRowIndex] = useState({});
    const [gradtionData, setGradtionData] = useState(props?.gradtionDataSet);
    const [dropDialog, setDropDialog] = useState(false);
    const [gradiationDialog, setGradiationDialog] = useState(false);
    const [tableColumnDetails, setTableColumnDetails] = useState({ "Table 1": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 2": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 3": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 4": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 5": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 6": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 7": { "columnNames": {}, "range": { "start": "", "end": "" } } });
    const [formTableColumnDetails, setFormTableColumnDetails] = useState({ "FormTable 1": { "columnNames": {}, "range": { "start": "", "end": "" } }, "FormTable 2": { "columnNames": {}, "range": { "start": "", "end": "" } }, "FormTable 3": { "columnNames": {}, "range": { "start": "", "end": "" } } });
    const [exTableColumnDetails, setExTableColumnDetails] = useState({ "ExTable 1": { "columnNames": {}, "range": { "start": "", "end": "" } } });
    const [qacTableColumnDetails, setQacTableColumnDetails] = useState({});
    let token = sessionStorage.getItem('token');
    const brokerId = jobId.slice(0, 4);
    const AODDBrokerIds = ["1003", "1150", "1165", "1167"];    //broker config for formsCmp & exclusion broker to be shown;
    const reqEndorsementColHideBrokerId = ["1162"];
    const renderBrokerIdsQac = ["1162"];
    const [endorsementColflag, setEndorsementColflag] = useState("");
    const [brokerData, setBrokerData] = useState([]);
    const [reviewData, setReviewData] = useState([]);
    const [sendPolicyInsuredData, setSendPolicyInsuredData] = useState([]);
    localStorage.removeItem('removedArrays');//for icon grouping dont remove 
    localStorage.removeItem('hitedselectedValues');//for icon grouping dont remove
    localStorage.removeItem('IconShownIndex');//for icon grouping dont remove
    const [userName, setUserName] = useState("");
    const sessionUserName = sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName");
    const [activeUserName, SetActiveUserName] = useState(props?.defaultSheetUserNameData);


    const Policy_appDataConfig = {
        demo: {
            name: "PolicyReviewChecklist",
            color: "",
            config: {
                merge: {
                    "0_1": {
                        "rs": 1,
                        "cs": 6,
                        "r": 0,
                        "c": 1
                    },
                    "1_1": {
                        "rs": 1,
                        "cs": 2,
                        "r": 0,
                        "c": 1
                    },
                },
                borderInfo: [],
                rowlen: {
                    "0": 20,
                    "1": 20,
                    "2": 20,
                    "3": 35,
                    "4": 50,
                    "5": 35,
                    "6": 35,
                    "7": 35,
                    "8": 35,
                    "9": 35,
                    "10": 50,
                    "11": 50,
                    "12": 60,
                    "13": 20,
                    "14": 20,
                    "15": 20,
                    "16": 20,
                    "17": 31
                },
                columnlen: {
                    "0": 63,
                    "1": 280,
                    "2": 250,
                    "3": 250,
                    "4": 250,
                    "5": 250,
                    "6": 250,
                    "7": 250,
                    "8": 250,
                    "9": 250,
                    "10": 250,
                    "11": 250,
                    "12": 250,
                    "13": 250,
                    "14": 250
                },
                "curentsheetView": "viewPage",
                "sheetViewZoom": {
                    "viewNormalZoomScale": 0.6,
                    "viewPageZoomScale": 0.6,
                },
            },

            chart: [],
            status: "1",
            order: "0",
            hide: 0,
            column: 20,
            celldata: [],
            ch_width: 2322,
            rh_height: 949,
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [],
            calcChain: [],
            luckysheet_alternateformat_save: [],
            luckysheet_alternateformat_save_modelCustom: [],
            sheets: []
        }
    };
    const FormCompare_appconfigdata = {
        forms: {
            name: "Forms Compare",
            color: "",
            config: {
                merge: {
                    "1_1": {
                        "rs": 1,
                        "cs": 2,
                        "r": 0,
                        "c": 1
                    },
                },
                borderInfo: [],
                rowlen: {
                    "3": 20,
                    "4": 20,
                    "5": 35,
                    "6": 35,
                    "7": 35,
                    "8": 35,
                    "9": 35,
                    "10": 50,
                    "11": 50,
                    "12": 20,
                    "13": 20,
                    "14": 20,
                    "15": 20,
                    "16": 20,
                    "17": 31
                },
                columnlen: {
                    "0": 62,
                    "1": 250,
                    "2": 250,
                    "3": 250,
                    "4": 250,
                    "5": 250,
                    "6": 250,
                    "7": 250,
                    "8": 250,
                    "9": 250,
                    "10": 250,
                    "11": 250,
                    "12": 250,
                    "13": 250,
                    "14": 250
                },
                "curentsheetView": "viewPage",
                "sheetViewZoom": {
                    "viewNormalZoomScale": 0.6,
                    "viewPageZoomScale": 0.6,
                },
            },

            chart: [],
            order: "0",
            hide: 0,
            column: 20,
            celldata: [],
            ch_width: 2322,
            rh_height: 949,
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [],
            calcChain: [],
            luckysheet_alternateformat_save: [],
            luckysheet_alternateformat_save_modelCustom: [],
            sheets: [],
        },
    };
    const Exclusion_appDataConfig = {
        exclusion: {
            name: "Exclusion",
            config: {
                merge: {},
                borderInfo: [],
                columnlen: {
                    "0": 230,
                    "1": 230,
                    "2": 230,
                    "3": 230,
                    "4": 230,
                    "5": 230,
                    "6": 230,
                    "7": 230,
                },
                rowlen: {},
                "curentsheetView": "viewPage",
                "sheetViewZoom": {
                    "viewNormalZoomScale": 0.6,
                    "viewPageZoomScale": 0.6,
                },
            },
            status: "1",
            column: 20,
            row: 500,
            celldata: [],
            ch_width: 2322,
            rh_height: 949,
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [],
            calcChain: [],
            luckysheet_alternateformat_save: [],
            luckysheet_alternateformat_save_modelCustom: [],
            sheets: [],
        },
    };
    const qac_appDataConfig = {
        name: "QAC not answered questions",
        color: "",
        config: {
            merge: {
            },
            borderInfo: [],
            rowlen: {
            },
            columnlen: {
                "0": 550,
                "1": 300,
                "2": 300
            },
            "curentsheetView": "viewPage",
            "sheetViewZoom": {
                "viewNormalZoomScale": 0.6,
                "viewPageZoomScale": 0.6,
            },
        },
        chart: [],
        status: "0",
        order: "0",
        hide: 0,
        column: 50,
        row: 50,
        celldata: [],
        ch_width: 2322,
        rh_height: 949,
        scrollLeft: 0,
        scrollTop: 0,
        luckysheet_select_save: [],
        calcChain: [],
        isPivotTable: false,
        pivotTable: {},
        filter_select: null,
        filter: null,
        luckysheet_alternateformat_save: [],
        luckysheet_alternateformat_save_modelCustom: [],
        sheets: []
    };


    const luckyCss = {
        margin: '0px',
        padding: '0px',
        position: 'absolute',
        width: '100% !important',
        height: '50%',
        left: '0px',
        top: '0px',
    };

    const handleVisibilityChange = () => {
        if (luckysheet) {
            luckysheet?.refresh();
            luckysheet?.exitEditMode();
        }
    }

    const loadBrokerData = async () => {
        try {
            let brokerData = await brokerIdsGetData();
            localStorage.setItem('brokerDatas', JSON.stringify(brokerData))
            setBrokerData(brokerData);
        } catch (error) {
            console.error('Error fetching dat:', error);
        }
    };


    useEffect(() => {
        const mainData = state;
        const formCompareData = formstate;
        const exclusionData = exclusionstate;
        setUserName(sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName"));

        document.addEventListener('visibilitychange', handleVisibilityChange, false);

        try {
            getCsrReviewData(jobId);
        } catch (error) {
            console.log(error);
        }

        loadBrokerData();
        const renderTextBlock = () => {
            const textBlock = [
                {
                    ct: {
                        fa: "General",
                        t: "g"
                    },
                    fc: "#ff0000",
                    ff: "\"Tahoma\"",
                    m: Disclaimer,
                    v: Disclaimer,
                },
            ];

            return textBlock;
        };

        const renderList = () => {
            const listBlock = [
                {
                    ct: {
                        fa: "General",
                        t: "g"
                    },
                    fc: "#000000",
                    ff: "\"Tahoma\"",
                    m: Checklist,
                    v: Checklist,
                },
            ];

            return listBlock;
        };

        const renderForm = () => {
            const formBlock = [
                {
                    ct: {
                        fa: "General",
                        t: "g"
                    },
                    fc: "#000000",
                    ff: "\"Tahoma\"",
                    m: formCompare,
                    v: formCompare,
                },
            ];

            return formBlock;
        };

        const csrTable1 = () => {
            const tableData1 = mainData.find((data) => data.Tablename === "Table 1");
            if (tableData1) {
                const table1json = typeof tableData1.TemplateData === 'string' ? JSON.parse(tableData1.TemplateData) : tableData1.TemplateData;

                let sheetDataTable1 = [];
                let sheetDataTable2 = [];

                const rowIndexOfTable1 = 3
                const textBlockData = renderTextBlock();
                const listData = renderList();

                textBlockData.forEach((item, index) => {
                    const mergeConfig = Policy_appDataConfig.demo.config.merge["0_1"];

                    sheetDataTable2.push({
                        r: index + mergeConfig.r,
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            bl: 1,
                            fs: 8.5,
                            ff: item.ff,
                            merge: mergeConfig,
                            fc: item.fc,
                            // tb: '55',
                        }
                    });
                });

                listData.forEach((item, index) => {
                    const mergeConfig = Policy_appDataConfig.demo.config.merge["1_1"];

                    sheetDataTable2.push({
                        r: 1 + mergeConfig.r,
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            ff: item.ff,
                            fs: 13,
                            merge: mergeConfig,
                            fc: item.fc,
                        }
                    });
                });

                table1json.map((item, index) => {
                    if (item["Headers"] != null) {
                        sheetDataTable1.push({
                            r: rowIndexOfTable1 + index,
                            c: 1,
                            v: {
                                ct: { fa: "@", t: "inlineStr", s: [{ v: item["Headers"], ff: "Tahoma" }] },
                                m: item["Headers"],
                                v: item["Headers"],
                                fs: 10,
                                ff: "Tahoma",
                                merge: null,
                                bg: "rgb(139,173,212)",
                                tb: '2',
                            }
                        });

                        const tidleValue = item["(No column name)"] !== null ? item["(No column name)"].replace(/~~/g, "\n") : "";
                        if (item["Headers"] === "Exdion CSRDigiT") {
                            sheetDataTable1.push({
                                r: rowIndexOfTable1 + index,
                                c: 2,
                                v: {
                                    ct: { fa: "@", t: "inlineStr" },
                                    m: tidleValue,
                                    v: tidleValue,
                                    fs: 9,
                                    un: 1,
                                    fc: "#0b5394",
                                    ff: "\"Tahoma\"",
                                    merge: null,
                                    tb: '2',
                                }
                            });
                        } else {
                            sheetDataTable1.push({
                                r: rowIndexOfTable1 + index,
                                c: 2,
                                v: {
                                    ct: { fa: "@", t: "inlineStr" },
                                    m: tidleValue,
                                    v: tidleValue,
                                    fs: 9,
                                    ff: "\"Tahoma\"",
                                    merge: null,
                                    tb: '2',
                                }
                            });
                        }

                        let maxLength = 0;
                        const lengths = [];
                        Object.keys(item).forEach((key) => {
                            if (item[key]) {
                                lengths.push(item[key].length);
                            }
                        });
                        lengths.forEach((length) => {
                            if (length > maxLength) {
                                maxLength = length;
                            }
                        });
                        const rowHeight = parseInt(maxLength / 3 + 15);
                        Policy_appDataConfig.demo.config.rowlen[`${rowIndexOfTable1 + index}`] = rowHeight;
                    }
                });

                const dummyData = [];
                const matchedUnMatchedFilter = [
                    {
                        "r": 3,
                        "c": 4,
                        "v": {
                            "ct": {
                                "fa": "General",
                                "t": "inlineStr",
                                "s": [
                                    {
                                        "fs": "16",
                                        "v": "□ "
                                    },
                                    {
                                        "vt": "0",
                                        "ht": "1",
                                        "fs": "9",
                                        "un": 0,
                                        "bl": 1,
                                        "fc": "#0000ff",
                                        "ff": "\"Tahoma\"",
                                        "m": brokerId === "1167" ? "Full Variances" : "All Variances",
                                        "v": brokerId === "1167" ? "Full Variances" : "All Variances",
                                    }
                                ]
                            },
                            "merge": null,
                            "w": 55,
                            "tb": "2",
                            "fc": "#0000ff",
                            "fs": "16"
                        }
                    },
                    {
                        "r": 4,
                        "c": 4,
                        "v": {
                            "ct": {
                                "fa": "General",
                                "t": "inlineStr",
                                "s": [
                                    {
                                        "fs": "16",
                                        "v": "□ "
                                    },
                                    {
                                        "vt": "0",
                                        "ht": "1",
                                        "fs": "9",
                                        "un": 0,
                                        "bl": 1,
                                        "fc": "#0000ff",
                                        "ff": "\"Tahoma\"",
                                        "m": "Matched",
                                        "v": "Matched"
                                    }
                                ]
                            },
                            "merge": null,
                            "w": 55,
                            "tb": "2",
                            "fc": "#0000ff",
                            "fs": "16"
                        }
                    },
                    {
                        "r": 5,
                        "c": 4,
                        "v": {
                            "ct": {
                                "fa": "General",
                                "t": "inlineStr",
                                "s": [
                                    {
                                        "fs": "16",
                                        "v": "□ "
                                    },
                                    {
                                        "vt": "0",
                                        "ht": "1",
                                        "fs": "9",
                                        "un": 0,
                                        "bl": 1,
                                        "fc": "#0000ff",
                                        "ff": "\"Tahoma\"",
                                        "m": "Variances",
                                        "v": "Variances"
                                    }
                                ]
                            },
                            "merge": null,
                            "w": 55,
                            "tb": "2",
                            "fc": "#0000ff",
                            "fs": "16"
                        }
                    },
                    {
                        "r": 6,
                        "c": 4,
                        "v": {
                            "ct": {
                                "fa": "General",
                                "t": "inlineStr",
                                "s": [
                                    {
                                        "fs": "16",
                                        "v": "□ "
                                    },
                                    {
                                        "vt": "0",
                                        "ht": "1",
                                        "fs": "9",
                                        "un": 0,
                                        "bl": 1,
                                        "fc": "#0000ff",
                                        "ff": "\"Tahoma\"",
                                        "m": "Details not available in the document",
                                        "v": "Details not available in the document"
                                    }
                                ]
                            },
                            "merge": null,
                            "w": 55,
                            "tb": "2",
                            "fc": "#0000ff",
                            "fs": "16"
                        }
                    },
                ];

                const allRows = [...sheetDataTable1, ...matchedUnMatchedFilter, ...sheetDataTable2];
                if (sheetDataTable1 && sheetDataTable1?.length > 0) {
                    const tableColumnDetails1 = tableColumnDetails;
                    tableColumnDetails1["Table 1"] = { "columnNames": table1json.map((e) => e?.Headers), "range": { "start": 0, "end": sheetDataTable1[sheetDataTable1?.length - 1]?.r } }
                    setTableColumnDetails(tableColumnDetails1);
                }
                allRows.sort((a, b) => a.r - b.r);

                dummyData.push(...allRows);
                Policy_appDataConfig.demo.config.borderInfo.push({
                    "rangeType": "range",
                    "borderType": "border-all",
                    "color": "#000",
                    "style": "1",
                    "range": [
                        {
                            "left": 857,
                            "width": 250,
                            "top": 114,
                            "height": 50,
                            "left_move": 857,
                            "width_move": 250,
                            "top_move": 114,
                            "height_move": 122,
                            "row": [
                                3,
                                6
                            ],
                            "column": [
                                4,
                                4
                            ],
                            "row_focus": 4,
                            "column_focus": 4
                        }
                    ]
                });
                Policy_appDataConfig.demo.celldata = dummyData;
                allRows.forEach((row) => {
                    if (sheetDataTable1.includes(row)) {
                        Policy_appDataConfig.demo.config.borderInfo.push({
                            "rangeType": "cell",
                            "value": {
                                "row_index": row?.r,
                                "col_index": row?.c,
                                "l": {
                                    "style": 1,
                                    "color": "#000"
                                },
                                "r": {
                                    "style": 1,
                                    "color": "#000"
                                },
                                "t": {
                                    "style": 1,
                                    "color": "#000"
                                },
                                "b": {
                                    "style": 1,
                                    "color": "#000"
                                }
                            }
                        });
                    }
                });

                // const excludedTablenames = ["JobHeader", "JobCommonDeclaration", "JobCoverages", "Tbl_ChecklistForm1", "Tbl_ChecklistForm2", "Tbl_ChecklistForm3", "Tbl_ChecklistForm4"];
                let policyDataTracked = [];
                let positioningForPolicy = [];
                let xRAyDataMap = [];
                mainData.map((e, index) => {
                    if (e?.Tablename != 'Table 1' && e?.TemplateData?.length >= 1) {   //!excludedTablenames.includes(e?.Tablename) && 
                        let filteredData = Policy_appDataConfig.demo.celldata.filter((f, index) => f != null || !f);
                        csrTable2([...filteredData], e?.Tablename);
                        const trackedData = csrTable2([...filteredData], e?.Tablename);
                        policyDataTracked = [...policyDataTracked, ...trackedData?.dataForReturn];
                        positioningForPolicy = [...positioningForPolicy, ...trackedData?.tableColumnDetailsForPolicySave];
                        xRAyDataMap = [...xRAyDataMap, ...trackedData?.tableColumnDetailsForXRay];
                    }
                    if (mainData?.length === (index + 1)) {
                        setDataForSavePolicy(policyDataTracked);
                        localStorage.setItem('policyDataTracked', JSON.stringify(policyDataTracked))
                        setDataForSavePolicyPosition(positioningForPolicy);
                        localStorage.setItem('positioningForPolicy', JSON.stringify(positioningForPolicy))
                        setDataForXRayMapping(xRAyDataMap);
                    }
                });
                csrLuckySheet();
            }
        };

        const csrTable2 = (combinedata1, tableName) => {
            if (!Array.isArray(combinedata1)) {
                return;
            }
            const tableColumnNamesOfValid = {};
            const needDocumentViewer = false;
            let DefaultColumns;
            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                DefaultColumns = ["Actions on Discrepancy (from AMs)", "ActionOnDiscrepancy", "Notes", "NotesFreeFill"];
            } else {
                DefaultColumns = ["Actions on Discrepancy (from AMs)", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
            }
            const basedata = [...combinedata1];
            const parsedData = mainData;
            const parseTemplateData = (data) => {
                return data.map(item => ({
                    ...item,
                    TemplateData: typeof item.TemplateData == 'object' ? item.TemplateData : JSON.parse(item.TemplateData)
                }));
            };

            const inputData = parseTemplateData(parsedData);
            setCsrPolicydata(inputData);

            let defaultText = csrPolicyData2_3[0];
            if (tableName == 'Table 2' || tableName == 'Table 3') {
                let propsUpdateData = JSON.parse(defaultText.TemplateData);
                for (let i = 0; i < inputData.length; i++) {
                    if (inputData[i].TemplateData.length === 0) {
                        inputData[i].TemplateData = propsUpdateData;
                    }
                }
            }

            inputData.forEach(item => {
                if (item.Tablename != 'Table 1') {
                    if (Array.isArray(item.TemplateData)) {
                        item.TemplateData.forEach(data => {
                            if (typeof data === 'object' && data !== null) {
                                if (!data.hasOwnProperty('COVERAGE_SPECIFICATIONS_MASTER')) {
                                    data.CoverageSpecificationsMaster = null;
                                }
                            }
                        });
                    }
                }
            });

            const tableData2 = inputData.find((data) => data.Tablename === tableName);
            let applicaleLob = inputData[0].AvailableLobs;
            const policyLOBValuesForLobSplit = tableData2?.TemplateData?.filter(item =>
                applicaleLob.includes(item["POLICY LOB"] || item["Policy LOB"])
            );

            if (policyLOBValuesForLobSplit && policyLOBValuesForLobSplit?.length > 0) {
                tableName = 'Table 3';
                // tableData2.Tablename = 'Table 3';
            }

            if (!tableData2) {
                //console.error( "Table 2 data not found" );
                return;
            }

            const table22sonCopy = tableData2.TemplateData;

            let tableColumnKeys = [];
            if (table22sonCopy && table22sonCopy?.length > 0) {
                const allKeys = Object.keys(table22sonCopy[0]);
                allKeys.map((e) => {
                    if (e) {
                        let keyHasData = table22sonCopy?.filter((f) => (f[e] != null && f[e] !== "") || (e == "Lob" && tableData2?.isMultipleLobSplit) || (e == "COVERAGE_SPECIFICATIONS_MASTER" && (f[e] === null || f[e] === "")) || (e == "ActionOnDiscrepancy" && (f[e] === null || f[e] === ""))
                            || (e == "RequestEndorsement" && (f[e] === null || f[e] === "")) || (e == "Notes" && (f[e] === null || f[e] === "")) || (e == "NotesFreeFill" && (f[e] === null || f[e] === "")));
                        if (keyHasData?.length > 0) {
                            tableColumnKeys.push(e);
                        }
                    }
                });
                tableColumnKeys.push("Document Viewer");
            }

            const table2JsonCopy = table22sonCopy.map(obj => {
                let newObj = {};
                tableColumnKeys.forEach((key) => {
                    newObj[key] = obj[key];
                });
                return newObj;
            });

            const table2json = table2JsonCopy.map(item => {
                const {
                    JobId,
                    Jobid,
                    CreatedOn,
                    UpdatedOn,
                    Columnid,
                    columnid,
                    IsDataForSp,
                    ...filteredItem
                } = item;
                return filteredItem;

            });

            let header = Object.keys(table2json[0]);
            header = header.filter(f => !["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"]?.includes(f));
            const value = Object.values(table2json);

            const policyLOBValues = value.map(item => item["POLICY LOB"] || item["Policy LOB"]);
            let headerRows1 = [];
            let rowIndexForLOBStart = 0;
            let rowIndexForLOBEnd = 0;

            if (tableName === "Table 3") {
                rowIndexForLOBStart = basedata[basedata?.length - 1]?.r + 2;
                headerRows1 = [
                    {
                        r: basedata[basedata?.length - 1]?.r + 2,
                        rs: 1,
                        c: 1,
                        cs: header.length + 1,
                        v: {
                            ct: { fa: "General", t: "g" },
                            m: policyLOBValues[0],
                            v: policyLOBValues[0],
                            fs: 11,
                            ff: "\"Tahoma\"",
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                ]
            };


            const excludedColumns = ["Id", "POLICY LOB", "PolicyLob", "Policy LOB", "Checklist Questions", "Observation", "PageNumber", "OBSERVATION", "Page Number", "ActionOnDiscrepancy", "RequestEndorsement", "NotesforEndorsement", "NotesFreeFill", "Notes"];
            let headers = Object.keys(table2json[0]).filter(headerw => !excludedColumns.includes(headerw));
            if (policyLOBValues && policyLOBValues?.length > 0 && policyLOBValues[0] === 'Are the forms and endorsements attached, listed in current term policy?') {
                const indexListed = headers.indexOf("CurrentTermPolicyListed");
                const indexAttached = headers.indexOf("CurrentTermPolicyAttached");
                if (indexListed !== -1 && indexAttached !== -1 && indexAttached > indexListed) {
                    // Swap the elements at the identified indices
                    [headers[indexListed], headers[indexAttached]] = [headers[indexAttached], headers[indexListed]];
                }
            }
            const removalCode = headers.map(item => {
                if (tableName !== "Table 3" && (item === "COVERAGE_SPECIFICATIONS_MASTER" || item === "Coverage_Specifications_Master")) {
                    return policyLOBValues[0];
                } else if (tableName === "Table 3" && (item === "COVERAGE_SPECIFICATIONS_MASTER" || item === "Coverage_Specifications_Master")) {
                    return "COVERAGE SPECIFICATIONS";
                }
                else {
                    return item;
                }
            });


            headerRows1 = [
                ...headerRows1,
                ...removalCode.map((item, index) => {
                    if (index === 0) {
                        Policy_appDataConfig.demo.config.merge[`${tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 2 : basedata[basedata?.length - 1]?.r + 2}_${index}`] = {
                            "r": tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 2 : basedata[basedata?.length - 1]?.r + 2,
                            "c": index,
                            "rs": tableName === "Table 3" ? 3 : 2,
                            "cs": 1
                        }
                    }
                    Policy_appDataConfig.demo.config.merge[`${tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2}_${1 + index}`] = {
                        "r": tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2,
                        "c": 1 + index,
                        "rs": 2,
                        "cs": 1
                    }

                    return {
                        r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2, // Start from row 1 for headers
                        rs: 2,
                        c: 1 + index,
                        cs: 1,
                        v: {
                            ct: { fa: "General", t: "g" },
                            m: item,
                            v: item,
                            fs: 11,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }),
            ]
            if (headerRows1?.length > 0) {
                headerRows1.forEach((f, index) => {
                    if (tableName === "Table 3" && (index === 0 || index === 1)) {
                        tableColumnNamesOfValid["COVERAGE_SPECIFICATIONS_MASTER"] = f?.c;
                    } else if (index === 0) {
                        tableColumnNamesOfValid["COVERAGE_SPECIFICATIONS_MASTER"] = f?.c;
                    } else {
                        tableColumnNamesOfValid[f?.v?.v] = f?.c;
                    }
                });
            }
            //add documentviewer
            if (needDocumentViewer) {
                Policy_appDataConfig.demo.config.merge[`${tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2}_${1 + headerRows1[headerRows1?.length - 1]?.c}`] = {
                    "r": tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2,
                    "c": 1 + headerRows1[headerRows1?.length - 1]?.c,
                    "rs": 2,
                    "cs": 1
                }

                const DocumentViewer = [{
                    r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 2 : basedata[basedata?.length - 1]?.r + 1,
                    rs: 2,
                    c: 1 + headerRows1[headerRows1?.length - 1]?.c,
                    cs: 1,
                    v: {
                        ct: { fa: "General", t: "g" },
                        v: 'Document Viewer',
                        merge: null,
                        bg: "rgb(139,173,212)",
                        tb: '2',
                        w: 55,
                    }
                }];
                headerRows1 = [...headerRows1, ...DocumentViewer];
            }

            const defaultHeaderRows1 = DefaultColumns.map((item, index) => {
                if (DefaultColumns?.length === index + 1) {
                    rowIndexForLOBEnd = tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index;
                }
                if (index == 0) {
                    Policy_appDataConfig.demo.config.merge[`${tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2}_${tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1}`] = {
                        "r": tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2,
                        "c": tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1,
                        "rs": 1,
                        "cs": reqEndorsementColHideBrokerId.includes(brokerId) ? 3 : 4,
                    }
                    return {
                        r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 3 : basedata[basedata?.length - 1]?.r + 2,
                        rs: 1,
                        c: tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1,
                        cs: 1,
                        v: {
                            ht: 0,
                            ct: { fa: "General", t: "g" },
                            m: item,
                            v: item,
                            fs: 11,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }
                else {
                    return {
                        r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3,
                        rs: 1,
                        c: tableName === "Table 3" ? headerRows1.length + index - 1 : headerRows1.length + index,
                        cs: 1,
                        v: {
                            ct: { fa: "General", t: "g" },
                            m: item,
                            v: item,
                            fs: 11,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }
            });

            if (tableName === "Table 3") {
                Policy_appDataConfig.demo.config.merge[`${rowIndexForLOBStart}_${1}`] = {
                    "r": rowIndexForLOBStart,
                    "c": 1,
                    "rs": 1,
                    "cs": rowIndexForLOBEnd - 1,
                }
            }

            let headerRows1Values = [];
            let rowIndex = defaultHeaderRows1[defaultHeaderRows1.length - 1]?.r + 1;
            let actionColumnKeys;
            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                actionColumnKeys = ["ActionOnDiscrepancy", "Notes", "NotesFreeFill"];
            } else {
                actionColumnKeys = ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
            }
            headers = [...headers, ...actionColumnKeys];
            let dataToUpDate = table2json.map((item, cIndex) => {

                let rowIndexForStateData = 0; //by gokul
                let rowMaxValue = 30;
                let fs = 9;
                function escapeRegExp(str) {
                    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); // $& means the whole matched string
                }

                function splitWordsWithComma(array) {
                    if (!array || array.length === 0) {
                        return [];
                    }
                    let newArray = [];

                    array.forEach((word) => {
                        // Check if the word ends with a comma
                        word = word.trim();
                        if (word.endsWith(',')) {
                            const wordWithoutComma = word.slice(0, -1).trim();
                            // Add the word without the comma as a separate character, excluding leading spaces
                            if (wordWithoutComma !== '') {
                                newArray.push(wordWithoutComma);
                            }
                            newArray.push(',');
                        } else if (word.includes('(') && word.includes(')')) {
                            // If the word contains both '(' and ')', split them into separate characters
                            const openingParen = word.indexOf('(');
                            const closingParen = word.indexOf(')');
                            const beforeParen = word.slice(0, openingParen);
                            const insideParen = word.slice(openingParen + 1, closingParen);
                            const afterParen = word.slice(closingParen + 1);
                            if (beforeParen !== '') {
                                newArray.push(beforeParen);
                            }
                            newArray.push('(');
                            if (insideParen !== '') {
                                newArray.push(insideParen);
                            }
                            newArray.push(')');
                            if (afterParen !== '') {
                                newArray.push(afterParen);
                            }
                        } else if (word.includes('(')) {
                            // If the word contains an open parenthesis, split it into separate characters
                            const openingParen = word.indexOf('(');
                            const beforeParen = word.slice(0, openingParen);
                            const insideParen = word.slice(openingParen + 1);
                            if (beforeParen !== '') {
                                newArray.push(beforeParen);
                            }
                            newArray.push('(');
                            if (insideParen !== '') {
                                newArray.push(insideParen);
                            }
                        } else if (word.includes(')')) {
                            // If the word contains a closing parenthesis, split it into separate characters
                            let closingParen = word.indexOf(')');
                            let insideParen = word.slice(0, closingParen).trim();
                            const afterParen = word.slice(closingParen + 1);
                            if (insideParen !== '') {
                                newArray.push(insideParen);
                            }
                            newArray.push(')');
                            if (afterParen !== '') {
                                newArray.push(afterParen);
                            }
                        } else {
                            // If no comma, just add the word to the new array
                            newArray.push(word);
                        }
                    });

                    return newArray;
                }
                headers.map((key, rIndex) => {
                    rowIndex = headerRows1Values?.length == 0 ? rowIndex : headerRows1Values?.length > 0 && rIndex == 0 ? headerRows1Values[headerRows1Values.length - 1]?.r + 1 : headerRows1Values[headerRows1Values.length - 1]?.r;
                    rowIndexForStateData = rowIndex;
                    let text = item[key]?.toString()?.split('~~');
                    let ct = [];

                    if (key != 'Document Viewer' && text && text?.length > 0) {
                        text?.map((e) => {
                            if (e?.toLowerCase().includes('page #')) {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(68, 114, 196)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 1,
                                    "it": 0,
                                    "v": "\r\n" + e.trim() + "\r\n"
                                });
                            } else if (e?.toLowerCase().includes('endorsement page #')) {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(68, 114, 196)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 1,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            }
                            else if (key === "PageNumber") {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "#000000",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            }
                            else if (e?.toLowerCase().includes('current policy listed')) {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(68, 114, 196)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 1,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            } else if (e?.toLowerCase().includes('current policy endorsement listed')) {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(68, 114, 196)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 1,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            } else if (e?.toLowerCase().includes('current policy attached')) {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(68, 114, 196)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 1,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            } else if (e?.toLowerCase().includes('current policy endorsement attached')) {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(68, 114, 196)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 1,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            } else if (e === 'MATCHED') {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(0, 128, 0)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 1,
                                    "it": 0,
                                    "v": e.trim()
                                });
                            }
                            else if (key === "Prior Term Policy" && item["Prior Term Policy"]?.trim() != item["Current Term Policy"]?.trim()
                                && !(item["Prior Term Policy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Prior Term Policy"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Binder" && item["Binder"]?.trim() != item["Current Term Policy"]?.trim()
                                && !(item["Binder"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Binder"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Proposal" && item["Proposal"]?.trim() != item["Current Term Policy"]?.trim()
                                && !(item["Proposal"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Proposal"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Quote" && item["Quote"]?.trim() != item["Current Term Policy"]?.trim()
                                && !(item["Quote"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Quote"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Schedule" && item["Schedule"]?.trim() != item["Current Term Policy"]?.trim()
                                && !(item["Schedule"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Schedule"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Current Term Policy Attached" && item["Current Term Policy Attached"]?.trim() != item["Current Term Policy Listed"]?.trim()
                                && !(item["Current Term Policy Attached"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Current Term Policy Attached"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Application" && item["Application"]?.trim() != item["Current Term Policy"]?.trim()
                                && !(item["Application"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Application"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Application - Listed" && item["Application - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
                                && !(item["Application - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Application - Listed"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Quote - Listed" && item["Quote - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
                                && !(item["Quote - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Quote - Listed"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Proposal - Listed" && item["Proposal - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
                                && !(item["Proposal - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Proposal - Listed"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Binder - Listed" && item["Binder - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
                                && !(item["Binder - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Binder - Listed"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Schedule - Listed" && item["Schedule - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
                                && !(item["Schedule - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Schedule - Listed"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Prior Term Policy - Listed" && item["Prior Term Policy - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
                                && !(item["Prior Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Prior Term Policy - Listed"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Prior Term Policy - Listed" && item["Prior Term Policy - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
                                && !(item["Prior Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Prior Term Policy - Listed"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else if (key === "Current Term Policy Attached" && item["Current Term Policy Attached"]?.trim() != item["Current Term Policy Listed"]?.trim()
                                && !(item["Current Term Policy Attached"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                    || item["Current Term Policy Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                let ptpSplitArray = item["Current Term Policy Attached"]?.split('~~')[0]?.split(" ");
                                let ctpSplitArray = item["Current Term Policy Listed"]?.split('~~')[0]?.split(" ");

                                const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach((ptpe) => {
                                    let css = "#000000";
                                    let ctpText = ctpFlattenedArray.join(" ");

                                    if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                        css = "#000000";
                                    } else {
                                        let pattern = new RegExp(`\\b${escapeRegExp(ptpe.trim())}\\b`, 'i');
                                        const ctpWordsArray = ctpText.split(' ');

                                        // Check if each word in ptpe is present in ctpWordsArray
                                        ptpe.split(' ').forEach((word) => {
                                            if (!ctpWordsArray.includes(word.trim())) {
                                                css = "#ff0000";
                                            }
                                        });

                                        if (!pattern.test(ctpText)) {
                                            css = "#ff0000";
                                        }
                                    }
                                    ct.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": css,
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": ptpe.trim() + " "
                                    });
                                });
                            }
                            else {
                                ct.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "#000000",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            }
                        })
                    }
                    if (key === 'Document Viewer') {

                        const dvData = item[key];
                        if (dvData != undefined && dvData != null && dvData?.trim() != '') {
                            ct.push({
                                "ff": "\"Tahoma\"",
                                "fc": "rgb(61, 133, 198)",
                                "fs": `${fs}`,
                                "cl": 0,
                                "un": 0,
                                "bl": 0,
                                "it": 0,
                                "ht": "0",
                                "un": 1,
                                "v": "X-Ray"
                            });
                        } else {
                            ct.push({
                                "ff": "\"Tahoma\"",
                                "fc": "#000000",
                                "fs": `${fs}`,
                                "cl": 0,
                                "un": 0,
                                "bl": 0,
                                "it": 0,
                                "v": "  "
                            });
                        }
                    } else if (key === 'ActionOnDiscrepancy' || key === 'RequestEndorsement' || key === 'Notes') {
                        const dvData = item[key];
                        if (dvData == "" || dvData == undefined || dvData == null) {
                            ct.push({
                                "ff": "\"Tahoma\"",
                                "fc": "rgba(171, 160, 160, 0.957)",
                                "fs": `${fs}`,
                                "cl": 0,
                                "un": 0,
                                "bl": 0,
                                "it": 0,
                                "ht": "0",
                                "v": "Click here"
                            });
                        }
                    }
                    if (key === "PageNumber") {
                        const textOfct = ct[ct.length - 1]?.v?.replace('\r\n', ' ');
                        if (textOfct && ct?.length > 0) {
                            ct[ct.length - 1]["v"] = textOfct;
                        }
                    }
                    if (item[key]) {
                        const text = item[key];
                        if (typeof text === 'string') {
                            rowMaxValue = rowMaxValue < (text?.length / 2) + 10 ? (text?.length / 3 + 25) : rowMaxValue;
                        }
                    }

                    headerRows1Values.push({
                        r: rowIndex,
                        c: rIndex + 1,
                        v: {
                            ct: { fa: "General", t: "inlineStr", s: ct },
                            // m: item[key]?.replace(/~~/g, '\r\n'), 
                            // v: item[key]?.replace(/~~/g, '\r\n'), 
                            merge: null,
                            w: 55,
                            tb: '2',
                            "ht": key === 'Document Viewer' ? "0" : key === 'ActionOnDiscrepancy' ? "0" : key === 'RequestEndorsement' ? "0" :
                                key === 'Notes' ? "0" : null
                        }
                    });
                    let maxLength = 0;
                    const lengths = [];
                    Object.keys(item).forEach((key) => {
                        if (key != "Document Viewer") {
                            if (item[key]) {
                                lengths.push(item[key].length);
                            }
                        }
                    });
                    lengths.forEach((length) => {
                        if (length > maxLength) {
                            maxLength = length;
                        }
                    });
                    const rowHeight = parseInt(maxLength && maxLength > 30 ? maxLength / 3 + 15 : 20);
                    Policy_appDataConfig.demo.config.rowlen[`${rowIndex}`] = rowHeight;

                    if (rIndex == 0) {
                        Policy_appDataConfig.demo.config.rowlen[`${rowIndex}`] = rowHeight;
                    }
                });
                item["sheetPosition"] = rowIndexForStateData;
                return item;
            });
            // if ( dataForSavePolicy?.length > 0 ){
            //     const hasData = dataForSavePolicy?.filter( ( f ) => f?.tableName === tableName );
            //     if(hasData.length > 0){
            //         let trackingDataPolicy = dataForSavePolicy.map((e) => {
            //             if ( e?.TableName === tableName ){
            //                 e.data = dataToUpDate;
            //             }
            //             return e;
            //         });
            //         setDataForSavePolicy( trackingDataPolicy );
            //     }else{
            //         setDataForSavePolicy( [ ...dataForSavePolicy ,... [ { "TableName": tableName, "data": dataToUpDate } ]] );
            //     }
            // }else{
            //     setDataForSavePolicy( [ { "TableName": tableName, "data": dataToUpDate }]);
            // }

            const dataForReturn = [{ "TableName": tableName, "data": dataToUpDate }];
            const xRaydataForReturn = [{ dataToUpDate }];

            Policy_appDataConfig.demo.config.borderInfo.push({
                "rangeType": "range",
                "borderType": "border-all",
                "color": "#000",
                "style": "1",
                "range": [
                    {
                        "left": 74,
                        "width": 300,
                        "top": 470,
                        "height": 42,
                        "left_move": 74,
                        "width_move": 4213,
                        "top_move": 471,
                        "height_move": 1107,
                        "row": [
                            headerRows1[0]?.r,
                            headerRows1Values[headerRows1Values?.length - 1]?.r
                        ],
                        "column": [
                            headerRows1[0]?.c,
                            defaultHeaderRows1[defaultHeaderRows1?.length - 1]?.c
                        ],
                        "row_focus": headerRows1[0]?.r,
                        "column_focus": headerRows1[0]?.c
                    }
                ]
            });
            Policy_appDataConfig.demo.config.borderInfo.push({
                "rangeType": "range",
                "borderType": "border-all",
                "color": "#000",
                "style": "1",
                "range": [
                    {
                        "left": 74,
                        "width": 300,
                        "top": 470,
                        "height": 42,
                        "left_move": 74,
                        "width_move": 4213,
                        "top_move": 471,
                        "height_move": 1107,
                        "row": [
                            headerRows1[0]?.r,
                            headerRows1[0]?.r + (tableName == "Table 3" ? 2 : 1)
                        ],
                        "column": [
                            0,
                            0
                        ],
                        "row_focus": headerRows1[0]?.r,
                        "column_focus": headerRows1[0]?.c
                    }
                ]
            });
            defaultHeaderRows1.forEach(row => {
                if (row?.v && row?.v?.m && typeof row?.c === 'number' && row.v.m !== 'Actions on Discrepancy (from AMs)') {
                    tableColumnNamesOfValid[row.v.m] = row?.c;
                }
            });

            const groupingHeader = [{
                "r": headerRows1[0]?.r,
                "rs": tableName == "Table 3" ? 3 : 2,
                "c": 0,
                "cs": 1,
                "v": {
                    "ct": {
                        "fa": "General",
                        "t": "g"
                    },
                    "m": "Grouping",
                    "v": "Grouping",
                    "fs": 11,
                    "ff": "\"Tahoma\"",
                    "merge": null,
                    "bg": "rgb(139,173,212)",
                    "tb": "2",
                    "w": 55
                }
            }]

            headerRows1 = [...groupingHeader, ...headerRows1, ...defaultHeaderRows1, ...headerRows1Values];

            const allRows2 = [...headerRows1];
            allRows2.sort((a, b) => a.r - b.r);
            if (allRows2 && allRows2?.length > 0) {
                const tableColumnDetailss = tableColumnDetails;
                tableColumnDetailss[tableData2?.Tablename] = { "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[0]?.r, "end": allRows2[allRows2?.length - 1]?.r } }
                setTableColumnDetails(tableColumnDetailss);
            }


            basedata.push(...allRows2);
            Policy_appDataConfig.demo.celldata = basedata;
            return {
                dataForReturn, xRaydataForReturn, tableColumnDetailsForXRay: [{ "TableName": tableName, "Data": xRaydataForReturn, "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[0]?.r, "end": allRows2[allRows2?.length - 1]?.r } }],
                tableColumnDetailsForPolicySave: [{ "TableName": tableName, "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[0]?.r, "end": allRows2[allRows2?.length - 1]?.r } }]
            };
        };
        if (canRenderChecklist) {
            csrTable1();
        }

        const csrformTable1 = () => {
            const formTableData1 = formCompareData.find((data) => data.Tablename === "FormTable 1");
            if (formTableData1) {
                const formtable1 = JSON.parse(formTableData1.TemplateData);

                let sheetDataTable3 = [];
                let sheetDataTable4 = [];
                const rowIndexOfTable1 = 3
                const formData = renderForm();
                formData.forEach((item, index) => {
                    const mergeConfig = FormCompare_appconfigdata.forms.config.merge["1_1"];
                    sheetDataTable3.push({
                        r: 1 + mergeConfig.r,
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            fs: 13,
                            ff: item.ff,
                            merge: mergeConfig,
                            fc: item.fc,
                        }
                    });
                });

                formtable1.map((item, index) => {
                    if (item["Headers"] != null) {
                        // Calculate row height
                        let maxLength = 0;
                        const lengths = [];
                        Object.keys(item).forEach((key) => {
                            if (item[key]) {
                                lengths.push(item[key].length);
                            }
                        });
                        lengths.forEach((length) => {
                            if (length > maxLength) {
                                maxLength = length;
                            }
                        });
                        const rowHeight = parseInt(maxLength / 3 + 15);
                        FormCompare_appconfigdata.forms.config.rowlen[`${rowIndexOfTable1 + index}`] = rowHeight;

                        sheetDataTable4.push({
                            r: rowIndexOfTable1 + index,
                            c: 1,
                            v: {
                                ct: { fa: "@", t: "inlineStr", s: [{ v: item["Headers"], ff: "Tahoma" }] },
                                m: item["Headers"],
                                v: item["Headers"],
                                fs: 9,
                                ff: "Tahoma",
                                merge: null,
                                bg: "rgb(139,173,212)",
                                tb: '2',
                            }
                        });

                        const tidleValue = item["(No column name)"] !== "" ? item["(No column name)"].replace(/~~/g, "\n") : "";
                        sheetDataTable4.push({
                            r: rowIndexOfTable1 + index,
                            c: 2,
                            v: {
                                ct: { fa: "@", t: "inlineStr" },
                                m: tidleValue,
                                v: tidleValue,
                                ff: "\"Tahoma\"",
                                merge: null,
                                tb: '2',
                            }
                        });
                    }
                });

                const dummyData1 = [];
                const allFormRows = [...sheetDataTable4, ...sheetDataTable3];
                if (sheetDataTable4 && sheetDataTable4?.length > 0) {
                    const formTableColumnDetails1 = formTableColumnDetails;
                    formTableColumnDetails1["FormTable 1"] = { "columnNames": formtable1.map((e) => e?.Headers), "range": { "start": 0, "end": sheetDataTable4[sheetDataTable4?.length - 1]?.r } }
                    setFormTableColumnDetails(formTableColumnDetails1);
                }
                allFormRows.sort((a, b) => a.r - b.r);

                dummyData1.push(...allFormRows);
                FormCompare_appconfigdata.forms.celldata = dummyData1;

                allFormRows.forEach((row) => {
                    if (sheetDataTable4.includes(row)) {
                        FormCompare_appconfigdata.forms.config.borderInfo.push({
                            "rangeType": "cell",
                            "value": {
                                "row_index": row?.r,
                                "col_index": row?.c,
                                "l": {
                                    "style": 1,
                                    "color": "#000"
                                },
                                "r": {
                                    "style": 1,
                                    "color": "#000"
                                },
                                "t": {
                                    "style": 1,
                                    "color": "#000"
                                },
                                "b": {
                                    "style": 1,
                                    "color": "#000"
                                }
                            }
                        });
                    }
                });

                let formsCompareDataTracked = [];
                let positioningForFormsCompareData = [];
                let xRAyFormDataMap = [];
                formCompareData.map((e, index) => {
                    if (e?.Tablename != 'FormTable 1' && (e?.TemplateData?.length >= 3 || e?.TemplateData?.length < 3)) {
                        let filteredData = FormCompare_appconfigdata.forms.celldata.filter((f, index) => f != null || !f);
                        csrformTable2([...filteredData], e?.Tablename);
                        const trackedData = csrformTable2([...filteredData], e?.Tablename);
                        formsCompareDataTracked = [...formsCompareDataTracked, ...trackedData?.formsdataForReturn];
                        positioningForFormsCompareData = [...positioningForFormsCompareData, ...trackedData?.formTableColumnDetailsForSave];
                        xRAyFormDataMap = [...xRAyFormDataMap, ...trackedData?.formTableColumnDetailsForXRay];
                    }
                    if (formCompareData?.length === (index + 1)) {
                        setDataForSaveFormsCompare(formsCompareDataTracked);
                        localStorage.setItem('formsCompareDataTracked', JSON.stringify(formsCompareDataTracked))
                        setDataForSaveFormsComparePosition(positioningForFormsCompareData);
                        localStorage.setItem('positioningForFormsCompareData', JSON.stringify(positioningForFormsCompareData))
                        setFormDataForXRayMapping(xRAyFormDataMap);
                    }
                });
            }
            csrLuckySheet();
        }

        const csrformTable2 = (combinedata1, tableName) => {
            if (!Array.isArray(combinedata1)) {
                return;
            }
            const tableColumnNamesOfValid = {};
            const needDocumentViewer = true;
            const DefaultColumns = ["Actions on Discrepancy (from AMs)"];
            const basedata = [...combinedata1];

            const parsedData = formstate;
            const parseTemplateData = (data) => {
                return data.map(item => ({
                    ...item,
                    TemplateData: JSON.parse(item.TemplateData)
                }));
            };

            let formCompareData = parseTemplateData(parsedData);
            const isMatchedSection = formCompareData && formCompareData?.length > 0 ? formCompareData.filter((f) => f?.IsMatched === true)?.length > 0 : false;
            if (isMatchedSection) {
                formCompareData = formCompareData.map(({ "Document Viewer": _, ...rest }) => rest);
            }
            setFormsComparedata(formCompareData);

            let defaultText = updateData[0];
            let propsUpdateData = JSON.parse(defaultText.TemplateData);
            for (let i = 0; i < formCompareData.length; i++) {
                if (formCompareData[i].TemplateData.length === 0) {
                    formCompareData[i].TemplateData = propsUpdateData;
                }
            }

            formCompareData.forEach((data) => {
                if (data.TemplateData && Array.isArray(data.TemplateData)) {
                    data.TemplateData.forEach((template) => {
                        Object.keys(template).forEach((key) => {
                            if (template[key] === null) {
                                template[key] = '';
                            }
                        });
                    });
                }
            });

            const formTableData2 = formCompareData.find((data) => data.Tablename === tableName && data.TemplateData.length > 0);
            if (formTableData2?.TemplateData?.length > 0) {
                const headersKeys = Object.keys(formTableData2?.TemplateData[0]);
                headersKeys.forEach((column) => {
                    if (formTableData2?.TemplateData?.filter((f) => f[column] != null)?.length > 0 || (tableName === "FormTable 2" && formTableData2) || (tableName === "FormTable 3" && formTableData2)) {
                        tableColumnNamesOfValid[column] = 0
                    }
                });
            }

            if (!formTableData2) {
                return;
            }
            const formtable2copy = formTableData2.TemplateData;

            const formDataCopy = formtable2copy.map(obj => {
                let newObj = {};
                Object.keys(obj).forEach(key => {
                    if (obj[key] !== null) {
                        newObj[key] = obj[key];
                    }
                });
                return newObj;
            });

            const formtable2 = formDataCopy.map(item => {
                const {
                    JobId,
                    Jobid,
                    CreatedOn,
                    UpdatedOn,
                    columnid,
                    IsMatched,
                    ...filteredItem
                } = item;
                return filteredItem;

            });
            const header = Object.keys(formtable2[0]);
            const value = Object.values(formtable2);
            const policyLOBValues = value.map(item => item["Policy LOB"]);

            let headerRows1 = [];
            let rowIndexForLOBStart = 0;
            let rowIndexForLOBEnd = 0;

            if (tableName === "FormTable 2" || tableName === "FormTable 3") {
                rowIndexForLOBStart = basedata[basedata?.length - 1]?.r + 2;
                headerRows1 = [
                    {
                        r: basedata[basedata?.length - 1]?.r + 2,
                        rs: 1,
                        c: 1,
                        cs: header.length + 1,
                        v: {
                            ct: { fa: "General", t: "g" },
                            m: tableName === "FormTable 2" ? "Unmatched Forms" : "Matched Forms",
                            v: tableName === "FormTable 2" ? "Unmatched Forms" : "Matched Forms",
                            fs: 9,
                            ff: "\"Tahoma\"",
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                ]
            };

            let excludedColumns;
            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                excludedColumns = ["Policy LOB", "Checklist Questions", "Observation", "PageNumber", "OBSERVATION", "Page Number", "Id", "RequestEndorsement"];
            } else {
                excludedColumns = ["Policy LOB", "Checklist Questions", "Observation", "PageNumber", "OBSERVATION", "Page Number", "Id"];
            }

            let headers = Object.keys(formtable2[0]).filter(headerw => !excludedColumns.includes(headerw));

            const AODColumnKeys = ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];

            // const filteredHeadersForEndorsement = brokerId === "1003"  
            const filteredHeadersForEndorsement = AODDBrokerIds.includes(brokerId)
                ? headers :
                headers.filter(item => !AODColumnKeys.includes(item))

            const removalCode = filteredHeadersForEndorsement.map(item => (item === "COVERAGE_SPECIFICATIONS_MASTER") ? policyLOBValues[0] : item);

            headerRows1 = [
                ...headerRows1,
                ...removalCode.map((item, index) => {
                    // if(brokerId != "1003" && index === 0){
                    if (AODDBrokerIds.includes(brokerId) && index === 0) {
                        if (!AODColumnKeys?.includes(item)) {
                            FormCompare_appconfigdata.forms.config.merge[`${basedata[basedata?.length - 1]?.r + 3}_${index + 1}`] = {
                                "r": basedata[basedata?.length - 1]?.r + 3,
                                "c": index,
                                "rs": 3,
                                "cs": 1
                            }
                        }
                    }
                    if (!AODColumnKeys?.includes(item)) {
                        FormCompare_appconfigdata.forms.config.merge[`${basedata[basedata?.length - 1]?.r + 3}_${1 + index}`] = {
                            "r": basedata[basedata?.length - 1]?.r + 3,
                            "c": 1 + index,
                            "rs": 2,
                            "cs": 1
                        }
                    }

                    // if (removalCode?.length === index + 1) {
                    //     rowIndexForLOBEnd = headerRows1.length + index + 1;
                    // }
                    return {
                        r: basedata[basedata?.length - 1]?.r + 3 + (AODColumnKeys?.includes(item) ? 1 : 0),
                        rs: 2,
                        c: 1 + index,
                        cs: 1,
                        v: {
                            ct: { fa: "General", t: "g" },
                            m: item,
                            v: item,
                            fs: 9,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                })
            ];

            if (headerRows1?.length > 0) {
                headerRows1.forEach((f, index) => {
                    if (tableName === "FormTable 3" && (index === 0 || index === 1)) {
                        tableColumnNamesOfValid["COVERAGE_SPECIFICATIONS_MASTER"] = f?.c;
                    } else if (index === 0) {
                        tableColumnNamesOfValid["COVERAGE_SPECIFICATIONS_MASTER"] = f?.c;
                    } else {
                        tableColumnNamesOfValid[f?.v?.v] = f?.c;
                    }
                });
            }

            if (needDocumentViewer) {
                const DocumentViewer = [{
                    r: basedata[basedata?.length - 1]?.r + 3,
                    rs: 2,
                    c: 1 + headerRows1[headerRows1?.length - 1]?.c,
                    cs: 1,
                    v: {
                        ct: { fa: "General", t: "g" },
                        m: 'Document Viewer',
                        v: 'Document Viewer',
                        ff: "\"Tahoma\"",
                        merge: null,
                        bg: "rgb(139,173,212)",
                        tb: '2',
                        w: 55,
                    }
                }];

                // headerRows1 = [...headerRows1, ...DocumentViewer];
                headerRows1 = [...headerRows1];
            }

            const defaultHeaderRows1 = DefaultColumns.map((item, index) => {
                if (DefaultColumns?.length === index + 1) {
                    rowIndexForLOBEnd = headerRows1.length + index + 1;
                }
                // if (brokerId != "1003" && index == 0) {
                if (AODDBrokerIds.includes(brokerId) && index == 0) {
                    FormCompare_appconfigdata.forms.config.merge[`${basedata[basedata?.length - 1]?.r + 3}_${reqEndorsementColHideBrokerId.includes(brokerId) ? headerRows1.length - 3 : headerRows1.length - 4}`] = {
                        "r": basedata[basedata?.length - 1]?.r + 3,
                        "c": reqEndorsementColHideBrokerId.includes(brokerId) ? headerRows1.length - 3 : headerRows1.length - 4,
                        "rs": 1,
                        "cs": reqEndorsementColHideBrokerId.includes(brokerId) ? 3 : 4,
                    }
                    return {
                        r: basedata[basedata?.length - 1]?.r + 3,
                        rs: 1,
                        c: reqEndorsementColHideBrokerId.includes(brokerId) ? headerRows1.length - 3 : headerRows1.length - 4,
                        cs: 1,
                        v: {
                            ht: 0,
                            ct: { fa: "General", t: "g" },
                            m: item,
                            v: item,
                            fs: 9,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }
                else {
                    return {
                        r: basedata[basedata?.length - 1]?.r + 3,
                        rs: 1,
                        c: headerRows1.length + index - 1,
                        cs: 1,
                        v: {
                            ct: { fa: "General", t: "g" },
                            m: item,
                            v: item,
                            fs: 9,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }
            });

            if (tableName === "FormTable 2" || tableName === "FormTable 3") {
                FormCompare_appconfigdata.forms.config.merge[`${rowIndexForLOBStart}_${1}`] = {
                    "r": rowIndexForLOBStart,
                    "c": 1,
                    "rs": 1,
                    "cs": rowIndexForLOBEnd - 2
                }
            }

            let headerRows1Values = [];
            let fs = 9;

            function splitWordsWithComma(array) {
                if (!array || array.length === 0) {
                    return [];
                }
                let newArray = [];

                array.forEach((word) => {
                    // Check if the word ends with a comma
                    word = word.trim();
                    if (word.endsWith(',')) {
                        const wordWithoutComma = word.slice(0, -1).trim();
                        // Add the word without the comma as a separate character, excluding leading spaces
                        if (wordWithoutComma !== '') {
                            newArray.push(wordWithoutComma);
                        }
                        newArray.push(',');
                    } else if (word.includes('(') && word.includes(')')) {
                        // If the word contains both '(' and ')', split them into separate characters
                        const openingParen = word.indexOf('(');
                        const closingParen = word.indexOf(')');
                        const beforeParen = word.slice(0, openingParen);
                        const insideParen = word.slice(openingParen + 1, closingParen);
                        const afterParen = word.slice(closingParen + 1);
                        if (beforeParen !== '') {
                            newArray.push(beforeParen);
                        }
                        newArray.push('(');
                        if (insideParen !== '') {
                            newArray.push(insideParen);
                        }
                        newArray.push(')');
                        if (afterParen !== '') {
                            newArray.push(afterParen);
                        }
                    } else if (word.includes('(')) {
                        // If the word contains an open parenthesis, split it into separate characters
                        const openingParen = word.indexOf('(');
                        const beforeParen = word.slice(0, openingParen);
                        const insideParen = word.slice(openingParen + 1);
                        if (beforeParen !== '') {
                            newArray.push(beforeParen);
                        }
                        newArray.push('(');
                        if (insideParen !== '') {
                            newArray.push(insideParen);
                        }
                    } else if (word.includes(')')) {
                        // If the word contains a closing parenthesis, split it into separate characters
                        let closingParen = word.indexOf(')');
                        let insideParen = word.slice(0, closingParen).trim();
                        const afterParen = word.slice(closingParen + 1);
                        if (insideParen !== '') {
                            newArray.push(insideParen);
                        }
                        newArray.push(')');
                        if (afterParen !== '') {
                            newArray.push(afterParen);
                        }
                    } else {
                        // If no comma, just add the word to the new array
                        newArray.push(word);
                    }
                });

                return newArray;
            }

            let rowIndex = basedata[basedata?.length - 1]?.r + 5;
            let dataToUpDate = formtable2.map((item, cIndex) => {
                let rowIndexForStateData = 0;
                // let filteredHeadersForEndorsement = brokerId === "1003" ? headers.filter(item => !AODColumnKeys.includes(item)) : headers;
                let filteredHeadersForEndorsement = AODDBrokerIds.includes(brokerId) ? headers : headers.filter(item => !AODColumnKeys.includes(item));
                filteredHeadersForEndorsement.map((key, rIndex) => {
                    if (item[key] !== null) {
                        rowIndex = headerRows1Values?.length == 0 ? rowIndex : headerRows1Values?.length > 0 && rIndex == 0 ? headerRows1Values[headerRows1Values.length - 1]?.r + 1 : headerRows1Values[headerRows1Values.length - 1]?.r;
                        rowIndexForStateData = rowIndex;
                        let text = item[key].toString().split('~~');
                        let ss = [];
                        if (key != "Document Viewer" && text && text?.length > 0) {
                            text.map((e, splitIndex) => {
                                if (e.toLowerCase().includes('page')) {
                                    ss.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": "\r\n" + e.trim()
                                    });
                                }
                                else if (key === "Prior Term Policy Attached"
                                    && item["Prior Term Policy Attached"]?.trim() !== item["Current Term Policy Attached"]?.trim()
                                    && !(item["Prior Term Policy Attached"]?.toLowerCase().replace(/\\r\\n/g, '').includes("details not available in the document")
                                        || item["Current Term Policy Attached"]?.toLowerCase().replace(/\\r\\n/g, '').includes("details not available in the document"))) {

                                    let ptpSplitArray = item["Prior Term Policy Attached"]?.split('~~')[0]?.split(" ");
                                    let ctpSplitArray = item["Current Term Policy Attached"]?.split('~~')[0]?.split(" ");

                                    const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                                    const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                                    if (ctpFlattenedArray && ctpFlattenedArray.length > 0) {
                                        for (let i = 0; i < ptpFlattenedArray.length; i++) {
                                            let ptpe = ptpFlattenedArray[i].trim();
                                            let ctp = ctpFlattenedArray[i]?.trim();
                                            let css = "#000000"; // Default color

                                            // Apply specific CSS color if certain characters are present
                                            if (ptpe.includes("$||") || ptpe.includes("||") || ptpe.includes("(") || ptpe.includes(")")) {
                                                css = "#000000";
                                            } else {
                                                // Set color to red if there's a mismatch at the same index
                                                if (ptpe !== ctp) {
                                                    css = "#ff0000";
                                                }
                                            }

                                            ss.push({
                                                "ff": "\"Tahoma\"",
                                                "fc": css,
                                                "fs": `${fs}`,
                                                "cl": 0,
                                                "un": 0,
                                                "bl": 0,
                                                "it": 0,
                                                "v": ptpe + " "
                                            });
                                        }
                                    }
                                }
                                else {
                                    ss.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": "#000000",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": e.trim() + " "
                                    });
                                }
                            });
                        }
                        if (key === "Document Viewer") {
                            const dvData = item[key];
                            if (dvData != undefined && dvData != null && dvData?.trim() != '') {
                                ss.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(61, 133, 198)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "ht": "0",
                                    "v": "X-Ray"
                                });
                            }
                        }
                        // if ( brokerId != "1003" && (key === 'ActionOnDiscrepancy' || key === 'RequestEndorsement' || key === 'Notes')) {
                        if (AODDBrokerIds.includes(brokerId) && (key === 'ActionOnDiscrepancy' || key === 'RequestEndorsement' || key === 'Notes')) {
                            const dvData = item[key];
                            if (dvData == "" || dvData == undefined || dvData == null) {
                                ss.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgba(171, 160, 160, 0.957)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "ht": "0",
                                    "v": "Click here"
                                });
                            }
                        }
                        headerRows1Values.push({
                            r: rowIndex,
                            c: rIndex + 1,
                            v: {
                                ct: { fa: "General", t: "inlineStr", s: ss },
                                merge: null,
                                w: 55,
                                ff: "\"Tahoma\"",
                                tb: '2',
                                "ht": key === 'Document Viewer' ? "0" : key === 'ActionOnDiscrepancy' ? "0" : key === 'RequestEndorsement' ? "0" :
                                    key === 'Notes' ? "0" : null
                            }
                        });
                        let maxLength = 0;
                        const lengths = [];

                        Object.keys(item).forEach((key) => {
                            if (item[key]) {
                                lengths.push(item[key].length);
                            }
                        });
                        lengths.forEach((length) => {
                            if (length > maxLength) {
                                maxLength = length;
                            }
                        });
                        const len = maxLength > 100 ? 40 : maxLength;
                        const rowHeight = parseInt(len);
                        FormCompare_appconfigdata.forms.config.rowlen[`${rowIndex}`] = rowHeight;

                        if (rIndex == 0) {
                            FormCompare_appconfigdata.forms.config.rowlen[`${rowIndex}`] = rowHeight;
                        }
                    }
                })
                item["sheetPosition"] = rowIndexForStateData;
                return item;
            });

            const formsdataForReturn = [{ "TableName": tableName, "data": dataToUpDate }];
            const xRayformsdataForReturn = [{ dataToUpDate }];

            FormCompare_appconfigdata.forms.config.borderInfo.push({
                "rangeType": "range",
                "borderType": "border-all",
                "color": "#000",
                "style": "1",
                "range": [
                    {
                        "left": 74,
                        "width": 300,
                        "top": 470,
                        "height": 42,
                        "left_move": 74,
                        "width_move": 4213,
                        "top_move": 471,
                        "height_move": 1107,
                        "row": [
                            headerRows1[0]?.r,
                            headerRows1Values[headerRows1Values?.length - 1]?.r
                        ],
                        "column": [
                            headerRows1[0]?.c,
                            headerRows1Values[headerRows1Values?.length - 1]?.c
                        ],
                        "row_focus": headerRows1[0]?.r,
                        "column_focus": headerRows1[0]?.c
                    }
                ]
            });


            // if(brokerId === "1003") {
            if (!AODDBrokerIds.includes(brokerId)) {
                headerRows1 = [...headerRows1, ...headerRows1Values];
            } else {
                headerRows1 = [...headerRows1, ...defaultHeaderRows1, ...headerRows1Values];
            }

            const allRows2 = [...headerRows1];
            if (allRows2 && allRows2?.length > 0) {
                const formTableColumnDetailss = formTableColumnDetails;
                formTableColumnDetailss[formTableData2?.Tablename] = { "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[0]?.r, "end": allRows2[allRows2?.length - 1]?.r } }
                setFormTableColumnDetails(formTableColumnDetailss);
            }
            allRows2.sort((a, b) => a.r - b.r);

            basedata.push(...allRows2);
            FormCompare_appconfigdata.forms.celldata = basedata;
            return {
                formsdataForReturn, xRayformsdataForReturn, formTableColumnDetailsForXRay: [{ "TableName": tableName, "Data": xRayformsdataForReturn, "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[0]?.r, "end": allRows2[allRows2?.length - 1]?.r } }],
                formTableColumnDetailsForSave: [{ "TableName": tableName, "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[0]?.r, "end": allRows2[allRows2?.length - 1]?.r } }]
            };
        };
        if (canRenderFormData) {
            csrformTable1();
        }

        const csrexclusionTable = () => {
            const basedata = [];
            const data = exclusionData;
            const dataMap = data;

            let DefaultColumns;
            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                DefaultColumns = ["Actions on Discrepancy (from AMs)", "ActionOnDiscrepancy", "Notes", "NotesFreeFill"];
            } else {
                DefaultColumns = ["Actions on Discrepancy (from AMs)", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
            }

            if (dataMap && dataMap?.length > 0 && !Array.isArray(dataMap[0])) {
                const exclusionjson = dataMap.map(item => {
                    const {
                        JobId,
                        CreatedOn,
                        UpdatedOn,
                        ConfidenceScore,
                        ...filteredItem
                    } = item;
                    return filteredItem;
                });

                let excludedColumns;
                if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                    excludedColumns = ["Id", "ActionOnDiscrepancy", "Notes", "NotesFreeFill"];
                } else {
                    excludedColumns = ["Id", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
                }

                let headers = Object.keys(exclusionjson[0]).filter(headerw => !excludedColumns.includes(headerw));
                headers = headers.filter(f => !["ActionOnDiscrepancy", "RequestEndorsement", "NotesFreeFill", "Notes"]?.includes(f));
                let headerRows1 = headers.map((item, index) => {
                    Exclusion_appDataConfig.exclusion.config.merge[`${0}_${index}`] = {
                        "r": 0,
                        "c": index,
                        "rs": 2,
                        "cs": 1
                    }
                    return {
                        r: 0,
                        rs: 2,
                        c: index,
                        cs: 1,
                        v: {
                            ct: { fa: "General", t: "g" },
                            m: item,
                            v: item,
                            fs: 10,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                });

                headerRows1 = [...headerRows1];

                const defaultHeaderRows1 = DefaultColumns.map((item, index) => {
                    if (index == 0) {
                        Exclusion_appDataConfig.exclusion.config.merge[`${headerRows1[headerRows1?.length - 1]?.r}_${headerRows1.length + index}`] = {
                            "r": headerRows1[headerRows1?.length - 1]?.r,
                            "c": reqEndorsementColHideBrokerId.includes(brokerId) ? 3 : 4,
                            "rs": 1,
                            "cs": reqEndorsementColHideBrokerId.includes(brokerId) ? 3 : 4,
                        }
                        return {
                            r: headerRows1[headerRows1?.length - 1]?.r,
                            rs: 1,
                            c: headerRows1.length,
                            cs: 1,
                            v: {
                                ht: 0,
                                ct: { fa: "General", t: "g" },
                                m: item,
                                v: item,
                                fs: 11,
                                ff: "\"Tahoma\"",
                                merge: null,
                                bg: "rgb(139,173,212)",
                                tb: '2',
                                w: 55,
                            }
                        }
                    }
                    else {
                        return {
                            r: headerRows1[headerRows1?.length - 1]?.r + 1,
                            rs: 1,
                            c: headerRows1.length + index - 1,
                            cs: 1,
                            v: {
                                ct: { fa: "General", t: "g" },
                                m: item,
                                v: item,
                                fs: 11,
                                ff: "\"Tahoma\"",
                                merge: null,
                                bg: "rgb(139,173,212)",
                                tb: '2',
                                w: 55,
                            }
                        }
                    }
                });

                let headerRows1Values = [];
                let rowIndex = defaultHeaderRows1[defaultHeaderRows1.length - 1]?.r + 1;
                let rowHeight = 60;
                let fs = 9;
                let dataToUpDate = exclusionjson.map((item, indexr) => {
                    let actionColumnKeys;
                    if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                        actionColumnKeys = ["ActionOnDiscrepancy", "Notes", "NotesFreeFill"];
                    } else {
                        actionColumnKeys = ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
                    }
                    // if(brokerId != "1003") { 
                    if (AODDBrokerIds.includes(brokerId)) {
                        headers = [...headers, ...actionColumnKeys];
                    } else {
                        headers = [...headers];
                    }
                    let uniqueHeaderSet = Array.from(new Set(headers))
                    let rowIndexForStateData = 0;
                    uniqueHeaderSet.map((key, rIndex) => {
                        // if(item[key] != null) {
                        rowIndexForStateData = indexr + 2;
                        let text = item[key] != null ? item[key].toString().split('~~') : [];
                        let ss = [];
                        if (text && text?.length > 0) {
                            text.map((e) => {
                                ss.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "#000000",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                });
                            }
                            );
                        }
                        if (AODDBrokerIds.includes(brokerId) && (key === 'ActionOnDiscrepancy' || key === 'RequestEndorsement' || key === 'Notes')) {
                            const dvData = item[key];
                            if (dvData == "" || dvData == undefined || dvData == null) {
                                ss.push({
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgba(171, 160, 160, 0.957)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "ht": "0",
                                    "v": "Click here"
                                });
                            }
                        }
                        headerRows1Values.push({
                            r: rowIndex + indexr,
                            c: rIndex,
                            v: {
                                ct: { fa: "General", t: "inlineStr", s: ss },
                                merge: null,
                                w: 55,
                                ff: "\"Tahoma\"",
                                tb: '2',
                                "ht": key === 'ActionOnDiscrepancy' ? "0" : key === 'RequestEndorsement' ? "0" :
                                    key === 'Notes' ? "0" : null
                            }
                        });

                        if (text && rowHeight < parseInt(item[key]?.length / 2 + 10)) {
                            rowHeight = parseInt(item[key]?.length / 2 + 10);
                            Exclusion_appDataConfig.exclusion.config.rowlen[`${rowIndex}`] = rowHeight;
                        }
                        // }
                    })
                    item["sheetPosition"] = rowIndexForStateData;
                    return item;
                });

                const exclusionTableForReturn = [{ "TableName": "ExTable 1", "data": dataToUpDate }];
                Exclusion_appDataConfig.exclusion.config.borderInfo.push({
                    "rangeType": "range",
                    "borderType": "border-all",
                    "color": "#000",
                    "style": "1",
                    "range": [
                        {
                            "left": 74,
                            "width": 300,
                            "top": 470,
                            "height": 42,
                            "left_move": 74,
                            "width_move": 4213,
                            "top_move": 471,
                            "height_move": 1107,
                            "row": [
                                headerRows1[0]?.r,
                                headerRows1Values[headerRows1Values?.length - 1]?.r
                            ],
                            "column": [
                                headerRows1[0]?.c,
                                // brokerId != "1003" ? 7 : 3,
                                reqEndorsementColHideBrokerId.includes(brokerId) ? 6 : AODDBrokerIds.includes(brokerId) ? 7 : 3,
                            ],
                            "row_focus": headerRows1[0]?.r,
                            "column_focus": headerRows1[0]?.c
                        }
                    ]
                });

                // if(brokerId != "1003") {
                if (AODDBrokerIds.includes(brokerId)) {
                    headerRows1 = [...headerRows1, ...defaultHeaderRows1, ...headerRows1Values];
                } else {
                    headerRows1 = [...headerRows1, ...headerRows1Values];
                }

                const allRows2 = [...headerRows1];
                if (headerRows1 && headerRows1?.length > 0) {
                    const ExTableColumnDetails = exTableColumnDetails;
                    const columnNames = Object.keys(exclusionjson[0]);
                    ExTableColumnDetails["ExTable 1"] = {
                        "columnNames": [columnNames],
                        "range": {
                            "start": 0,
                            "end": headerRows1[headerRows1.length - 1]?.r
                        }
                    };
                    setExTableColumnDetails(ExTableColumnDetails);
                }

                allRows2.sort((a, b) => a.r - b.r);
                basedata.push(...allRows2);
                Exclusion_appDataConfig.exclusion.celldata = basedata;

                let exclusionDataTracked = [];
                let positioningForexclusionData = [];

                exclusionDataTracked.push(exclusionTableForReturn);
                setDataForSaveExclusion([...exclusionDataTracked]);
                localStorage.setItem('exclusionDataTracked', JSON.stringify(exclusionDataTracked))

                positioningForexclusionData.push({
                    TableName: "ExTable 1",
                    columnNames: exTableColumnDetails["ExTable 1"]?.columnNames,
                    range: {
                        start: allRows2[0]?.r,
                        end: allRows2[allRows2?.length - 1]?.r
                    }
                });

                setDataForSaveExclusionPosition([...positioningForexclusionData]);
                localStorage.setItem('positioningForexclusionData', JSON.stringify(positioningForexclusionData))
            }
        };
        if (canRenderExclusionData) {
            csrexclusionTable();
        }

        const qacRender = () => {
            const basedata = [];
            const cellData = [];
            const apiData = canRenderQac ? typeof canRenderQac === 'string' ? JSON?.parse(canRenderQac) : {} : {};
            let finalQacData = apiData;
            let rowIndex = 0;
            let lobIndex = 0;
            const tableRangeState = {};

            Object.entries(apiData)?.forEach(([category, subCategories], clrIdx) => {
                tableRangeState[category] = [];

                cellData.push({
                    r: rowIndex,
                    c: 0,
                    v: {
                        ct: { fa: "@", t: "inlineStr" },
                        m: category,
                        v: category,
                        merge: null,
                        tb: "2",
                        fs: "10",
                        ff: "\"Tahoma\"",
                        bl: 1,
                    },
                });

                rowIndex += 1;

                Object.entries(subCategories)?.forEach(([lob, questions]) => {
                    const LobIndex = lobIndex;
                    const lobStartIndex = rowIndex;
                    const lobDataLength = questions?.length;

                    qac_appDataConfig.config.merge[`${lobStartIndex}_${0}`] = {
                        "r": lobStartIndex,
                        "c": 0,
                        "rs": 1,
                        "cs": 3
                    };

                    cellData.push({
                        r: rowIndex,
                        c: 0,
                        v: {
                            ct: { fa: "@", t: "inlineStr" },
                            m: lob,
                            v: lob,
                            merge: null,
                            tb: "2",
                            fs: 7,
                            ff: "\"Tahoma\"",
                            bl: 1,
                            bg: category?.toLowerCase()?.includes('high') ? "rgb(241, 169, 131)" : "rgb(139,173,212)",
                        },
                    });

                    rowIndex += 1;

                    const headers = Object?.keys(questions[0]);
                    headers?.forEach((header, index) => {
                        cellData.push({
                            r: rowIndex,
                            c: index,
                            v: {
                                ct: { fa: "@", t: "inlineStr" },
                                m: header,
                                v: header,
                                merge: null,
                                tb: "2",
                                fs: 7,
                                ff: "\"Tahoma\"",
                                bl: 1,
                                bg: category?.toLowerCase()?.includes('high') ? "rgb(241, 169, 131)" : "rgb(139,173,212)",
                            },
                        });
                    });
                    rowIndex += 1;

                    questions?.forEach((question, dIndex) => {
                        headers?.forEach((header, index) => {
                            const questionValue = question[header] !== null ? question[header].replace(/~~/g, "\n") : "";
                            cellData.push({
                                r: rowIndex,
                                c: index,
                                v: {
                                    ct: { fa: "@", t: "inlineStr" },
                                    m: questionValue,
                                    v: questionValue,
                                    merge: null,
                                    tb: "2",
                                    fs: 7,
                                    ff: "\"Tahoma\"",
                                },
                            });
                        });
                        finalQacData[category][lob][dIndex]["RowIndex"] = rowIndex;
                        rowIndex += 1;
                    });

                    if (!tableRangeState[category]) {
                        tableRangeState[category] = [];
                    }

                    tableRangeState[category].push({
                        [`qacTable ${LobIndex}`]: {
                            columnNames: headers.reduce((acc, header, index) => {
                                acc[header] = index;
                                return acc;
                            }, {}),
                            range: {
                                start: lobStartIndex,
                                end: rowIndex - 1,
                            },
                            policyLob: [lob],
                        }

                    });


                    if (lobDataLength == ((questions?.length - 1) + 1)) {
                        qac_appDataConfig["config"]["borderInfo"][LobIndex] = {
                            "rangeType": "range",
                            "borderType": "border-all",
                            "color": "#000",
                            "style": "1",
                            "range": [
                                {
                                    "left": 0,
                                    "width": 830,
                                    "top": 0,
                                    "height": 19,
                                    "left_move": 0,
                                    "width_move": 830,
                                    "top_move": 0,
                                    "height_move": 199,
                                    "row": [lobStartIndex, (lobStartIndex + lobDataLength) + 1],
                                    "column": [0, headers?.length - 1],
                                    "row_focus": 0,
                                    "column_focus": 0
                                }
                            ]
                        }
                        rowIndex += 2;
                        lobIndex += 1;
                    }
                });
            });
            const allRows2 = [...cellData]
            allRows2.sort((a, b) => a.r - b.r);
            basedata.push(...allRows2);
            qac_appDataConfig.celldata = basedata;
            setSheetDataRowIndex((prev) => {
                const updatedData = { ...prev, ...finalQacData };
                sheetDataRowIndexRef.current = updatedData; // Store the latest data in ref
                return updatedData;
            });
            sessionStorage.setItem("qacTblRange", JSON.stringify(tableRangeState))
            setQacTableColumnDetails(tableRangeState);
        };
        if (canRenderQac && renderBrokerIdsQac?.includes(brokerId)) {
            qacRender();
        }

        csrLuckySheet();

        setTimeout(() => {
            dataGrouping();
        }, 2000);

        let interval;
        if (activeUserName === sessionUserName) {
            setTimeout(() => {
                const { UpdatePeriod } = autoupdate();
                interval = setInterval(() => {
                    const { UpdateEnable } = autoupdate();
                    if (UpdateEnable && issavessheet === false) {
                        Autoupdateclick(true);
                    }
                }, UpdatePeriod);
            }, 2000);
        }

        return () => {
            clearInterval(interval)
            if (activeUserName === sessionUserName) {
                const allowautoup = sessionStorage.getItem("IsAutoUpdate")
                if (autoprogress && issavessheet == false && (allowautoup == "true" || allowautoup == true)) {
                    Autoupdateclick(true);
                }
                if (issavessheet == false && (allowautoup == "true" || allowautoup == true)) {
                    Autoupdateclick(true);
                }
            }
        };
    }, [props.data, activeUserName, sessionUserName]);

    const getCsrReviewData = async (jobId) => {
        try {
            let token = sessionStorage.getItem("token");
            token = await processAndUpdateToken(token);
            const response = await axios.get(
                baseUrl + `/api/JobConfiguration/GetJobReviewData?jobId=${jobId}`,
                {
                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json",
                    },
                }
            );
            if (response?.data?.length > 0) {
                setReviewData(response?.data);
                setSendPolicyInsuredData(response?.data);
            } else {
                setReviewData([]);
                setSendPolicyInsuredData([]);
            }
        } catch (error) {
            setReviewData([]);
            setSendPolicyInsuredData([]);
        }
    }

    const setCellValue = (row, column, data) => {
        if (luckysheet && (luckysheet != undefined || luckysheet != null)) {
            luckysheet?.setcellvalue(row, column, luckysheet?.flowdata(), data);
            luckysheet?.jfrefreshgrid();
        }
    }

    const setCsrCellValue = (row, column, data) => {
        if (luckysheet && (luckysheet != undefined || luckysheet != null)) {
            luckysheet?.setcellvalue(row, column, luckysheet.flowdata(), data);
            luckysheet?.jfrefreshgrid();
            luckysheet?.setCellFormat(row, column, "ct", { fa: "General", t: "g" })
        }
    }

    const toggleDropDialog = () => {
        let flagCheck = luckysheet.getSheet().name;
        if (flagCheck != 'Red' && flagCheck != 'Green') {
            setDropDialog(!dropDialog);
            const propsData = luckysheet.getSheetData();
            setsheetState(propsData);
        }
        else {
            setGradiationDialog(!gradiationDialog);
        }
    };


    const tbl1HyperFun = (range, flagCheck) => {
        if (range) {
            const rowIdx = range[0].row[0];
            const columnIdx = range[0].column[0];

            if (flagCheck === "PolicyReviewChecklist") {
                let checklistData = [...props?.data];
                if (checklistData && checklistData?.length > 0) {
                    checklistData = checklistData.map((e) => {
                        if (e?.TemplateData && typeof e?.TemplateData != 'object' && typeof e?.TemplateData === 'string') {
                            let templateData = JSON.parse(e.TemplateData);
                            e["TemplateData"] = templateData;
                        }
                        return e;
                    })
                }

                const tableDetails = tableColumnDetails;
                let keys = Object.keys(tableDetails);
                if (keys && keys?.length > 0 && checklistData && checklistData.length > 0 && rowIdx > 0 && columnIdx > 0) {
                    let tableName = '';
                    keys.forEach((f) => {
                        const tableNameData = tableDetails[f];
                        const tblRangeData = tableNameData?.range;
                        if (tableNameData && tblRangeData && tblRangeData?.start <= rowIdx && tblRangeData?.end >= rowIdx) {
                            tableName = f;
                        }
                    });
                    if (tableName && tableName == 'Table 1') {
                        const filterTbl1 = tableDetails[tableName];
                        const hyperLinkColumn = filterTbl1?.columnNames.indexOf("Exdion CSRDigiT");
                        let tbl1Data = {};
                        const filteredTbl1Data = checklistData.filter((f) => f?.Tablename === tableName);
                        if (filteredTbl1Data && filteredTbl1Data?.length > 0) {
                            tbl1Data = filteredTbl1Data[0];
                            const selectedRecordData = tbl1Data?.TemplateData[hyperLinkColumn];
                            if (selectedRecordData != undefined) {
                                if (selectedRecordData['Headers'] == 'Exdion CSRDigiT' && columnIdx === 2 && filterTbl1?.range?.end === rowIdx) {
                                    const tbl1HyperLink = selectedRecordData['(No column name)'];
                                    if (tbl1HyperLink) {
                                        window.open(tbl1HyperLink, '_blank', 'noopener');
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }


    const XRayRoute = (range, flag) => {
        if (range) {
            const row = range[0].row[0];
            const column = range[0].column[0];

            if (flag === "PolicyReviewChecklist") {

                let checklistData = [...props?.data];
                if (checklistData && checklistData?.length > 0) {
                    checklistData = checklistData.map((e) => {
                        if (e?.TemplateData && typeof e?.TemplateData != 'object' && typeof e?.TemplateData === 'string') {
                            let templateData = JSON.parse(e.TemplateData);
                            e["TemplateData"] = templateData;
                        }
                        return e;
                    })
                }
                const tableDetails = tableColumnDetails;
                let keys = Object.keys(tableDetails);
                if (keys && keys?.length > 0 && checklistData && checklistData.length > 0 && row > 0 && column > 0) {
                    let tableName = '';
                    keys.forEach((f) => {
                        const tableCData = tableDetails[f];
                        const tRangeData = tableCData?.range;
                        if (tableCData && tRangeData && tRangeData?.start <= row && tRangeData?.end >= row) {
                            tableName = f;
                        }
                    });
                    if (tableName && tableName != 'Table 1') {
                        const activeTableData = tableDetails[tableName];
                        const documentviewerIndex = activeTableData?.columnNames["Document Viewer"];
                        if (documentviewerIndex && documentviewerIndex === column) {
                            let activeTableDBData = {};
                            const filteredData = checklistData.filter((f) => f?.Tablename === tableName);
                            if (filteredData && filteredData?.length > 0) {
                                activeTableDBData = filteredData[0];
                                const activeTableRange = activeTableData?.range;
                                const configureStartFrom = (activeTableRange?.end - activeTableRange.start) - activeTableDBData?.TemplateData?.length;
                                const startT = activeTableRange?.start + configureStartFrom + 1;
                                const endT = activeTableRange?.end;
                                const indexSet = Array.from({ length: endT - startT + 1 }, (_, i) => startT + i);
                                // console.log(indexSet);
                                if (indexSet && indexSet?.length > 0) {
                                    const recordIndexToGet = indexSet.indexOf(row);
                                    // console.log("recordIndexToGet : ", recordIndexToGet);
                                    if (recordIndexToGet >= 0) {
                                        const selectedRecordData = activeTableDBData?.TemplateData[recordIndexToGet];
                                        const XRrayURL = selectedRecordData["Document Viewer"];
                                        if (XRrayURL) {
                                            window.open(XRrayURL, '_blank', 'noopener');
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
            } else if (flag === "Red" || flag === "Green") {
                let gradtionSheetData = flag == "Red" ? sessionStorage.getItem('redSheetData') : sessionStorage.getItem('greenSheetData');
                let data = JSON.parse(gradtionSheetData);
                let filterData = data.filter(item => Object.keys(item.data).length !== 0);
                filterData.forEach((item, index) => {
                    item.TableName = `Table ${index + 2}`;
                });
                let gradtionData = [...filterData];
                let tblRangeData = flag == "Red" ? sessionStorage.getItem('redTableRangeData') : sessionStorage.getItem('greenTableRangeData');
                let parsedSet = JSON.parse(tblRangeData);
                const tableDetails = parsedSet;
                let keys = Object.keys(tableDetails);

                if (keys && keys?.length > 0 && gradtionData && gradtionData.length > 0 && row > 0 && column > 0) {
                    let tableName = '';
                    keys.forEach((f) => {
                        const tableCData = tableDetails[f];
                        const tRangeData = tableCData?.range;
                        if (tableCData && tRangeData && tRangeData?.start <= row && tRangeData?.end >= row) {
                            tableName = f;
                        }
                    });
                    if (tableName && tableName != 'Table 1') {
                        const activeTableData = tableDetails[tableName];
                        const documentviewerIndex = activeTableData?.columnNames["Document Viewer"];
                        if (documentviewerIndex && documentviewerIndex === column) {
                            let activeTableDBData = {};
                            const filteredData = gradtionData.filter((f) => f?.TableName === tableName);
                            if (filteredData && filteredData?.length > 0) {
                                activeTableDBData = filteredData[0];
                                const activeTableRange = activeTableData?.range;
                                const configureStartFrom = (activeTableRange?.end - activeTableRange.start) - activeTableDBData?.data?.length;
                                const startT = activeTableRange?.start + configureStartFrom + 1;
                                const endT = activeTableRange?.end;
                                const indexSet = Array.from({ length: endT - startT + 1 }, (_, i) => startT + i);
                                // console.log(indexSet);
                                if (indexSet && indexSet?.length > 0) {
                                    const recordIndexToGet = indexSet.indexOf(row);
                                    // console.log("recordIndexToGet : ", recordIndexToGet);
                                    if (recordIndexToGet >= 0) {
                                        const selectedRecordData = activeTableDBData?.data[recordIndexToGet];
                                        const XRrayURL = selectedRecordData["Document Viewer"];
                                        if (XRrayURL) {
                                            window.open(XRrayURL, '_blank', 'noopener');
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
            } else if (flag === "Forms Compare") {
                const formData = [...props?.formCompareData];
                const formSectionDetails = formTableColumnDetails;
                const unMatchedSectionDetails = formSectionDetails["FormTable 2"];
                if (unMatchedSectionDetails && unMatchedSectionDetails?.range && unMatchedSectionDetails?.columnNames) {
                    const rangeData = unMatchedSectionDetails?.range;
                    const columnDetails = unMatchedSectionDetails?.columnNames;
                    const documentviewerPosition = columnDetails["Document Viewer"];
                    const start = rangeData?.start;
                    const end = rangeData?.end;
                    const modifiedStart = start + 3;
                    const indexSet = Array.from({ length: end - modifiedStart + 1 }, (_, i) => modifiedStart + i);
                    if (indexSet && indexSet?.length > 0 && start && end && documentviewerPosition === column && start <= row && end >= row && formData && formData?.length > 0) {
                        const unmatchedFormSectionData = formData[1];
                        if (unmatchedFormSectionData && unmatchedFormSectionData?.TemplateData) {
                            const unMatchedSectionParsedData = typeof unmatchedFormSectionData?.TemplateData === "string" ? JSON.parse(unmatchedFormSectionData?.TemplateData) : unmatchedFormSectionData?.TemplateData;
                            if (Array.isArray(unMatchedSectionParsedData) && unMatchedSectionParsedData?.length > 0) {
                                const recordIndexToGet = indexSet.indexOf(row);
                                // console.log("recordIndexToGet : ", recordIndexToGet);
                                if (recordIndexToGet >= 0) {
                                    const selectedRecordData = unMatchedSectionParsedData[recordIndexToGet];
                                    const XRrayURL = selectedRecordData["Document Viewer"];
                                    if (XRrayURL) {
                                        luckysheet.exitEditMode();
                                        setTimeout(() => {
                                            luckysheet.setRangeShow({ row: [row, row], column: [column - 1, column - 1] });
                                            setTimeout(() => {
                                                luckysheet.enterEditMode();
                                                setTimeout(() => {
                                                    window.open(XRrayURL, '_blank', 'noopener');
                                                }, 50);
                                            }, 100);
                                        }, 100);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    const funForDiscrepancyCol = (e) => {
        let range = luckysheet.getRange();
        let flagCheck = luckysheet.getSheet().name;
        if (flagCheck == "PolicyReviewChecklist" || flagCheck == "Forms Compare" || flagCheck == 'Exclusion') {
            let selectedIndex = range[0].row[0];
            let tabledata = flagCheck == "PolicyReviewChecklist" ? tableColumnDetails : flagCheck == "Forms Compare" ? formTableColumnDetails : exTableColumnDetails;
            const excludedColumns = ["Id", "sheetPosition", "columnid"];
            const selectedTable = findTableForIndex(selectedIndex, tabledata, excludedColumns);

            let actionColumnTable = flagCheck == "PolicyReviewChecklist" ? tableColumnDetails[selectedTable] : flagCheck == "Forms Compare" ? formTableColumnDetails[selectedTable] : exTableColumnDetails[selectedTable];
            let values = Object.values(actionColumnTable.columnNames);
            let largestIndex = selectedTable == "ExTable 1" ? Math.max(...Object.keys(values[0])) : Math.max(...values);
            let Actioncolumnindex = selectedTable == "ExTable 1" ? largestIndex - 5 : reqEndorsementColHideBrokerId.includes(brokerId) ? largestIndex - 2 : largestIndex - 3;
            let Requestcolumnindex = selectedTable == "ExTable 1" ? largestIndex - 4 : reqEndorsementColHideBrokerId.includes(brokerId) ? largestIndex - 1 : largestIndex - 2;
            let Notescolumnindex = selectedTable == "ExTable 1" ? (reqEndorsementColHideBrokerId.includes(brokerId) ? (largestIndex - 4) : (largestIndex - 3)) : largestIndex - 1;
            let r = range[0].row[0];
            let c = range[0].column[0];

            if (e?.hasData) {
                if (e.selectedOption1 != null && e.selectedOption2 != null && e.selectedOption3 != null) {
                    if (c == Actioncolumnindex) {
                        setCellValue(r, c, e.selectedOption1.text);
                        setCellValue(r, c + 1, e.selectedOption2.text);
                        setCellValue(r, c + 2, e.selectedOption3.text);
                    } else if (c == Requestcolumnindex) {
                        setCellValue(r, c - 1, e.selectedOption1.text);
                        setCellValue(r, c, e.selectedOption2.text);
                        setCellValue(r, c + 1, e.selectedOption3.text);
                    } else if (c == Notescolumnindex) {
                        setCellValue(r, c - 2, e.selectedOption1.text);
                        setCellValue(r, c - 1, e.selectedOption2.text);
                        setCellValue(r, c, e.selectedOption3.text);
                    }
                } else {
                    // Handle each option separately if not all are selected
                    if (e && e.selectedOption1 && e.selectedOption1 != null) {
                        let actionText = e.selectedOption1.text;
                        setCellValue(r, Actioncolumnindex, actionText);
                        luckysheet.exitEditMode();
                    }
                    if (e && e.selectedOption2 && e.selectedOption2 != null) {
                        let requestText = e.selectedOption2.text;
                        setCellValue(r, Requestcolumnindex, requestText);
                        luckysheet.exitEditMode();
                    }
                    if (e && e.selectedOption3 && e.selectedOption3 != null) {
                        let notesText = e.selectedOption3.text;
                        setCellValue(r, Notescolumnindex, notesText);
                        luckysheet.exitEditMode();
                    }
                }
            }
            setDropDialog(false);
        } else if (flagCheck == 'Red' || flagCheck == 'Green') {
            let r = range[0].row[0];
            let c = range[0].column[0];
            const gradiationSet = flagCheck == 'Red' ? sessionStorage.getItem('redTableRangeData') : flagCheck == 'Green' ? sessionStorage.getItem('greenTableRangeData') : null;
            let parsedSet = JSON.parse(gradiationSet);
            const excludedColumns = ["Actions on Discrepancy (from AMs)"];
            const selectedTable = findTableForIndex(r, parsedSet, excludedColumns);
            let actionColumnTable = parsedSet[selectedTable];
            let values = actionColumnTable ? Object.keys(actionColumnTable.columnNames) : [];
            let largestIndex = values.indexOf("Document Viewer") + 1;
            let Actioncolumnindex = largestIndex + 1;
            let Requestcolumnindex = largestIndex + 2;
            let Notescolumnindex = largestIndex + 3;

            if (e?.hasData) {
                if (e.selectedOption1 != null && e.selectedOption2 != null && e.selectedOption3 != null) {
                    if (c == Actioncolumnindex) {
                        setCsrCellValue(r, c, e.selectedOption1.text);
                        setCsrCellValue(r, c + 1, e.selectedOption2.text);
                        setCsrCellValue(r, c + 2, e.selectedOption3.text);
                    } else if (c == Requestcolumnindex) {
                        setCsrCellValue(r, c - 1, e.selectedOption1.text);
                        setCsrCellValue(r, c, e.selectedOption2.text);
                        setCsrCellValue(r, c + 1, e.selectedOption3.text);
                    } else if (c == Notescolumnindex) {
                        setCsrCellValue(r, c - 2, e.selectedOption1.text);
                        setCsrCellValue(r, c - 1, e.selectedOption2.text);
                        setCsrCellValue(r, c, e.selectedOption3.text);
                    }
                } else {
                    // Handle each option separately if not all are selected
                    if (e && e.selectedOption1 && e.selectedOption1 != null) {
                        let actionText = e.selectedOption1.text;
                        setCsrCellValue(r, Actioncolumnindex, actionText);
                        luckysheet.exitEditMode();
                    }
                    if (e && e.selectedOption2 && e.selectedOption2 != null) {
                        let requestText = e.selectedOption2.text;
                        setCsrCellValue(r, Requestcolumnindex, requestText);
                        luckysheet.exitEditMode();
                    }
                    if (e && e.selectedOption3 && e.selectedOption3 != null) {
                        let notesText = e.selectedOption3.text;
                        setCsrCellValue(r, Notescolumnindex, notesText);
                        luckysheet.exitEditMode();
                    }
                }
            }
            setGradiationDialog(false);
        }
    };
    // luckysheet.exitEditMode();

    let isLuckysheetRendered = false;
    const csrLuckySheet = async () => {

        let garadationDataForRedSheet = {};
        let garadationDataForGreenSheet = {};
        // const qacDataSet = await getQACData(jobId, token);
        if (props?.gradtionDataSet && canRenderGradation) {
            const grData = await gradiationConverter(props?.gradtionDataSet, '', true);
            if (grData?.response1?.celldata?.length > 0) {
                garadationDataForRedSheet = grData?.response1;
            }
            if (grData?.response2?.celldata?.length > 0) {
                garadationDataForGreenSheet = grData?.response2;
            }

        }
        let dataSet = [];
        if (canRenderChecklist) {
            dataSet.push(Policy_appDataConfig.demo);
        }
        if (canRenderGradation && garadationDataForRedSheet && garadationDataForRedSheet != null && garadationDataForRedSheet != undefined && Object.keys(garadationDataForRedSheet).length > 0) {
            dataSet = [...dataSet, garadationDataForRedSheet]
        }
        if (canRenderGradation && garadationDataForGreenSheet && garadationDataForGreenSheet != null && garadationDataForGreenSheet != undefined && Object.keys(garadationDataForGreenSheet).length > 0) {
            dataSet = [...dataSet, garadationDataForGreenSheet]
        }
        if (canRenderFormData) {
            dataSet.push(FormCompare_appconfigdata.forms);
        }
        if (canRenderExclusionData) {
            dataSet.push(Exclusion_appDataConfig.exclusion);
        }
        if (canRenderQac && renderBrokerIdsQac?.includes(brokerId)) {
            dataSet.push(qac_appDataConfig);
        }
        if (!isLuckysheetRendered && luckysheet) {
            isLuckysheetRendered = true;
            if (luckysheet) {
                const options = {
                    container: "luckysheet2", // Container ID
                    showinfobar: false,
                    showsheetbar: true,
                    lang: 'en',
                    data: dataSet,
                    enableAddRow: true,
                    showtoolbar: true,
                    row: 2,
                    column: 3,
                    allowUpdate: true,
                    enableAddBackTop: true,
                    sheetRightClickConfig: {
                        delete: false,
                        copy: false,
                        rename: false,
                        color: false,
                        hide: false,
                        move: false,
                    },
                    showsheetbarConfig: {
                        add: false,
                        menu: false,
                    },

                    hook: {
                        workbookCreateAfter(json) {
                            luckysheet.setSheetZoom(1);
                        },
                        cellEditBefore(range) {
                            let sheetflagCheck = luckysheet?.getSheet()?.name;

                            //code for variance columns options
                            if (range && range.length > 0 && range != undefined) {
                                if (luckysheet?.getSheet()?.name === 'PolicyReviewChecklist') {
                                    if (range && range[0]?.row && range[0]?.column && range[0]?.row[0] === range[0]?.row[1] &&
                                        range[0]?.column[0] === range[0]?.column[1] && [3, 4, 5, 6].includes(range[0].row[0]) &&
                                        range[0].column[0] === 4) {
                                        setTimeout(() => {
                                            luckysheet.exitEditMode();
                                            matchedOrUnMatchedFilter(range[0]?.row[0]);
                                            container.current.showSnackbar(range[0]?.row[0] === 4 ?
                                                "Matched Records Filtered" : range[0]?.row[0] === 5 ? "Variances Records Filtered" :
                                                    range[0]?.row[0] === 6 ? "Details not available Questions filtered" : "Filter Removed", "info", true);
                                        }, 100);


                                        // checkbox setvalue for Variances columns
                                        if (range && range.length > 0 && range != undefined) {
                                            handleCellSelection(range, luckysheet?.getSheet()?.name);
                                        }
                                        return;
                                    } 
                                    if (range && range[0]?.row && range[0]?.column && range[0]?.row[0] === range[0]?.row[1] &&
                                        range[0]?.column[0] === range[0]?.column[1] && range[0].column[0] === 0) {
                                        if (luckysheet?.getSheet()?.name != 'QAC not answered questions' && luckysheet?.getSheet()?.name != "Exclusion") {
                                            setTimeout(() => {
                                                luckysheet.exitEditMode();
                                                dataGrouping(range[0]?.row[0]);
                                            }, 100);
                                        }
                                        return;
                                    }
                                }
                            }

                            // Denied the action of selection of all the cells
                            if (range && range.length > 0 && range != undefined) {
                                let selectedRowIndex = range[0].row[0];
                                let selectedColumnIndex = range[0].column[0];
                                const redTableData = sessionStorage.getItem('redTableRangeData');
                                const greenTableData = sessionStorage.getItem('greenTableRangeData');
                                const parsedRedTableData = JSON.parse(redTableData);
                                const parsedGreenTableData = JSON.parse(greenTableData);

                                if (luckysheet?.getSheet()?.name != 'QAC not answered questions') {
                                    const tabledata = luckysheet?.getSheet()?.name == 'PolicyReviewChecklist' ? tableColumnDetails : luckysheet?.getSheet()?.name == 'Forms Compare' ? formTableColumnDetails : luckysheet?.getSheet()?.name == 'Exclusion' ? exTableColumnDetails : luckysheet?.getSheet()?.name == 'Red' ? parsedRedTableData : parsedGreenTableData;
                                    const excludedColumns = ["columnid"];
                                    const selectedTable = findTableForIndex(selectedRowIndex, tabledata, excludedColumns);
                                    const tblSelectedRow = tabledata[selectedTable];
                                    if (selectedTable != "ExTable 1" && tblSelectedRow != undefined) {
                                        for (const key in tblSelectedRow.columnNames) {
                                            if (tblSelectedRow.columnNames[key] === 0) {
                                                delete tblSelectedRow.columnNames[key];
                                            }
                                        }

                                        let values = selectedTable == "ExTable 1" ? Object.values(tblSelectedRow.columnNames[0]) : Object.keys(tblSelectedRow.columnNames);
                                        let docViewerIndex = selectedTable == "ExTable 1" ? values.indexOf("PageNumber") + 1 : values.indexOf("Document Viewer") + 1;

                                        if (selectedColumnIndex < docViewerIndex) {
                                            setTimeout(() => {
                                                luckysheet.exitEditMode();
                                            }, 100);
                                            return false;
                                        }
                                    }
                                } else if (luckysheet?.getSheet()?.name == 'QAC not answered questions') {
                                    const sessionQacData = sessionStorage.getItem("qacTblRange");
                                    const tabledata = JSON?.parse(sessionQacData);
                                    const keys = Object.keys(tabledata);
                
                                    if (keys?.length > 0) {
                                        keys?.forEach((key) => {
                                            const tableData = tabledata[key];
                                            const structuredTblObj = qacTblRangeStructureFun(tableData);
                                            const selectedTable = findTableForIndex(selectedRowIndex, structuredTblObj, "");
                                            const tblSelectedRow = structuredTblObj[selectedTable];
                                            if (tblSelectedRow != undefined) {
                                                const selectedValue = luckysheet?.getCellValue(selectedRowIndex ,selectedColumnIndex);
                                                const lobRow = tblSelectedRow?.policyLob[0];
                                                const headerArray = Object.keys(tblSelectedRow?.columnNames);
                                                const coverageQueCol = tblSelectedRow?.range;
                                                if(selectedValue === lobRow) {
                                                    setTimeout(() => {
                                                        luckysheet.exitEditMode();
                                                    }, 100);
                                                    return false;
                                                } else if(headerArray?.includes(selectedValue)) {
                                                    setTimeout(() => {
                                                        luckysheet.exitEditMode();
                                                    }, 100);
                                                    return false;
                                                } else if((selectedRowIndex <= coverageQueCol?.end && coverageQueCol?.start + 2 <= selectedRowIndex) && selectedColumnIndex === 0) {
                                                    setTimeout(() => {
                                                        luckysheet.exitEditMode();
                                                    }, 100);
                                                    return false;
                                                }
                                            }
                                        })
                                    }
                                }
                            }

                            // code for trigger X-Ray
                            if ((luckysheet?.getSheet()?.name === "PolicyReviewChecklist" || luckysheet?.getSheet()?.name === "Forms Compare" || luckysheet?.getSheet()?.name === "Red" || luckysheet?.getSheet()?.name === "Green") && range[0].row[0] === range[0].row[1] && range[0].column[0] === range[0].column[1]) {
                                XRayRoute(range, luckysheet?.getSheet()?.name);
                            }

                            //table1 hyperlink 
                            if ((luckysheet?.getSheet()?.name === "PolicyReviewChecklist") && range[0].row[0] === range[0].row[1] && range[0].column[0] === range[0].column[1]) {
                                tbl1HyperFun(range, luckysheet?.getSheet()?.name);
                            }

                            // code to pop-Up the endorsement dialog
                            if (range && range.length > 0 && range != undefined) {
                                if (luckysheet?.getSheet()?.name != 'QAC not answered questions' && luckysheet?.getSheet()?.name != 'Red' && luckysheet?.getSheet()?.name != 'Green') {
                                    let selectedRowIndex = range[0].row[0];
                                    let nullcolumncheck = luckysheet.getSheetData()[selectedRowIndex];
                                    const isAllNull = nullcolumncheck.every(element => element === null);
                                    if (!isAllNull) {
                                        let selectedRowIndex = range[0].row[0];
                                        let tabledata = luckysheet?.getSheet()?.name === "PolicyReviewChecklist" ? tableColumnDetails : luckysheet?.getSheet()?.name === "Forms Compare" ? formTableColumnDetails : exTableColumnDetails;
                                        const excludedColumns = ["Id", "sheetPosition", "columnid"];
                                        const selectedTable = findTableForIndex(selectedRowIndex, tabledata, excludedColumns);

                                        let ranges = luckysheet.getRange();
                                        let selectedcolumnindex = ranges[0].column[0];
                                        let actionColumnTable = luckysheet?.getSheet()?.name === "PolicyReviewChecklist" ? tableColumnDetails[selectedTable] : luckysheet?.getSheet()?.name === "Forms Compare" ? formTableColumnDetails[selectedTable] : exTableColumnDetails[selectedTable];
                                        let values = actionColumnTable ? Object.values(actionColumnTable.columnNames) : [];
                                        let largestIndex = selectedTable == "ExTable 1" ? Math.max(...Object.keys(values[0])) : Math.max(...values);
                                        let Actioncolumnindex = selectedTable == "ExTable 1" ? (largestIndex - 5) : (reqEndorsementColHideBrokerId.includes(brokerId) ? (largestIndex - 2) : (largestIndex - 3));
                                        let Requestcolumnindex = selectedTable == "ExTable 1" ? (largestIndex - 4) : (reqEndorsementColHideBrokerId.includes(brokerId) ? (largestIndex - 1) : (largestIndex - 2));
                                        let Notescolumnindex = selectedTable == "ExTable 1" ? (reqEndorsementColHideBrokerId.includes(brokerId) ? (largestIndex - 4) : (largestIndex - 3)) : largestIndex - 1;
                                        if (actionColumnTable !== undefined) {
                                            if (luckysheet?.getSheet()?.name === "PolicyReviewChecklist") {
                                                if (selectedTable != 'Table 3' && selectedRowIndex >= actionColumnTable.range.start + 2 || selectedTable == 'Table 3' && selectedRowIndex >= actionColumnTable.range.start + 3) {
                                                    if ((Actioncolumnindex == selectedcolumnindex || Requestcolumnindex == selectedcolumnindex || Notescolumnindex == selectedcolumnindex)) {
                                                        toggleDropDialog()
                                                        return false;
                                                    }
                                                }
                                                // } else if ( brokerId != "1003" && luckysheet?.getSheet()?.name === "Exclusion"){
                                            } else if (AODDBrokerIds.includes(brokerId) && luckysheet?.getSheet()?.name === "Exclusion") {
                                                if (selectedTable == "ExTable 1" && selectedRowIndex >= actionColumnTable.range.start + 2) {
                                                    if ((Actioncolumnindex == selectedcolumnindex || Requestcolumnindex == selectedcolumnindex || Notescolumnindex == selectedcolumnindex)) {
                                                        toggleDropDialog()
                                                        return false;
                                                    }
                                                }
                                                // } else if ( brokerId != "1003" && luckysheet?.getSheet()?.name === "Forms Compare"){
                                            } else if (AODDBrokerIds.includes(brokerId) && luckysheet?.getSheet()?.name === "Forms Compare") {
                                                if ((Actioncolumnindex == selectedcolumnindex || Requestcolumnindex == selectedcolumnindex || Notescolumnindex == selectedcolumnindex)) {
                                                    toggleDropDialog()
                                                    return false;
                                                }
                                            }
                                        } else {
                                            return false;
                                        }
                                    }
                                } else if (luckysheet?.getSheet()?.name == 'Red' || luckysheet?.getSheet()?.name == 'Green') {
                                    let selectedRowIndex = range[0].row[0];
                                    let selectedcolumnindex = range[0].column[0];
                                    let nullcolumncheck = luckysheet.getSheetData()[selectedRowIndex];
                                    const isAllNull = nullcolumncheck.every(element => element === null);
                                    if (!isAllNull) {
                                        if (luckysheet?.getSheet()?.name == 'Red') {
                                            const gradiationSet = sessionStorage.getItem('redTableRangeData');
                                            let parsedSet = JSON.parse(gradiationSet);
                                            const excludedColumns = ["Actions on Discrepancy (from AMs)"];
                                            const selectedTable = findTableForIndex(selectedRowIndex, parsedSet, excludedColumns);
                                            let actionColumnTable = parsedSet[selectedTable];
                                            let values = actionColumnTable ? Object.keys(actionColumnTable.columnNames) : [];
                                            let largestIndex = values.indexOf("Document Viewer") + 1;
                                            let Actioncolumnindex = largestIndex + 1;
                                            let Requestcolumnindex = largestIndex + 2;
                                            let Notescolumnindex = largestIndex + 3;
                                            if (actionColumnTable !== undefined) {
                                                if (Actioncolumnindex != undefined && Requestcolumnindex != undefined && Notescolumnindex != undefined && nullcolumncheck[Actioncolumnindex].bg != 'rgb(139,173,212)') {
                                                    if ((Actioncolumnindex == selectedcolumnindex || Requestcolumnindex == selectedcolumnindex || Notescolumnindex == selectedcolumnindex)) {
                                                        toggleDropDialog();
                                                        return false;
                                                    }
                                                }
                                            }
                                        }
                                        if (luckysheet?.getSheet()?.name == 'Green') {
                                            const gradiationSet = sessionStorage.getItem('greenTableRangeData');
                                            let parsedSet = JSON.parse(gradiationSet);
                                            const excludedColumns = ["Actions on Discrepancy (from AMs)"];
                                            const selectedTable = findTableForIndex(selectedRowIndex, parsedSet, excludedColumns);
                                            let actionColumnTable = parsedSet[selectedTable];
                                            let values = actionColumnTable ? Object.keys(actionColumnTable.columnNames) : [];
                                            let largestIndex = values.indexOf("Document Viewer") + 1;
                                            let Actioncolumnindex = largestIndex + 1;
                                            let Requestcolumnindex = largestIndex + 2;
                                            let Notescolumnindex = largestIndex + 3;
                                            if (actionColumnTable !== undefined) {
                                                if (Actioncolumnindex != undefined && Requestcolumnindex != undefined && Notescolumnindex != undefined && nullcolumncheck[Actioncolumnindex].bg != 'rgb(139,173,212)') {
                                                    if ((Actioncolumnindex == selectedcolumnindex || Requestcolumnindex == selectedcolumnindex || Notescolumnindex == selectedcolumnindex)) {
                                                        toggleDropDialog();
                                                        return false;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        },
                        cellUpdated: function (r, c, oldValue, newValue, isRefresh) {
                            // Setting each cell rowlength as per the largest characters in a row 
                            let range = luckysheet.getRange();

                            if (r && c && r != undefined && c != undefined) {
                                let rowData = luckysheet.getcellvalue(r);
                                if (rowData && rowData.length > 0) {
                                    rowData = rowData.filter((f) => f != null);
                                    let length = [];
                                    let maxLength = 0;
                                    rowData.forEach((f) => {
                                        if (f?.ct?.s) {
                                            if (f?.ct?.s?.length > 1) {
                                                var text = '';
                                                f?.ct?.s?.forEach((e) => { text += e?.v })
                                                length.push(text?.length);
                                            } else { length.push(f?.ct?.s[0]?.v?.length) }
                                        }
                                    });
                                    length = Array.from(new Set(length));
                                    // let minLength = Math.min(...length) + 15;

                                    length.forEach((f) => {
                                        if (f > maxLength) {
                                            maxLength = f;
                                        }
                                    });

                                    let config = luckysheet.getConfig();
                                    config.rowlen[r] = maxLength && maxLength > 5 ? maxLength / 2 + 20 : 30;
                                    luckysheet.setConfig(config);
                                }
                            }

                            // Always revert the cell value to oldValue except the endorsement columns
                            if (range && range.length > 0 && range != undefined) {
                                let selectedColumnIndex = range[0].column[0];
                                const redTableData = sessionStorage.getItem('redTableRangeData');
                                const greenTableData = sessionStorage.getItem('greenTableRangeData');
                                const parsedRedTableData = JSON.parse(redTableData);
                                const parsedGreenTableData = JSON.parse(greenTableData);

                                let sheetflag = luckysheet?.getSheet()?.name;
                                if (luckysheet?.getSheet()?.name != 'QAC not answered questions') {
                                    const tabledata = luckysheet?.getSheet()?.name == 'PolicyReviewChecklist' ? tableColumnDetails : luckysheet?.getSheet()?.name == 'Forms Compare' ? formTableColumnDetails : luckysheet?.getSheet()?.name == 'Exclusion' ? exTableColumnDetails : luckysheet?.getSheet()?.name == 'Red' ? parsedRedTableData : parsedGreenTableData;
                                    const excludedColumns = ["columnid"];
                                    const selectedTable = findTableForIndex(r, tabledata, excludedColumns);
                                    const tblSelectedRow = tabledata[selectedTable];
                                    if (tblSelectedRow != undefined) {
                                        if (selectedTable != 'Table 1' && selectedTable != "ExTable 1" && selectedTable != 'FormTable 1') {
                                            for (const key in tblSelectedRow.columnNames) {
                                                if (tblSelectedRow.columnNames[key] === 0) {
                                                    delete tblSelectedRow.columnNames[key];
                                                }
                                            }
                                            let values = tblSelectedRow ? Object.keys(tblSelectedRow.columnNames) : [];
                                            let docViewerIndex = values.indexOf("Document Viewer") + 2;
                                            if (oldValue && oldValue.ct && oldValue.ct.s && oldValue.ct.s[0] && oldValue != undefined) {
                                                if (selectedColumnIndex <= docViewerIndex) {
                                                    setCellValue(r, c, oldValue);
                                                }
                                            } else if (oldValue && oldValue.m && oldValue.v && oldValue != undefined) {
                                                if (selectedColumnIndex <= docViewerIndex) {
                                                    setCellValue(r, c, oldValue);
                                                }
                                            }
                                            luckysheet.exitEditMode();
                                            return false;
                                        } else if ((luckysheet?.getSheet()?.name != 'Red' && luckysheet?.getSheet()?.name != 'Green') && (selectedTable == 'Table 1' || selectedTable == "ExTable 1" || selectedTable == 'FormTable 1')) {
                                            let values = Object.values(tblSelectedRow.columnNames[0]);
                                            let pageNumberColIndex = values.indexOf("PageNumber");
                                            let colCheck = luckysheet.find("Named Insured");
                                            if (luckysheet?.getSheet()?.name == 'PolicyReviewChecklist' && brokerId == "1167") {
                                                if (colCheck) {
                                                    if (colCheck[0]?.v == "Named Insured" && colCheck[0]?.column == 1) {
                                                        if (colCheck[0]?.row == r)
                                                            return true;
                                                    }
                                                    else if (oldValue && oldValue.m && oldValue.v && oldValue != undefined) {
                                                        setCellValue(r, c, oldValue);
                                                    }
                                                }
                                            }
                                            if (oldValue && oldValue.m && oldValue.v && oldValue != undefined) {
                                                setCellValue(r, c, oldValue);
                                            } else if (oldValue && oldValue.ct && oldValue.ct.s && oldValue.ct.s[0] && oldValue != undefined) {
                                                if (selectedColumnIndex <= pageNumberColIndex) {
                                                    setCellValue(r, c, oldValue);
                                                }
                                            }
                                            else {
                                                if (luckysheet?.getSheet()?.name == 'PolicyReviewChecklist') {
                                                    setCellValue(r, c, '');
                                                }

                                            }
                                        } else {
                                            setCellValue(r, c, oldValue || '');
                                        }
                                    }
                                } 
                                // else if (sheetflag == 'QAC not answered questions') {
                                //     const tabledata = qacTableColumnDetails;
                                //     const keys = Object.keys(tabledata);

                                //     if (keys?.length > 0) {
                                //         keys?.forEach((key) => {
                                //             const tableData = qacTableColumnDetails[key];
                                //             const structuredTblObj = qacTblRangeStructureFun(tableData);
                                //             console.log(structuredTblObj)
                                //         })
                                //     }
                                // }
                            }
                        },
                        rangePasteBefore: function (range, data) {
                            // Denied the action of Copy/Paste for all the cells 
                            if (range && range.length > 0 && range != undefined) {
                                let selectedRowIndex = range[0].row[0];
                                let selectedColumnIndex = range[0].column[0];
                                const redTableData = sessionStorage.getItem('redTableRangeData');
                                const greenTableData = sessionStorage.getItem('greenTableRangeData');
                                const parsedRedTableData = JSON.parse(redTableData);
                                const parsedGreenTableData = JSON.parse(greenTableData);

                                let flagName = luckysheet?.getSheet()?.name;
                                if (luckysheet?.getSheet()?.name != 'QAC not answered questions') {
                                    let tabledata = luckysheet?.getSheet()?.name == 'PolicyReviewChecklist' ? tableColumnDetails : luckysheet?.getSheet()?.name == 'Forms Compare' ? formTableColumnDetails : luckysheet?.getSheet()?.name == 'Exclusion' ? exTableColumnDetails : luckysheet?.getSheet()?.name == 'Red' ? parsedRedTableData : parsedGreenTableData;
                                    const selectedTable = findTblRowAllIndex(selectedRowIndex, tabledata);
                                    const tblSelectedRow = tabledata[selectedTable];
                                    const bgColorIndexes = selectedTable == 'Table 3' ? tblSelectedRow?.range?.start + 1 : selectedTable == "ExTable 1" ? (tblSelectedRow?.range?.start || tblSelectedRow?.range?.start + 1) : tblSelectedRow?.range?.start;
                                    const filteredIndex = filterSelectedRowIndexForCopyPaste(tblSelectedRow, selectedRowIndex);
                                    let values = tblSelectedRow ? Object.values(tblSelectedRow.columnNames) : [];
                                    let largestIndex = selectedTable == "ExTable 1" ? Math.max(...Object.keys(values[0])) : Math.max(...values);
                                    let Actioncolumnindex = selectedTable == "ExTable 1" ? largestIndex - 5 : largestIndex - 3;
                                    let Requestcolumnindex = selectedTable == "ExTable 1" ? largestIndex - 4 : largestIndex - 2;
                                    let Notescolumnindex = selectedTable == "ExTable 1" ? largestIndex - 3 : largestIndex - 1;
                                    let NotesFreeFillcolumnindex = selectedTable == "ExTable 1" ? largestIndex - 2 : largestIndex;

                                    if (tblSelectedRow !== undefined) {
                                        if (luckysheet?.getSheet()?.name != "Exclusion") {
                                            if ((selectedTable != 'Table 3' && selectedRowIndex >= tblSelectedRow.range.start + 2 || selectedTable == 'Table 3' && selectedRowIndex >= tblSelectedRow.range.start + 3) || selectedRowIndex == bgColorIndexes) {
                                                if ((Actioncolumnindex != selectedColumnIndex && Requestcolumnindex != selectedColumnIndex && Notescolumnindex != selectedColumnIndex && NotesFreeFillcolumnindex != selectedColumnIndex) || selectedRowIndex == bgColorIndexes) {
                                                    let isAction = false;
                                                    range.forEach(item => {
                                                        const targetRow = item.row[0];
                                                        if (targetRow === filteredIndex) {
                                                            isAction = true;
                                                        }
                                                    });

                                                    if (isAction == true) {
                                                        return false;
                                                    }
                                                }
                                            }
                                        } else if (luckysheet?.getSheet()?.name == "Exclusion") {
                                            if ((selectedTable == "ExTable 1" && selectedRowIndex >= tblSelectedRow.range.start + 2) || selectedRowIndex == bgColorIndexes) {
                                                if ((Actioncolumnindex != selectedColumnIndex && Requestcolumnindex != selectedColumnIndex && Notescolumnindex != selectedColumnIndex && NotesFreeFillcolumnindex != selectedColumnIndex) || selectedRowIndex == bgColorIndexes) {
                                                    let isAction = false;
                                                    range.forEach(item => {
                                                        const targetRow = item.row[0];
                                                        if (targetRow === filteredIndex) {
                                                            isAction = true;
                                                        }
                                                    });

                                                    if (isAction == true) {
                                                        return false;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            //Denied the action of Copy/Paste the data in the Header Section for Exclusion
                            let flagName = luckysheet?.getSheet()?.name;
                            if (luckysheet?.getSheet()?.name == 'Exclusion') {
                                let tabledata = exTableColumnDetails;
                                const selectedTable = "ExTable 1";
                                const exclusionHeader = tabledata[selectedTable]?.range?.start;
                                let isHeader = false;

                                range.forEach(item => {
                                    const targetRow = item.row[0];
                                    if (targetRow === exclusionHeader) {
                                        isHeader = true;
                                    }
                                });

                                if (isHeader == true) {
                                    return false;
                                }
                            }
                        },
                        rangeSelect: function (index, sheet) {
                            //Handled the functionality for Del key operation and CTrl+ d shortcut options
                            let range = luckysheet?.getRange();
                            let currentFlag = luckysheet?.getSheet()?.name;

                            if (range && range.length > 0 && range != undefined) {
                                // if ( e.which != 40 ) {
                                let selectedRowIndex = range[0].row[0];
                                let selectedColumnIndex = range[0].column[0];
                                const targetrow = range[0].row[0];
                                const targetcolumn = range[0].column;
                                const sheetdatas = luckysheet.getSheetData();
                                const getrowdata = sheetdatas[targetrow - 1];
                                const redTableData = sessionStorage.getItem('redTableRangeData');
                                const greenTableData = sessionStorage.getItem('greenTableRangeData');
                                const parsedRedTableData = JSON.parse(redTableData);
                                const parsedGreenTableData = JSON.parse(greenTableData);

                                if (luckysheet?.getSheet()?.name != 'QAC not answered questions' && luckysheet?.getSheet()?.name != "Exclusion") {
                                    const tabledata = luckysheet?.getSheet()?.name == 'PolicyReviewChecklist' ? tableColumnDetails : luckysheet?.getSheet()?.name == 'Forms Compare' ? formTableColumnDetails : luckysheet?.getSheet()?.name == 'Red' ? parsedRedTableData : parsedGreenTableData;
                                    if (tabledata != null) {
                                        const excludedColumns = luckysheet?.getSheet()?.name == 'PolicyReviewChecklist' ? ["columnid"] : luckysheet?.getSheet()?.name == 'Red' ? ["columnid"]
                                            : luckysheet?.getSheet()?.name == 'Green' ? ["columnid"] : ["Id", "Checklist Questions", "OBSERVATION", "Policy LOB", "Page Number", "IsMatched", "columnid", "Attached Forms"];
                                        const selectedTable = findTableForIndex(selectedRowIndex, tabledata, excludedColumns);
                                        const tblSelectedRow = tabledata[selectedTable];
                                        if (luckysheet?.getSheet()?.name == 'Forms Compare' && selectedTable != null) {
                                            for (const key in tblSelectedRow.columnNames) {
                                                if (tblSelectedRow.columnNames[key] === 0) {
                                                    delete tblSelectedRow.columnNames[key];
                                                }
                                            }
                                        }
                                        let values = tblSelectedRow ? Object.keys(tblSelectedRow.columnNames) : [];
                                        let docViewerIndex = values.indexOf("Document Viewer") + 2;
                                        let ctrlDDocViewerIndex = values.indexOf("Document Viewer") + 1;
                                        let Actioncolumnindex = ctrlDDocViewerIndex + 1;
                                        let Requestcolumnindex = ctrlDDocViewerIndex + 2;
                                        let Notescolumnindex = ctrlDDocViewerIndex + 3;
                                        let NotesFreeFillcolumnindex = ctrlDDocViewerIndex + 4;
                                        if (index && sheet && sheet.length > 0) {
                                            document.onkeyup = function (e) {
                                                if (e.which != 40 && luckysheet?.getSheet()?.name != "QAC not answered questions") {
                                                    if ((e.which == 46 || e.which == 8) && luckysheet?.getSheet()?.name != "QAC not answered questions") {
                                                        if ((selectedTable != "Table 1" || selectedTable != "FormTable 1")) {
                                                            if ((selectedColumnIndex < docViewerIndex)) {
                                                                luckysheet.undo()
                                                            }
                                                        } else if (selectedTable == "Table 1" || selectedTable == "FormTable 1") {
                                                            const getRange = tableColumnDetails || formTableColumnDetails
                                                            const rangeEnd = getRange["Table 1"].range.end
                                                            const range = luckysheet.getRange();
                                                            const sheetrange = range[0].row[1];
                                                            if ((rangeEnd >= sheetrange || rangeEnd <= sheetrange)) {
                                                                luckysheet.undo()
                                                            }
                                                        }
                                                    } else if (e.ctrlKey && e.which == 68) {
                                                        if (range[0].column[0] == range[0].column[1]) {
                                                            const getrowdata = sheetdatas[targetrow - 1];
                                                            const endorsementColumns = [1, 2, 3, 4];
                                                            endorsementColumns.forEach((aod) => {
                                                                if (getrowdata[targetcolumn[0]]) {
                                                                    if (Actioncolumnindex != undefined && Requestcolumnindex != undefined && Notescolumnindex != undefined && NotesFreeFillcolumnindex != undefined) {
                                                                        if ((Actioncolumnindex == selectedColumnIndex || Requestcolumnindex == selectedColumnIndex || Notescolumnindex == selectedColumnIndex || NotesFreeFillcolumnindex == selectedColumnIndex)) {
                                                                            setTimeout(() => {
                                                                                setCellValue(targetrow, (ctrlDDocViewerIndex + aod), getrowdata[ctrlDDocViewerIndex + aod]);
                                                                            }, 100);
                                                                        }
                                                                    }
                                                                }
                                                            });
                                                        } else {
                                                            const startColumn = ctrlDDocViewerIndex + 1;
                                                            const endColumn = Math.max(range[0].column[0], range[0].column[1]);
                                                            for (let idx = startColumn; idx <= endColumn; idx++) {
                                                                if (getrowdata[idx]) {
                                                                    if (Actioncolumnindex != undefined && Requestcolumnindex != undefined && Notescolumnindex != undefined && NotesFreeFillcolumnindex != undefined) {
                                                                        if ((Actioncolumnindex == selectedColumnIndex || Requestcolumnindex == selectedColumnIndex || Notescolumnindex == selectedColumnIndex || NotesFreeFillcolumnindex == selectedColumnIndex)) {
                                                                            setCellValue(targetrow, idx, getrowdata[idx]);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } else if (luckysheet?.getSheet()?.name == "Exclusion") {
                                    const tabledata = exTableColumnDetails;
                                    const tblSelectedRow = tabledata["ExTable 1"];
                                    let values = Object.values(tblSelectedRow.columnNames[0]);
                                    let pageNoIndex = values.indexOf("PageNumber");
                                    document.onkeyup = function (e) {
                                        if (e.which != 40 && luckysheet?.getSheet()?.name != "QAC not answered questions") {
                                            if ((e.which == 46 || e.which == 8) && luckysheet?.getSheet()?.name != "QAC not answered questions") {
                                                if ((selectedColumnIndex < pageNoIndex)) {
                                                    luckysheet.undo();
                                                }
                                            }
                                        }
                                    }
                                } 
                            }
                        },
                        cellAllRenderBefore: function (data, sheetFile, ctx) {
                            if (sheetFile && sheetFile != undefined) {
                                let endorsementColflag = sheetFile?.name;
                                setEndorsementColflag(endorsementColflag);
                            }
                        }
                    },
                    cellRightClickConfig: {
                        copy: false, // copy
                        copyAs: false, // copy as
                        paste: false, // paste
                        insertRow: false, // insert row
                        insertColumn: false, // insert column
                        deleteRow: false, // delete the selected row
                        deleteColumn: false, // delete the selected column
                        deleteCell: false, // delete cell
                        hideRow: false, // hide the selected row and display the selected row
                        hideColumn: false, // hide the selected column and display the selected column
                        rowHeight: false, // row height
                        columnWidth: false, // column width
                        clear: false, // clear content
                        matrix: false, // matrix operation selection
                        sort: false, // sort selection
                        filter: false, // filter selection
                        chart: false, // chart generation
                        image: false, // insert picture
                        link: false, // insert link
                        data: false, // data verification
                        cellFormat: false // Set cell format
                    },
                    showtoolbarConfig: {
                        moreFormats: false,
                        sortAndFilter: false,
                        link: false,
                        chart: false,
                        print: false,
                        textRotateMode: false,
                        image: false,
                        postil: false,
                        dataVerification: false,
                        splitColumn: false,
                        screenshot: false,
                        findAndReplace: false,
                    }
                };
                luckysheet.create(options);
            }
        }
    }

    // Function to handle cell selection and update values
    let previousCell = { row: null, col: null, text: 'string' };
    const handleCellSelection = async (range, flagCheck) => {
        if (flagCheck === 'PolicyReviewChecklist') {
            if (range && range[0]?.row && range[0]?.column && range[0]?.row[0] === range[0]?.row[1] &&
                range[0]?.column[0] === range[0]?.column[1] && [3, 4, 5, 6].includes(range[0]?.row[0]) &&
                range[0]?.column[0] === 4) {

                const currentRow = range[0]?.row[0];
                const currentCol = 4;

                let getRowValue = luckysheet.getcellvalue(currentRow);
                const objectAtIndex4 = getRowValue[4];
                // console.log(objectAtIndex4);
                if (objectAtIndex4?.ct?.s && Array.isArray(objectAtIndex4?.ct?.s)) {
                    let varianceText = objectAtIndex4?.ct?.s[1]?.v
                    if (varianceText === 'Matched' || varianceText === 'All Variances' || varianceText === 'Full Variances' || varianceText === 'Variances' || varianceText === 'Details not available in the document') {
                        const revertBoxCode = {
                            "ct": {
                                "fa": "General",
                                "t": "inlineStr",
                                "s": [
                                    {
                                        "fs": "16",
                                        "v": "■ "
                                    },
                                    {
                                        "vt": "0",
                                        "ht": "1",
                                        "fs": "9",
                                        "un": 0,
                                        "bl": 1,
                                        "fc": "#0000ff",
                                        "ff": "\"Tahoma\"",
                                        "m": varianceText,
                                        "v": varianceText
                                    }
                                ]
                            },
                            "merge": null,
                            "w": 55,
                            "tb": "2",
                            "fc": "#0000ff",
                            "fs": "16"
                        }

                        // Update the current cell with revertBoxCode
                        setCellValue(currentRow, currentCol, revertBoxCode);

                        // Restore the previous cell with previouscode if there was a previous selection
                        if (previousCell.row !== null && previousCell.col !== null && previousCell.text !== undefined) {
                            const previousBoxcode = {
                                "ct": {
                                    "fa": "General",
                                    "t": "inlineStr",
                                    "s": [
                                        {
                                            "fs": "16",
                                            "v": "□ "
                                        },
                                        {
                                            "vt": "0",
                                            "ht": "1",
                                            "fs": "9",
                                            "un": 0,
                                            "bl": 1,
                                            "fc": "#0000ff",
                                            "ff": "\"Tahoma\"",
                                            "m": previousCell.text,
                                            "v": previousCell.text
                                        }
                                    ]
                                },
                                "merge": null,
                                "w": 55,
                                "tb": "2",
                                "fc": "#0000ff",
                                "fs": "16"
                            }
                            setCellValue(previousCell.row, previousCell.col, previousBoxcode);
                        }

                        // Update the previousCell to the current cell
                        previousCell.row = currentRow;
                        previousCell.col = currentCol;
                        previousCell.text = varianceText
                    }
                }
            }
        }
    }

    const ExportClick = async (hasExport) => {
        //Broker configuration for exporting to Excel: If anything changes in this function keep an eye on the broker configuration
        const mergedData = [];
        let sheetnames = []
        let flagcheck = luckysheet.getAllSheets();

        for (let i = 0; i < flagcheck.length; i++) {
            sheetnames.push(flagcheck[i].name);
        }

        if (sheetnames?.includes("PolicyReviewChecklist")) {
            generateExcelData(true, "PolicyReviewChecklist", (dataFromOnUpdateClick) => {
                const Tabledata1 = dataFromOnUpdateClick;
                mergedData.push(...Tabledata1);
            });
        }
        if (sheetnames?.includes("Forms Compare")) {
            generateExcelData(true, "Forms Compare", (dataFromOnUpdateClick) => {
                const Tabledata2 = dataFromOnUpdateClick;
                mergedData.push(...Tabledata2);
            });
        }
        if (brokerId == '1150' || brokerId == '1003') {
            if (sheetnames?.includes("Exclusion")) {
                generateExcelData(true, "Exclusion", (dataFromOnUpdateClick) => {
                    const Tabledata3 = dataFromOnUpdateClick;
                    mergedData.push(...Tabledata3);
                });
            }
        }

        if (sheetnames?.includes("Red")) {
            generateExcelData(true, "Red", (dataFromOnUpdateClick) => {
                const Tabledata4 = dataFromOnUpdateClick;
                mergedData.push(...Tabledata4);
            });
        }
        if (sheetnames?.includes("Green")) {
            generateExcelData(true, "Green", (dataFromOnUpdateClick) => {
                const Tabledata4 = dataFromOnUpdateClick;
                mergedData.push(...Tabledata4);
            });
        }
        if (sheetnames?.includes("QAC not answered questions")) {
            if (brokerId == '1162') {
                generateExcelData(true, "QAC not answered questions", (dataFromOnUpdateClick) => {
                let Tabledata5 = dataFromOnUpdateClick;
                if(Tabledata5?.length > 0) {
                    mergedData.push(...Tabledata5);
                } else {
                    let mergedStaticData =  staticDataforQacExport();
                    Tabledata5 = mergedStaticData;
                    mergedData.push(...Tabledata5);
                }
                });
            }
        }

        const uniqueMergedData = mergedData.reduce((acc, curr) => {
            if (!acc.find(item => item?.TableName === curr?.TableName)) {
                acc.push(curr);
            }
            return acc;
        }, []);
        const filteredData = uniqueMergedData?.filter(item => {
            const data = JSON.parse(item.Data);
            return !Array.isArray(data) || data.length > 0;
        });

        let formTableData = [];

        filteredData?.forEach(item => {
            if (item?.TableName === "CsrTable 1" || item?.TableName === "ExclusionTable" || item?.TableName === "CsrFormTable 1" || item?.TableName === "RedTable 1" || item?.TableName === "GreenTable 1" || item?.TableName === "HighVolumeTable1") {
                let k = item?.TableName == "CsrTable 1" ? "Table 1" : item.TableName;
                formTableData.push(k);
            }
        });
        if (brokerId == '1162') {

        if (!formTableData.includes("HighVolumeTable1") ? true : false) {
            // if (luckysheet?.getSheet()?.name == 'QAC not answered questions'){
                formTableData.push("LowVolumeTable1")
            // }
        }
    }

        formTableData.push("csrflag");
        // Convert formTableData to JSON string and wrap it in double quotes
        const formTableDataJson = '"' + JSON.stringify(formTableData) + '"';

        if (hasExport == true) {
            const response = await exportExcelData(filteredData, formTableDataJson);
            // if(reqEndorsementColHideBrokerId.includes(brokerId)) {
            const getLocalStgRowDatas = localStorage.getItem('endorsementRowData');
            await CsrPendingReport(props?.selectedJob, sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName"), getLocalStgRowDatas, 'CSrExport-Save');
            // }
            return response;
        }
    }

    const exportExcelData = async (Tabledata, TableNames) => {
        document.body.classList.add('loading-indicator');
        const Token = await processAndUpdateToken(token);
        updateGridAuditLog(jobId, "CSR-ExportData", "CSR-ExportExcel", (sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName")));
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${baseUrl}/api/Excel/ExportExcel`;

        try {
            const response = await axios({
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    Data: TableNames,
                    Tabledata: Tabledata
                },
                responseType: 'blob'
            });
            if (response.status !== 200) {
                return "error";
            }

            const url = window.URL.createObjectURL(new Blob([response.data]));
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', `${Tabledata[0].Id}GridExcel.xlsx`);
            document.body.appendChild(link);
            link.click();

            return "success";
        } catch (error) {
            // console.error('Error:', error);
            return "error";
        } finally {
            updateGridAuditLog(jobId, "CSR-ExportData-Success", "CSR-ExportExcel-Success1", (sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName")));
            document.body.classList.remove('loading-indicator');
            return "success";
        }
    };

    const generateExcelData = (isExport, sheetName, callback) => {
        let flagCheck = sheetName;
        const propsData = luckysheet.getAllSheets()?.filter(f => f.name.includes(flagCheck))[0].data;
        setsheetState(propsData);
        const gradiationSet = flagCheck == 'Red' ? sessionStorage.getItem('redTableRangeData')
            : flagCheck == 'Green' ? sessionStorage.getItem('greenTableRangeData') : null;

        let parsedSet = JSON.parse(gradiationSet);
        const keys = flagCheck == "PolicyReviewChecklist" ? Object.keys(tableColumnDetails)
            : flagCheck == "Forms Compare" ? Object.keys(formTableColumnDetails) : flagCheck == 'Red' ? Object.keys(parsedSet) : flagCheck == 'Green' ? Object.keys(parsedSet)
                : flagCheck == "Exclusion" ? "" : "";

        const findTable = [];
        let policySheetData = [];
        let formCompareSheetData = [];
        let exclusionSheetData = [];
        let qacSheetDataSet = [];
        let redSheetData = [];
        let greenSheetData = [];
        let exclusionSetArray;

        if (flagCheck == "PolicyReviewChecklist") {
            if (keys?.length > 0 && propsData && propsData != undefined) {
                keys?.forEach((f) => {
                    const tableData = tableColumnDetails[f];
                    if (tableData && tableData?.range && tableData?.range?.start != undefined && tableData?.range?.end != undefined && tableData?.columnNames && tableData?.range?.end != '') {
                        let slicedData = propsData?.slice(tableData?.range?.start, tableData?.range?.end + 1);
                        if (f == 'Table 1') {
                            slicedData = slicedData?.slice(3);
                            slicedData = slicedData?.map(sublist => sublist?.filter(item => item !== null));
                        }
                        findTable.push(slicedData);
                    }
                });
            }

            let headerSpilttedArray = [];
            let tableDataSetArray = [];
            findTable?.map((e, index) => {
                if (index > 0) {
                    let tblIndex = [];
                    let tblColumnName = [];
                    let hasReachedLimit = false;

                    const lobMap = props?.data?.map(r => r.AvailableLobs)[0];
                    const indexMap = e[0][1].m;
                    const isTextInLobMap = lobMap?.includes(indexMap);
                    if (isTextInLobMap) {
                        index = 2;
                    }
                    const data = index == 2 ? e[1] : e[0];
                    data?.map((e1, index1) => {
                        if (index1 > 0 && !hasReachedLimit && e1?.m?.toLowerCase() != "document viewer") {
                            tblIndex.push(index1);
                            tblColumnName.push(e1?.m || e1?.v);
                        }
                        if (index1 > 0 && e1?.m?.toLowerCase() == "document viewer") {
                            hasReachedLimit = true;
                            let largeIndex = Math.max(...tblIndex);
                            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                                tblColumnName = [...tblColumnName, "Document Viewer", "ActionOnDiscrepancy", "Notes", "NotesFreeFill"]
                                tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4];
                            } else {
                                tblColumnName = [...tblColumnName, "Document Viewer", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"]
                                tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4, largeIndex + 5];
                            }
                            headerSpilttedArray.push({ "Table": `Table ${index + 1}`, index: tblIndex, tblColumnName });
                        }

                        headerSpilttedArray?.forEach(table => {
                            const tableNumber = parseInt(table?.Table?.split(" ")[1]);
                            if (tableNumber === 2 || tableNumber === 4) {
                                table.tblColumnName[0] = "COVERAGE_SPECIFICATIONS_MASTER";
                            } else if (tableNumber >= 5 && tableNumber <= 70) {
                                table.tblColumnName[0] = "Coverage_Specifications_Master";
                            }
                        });
                    });
                }
            });
            headerSpilttedArray.forEach((item, index) => {
                item.Table = `Table ${index + 2}`;
            });
            // console.log(headerSpilttedArray)

            findTable?.map((f, index) => {
                if (index > 0) {
                    let keyValuePair = [];
                    f?.map((e, index1) => {
                        if (headerSpilttedArray[index - 1]?.Table == "Table 3" ? index1 > 2 : index1 > 1) {
                            let tableValueMap = headerSpilttedArray[index - 1];
                            if (tableValueMap) {
                                let object = {};
                                tableValueMap.index.map((i, index2) => {
                                    if (e[i]?.ct?.s && Array.isArray(e[i]?.ct?.s)) {
                                        let arrayS = e[i]?.ct?.s.filter((f) => f != null);
                                        let concatenatedValues = arrayS?.map(item => item?.v)?.join('');
                                        concatenatedValues = concatenatedValues?.replace(/\r\n/g, '~~');
                                        // if (concatenatedValues.endsWith('~~')) {
                                        //     concatenatedValues = concatenatedValues.slice(0, -2);
                                        // }
                                        // concatenatedValues = concatenatedValues.trimEnd();
                                        const finalValue = e[i]?.m || e[i]?.v || concatenatedValues;
                                        object[`${tableValueMap?.tblColumnName[index2]}`] = finalValue;
                                    } else {
                                        object[`${tableValueMap?.tblColumnName[index2]}`] = e[i]?.m || e[i]?.v || e[i]?.ct?.s;
                                    }
                                    if (tableValueMap.index?.length == index2 + 1) {
                                        keyValuePair.push(object);
                                    }
                                })
                            }
                        }
                        if (f?.length == index1 + 1) {
                            tableDataSetArray.push({ "Table": `CsrTable ${index + 1}`, DiscrepancyData: keyValuePair });
                        }
                    });
                }
            });
            // console.log(tableDataSetArray);

            let cellDatas = csrPolicyData;
            cellDatas?.forEach((t, index) => { t.Tablename = `CsrTable ${index + 1}` })
            let lobResult = [];
            let uniqueLobSet = new Set();

            if (tableDataSetArray && cellDatas && cellDatas?.length > 0 && tableDataSetArray?.length > 0) {
                cellDatas.forEach(table => {
                    if ((table?.Tablename !== 'CsrTable 1') && !uniqueLobSet.has(table?.Tablename)) {
                        let policyLOB = table?.TemplateData[0]["POLICY LOB"] || table?.TemplateData[0]["Policy LOB"];
                        lobResult.push({
                            Table: table.Tablename,
                            "POLICY LOB": policyLOB
                        });
                        uniqueLobSet.add(table.Tablename);
                    }
                });
                setLobResult(lobResult);

                // Create a mapping from Table to POLICY LOB
                let lobMapping = {};
                lobResult.forEach(item => {
                    lobMapping[item.Table] = item["POLICY LOB"];
                });

                // Assign the POLICY LOB value to the corresponding Table in tableDataSetArray
                tableDataSetArray?.forEach(table => {
                    if (lobMapping[table.Table]) {
                        table["POLICY_LOB"] = lobMapping[table.Table];
                        delete table.Table;
                    }
                });

                const documentViewerMap = {};
                cellDatas?.forEach(tbl => {
                    const tableName = tbl.Tablename;
                    if (tableName != 'CsrTable 1') {
                        if (tbl?.TemplateData && tbl?.TemplateData?.length > 0) {
                            documentViewerMap[tableName] = tbl?.TemplateData?.map(row => row["Document Viewer"]);
                        }
                    }
                });

                tableDataSetArray.forEach((f, index) => {
                    const TableName = "CsrTable " + (index + 2)
                    let data = f?.DiscrepancyData;
                    if (data && data?.length > 0) {
                        data = (data?.map((item, index) => {
                            item["policyLob"] = f?.POLICY_LOB;

                            // Replace Document Viewer value if mapping exists
                            if (documentViewerMap[TableName] && Array.isArray(documentViewerMap[TableName])) {
                                if (index < documentViewerMap[TableName].length) {
                                    let docItem = item["Document Viewer"] == " " ? item["Document Viewer"].trim() : item["Document Viewer"];
                                    if (docItem != "") {
                                        item["Document Viewer"] = documentViewerMap[TableName][index];
                                    }
                                }
                            }

                            // Replace 'Click here' msg to an empty string
                            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                                if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["Notes"] === "~~Click here" && item["NotesFreeFill"] === "~~Click here"
                                ) || (item["ActionOnDiscrepancy"] === "Click here" && item["Notes"] === "Click here" && item["NotesFreeFill"] === "Click here"
                                    ) ||
                                    (item["ActionOnDiscrepancy"] === " Click here" && item["Notes"] === " Click here" && item["NotesFreeFill"] === " Click here"
                                    )) {
                                    item["ActionOnDiscrepancy"] = "";
                                    item["Notes"] = "";
                                    item["NotesFreeFill"] = "";
                                }
                            } else if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["RequestEndorsement"] == "~~Click here"
                                && item["Notes"] == "~~Click here" && item["NotesFreeFill"] == "~~Click here"
                            ) || (item["ActionOnDiscrepancy"] === "Click here" && item["RequestEndorsement"] == "Click here"
                                && item["Notes"] == "Click here" && item["NotesFreeFill"] == "Click here"
                                ) || (item["ActionOnDiscrepancy"] === " Click here" && item["RequestEndorsement"] == " Click here"
                                    && item["Notes"] == " Click here" && item["NotesFreeFill"] == " Click here"
                                )) {
                                item["ActionOnDiscrepancy"] = "";
                                item["RequestEndorsement"] = "";
                                item["Notes"] = "";
                                item["NotesFreeFill"] = "";
                            } else {
                                ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"].forEach(key => {
                                    if ((item[key] === "~~Click here") || (item[key] === "Click here") || (item[key] === " Click here")) {
                                        item[key] = "";
                                    }
                                })
                            }
                            return item;
                        }));
                    }


                    const cleanedData = data?.filter(entry =>
                        !Object.values(entry).includes(undefined)
                    );

                    policySheetData.push({ TableName, Data: JSON.stringify(cleanedData) });
                });
            }
        } else if (flagCheck == "Forms Compare") {
            keys?.forEach((f) => {
                const tableData = formTableColumnDetails[f];
                if (tableData && tableData?.range && tableData?.range?.start != undefined && tableData?.range?.end != undefined && tableData?.columnNames && tableData?.range?.end != '') {
                    let slicedData = propsData.slice(tableData?.range?.start, tableData?.range?.end + 1);
                    if (f == 'FormTable 1') {
                        slicedData = slicedData?.slice(3);
                        slicedData = slicedData?.map(sublist => sublist?.filter(item => item !== null));
                    }
                    findTable.push(slicedData);
                }
            });

            let headerSpilttedArray = [];
            let tableDataSetArray = [];

            let limitReached = false;
            findTable?.map((e, index) => {
                if (index >= 1 && !limitReached) {
                    let tblIndex = [];
                    let tblColumnName = [];
                    let hasReachedLimit = false;
                    let tabledata = formTableColumnDetails;
                    for (let tableName in tabledata) {
                        if (tableName != "FormTable 1") {
                            let tableInfo = tabledata[tableName];
                            for (let columnName in tableInfo.columnNames) {
                                let columnValue = tableInfo.columnNames[columnName];
                                if (columnValue === 0) {
                                    delete tableInfo.columnNames[columnName];
                                    if (tableInfo.columnNames.hasOwnProperty("Attached Forms")) {
                                        delete tableInfo.columnNames["Attached Forms"];
                                    }
                                }
                            }
                        }
                    }

                    const data = index == 1 ? e[1] : e[1];
                    const filteredData = data?.filter(item => item !== null);
                    filteredData?.forEach((e1, index1) => {
                        if (e1 && e1?.m) {
                            if (index1 >= 0 && !hasReachedLimit && e1?.m?.toLowerCase() != "document viewer" && e1?.m != "Actions on Discrepancy (from AMs)") {
                                tblIndex.push(index1);
                                tblColumnName.push(e1?.m || e1?.v);
                            }
                        }
                        if (index == 1 && e1?.m?.toLowerCase() == "document viewer" && keys[1] == 'FormTable 2') {
                            tblColumnName = tblColumnName.filter(column => column !== undefined);
                            let largeIndex = Math.max(...tblIndex);
                            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                                tblColumnName = [...tblColumnName, "Document Viewer", "ActionOnDiscrepancy", "Notes", "NotesFreeFill"]
                                tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4];
                            } else {
                                tblColumnName = [...tblColumnName, "Document Viewer", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"]
                                tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4, largeIndex + 5];
                            }
                            headerSpilttedArray.push({ "Table": `FormTable ${index + 1}`, index: tblIndex, tblColumnName });
                        }
                    });
                    if (index == 2 && keys[2] == 'FormTable 3') {
                        tblColumnName = tblColumnName.filter(column => column !== undefined);
                        let largeIndex = Math.max(...tblIndex);
                        if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                            tblColumnName = [...tblColumnName, "ActionOnDiscrepancy", "Notes", "NotesFreeFill"]
                            tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3];
                        } else {
                            tblColumnName = [...tblColumnName, "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"]
                            tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4];
                        }
                        headerSpilttedArray.push({ "Table": `FormTable ${index + 1}`, index: tblIndex, tblColumnName });
                    }
                }
                headerSpilttedArray?.forEach(table => {
                    if (table.Table === "FormTable 2" || table.Table === "FormTable 3") {
                        table.tblColumnName[0] = "COVERAGE_SPECIFICATIONS_MASTER";
                    }
                });
                headerSpilttedArray?.forEach(item => {
                    const tableDetails = formTableColumnDetails[item.Table];
                    if (tableDetails && tableDetails.columnNames) {
                        const validColumns = Object.keys(tableDetails.columnNames);
                        item.tblColumnName = item?.tblColumnName?.filter(column => validColumns?.includes(column));
                        item.index = Array.from({ length: item.tblColumnName.length }, (_, i) => i);
                    }
                });
            });

            findTable?.map((f, index) => {
                if (index >= 1) {
                    let keyValuePair = [];
                    f?.map((e, index1) => {
                        if (index1 > 0) {
                            let tableValueMap = headerSpilttedArray[index - 1];
                            if (tableValueMap) {
                                let object = {};
                                tableValueMap.index.map((i, index2) => {
                                    i = i + 1;
                                    if (e[i]?.ct?.s && Array.isArray(e[i]?.ct?.s)) {
                                        let filteredS = e[i]?.ct?.s.filter((f) => f != null);
                                        let concatenatedValues = filteredS?.map(item => item?.v)?.join('');
                                        concatenatedValues = concatenatedValues?.replace(/\r\n/g, '~~');
                                        // if (concatenatedValues.endsWith('~~')) {
                                        //     concatenatedValues = concatenatedValues.slice(0, -2);
                                        // }
                                        // concatenatedValues = concatenatedValues.trimEnd();
                                        const finalValue = e[i]?.m || e[i]?.v || concatenatedValues;
                                        object[`${tableValueMap?.tblColumnName[index2]}`] = finalValue;
                                    } else {
                                        object[`${tableValueMap?.tblColumnName[index2]}`] = e[i]?.m || e[i]?.v || e[i]?.ct?.s;
                                    }
                                    if (tableValueMap.index?.length == index2 + 1) {
                                        keyValuePair.push(object);
                                    }
                                })
                            }
                        }
                    });
                    let slicedKeyValuePair = keyValuePair.slice(2);
                    tableDataSetArray.push({ Table: `FormTable ${index + 1}`, DiscrepancyData: slicedKeyValuePair });
                }
            });

            let cellDatas = formsComparedata;
            let lobResult = [];
            let uniqueLobSet = new Set();

            if (tableDataSetArray && cellDatas && cellDatas?.length > 0 && tableDataSetArray?.length > 0) {
                cellDatas?.forEach(table => {
                    if ((table.Tablename !== 'FormTable 1') && !uniqueLobSet.has(table.Tablename)) {
                        let policyLOB = table?.TemplateData[0]["POLICY LOB"] || table?.TemplateData[0]["Policy LOB"];
                        lobResult.push({
                            Table: table.Tablename,
                            "POLICY LOB": policyLOB
                        });
                        uniqueLobSet.add(table.Tablename);
                    }
                });
                setLobResult(lobResult);

                // Create a mapping from Table to POLICY LOB
                let lobMapping = {};
                lobResult?.forEach(item => {
                    lobMapping[item.Table] = item["POLICY LOB"];
                });

                // Assign the POLICY LOB value to the corresponding Table in tableDataSetArray
                tableDataSetArray?.forEach(table => {
                    if (lobMapping[table.Table]) {
                        table["POLICY_LOB"] = lobMapping[table.Table];
                        delete table.Table;
                    }
                });

                const documentViewerMap = {};
                formsComparedata?.forEach(tbl => {
                    let tableName = tbl.Tablename;
                    // Adjust tableName from 'FormTable' to 'CsrFormTable'
                    if (tableName.startsWith('FormTable ')) {
                        const numberPart = tableName.split(' ')[1];
                        tableName = `CsrFormTable ${numberPart}`;
                    }
                    if (tableName != 'CsrFormTable 1' && tableName != 'CsrFormTable 3') {
                        if (tbl?.TemplateData && tbl?.TemplateData?.length > 0) {
                            documentViewerMap[tableName] = tbl?.TemplateData?.map(row => row["Document Viewer"]);
                        }
                    }
                });

                tableDataSetArray?.forEach((f, index) => {
                    const TableName = "CsrFormTable " + (index + 2);
                    let data = f?.DiscrepancyData;
                    if (data && data?.length > 0) {
                        data = (data.map((item, index) => {
                            item["policyLob"] = f?.POLICY_LOB;

                            // Replace Document Viewer value if mapping exists
                            if (documentViewerMap[TableName] && Array.isArray(documentViewerMap[TableName])) {
                                if (index < documentViewerMap[TableName].length) {
                                    let docItem = item["Document Viewer"] == " " ? item["Document Viewer"].trim() : item["Document Viewer"];
                                    if (docItem != "") {
                                        item["Document Viewer"] = documentViewerMap[TableName][index];
                                    }
                                }
                            }

                            // Replace 'Click here' msg to an empty string
                            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                                if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["Notes"] === "~~Click here" && item["NotesFreeFill"] === "~~Click here"
                                ) || (item["ActionOnDiscrepancy"] === "Click here" && item["Notes"] === "Click here" && item["NotesFreeFill"] === "Click here"
                                    ) ||
                                    (item["ActionOnDiscrepancy"] === " Click here" && item["Notes"] === " Click here" && item["NotesFreeFill"] === " Click here"
                                    )) {
                                    item["ActionOnDiscrepancy"] = "";
                                    item["Notes"] = "";
                                    item["NotesFreeFill"] = "";
                                }
                            } else if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["RequestEndorsement"] == "~~Click here"
                                && item["Notes"] == "~~Click here" && item["NotesFreeFill"] == "~~Click here"
                            ) || (item["ActionOnDiscrepancy"] === "Click here" && item["RequestEndorsement"] == "Click here"
                                && item["Notes"] == "Click here" && item["NotesFreeFill"] == "Click here"
                                ) || (item["ActionOnDiscrepancy"] === " Click here" && item["RequestEndorsement"] == " Click here"
                                    && item["Notes"] == " Click here" && item["NotesFreeFill"] == " Click here"
                                )) {
                                item["ActionOnDiscrepancy"] = "";
                                item["RequestEndorsement"] = "";
                                item["Notes"] = "";
                                item["NotesFreeFill"] = "";
                            } else {
                                ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"].forEach(key => {
                                    if ((item[key] === "~~Click here") || (item[key] === "Click here") || (item[key] === " Click here")) {
                                        item[key] = "";
                                    }
                                })
                            }
                            return item;
                        }));
                    }

                    const cleanedData = data.filter(entry =>
                        !Object.values(entry).includes(undefined)
                    );
                    formCompareSheetData.push({ TableName, Data: JSON.stringify(cleanedData) });
                });
            }
        } else if (flagCheck == 'Red' || flagCheck == 'Green') {
            keys?.forEach((f) => {
                const tableData = parsedSet[f];
                if (tableData && tableData?.range && tableData?.range?.start != undefined && tableData?.range?.end != undefined && tableData?.range?.end != '') {
                    let slicedData = propsData?.slice(tableData?.range?.start, tableData?.range?.end + 1);
                    if (f == 'Table 1') {
                        slicedData = slicedData?.map(sublist => sublist.filter(item => item !== null));
                    }
                    findTable.push(slicedData);
                }
            });

            let headerSpilttedArray = [];
            let tableDataSetArray = [];

            findTable?.map((e, index) => {
                if (index > 0) {
                    let tblIndex = [];
                    let tblColumnName = [];
                    let hasReachedLimit = false;
                    let data;

                    data = index == 2 ? e[1] : e[0];
                    if (flagCheck == 'Green') {
                        if (data[0] === null && data[1] !== null && data[1]?.m !== 'Are the forms and endorsements attached, listed in current term policy?') {
                            if (e[1]) {
                                data = e[1];
                            }
                        } else if (data[0] === null) {
                            data = e[0];
                        } else if (data[1]?.m === 'Are the forms and endorsements attached, listed in current term policy?') {
                            data = e[0];
                        }
                    } else {
                        data = index == 2 ? e[1] : e[0];
                    }

                    const containsActionsOnDiscrepancy = e[1]?.some(obj => obj && obj.m === "Actions on Discrepancy");
                    data = containsActionsOnDiscrepancy ? e[0] : e[1];

                    data?.map((e1, index1) => {
                        if (index1 > 0 && !hasReachedLimit && e1?.m?.toLowerCase() != "document viewer") {
                            tblIndex.push(index1);
                            tblColumnName.push(e1?.m || e1?.v);
                        }
                        if (index1 > 0 && e1?.m?.toLowerCase() == "document viewer") {
                            hasReachedLimit = true;
                            let largeIndex = Math.max(...tblIndex);
                            if (reqEndorsementColHideBrokerId.includes(brokerId)) {
                                tblColumnName = [...tblColumnName, "Document Viewer", "ActionOnDiscrepancy", "Notes", "NotesFreeFill"]
                                tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4];
                            } else {
                                tblColumnName = [...tblColumnName, "Document Viewer", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"]
                                tblIndex = [...tblIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4, largeIndex + 5];
                            }
                            headerSpilttedArray.push({ "Table": `Table ${index + 1}`, index: tblIndex, tblColumnName });
                        }
                        headerSpilttedArray.forEach((table, idx) => {
                            if (table.Table.startsWith('Table ')) {
                                const tableNumber = parseInt(table.Table.replace('Table ', ''), 10);

                                if (!isNaN(tableNumber)) {
                                    if (tableNumber <= 4) {
                                        // For tables up to Table 4
                                        table.tblColumnName[0] = "COVERAGE_SPECIFICATIONS_MASTER";
                                    } else {
                                        // For tables after Table 4
                                        table.tblColumnName[0] = "Coverage_Specifications_Master";
                                    }
                                }
                            }
                        });
                    });
                }
            });

            findTable.map((f, index) => {
                if (index > 0) {
                    let keyValuePair = [];
                    f?.map((e, index1) => {
                        if (index1 > 1) {
                            let tableValueMap = headerSpilttedArray[index - 1];
                            if (tableValueMap) {
                                let object = {};
                                tableValueMap.index.map((i, index2) => {
                                    if (e[i]?.ct?.s && Array.isArray(e[i]?.ct?.s)) {
                                        let arrayS = e[i]?.ct?.s.filter((f) => f != null);
                                        let concatenatedValues = arrayS?.map(item => item?.v)?.join('');
                                        concatenatedValues = concatenatedValues?.replace(/\r\n/g, '~~');
                                        // if (concatenatedValues.endsWith('~~')) {
                                        //     concatenatedValues = concatenatedValues.slice(0, -2);
                                        // }
                                        // concatenatedValues = concatenatedValues.trimEnd();
                                        const finalValue = e[i]?.m || e[i]?.v || concatenatedValues;
                                        object[`${tableValueMap?.tblColumnName[index2]}`] = finalValue;
                                    } else {
                                        object[`${tableValueMap?.tblColumnName[index2]}`] = e[i]?.m || e[i]?.v || e[i]?.ct?.s;
                                    }
                                    if (tableValueMap.index?.length == index2 + 1) {
                                        keyValuePair.push(object);
                                    }
                                })
                            }
                        }
                        if (f?.length == index1 + 1) {
                            tableDataSetArray.push({ "Table": `Table ${index + 1}`, DiscrepancyData: keyValuePair });
                        }
                    });
                }
            });

            let cellDatas = gradtionData;
            let gradtionSheetData = flagCheck == "Red" ? sessionStorage.getItem('redSheetData') : sessionStorage.getItem('greenSheetData');
            let gradiationPropsData = JSON.parse(gradtionSheetData);
            let lobResult = [];
            let uniqueLobSet = new Set();

            if (tableDataSetArray && cellDatas && cellDatas?.length > 0 && tableDataSetArray?.length > 0) {

                gradiationPropsData?.forEach((table) => {
                    if (table?.data && Object.keys(table?.data)?.length > 0) {
                        const policy_Lob = table?.data[0]["POLICY LOB"] || table?.data[0]["Policy LOB"];
                        lobResult.push({
                            Table: table?.TableName,
                            "POLICY LOB": policy_Lob
                        });
                        uniqueLobSet.add(table?.TableName);
                    }
                });
                lobResult = lobResult?.map((item, index) => {
                    return {
                        ...item,
                        Table: `Table ${index + 2}`
                    };
                });

                setLobResult(lobResult);

                let lobMapping = {};
                lobResult.forEach(item => {
                    lobMapping[item.Table] = item["POLICY LOB"];
                });

                // Assign the POLICY LOB value to the corresponding Table in tableDataSetArray
                tableDataSetArray?.forEach(table => {
                    if (lobMapping[table.Table]) {
                        table["POLICY_LOB"] = lobMapping[table.Table];
                        delete table.Table;
                    }
                });

                let documentViewerMap = {};
                let tblStructuredDocumentViewerMap = {};
                gradiationPropsData?.forEach(tbl => {
                    let tableName = tbl?.TableName;
                    if (flagCheck == 'Red') {
                        // Adjust tableName from 'Table' to 'RedTable'
                        if (tableName.startsWith('Table ')) {
                            const numberPart = tableName.split(' ')[1];
                            tableName = `RedTable ${numberPart}`;
                        }
                    }
                    if (flagCheck == 'Green') {
                        // Adjust tableName from 'Table' to 'GreenTable'
                        if (tableName.startsWith('Table ')) {
                            const numberPart = tableName.split(' ')[1];
                            tableName = `GreenTable ${numberPart}`;
                        }
                    }
                    if (tableName != 'RedTable 1' && tableName != 'GreenTable 1') {
                        if (tbl.data && tbl.data.length > 0) {
                            documentViewerMap[tableName] = tbl?.data.map(row => row["Document Viewer"]);
                        }
                    }
                    Object.keys(documentViewerMap).forEach((item, index) => {
                        const tbl = flagCheck == "Red" ? `RedTable ${index + 2}` : `GreenTable ${index + 2}`;;
                        tblStructuredDocumentViewerMap[tbl] = documentViewerMap[item];
                    });
                });

                tableDataSetArray?.forEach((f, index) => {
                    const TableName = flagCheck == 'Red' ? "RedTable " + (index + 2) : "GreenTable " + (index + 2);
                    let data = f?.DiscrepancyData;

                    if (data && data?.length > 0) {
                        data = (data.map((item, index) => {
                            item["policyLob"] = f?.POLICY_LOB;

                            // Replace Document Viewer value if mapping exists
                            if (tblStructuredDocumentViewerMap[TableName] && Array.isArray(tblStructuredDocumentViewerMap[TableName])) {
                                if (index < tblStructuredDocumentViewerMap[TableName].length) {
                                    let docItem = item["Document Viewer"] == " " ? item["Document Viewer"].trim() : item["Document Viewer"];
                                    if (docItem != "") {
                                        item["Document Viewer"] = tblStructuredDocumentViewerMap[TableName][index];
                                    }
                                }
                            }

                            // Replace 'Click here' msg to an empty string
                            if (reqEndorsementColHideBrokerId?.includes(brokerId)) {
                                if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["Notes"] === "~~Click here" && item["NotesFreeFill"] === "~~Click here"
                                ) || (item["ActionOnDiscrepancy"] === "Click here" && item["Notes"] === "Click here" && item["NotesFreeFill"] === "Click here"
                                    ) ||
                                    (item["ActionOnDiscrepancy"] === " Click here" && item["Notes"] === " Click here" && item["NotesFreeFill"] === " Click here"
                                    )) {
                                    item["ActionOnDiscrepancy"] = "";
                                    item["Notes"] = "";
                                    item["NotesFreeFill"] = "";
                                }
                            } else if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["RequestEndorsement"] == "~~Click here"
                                && item["Notes"] == "~~Click here" && item["NotesFreeFill"] == "~~Click here"
                            ) || (item["ActionOnDiscrepancy"] === "Click here" && item["RequestEndorsement"] == "Click here"
                                && item["Notes"] == "Click here" && item["NotesFreeFill"] == "Click here"
                                ) || (item["ActionOnDiscrepancy"] === " Click here" && item["RequestEndorsement"] == " Click here"
                                    && item["Notes"] == " Click here" && item["NotesFreeFill"] == " Click here"
                                )) {
                                item["ActionOnDiscrepancy"] = "";
                                item["RequestEndorsement"] = "";
                                item["Notes"] = "";
                                item["NotesFreeFill"] = "";
                            } else {
                                ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"].forEach(key => {
                                    if ((item[key] === "~~Click here") || (item[key] === "Click here") || (item[key] === " Click here")) {
                                        item[key] = "";
                                    }
                                })
                            }
                            return item;
                        }));
                    }

                    const cleanedData = data?.filter(entry =>
                        !Object.values(entry)?.includes(undefined)
                    );

                    if (flagCheck == 'Red') {
                        redSheetData.push({ TableName, Data: JSON.stringify(cleanedData) });
                    } else if (flagCheck == 'Green') {
                        greenSheetData.push({ TableName, Data: JSON.stringify(cleanedData) });
                    }
                });
            }
        } else if (flagCheck == "Exclusion") {
            let headerSpilttedArray = [];
            let mergingArrayData = [];
            let alldatas = luckysheet.getAllSheets();
            let filterdatasheet = alldatas?.filter(f => f.name?.includes("Exclusion"));
            if (filterdatasheet.length != undefined && filterdatasheet.length > 0) {
                let sheetData = filterdatasheet[0].data;
                let limitReached = false;
                sheetData && sheetData?.length > 0 && sheetData.map((e, index) => {
                    if (index == 0 && !limitReached) {
                        let tblIndex = [];
                        let tblColumnName = [];
                        let hasReachedLimit = false;
                        const filteredData = e.filter(item => item !== null);
                        filteredData?.forEach((e1, index1) => {
                            if (index1 >= 0 && !hasReachedLimit && e1?.m != "Actions on Discrepancy (from AMs)") {
                                tblIndex.push(index1);
                                tblColumnName.push(e1?.m || e1?.v);
                            }
                        });

                        tblColumnName = tblColumnName.filter(column => column !== undefined);
                        let slicedIndex = tblIndex.slice(0, tblColumnName.length)
                        let largeIndex = Math.max(...slicedIndex);
                        if (reqEndorsementColHideBrokerId?.includes(brokerId)) {
                            tblColumnName = [...tblColumnName, "ActionOnDiscrepancy", "Notes", "NotesFreeFill"]
                            slicedIndex = [...slicedIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3];
                        } else {
                            tblColumnName = [...tblColumnName, "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"]
                            slicedIndex = [...slicedIndex, largeIndex + 1, largeIndex + 2, largeIndex + 3, largeIndex + 4];
                        }
                        slicedIndex.slice(0, tblColumnName.length);
                        headerSpilttedArray.push({ "Table": `ExTable ${index + 1}`, index: slicedIndex, tblColumnName });
                        hasReachedLimit = true;
                    }
                });

                sheetData && sheetData?.length > 0 && sheetData.forEach((e, index) => {
                    let keyValuePair = [];
                    if (index > 1) {
                        let tableValueMap = headerSpilttedArray[0];
                        if (tableValueMap) {
                            let object = {};
                            tableValueMap?.index?.forEach((i, index2) => {
                                if (e[i]?.ct?.s && Array.isArray(e[i]?.ct?.s)) {
                                    let filteredS = e[i]?.ct?.s.filter((f) => f != null);
                                    let concatenatedValues = filteredS?.map(item => item?.v)?.join('');
                                    concatenatedValues = concatenatedValues?.replace(/\r\n/g, ' ');
                                    concatenatedValues = concatenatedValues?.trim();
                                    const finalValue = e[i]?.m || e[i]?.v || concatenatedValues;
                                    object[`${tableValueMap?.tblColumnName[index2]}`] = finalValue;
                                } else {
                                    object[`${tableValueMap?.tblColumnName[index2]}`] = e[i]?.m || e[i]?.v || e[i]?.ct?.s;
                                }
                            })
                            keyValuePair.push(object);
                        }
                        mergingArrayData.push({ DiscrepancyData: keyValuePair });
                    }
                });
                let combineDiscrepancyTempData = [];

                mergingArrayData?.forEach(item => {
                    const newTemplateData = item?.DiscrepancyData?.filter(data => {
                        return Object.values(data).some(value => value !== undefined);
                    });
                    if (newTemplateData?.length > 0) {
                        combineDiscrepancyTempData = combineDiscrepancyTempData.concat(newTemplateData);
                    }
                });

                if (combineDiscrepancyTempData?.length > 0) {
                    exclusionSetArray = {
                        "DiscrepancyData": combineDiscrepancyTempData
                    };
                }

                let data = exclusionSetArray?.DiscrepancyData;

                if (data && data?.length > 0) {
                    data = (data.map((item, index) => {

                        // Replace 'Click here' msg to an empty string
                        if (reqEndorsementColHideBrokerId?.includes(brokerId)) {
                            if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["Notes"] === "~~Click here" && item["NotesFreeFill"] === "~~Click here"
                            ) || (item["ActionOnDiscrepancy"] === "Click here" && item["Notes"] === "Click here" && item["NotesFreeFill"] === "Click here"
                                ) ||
                                (item["ActionOnDiscrepancy"] === " Click here" && item["Notes"] === " Click here" && item["NotesFreeFill"] === " Click here"
                                )) {
                                item["ActionOnDiscrepancy"] = "";
                                item["Notes"] = "";
                                item["NotesFreeFill"] = "";
                            }
                        } else if ((item["ActionOnDiscrepancy"] === "~~Click here" && item["RequestEndorsement"] == "~~Click here"
                            && item["Notes"] == "~~Click here" && item["NotesFreeFill"] == "~~Click here"
                        ) || (item["ActionOnDiscrepancy"] === "Click here" && item["RequestEndorsement"] == "Click here"
                            && item["Notes"] == "Click here" && item["NotesFreeFill"] == "Click here"
                            ) || (item["ActionOnDiscrepancy"] === " Click here" && item["RequestEndorsement"] == " Click here"
                                && item["Notes"] == " Click here" && item["NotesFreeFill"] == " Click here"
                            )) {
                            item["ActionOnDiscrepancy"] = "";
                            item["RequestEndorsement"] = "";
                            item["Notes"] = "";
                            item["NotesFreeFill"] = "";
                        } else {
                            ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"].forEach(key => {
                                if ((item[key] === "~~Click here") || (item[key] === "Click here") || (item[key] === " Click here")) {
                                    item[key] = "";
                                }
                            })
                        }
                        return item;
                    }));
                }
                if (flagCheck === 'Exclusion') {
                    exclusionSheetData.push({
                        TableName: "ExclusionTable",
                        Data: JSON.stringify(exclusionSetArray["DiscrepancyData"])
                    });
                }
            }
        } else if (flagCheck == "QAC not answered questions") {
            const keys = Object.keys(qacTableColumnDetails);
            const tableDataSetArray = [];
            if (keys?.length > 0) {
                keys?.forEach((key) => {
                    const slicedTables = [];
                    const tableData = qacTableColumnDetails[key];
                    const structuredTblObj = qacTblRangeStructureFun(tableData);

                    Object.keys(structuredTblObj)?.forEach((tableKey) => {
                        const table = structuredTblObj[tableKey];
                        if (table?.range?.start !== undefined && table?.range?.end !== undefined) {
                            const slicedData = propsData?.slice(table?.range?.start, table?.range?.end + 1);
                            slicedTables.push(slicedData);
                        }
                    });


                    const headerSpilttedArray = slicedTables?.map((tableData, index) => {
                        if (tableData[1]) {
                            const tblIndex = [];
                            const tblColumnName = [];
                            tableData[1]?.forEach((cell, cellIndex) => {
                                if (cell) {
                                    tblIndex.push(cellIndex);
                                    tblColumnName.push(cell?.m || cell?.v);
                                }
                            });

                            return {
                                index: tblIndex?.slice(0, tblColumnName.length),
                                tblColumnName: tblColumnName?.filter(Boolean),
                            };
                        }
                        return null;
                    }).filter(Boolean);

                    slicedTables?.forEach((tableData, index) => {
                        const keyValuePair = [];
                        tableData?.forEach((row, rowIndex) => {
                            if (rowIndex > 1) {
                                const tableValueMap = headerSpilttedArray[index];
                                if (tableValueMap) {
                                    const rowObject = {};
                                    tableValueMap?.index?.forEach((i, colIndex) => {
                                        const cell = row[i];
                                        const finalValue = cell?.m || cell?.v || cell?.ct?.s?.map((s) => s.v)?.join('') || "";
                                        rowObject[tableValueMap.tblColumnName[colIndex]] = finalValue;
                                    });
                                    keyValuePair.push(rowObject);
                                }
                            }
                            if (rowIndex === 0) {
                                const policyLobValue = row?.[0]?.m || "";
                                keyValuePair.push({
                                    policyLob: policyLobValue,
                                    IsHighVolume: key?.toLowerCase()?.includes("high") ? true : false
                                });
                            }
                        });

                        tableDataSetArray.push({
                            Data: keyValuePair,
                            TableName: key?.toLowerCase()?.includes("high") ? `HighVolumeTable${index + 1}` : `LowVolumeTable${index + 1}`
                        });
                    });
                });

                let updatedTableDataSetArray = tableDataSetArray?.map(({ Data, TableName }) => {
                    const { policyLob, IsHighVolume } = Data[0];
                    return {
                        TableName,
                        Data: Data?.slice(1).map((item) => ({
                            ...item,
                            policyLob,
                            IsHighVolume,
                        })),
                    };
                });

                if (flagCheck == 'QAC not answered questions') {
                    const filteredData = updatedTableDataSetArray?.filter(item =>
                        item?.Data?.some(
                            entry =>
                                entry["Discrepancies to note on Intranet Form:"] !== "" ||
                                entry["Checked By (Initials):"] !== ""
                        )
                    );

                    let highVolumeCount = 1;
                    let lowVolumeCount = 1;

                    const reseededTblData = filteredData?.map((table) => {
                        if (table.Data[0].IsHighVolume) {
                            table.TableName = `HighVolumeTable${highVolumeCount}`;
                            highVolumeCount++;
                        } else {
                            table.TableName = `LowVolumeTable${lowVolumeCount}`;
                            lowVolumeCount++;
                        }
                        return table;
                    });

                    reseededTblData?.forEach(({ TableName, Data }) => {
                        qacSheetDataSet.push({
                            TableName,
                            Data: JSON.stringify(Data),
                        });
                    });
                }
            }
        }

        if (flagCheck != 'Exclusion' && flagCheck != 'QAC not answered questions') {
            if (findTable && findTable?.length > 0) {
                const result = findTable[0]?.map(([index1, index2]) => ({
                    index1,
                    index2,
                }));
                const csrTable1Data = [];
                for (let row in result) {
                    let jobid = props?.selectedJob;
                    const cellData = result[row];
                    const inputData = flagCheck == 'PolicyReviewChecklist' ? csrPolicyData : flagCheck == 'Forms Compare' ? formstate : gradtionData;
                    const csrTable1data = flagCheck == 'PolicyReviewChecklist' ? inputData.find((data) => data.Tablename === "CsrTable 1")
                        : flagCheck == 'Forms Compare' ? inputData.find((data) => data.Tablename === "FormTable 1")
                            : inputData.find((data) => data.Tablename === "Table 1");

                    if (csrTable1data && csrTable1data?.TemplateData?.length > 0) {
                        const csrTable1json = flagCheck == 'PolicyReviewChecklist' ? csrTable1data.TemplateData
                            : flagCheck == 'Forms Compare' ? JSON.parse(csrTable1data.TemplateData) : csrTable1data.TemplateData;

                        const lobValue = csrTable1json.map(item => item["Policy LOB"]);
                        const cell1Text = cellData.index1?.v || cellData.index1?.ct || '';
                        const cell2Text = cellData.index2?.v !== undefined ? (cellData.index2.v || cellData.index2?.ct) : (cellData.index2?.ct?.fa === "@" ? "" : cellData.index2?.ct?.fa);

                        if (cell1Text && cell1Text.s && cell1Text.s.length > 0 && cell2Text || cell1Text) {
                            const headerValue = Array.isArray(cell1Text.s) ? cell1Text.s.map(item => item.v || '').join(',') : cell1Text;
                            const concatenatedValues = Array.isArray(cell2Text.s) ? cell2Text.s.map(item => item.v).join('') : cell2Text;

                            const objStructure = {
                                HeaderID: row,
                                JOBID: jobid,
                                'Policy LOB': lobValue[0],
                                Headers: headerValue,
                                '(No column name)': concatenatedValues,
                            }

                            const handleNonwhitespaceCharacter = (value) => {
                                if (typeof value === 'string') {
                                    return value.replace(/\n/g, '~~').replace(/"/g, '\\"');
                                }
                                return value;
                            };

                            const jsonString = `{${Object.entries(objStructure).map(([key, value]) => {
                                if (key === 'HeaderID') {
                                    const updatedValue = Number(value).toString();
                                    return `"${key}":${updatedValue}`;
                                } else if (key === '(No column name)') {
                                    return `"${key}":"${handleNonwhitespaceCharacter(value)}"`;
                                } else if (key === 'Headers') {
                                    const sValue = Array.isArray(value) ? `"${handleNonwhitespaceCharacter(value.join(', '))}"` : `"${handleNonwhitespaceCharacter(value)}"`;
                                    return `"${key}":${sValue}`;
                                }
                                return `"${key}":"${handleNonwhitespaceCharacter(value)}"`;
                            }).join(',')}}`;
                            csrTable1Data.push(jsonString);
                        }
                    }
                }
                const Json = `[${csrTable1Data.join(',')}]`;
                const sanitizejson = Json?.replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
                const parsedTable1Data = JSON.parse(sanitizejson);
                if (flagCheck == 'PolicyReviewChecklist' || flagCheck == 'Red' || flagCheck == 'Green') {
                    const tablestring = flagCheck == 'PolicyReviewChecklist' ? "CsrTable 1" : flagCheck == 'Red' ? "RedTable 1" : "GreenTable 1";
                    const transformedJson = [
                        {
                            "TableName": tablestring,
                            "Data": JSON.stringify(parsedTable1Data)
                        }
                    ];
                    const findingTable2Index = flagCheck == 'PolicyReviewChecklist' ? policySheetData?.findIndex(table => table.TableName === "CsrTable 2")
                        : flagCheck == 'Red' ? redSheetData?.findIndex(table => table.TableName === "RedTable 2")
                            : flagCheck == 'Green' ? greenSheetData?.findIndex(table => table.TableName === "GreenTable 2") : "";

                    if (flagCheck == 'PolicyReviewChecklist' && findingTable2Index !== -1) {
                        policySheetData.splice(findingTable2Index, 0, transformedJson[0]);
                        // console.log(policySheetData);
                    }
                    if (flagCheck == 'Red') {
                        if (findingTable2Index !== -1) {
                            redSheetData.splice(findingTable2Index, 0, transformedJson[0]);
                            // console.log(redSheetData);
                        } else {
                            redSheetData.push(transformedJson[0]);
                        }
                    }
                    if (flagCheck == 'Green') {
                        if (findingTable2Index !== -1) {
                            greenSheetData.splice(findingTable2Index, 0, transformedJson[0]);
                            // console.log(greenSheetData);
                        } else {
                            greenSheetData.push(transformedJson[0]);
                        }
                    }
                } else if (flagCheck == 'Forms Compare') {
                    const transformedJson = [
                        {
                            "TableName": "CsrFormTable 1",
                            "Data": JSON.stringify(parsedTable1Data)
                        }
                    ];
                    const findingTable2Index = formCompareSheetData?.findIndex(table => table.TableName === "CsrFormTable 2");
                    if (findingTable2Index !== -1) {
                        formCompareSheetData.splice(findingTable2Index, 0, transformedJson[0]);
                        // console.log(formCompareSheetData);
                    }
                }
            }
        }

        if (isExport == true) {
            const sheetDatas = flagCheck == "PolicyReviewChecklist" ? policySheetData : flagCheck == "Forms Compare" ? formCompareSheetData : flagCheck == "Exclusion" ? exclusionSheetData : flagCheck == "Red" ? redSheetData : flagCheck == "Green" ? greenSheetData : qacSheetDataSet;

            sheetDatas.forEach(item => {
                const sanitizedData = item?.Data  //sanitize the JSON string by removing any problematic control characters before parsing it.  so dont remove this
                let parsedData;
                try {
                    parsedData = sanitizedData;
                } catch (error) {
                    // console.error("parsing json error catch:", error);
                    return;
                }
                // parsedData.forEach(obj => {
                //     if (obj[""] !== undefined) {
                //         obj["NoColumnName"] = obj[""];
                //         delete obj[""];
                //     }
                // });
                // item.Data = JSON.stringify(parsedData);
            });

            const modifiedTabledata = sheetDatas?.map(item => ({
                Id: props?.selectedJob,
                TableName: item?.TableName,
                Data: item.Data
            }));

            const dataFromOnUpdateClick = modifiedTabledata;

            if (typeof callback === "function") {
                callback(dataFromOnUpdateClick);
            }
        } else {
            // const response = csrSheetPolicyGradiationDataSaveApi(id, jobid, CheckListType, wholeData);
            // return response;
        }
    };

    const staticDataforQacExport = () => {
        const qacSheetDataSet = [];
        const staticSheetData = luckysheet.getAllSheets()?.filter(f => f.name.includes("QAC not answered questions"))[0]?.data;
            const keys = Object.keys(qacTableColumnDetails);

            const tableDataSetArray = [];
            if (keys?.length > 0) {
                keys?.forEach((key) => {
                    const slicedTables = [];

                    const tableData = qacTableColumnDetails[key];
                    const structuredTblObj = qacTblRangeStructureFun(tableData);

                    Object.keys(structuredTblObj)?.forEach((tableKey) => {
                        const table = structuredTblObj[tableKey];
                        if (table?.range?.start !== undefined && table?.range?.end !== undefined) {
                            const slicedData = staticSheetData?.slice(table?.range?.start, table?.range?.end + 1);
                            slicedTables.push(slicedData);
                        }
                    });


                    const headerSpilttedArray = slicedTables?.map((tableData, index) => {
                        if (tableData[1]) {
                            const tblIndex = [];
                            const tblColumnName = [];
                            tableData[1]?.forEach((cell, cellIndex) => {
                                if (cell) {
                                    tblIndex.push(cellIndex);
                                    tblColumnName.push(cell?.m || cell?.v);
                                }
                            });

                            return {
                                index: tblIndex?.slice(0, tblColumnName.length),
                                tblColumnName: tblColumnName?.filter(Boolean),
                            };
                        }
                        return null;
                    }).filter(Boolean);

                    slicedTables?.forEach((tableData, index) => {
                        const keyValuePair = [];
                        tableData?.forEach((row, rowIndex) => {
                            if (rowIndex > 1) {
                                const tableValueMap = headerSpilttedArray[index];
                                if (tableValueMap) {
                                    const rowObject = {};
                                    tableValueMap?.index?.forEach((i, colIndex) => {
                                        const cell = row[i];
                                        const finalValue = cell?.m || cell?.v || cell?.ct?.s?.map((s) => s.v)?.join('') || "";
                                        rowObject[tableValueMap.tblColumnName[colIndex]] = finalValue;
                                    });
                                    keyValuePair.push(rowObject);
                                }
                            }
                            if (rowIndex === 0) {
                                const policyLobValue = row?.[0]?.m || "";
                                keyValuePair.push({
                                    policyLob: policyLobValue,
                                    IsHighVolume: key?.toLowerCase()?.includes("high") ? true : false
                                });
                            }
                        });

                        tableDataSetArray.push({
                            Data: keyValuePair,
                            TableName: key?.toLowerCase()?.includes("high") ? `HighVolumeTable${index + 1}` : `LowVolumeTable${index + 1}`
                        });
                    });
                });

                let updatedTableDataSetArray = tableDataSetArray?.map(({ Data, TableName }) => {
                    const { policyLob, IsHighVolume } = Data[0];
                    return {
                        TableName,
                        Data: Data?.slice(1).map((item) => ({
                            ...item,
                            policyLob,
                            IsHighVolume,
                        })),
                    };
                });

                    updatedTableDataSetArray?.forEach(({ TableName, Data }) => {
                        qacSheetDataSet.push({
                            Id:props?.selectedJob,
                            TableName,
                            Data: JSON.stringify(Data),
                        });
                    });

            }
            return qacSheetDataSet;
    };

    const csrSheetPolicyGradiationDataSaveApi = async (id, jobId, CheckListType, dataForSaveApi, isAutoUpdate) => {
        if (!isAutoUpdate) {
            document.body.classList.add('loading-indicator');
        }
        const Token = await processAndUpdateToken(token);

        let message;
        let log;
        if(CheckListType === "PreviewCheckList") {
            message = isAutoUpdate ? "CSR-AOD-AutoUpdate-PolicyCheck-initiated" : "CSR-AOD-Update-PolicyCheck-initiated";
            log = isAutoUpdate ? "CSR-AutoSave-PolicyCheck" : "CSR-Save-PolicyCheck";
        } else if(CheckListType === "GradiationSheet") {
            message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Gradiation-initiated" : "CSR-AOD-Update-Gradiation-initiated";
            log = isAutoUpdate ? "CSR-AutoSave-Gradiation" : "CSR-Save-Gradiation";
        }
        
        updateGridAuditLog(jobId, message, log, (userName || sessionUserName));
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${baseUrl}/api/ProcedureData/AddGridAoddMapping`;
        try {
            const response = await axios({
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    Id: id,
                    JobId: jobId,
                    CheckListType: CheckListType,
                    Data: dataForSaveApi,
                }
            });
            if (response.status !== 200) {
                if(CheckListType === "PreviewCheckList") {
                    message = isAutoUpdate ? "CSR-AOD-AutoUpdate-PolicyCheck-failed" : "CSR-AOD-Update-PolicyCheck-failed";
                } else if (CheckListType === "GradiationSheet"){
                    message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Gradiation-failed" : "CSR-AOD-Update-Gradiation-failed";
                }
                updateGridAuditLog(jobId, message, typeof response === 'object' ? JSON.stringify(response) : response, (userName || sessionUserName));
                return "error";
            }

            return response.data;
        } catch (error) {
            if(CheckListType === "PreviewCheckList") {
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-PolicyCheck-failed" : "CSR-AOD-Update-PolicyCheck-failed";
            } else if (CheckListType === "GradiationSheet"){
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Gradiation-failed" : "CSR-AOD-Update-Gradiation-failed";
            }
            updateGridAuditLog(jobId, message, typeof error === 'object' ? JSON.stringify(error) : error, (userName || sessionUserName));
            return "error";
        } finally {
            if(CheckListType === "PreviewCheckList") {
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-PolicyCheck-completed" : "CSR-AOD-Update-PolicyCheck-completed";
                log = isAutoUpdate ? "CSR-AutoSave-PolicyCheck-Success" : "CSR-Save-PolicyCheck-Success";
            } else if(CheckListType === "GradiationSheet") {
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Gradiation-completed" : "CSR-AOD-Update-Gradiation-completed";
                log = isAutoUpdate ? "CSR-AutoSave-Gradiation-Success" : "CSR-Save-Gradiation-Success";
            }
            updateGridAuditLog(jobId, message, log, (userName || sessionUserName));
            if (!isAutoUpdate) {
                document.body.classList.remove('loading-indicator');
            }
            document.body.classList.remove('loading-indicator');
            return "success";
        }
    };

    const qacSheetSaveCall = async (jobId, dataForSaveApi, isAutoUpdate) => {
        if (!isAutoUpdate) {
            document.body.classList.add('loading-indicator');
        }
        document.body.classList.add('loading-indicator');
        const Token = await processAndUpdateToken(token);
        const message = isAutoUpdate ? "CSR-QacChecklistData-AutoUpdate-initiated" : "CSR-QacChecklistData-Update-initiated";
        const log = isAutoUpdate ? "CSR-QacChecklistData-AutoSave" : "CSR-QacChecklistData-Save";
        updateGridAuditLog(jobId, message, log, (userName || sessionUserName));
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${baseUrl}/api/Defaultdatum/AddOrUpdateQacChecklistData`;
        try {
            const response = await axios({
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    JobId: jobId,
                    JsonData: dataForSaveApi,
                    UpdatedBy:sessionStorage.getItem("csrAuthUserName")
                }
            });
            if (response.status !== 200) {
                const message = isAutoUpdate ? "CSR-QacChecklistData-AutoUpdate-failed" : "CSR-QacChecklistData-Update-failed";
                updateGridAuditLog(jobId, message, typeof response === 'object' ? JSON.stringify(response) : response, (userName || sessionUserName));
                return "error";
            }

            return response.data;
        } catch (error) {
            const message = isAutoUpdate ? "CSR-QacChecklistData-AutoUpdate-failed" : "CSR-QacChecklistData-Update-failed";
            updateGridAuditLog(jobId, message, typeof error === 'object' ? JSON.stringify(error) : error, (userName || sessionUserName));
            return "error";
        } finally {
            const message = isAutoUpdate ? "CSR-QacChecklistData-AutoUpdate-completed" : "CSR-QacChecklistData-Update-completed";
            const log = isAutoUpdate ? "CSR-AutoSave-QacChecklistData-Success" : "CSR-Save-QacChecklistData-Success";
            container.current.showSnackbar('QacChecklistData updated successfully', "success", true);
            updateGridAuditLog(jobId, message, log, (userName || sessionUserName));
            if (!isAutoUpdate) {
                document.body.classList.remove('loading-indicator');
            }
            document.body.classList.remove('loading-indicator');
            return "success";
        }
    };

    const csrSheetCommonSaveApi = async (jobId, tableName, dataForSaveApi, isAutoUpdate) => {
        if (!isAutoUpdate) {
            document.body.classList.add('loading-indicator');
        }
        document.body.classList.add('loading-indicator');
        const Token = await processAndUpdateToken(token);

        let message;
        let log;
        if(tableName === "FormTable") {
            message = isAutoUpdate ? "CSR-AOD-AutoUpdate-FormsCompare-initiated" : "CSR-AOD-Update-FormsCompare-initiated";
            log = isAutoUpdate ? "CSR-AutoSave-FormsCompare" : "CSR-Save-FormsCompare";
        } else if(tableName === "ExclusionTable") {
            message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Exclusion-initiated" : "CSR-AOD-Update-Exclusion-initiated";
            log = isAutoUpdate ? "CSR-AutoSave-Exclusion" : "CSR-Save-Exclusion";
        }
        updateGridAuditLog(jobId, message, log, (userName || sessionUserName));
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${baseUrl}/api/ProcedureData/UpdatecsrFormTable`;
        try {
            const response = await axios({
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    JobId: jobId,
                    TableName: tableName,
                    // CheckListType: CheckListType,
                    NewTemplateData: dataForSaveApi,
                }
            });
            if (response.status !== 200) {
                if(tableName === "FormTable") {
                    message = isAutoUpdate ? "CSR-AOD-AutoUpdate-FormsCompare-failed" : "CSR-AOD-Update-FormsCompare-failed";
                } else if (tableName === "ExclusionTable"){
                    message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Exclusion-failed" : "CSR-AOD-Update-Exclusion-failed";
                }
                updateGridAuditLog(jobId, message, typeof response === 'object' ? JSON.stringify(response) : response, (userName || sessionUserName));
                return "error";
            }

            return response.data;
        } catch (error) {
            if(tableName === "FormTable") {
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-FormsCompare-failed" : "CSR-AOD-Update-FormsCompare-failed";
            } else if (tableName === "ExclusionTable"){
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Exclusion-failed" : "CSR-AOD-Update-Exclusion-failed";
            }
            updateGridAuditLog(jobId, message, typeof error === 'object' ? JSON.stringify(error) : error, (userName || sessionUserName));
            return "error";
        } finally {
            if(tableName === "FormTable") {
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-FormsCompare-completed" : "CSR-AOD-Update-FormsCompare-completed";
                log = isAutoUpdate ? "CSR-AutoSave-FormsCompare-Success" : "CSR-Save-FormsCompare-Success";
            } else if(tableName === "ExclusionTable") {
                message = isAutoUpdate ? "CSR-AOD-AutoUpdate-Exclusion-completed" : "CSR-AOD-Update-Exclusion-completed";
                log = isAutoUpdate ? "CSR-AutoSave-Exclusion-Success" : "CSR-Save-Exclusion-Success";
            }
            updateGridAuditLog(jobId, message, log, (userName || sessionUserName));
            document.body.classList.remove('loading-indicator');
            return "success";
        }
    };

    const handleYesDailog = () => {
        setYesDialog(false);
    }

    const toggleYesDialog = async () => {
        setYesDialog(!yesDialog);
        const filterAllSheets = luckysheet.getAllSheets().filter(f => f.name);
        const policySheet = filterAllSheets.find(sheet => sheet.name === "PolicyReviewChecklist");
        const formSheet = filterAllSheets.find(sheet => sheet.name === "Forms Compare");
        const exclusionSheet = filterAllSheets.find(sheet => sheet.name === "Exclusion");
        const redSheet = filterAllSheets.find(sheet => sheet.name === "Red");
        const greenSheet = filterAllSheets.find(sheet => sheet.name === "Green");

        const policyData = policySheet ? policySheet.data : null;
        const formData = formSheet ? formSheet.data : null;
        const exclusionData = exclusionSheet ? exclusionSheet.data : null;
        const redData = redSheet ? redSheet.data : null;
        const greenData = greenSheet ? greenSheet.data : null;

        let sheetCheck = luckysheet?.getSheet().name;
        if (policyData?.length > 0 && sheetCheck == "PolicyReviewChecklist") {
            setsheetState(policyData);
        } else if (formData?.length > 0 && sheetCheck == "Forms Compare") {
            setsheetState(formData);
        } else if (exclusionData?.length > 0 && sheetCheck == "Exclusion") {
            setsheetState(exclusionData);
        } else if (redData?.length > 0 && greenData?.length > 0) {
            const combinedData = {
                redData,
                greenData
            };
            setsheetState(combinedData);
        }
    };

    const matchedOrUnMatchedFilter = (rowIndex) => {
        if (rowIndex === 6 || rowIndex === 4 || rowIndex === 5) {
            let checklistData = [...props?.data];
            if (checklistData && checklistData?.length > 0) {
                checklistData = checklistData.map((e) => {
                    if (e?.TemplateData && typeof e?.TemplateData != 'object' && typeof e?.TemplateData === 'string') {
                        let templateData = JSON.parse(e.TemplateData);
                        e["TemplateData"] = templateData;
                    }
                    return e;
                })
            }
            const tableDetails = tableColumnDetails;
            let keys = Object.keys(tableDetails)?.filter((f) => f !== "Table 1");
            let recordsToHide = [];
            keys.forEach((f) => {
                const tableConfigData = tableDetails[f];
                const columnKeys = tableConfigData?.columnNames;
                const recordRange = tableConfigData?.range;
                const sourceColumns = Object.keys(columnKeys).filter((key) => columnKeys[key] > columnKeys["COVERAGE_SPECIFICATIONS_MASTER"] &&
                    columnKeys[key] < columnKeys["Document Viewer"]);
                if (sourceColumns && sourceColumns?.length > 0) {
                    const findData = checklistData.find((fi) => fi?.Tablename == f);
                    if (findData?.TemplateData && findData?.TemplateData?.length > 0) {
                        findData?.TemplateData.forEach((item, itemIndex) => {
                            if (rowIndex === 5) { //for unmatched(variance)
                                let needToHide = false;
                                let needToHideInCount = 0;
                                sourceColumns.forEach((srItem) => {
                                    const srCData = item[srItem];
                                    if (!needToHide && srCData && srCData?.trim()?.toLowerCase() == "matched") {
                                        needToHide = true;
                                    } else {
                                        if (srCData && srCData?.trim()?.toLowerCase() == "details not available in the document") {
                                            needToHideInCount = (needToHideInCount + 1);
                                        }
                                    }
                                });
                                if (needToHide || (needToHideInCount === sourceColumns?.length)) {
                                    recordsToHide.push(recordRange?.start + itemIndex + (f === "Table 3" ? 4 : 3));
                                }
                            } else if (rowIndex === 4) {
                                let needToHide = true;
                                sourceColumns.forEach((srItem) => {
                                    const srCData = item[srItem];
                                    if (needToHide && srCData && srCData?.trim()?.toLowerCase() == "matched") {
                                        needToHide = false;
                                    }
                                });
                                if (needToHide) {
                                    recordsToHide.push(recordRange?.start + itemIndex + (f === "Table 3" ? 4 : 3));
                                }
                            } else if (rowIndex === 6) {
                                let needToHide = 0;
                                sourceColumns.forEach((srItem) => {
                                    const srCData = item[srItem];
                                    if (srCData && srCData?.trim()?.toLowerCase() == "details not available in the document") {
                                        needToHide = (needToHide + 1);
                                    }
                                });
                                if (needToHide != sourceColumns?.length) {
                                    recordsToHide.push(recordRange?.start + itemIndex + (f === "Table 3" ? 4 : 3));
                                }
                            }
                        });
                    }
                }
            });
            if (recordsToHide?.length > 0) {
                recordsToHide = groupNumbers(recordsToHide);
            }
            showOrHideRecords(recordsToHide);
            let topConfig = 0;
            $("#luckysheet-scrollbar-y").scrollTop(topConfig + 400);
        } else {
            showOrHideRecords([]);
            let topConfig = 0;
            $("#luckysheet-scrollbar-y").scrollTop(topConfig + 400);
            if (luckysheet?.getSheet()?.name != 'QAC not answered questions' && luckysheet?.getSheet()?.name != "Exclusion") {
                dataGrouping(0);
            }
        }
    }

    const handleIconClick = () => {
        const currentSheetData = luckysheet.getSheet();
        if (currentSheetData?.name === 'PolicyReviewChecklist' || currentSheetData?.name === 'Red' || currentSheetData?.name === 'Green') {
            toggleFilterDialog();
        }
    }

    const toggleFilterDialog = () => {
        setOpenFilterDialog(!openFilterDialog);
    };

    const handleFilterDialogClose = (e) => {
        setOpenFilterDialog(false);
        if (e?.filterData?.selectedOption1 && e?.filterData?.selectedOption2) {
            setFilterSelectionData(e?.filterData);
        } else {
            setFilterSelectionData(null);
            if (luckysheet?.getSheet()?.name != 'QAC not answered questions' && luckysheet?.getSheet()?.name != "Exclusion") {
                dataGrouping(0);
            }
        }
    };

    function iconvaluestorage(newValue) {
        let existingValues = JSON.parse(localStorage.getItem('IconShownIndex')) || [];
        existingValues.push(newValue);
        localStorage.setItem('IconShownIndex', JSON.stringify(existingValues));
    };

    const dataGrouping = (ToBeShownIndex) => {
    if (luckysheet?.getSheet()?.name != 'QAC not answered questions' && luckysheet?.getSheet()?.name != "Exclusion") {
        localStorage.setItem('selectedgrouping', ToBeShownIndex);
        if (ToBeShownIndex != undefined) {
            iconvaluestorage(ToBeShownIndex);
        }
        let colRecordTect = '';
        if (ToBeShownIndex > 0) {
            const rowRecord = luckysheet.getcellvalue(ToBeShownIndex);
            const columnRecord = rowRecord[0];
            colRecordTect = getText(columnRecord, false);
            if (colRecordTect != "-" && colRecordTect != "+") {
                return;
            }
        }
        let checklistData = [...props?.data];
        if (checklistData && checklistData?.length > 0) {
            checklistData = checklistData?.map((e) => {
                if (e?.TemplateData && typeof e?.TemplateData != 'object' && typeof e?.TemplateData === 'string') {
                    let templateData = JSON.parse(e.TemplateData);
                    e["TemplateData"] = templateData;
                }
                return e;
            })
        }

        const tableDetails = tableColumnDetails;

        let keys = Object.keys(tableDetails)?.filter((f) => f !== "Table 1");
        let recordsToHide = [];
        keys?.forEach((f) => {
            const tableConfigData = tableDetails[f];
            const recordRange = tableConfigData?.range;
            const shortQuestion = [];
            const findData = checklistData?.find((fi) => fi?.Tablename == f);
            if (findData && findData?.TemplateData?.length > 0) {
                const backUpTemplateData = findData?.TemplateData;
                findData?.AvailableLobs?.forEach((lob) => {
                    backUpTemplateData?.forEach((item) => {
                        if (item["POLICY LOB"] && item["POLICY LOB"]?.toUpperCase()?.includes(lob.toUpperCase())) {
                            f = "Table 3";
                        }
                    });
                });
                let needIgnorance = false;
                const ignoranceShortCode = [];
                if (ToBeShownIndex > 0 && recordRange?.start < ToBeShownIndex && recordRange?.end >= ToBeShownIndex) {
                    needIgnorance = true;
                }
                const ignoreQCSet = [];
                findData?.TemplateData?.forEach((item, itemIndex) => {
                    let questionText = item["Checklist Questions"];
                    if (questionText) {
                        questionText = questionText?.toUpperCase()?.trim();
                        const questionCode = questionText?.substring(0, 3);

                        let isIgnorancecheck = false;
                        if (needIgnorance) {
                            isIgnorancecheck = ((recordRange?.start + itemIndex + (f == "Table 3" ? 3 : 2)) === ToBeShownIndex) && colRecordTect == "+";
                            if (isIgnorancecheck) {

                                ignoranceShortCode.push(questionCode);
                            }
                        } else if (ToBeShownIndex && ToBeShownIndex > 0) {
                            // ignoreQCSet, isIgnorancecheck
                            if (questionCode && !ignoreQCSet?.includes(questionCode)) {
                                let colRecordText1 = '';
                                const rowRecord1 = luckysheet.getcellvalue(recordRange?.start + itemIndex + (f == "Table 3" ? 3 : 2));
                                const columnRecord1 = rowRecord1[0];
                                colRecordText1 = getText(columnRecord1, false);
                                if (colRecordText1 == "-") {
                                    ignoreQCSet.push(questionCode);
                                }
                            }
                        }
                        if (shortQuestion?.includes(questionCode)) {
                            if ((!ignoranceShortCode || (ignoranceShortCode && ignoranceShortCode?.length === 0)) && (ignoreQCSet?.length === 0)) {
                                recordsToHide.push(recordRange?.start + itemIndex + (f === "Table 3" ? 4 : 3));
                            }
                            else if (!ignoranceShortCode?.includes(questionCode) && !ignoreQCSet?.includes(questionCode)) {
                                recordsToHide.push(recordRange?.start + itemIndex + (f === "Table 3" ? 4 : 3));
                            }
                        } else {
                            const hasMoreThanOne = backUpTemplateData.filter((qcF) => {
                                const cqTextForFilter = qcF["Checklist Questions"];
                                if (cqTextForFilter) {
                                    const cqFormatFixText = cqTextForFilter?.toUpperCase()?.trim();
                                    if (cqFormatFixText?.substring(0, 3) == questionCode) {
                                        return qcF;
                                    }
                                }
                            });
                            if (hasMoreThanOne && hasMoreThanOne?.length > 1) {
                                const dataToPopulate = {
                                    "m": ignoreQCSet?.includes(questionCode) ? "-" : ToBeShownIndex && ToBeShownIndex > 0 && isIgnorancecheck ? colRecordTect == "-" ? "+" : "-" : "+",
                                    "ct": {
                                        "fa": "General",
                                        "t": "g"
                                    },
                                    "v": ignoreQCSet?.includes(questionCode) ? "-" : ToBeShownIndex && ToBeShownIndex > 0 && isIgnorancecheck ? colRecordTect == "-" ? "+" : "-" : "+",
                                    "ht": "0",
                                    "fs": "17",
                                    "fc": "#000000",
                                    "bl": 1
                                };

                                // luckysheet.setcellvalue( recordRange?.start + itemIndex + ( f === "Table 3" ? 3 : 2), 0, luckysheet.flowdata(), dataToPopulate );
                                setCellValue(recordRange?.start + itemIndex + (f === "Table 3" ? 3 : 2), 0, dataToPopulate)

                                let storedvalues = JSON.parse(localStorage.getItem('IconShownIndex'));
                                if (storedvalues != null) {
                                    const frequencyMap = {};
                                    storedvalues?.forEach(value => {
                                        frequencyMap[value] = (frequencyMap[value] || 0) + 1;
                                    });

                                    const finalset = storedvalues?.filter(value => frequencyMap[value] === 1);

                                    const dataPopulate = {
                                        "m": "-",
                                        "ct": {
                                            "fa": "General",
                                            "t": "g"
                                        },
                                        "v": "-",
                                        "ht": "0",
                                        "fs": "17",
                                        "fc": "#000000",
                                        "bl": 1
                                    };
                                    localStorage.setItem('IconShownIndex', JSON.stringify(finalset));
                                    finalset?.forEach(value => {
                                        setCellValue(value, 0, dataPopulate);
                                    });

                                }
                            }
                            shortQuestion.push(questionCode);
                        }
                    }
                })
            }
        });
        if (recordsToHide?.length > 0) {
            recordsToHide = groupNumbers(recordsToHide);
        }
        showOrHideRecords(recordsToHide, ToBeShownIndex);
    }
    }

    const showOrHideRecords = (rowSet, setid) => {

        var groupingData = localStorage.getItem('selectedgrouping');

        if (groupingData !== "undefined") {
            let beforeselectedvalue = +groupingData + 2;
            updateLocalStorage(beforeselectedvalue);
            let storedValues = JSON.parse(localStorage.getItem('hitedselectedValues')) || [];
            const removedArrays = [];
            rowSet = rowSet?.filter(innerArray => {
                const containsK = innerArray?.includes(beforeselectedvalue);
                const containsStoredValue = innerArray?.some(value => storedValues?.includes(value));
                if (containsK || containsStoredValue) {
                    removedArrays.push(innerArray);
                    return false;
                }
                return true;
            });

            let existingRemovedArrays = JSON.parse(localStorage.getItem('removedArrays')) || [];
            existingRemovedArrays.push(...removedArrays);
            localStorage.setItem('removedArrays', JSON.stringify(existingRemovedArrays));
            localStorage.removeItem('selectedgrouping');

            // Check if beforeselectedvalue is in storedValues
            //without reference don't remove any part of this code
            if (storedValues?.includes(beforeselectedvalue)) {
                const getremovearray = JSON.parse(localStorage.getItem('removedArrays'));
                const arrayToAddBack = getremovearray?.find(innerArray => innerArray?.includes(beforeselectedvalue));

                if (arrayToAddBack) {
                    rowSet.push(arrayToAddBack);
                    const indexInRemovedArrays = existingRemovedArrays?.findIndex(removedArray =>
                        JSON.stringify(removedArray) === JSON.stringify(arrayToAddBack)
                    );

                    if (indexInRemovedArrays !== -1) {
                        existingRemovedArrays.splice(indexInRemovedArrays, 1);
                        let iconIndexset = JSON.parse(localStorage.getItem('IconShownIndex'));
                        if (iconIndexset.includes(beforeselectedvalue - 2)) {
                            const shouldRemove = (arr) => {
                                return getremovearray.some(removed =>
                                    arr.length === removed.length && arr.every(value => removed.includes(value)) &&
                                    removed.includes(beforeselectedvalue)
                                );
                            };
                            const filteredRowSet = rowSet.filter(arr => !shouldRemove(arr));
                            rowSet.length = 0;
                            Array.prototype.push.apply(rowSet, filteredRowSet);
                            existingRemovedArrays.splice(indexInRemovedArrays, 1);
                        }

                        localStorage.setItem('removedArrays', JSON.stringify(existingRemovedArrays));
                    }
                    var hiteval = JSON.parse(localStorage.getItem('hitedselectedValues'))
                    const hittedvalues = hiteval.filter(value => value !== beforeselectedvalue);

                    localStorage.setItem('hitedselectedValues', JSON.stringify(hittedvalues));
                }
            }
        }

        const config = luckysheet.getConfig();
        let range = luckysheet.getRange();
        let zoomConfig = luckysheet.getSheet();
        let rowRange = range[0].row[0] + 1;
        let rowRangee = range[0].row[0] + 2;
        const hiddenRows = config?.rowhidden ? Object.keys(config?.rowhidden) : [];
        if (hiddenRows && hiddenRows?.length > 0) {
            const parsedSet = hiddenRows.map((f) => parseInt(f));
            const grouppedSet = groupNumbers(parsedSet);
            grouppedSet.forEach((f) => {
                luckysheet.showRow(f[0], f[f?.length - 1]);
            });
            luckysheet.scroll({
                targetRow: setid,
                targetColumn: 0
            });
        }
        if (rowSet?.length > 0) {
            rowSet.forEach((f) => {
                if (f?.length > 0) {
                    luckysheet.hideRow((f[0] - 1), (f[f?.length - 1] - 1));
                } else {
                    luckysheet.hideRow((f[0] - 1), (f[f?.length - 1] - 1));
                }
            });
            luckysheet.scroll({
                targetRow: setid,
                targetColumn: 0
            });
        }
    }

    function updateLocalStorage(beforeValue) {
        let basedatas = JSON.parse(localStorage.getItem('hitedselectedValues')) || [];
        if (basedatas.length === 0 || !basedatas.includes(beforeValue)) {
            basedatas.push(beforeValue);
        }
        localStorage.setItem('hitedselectedValues', JSON.stringify(basedatas));
    }

    const groupNumbers = (data) => {
        data = data.sort((a, b) => a - b)
        const groupedData = [];

        if (data.length === 0) {
            return;
        }

        let currentGroup = [data[0]];

        for (let i = 1; i < data.length; i++) {
            if (data[i] === data[i - 1] || data[i] === data[i - 1] + 1) {
                currentGroup.push(data[i]);
            } else {
                groupedData.push(currentGroup);
                currentGroup = [data[i]];
            }
        }
        groupedData.push(currentGroup);
        return groupedData;
    };

    const updateNameInsuredApi = async (jobId, noColumnValue) => {
        try {
            const token = sessionStorage.getItem("token");
            const Token = await processAndUpdateToken(token);
            const data = {
                JobId: jobId,
                NoColumnValue: noColumnValue,
            };
            const headers = {
                Authorization: `Bearer ${Token}`,
                "Content-Type": "application/json",
            };
            const response = await axios.put(
                baseUrl + "/api/ProcedureData/update-name-insured",
                data,
                { headers }
            );
            return response?.data;
        } catch (error) {
            throw error;
        }
    };

    const table1Datamap = async (jobId, noColumnValue, shouldCallApi) => {
        const tblRange = tableColumnDetails["Table 1"];
        const sliceData = noColumnValue.slice(tblRange?.range?.start, tblRange?.range?.end + 1);
        const defaultTextValue = 'Named Insured';
        const funToFindnoColumnValue = (sliceData) => {
            for (const row of sliceData) {
                for (const obj of row) {
                    if (obj && (obj.m === defaultTextValue || obj.v === defaultTextValue)) {
                        return row;
                    }
                }
            }
        }
        let rowData = funToFindnoColumnValue(sliceData);
        let filterData = rowData.filter(e => e != null);
        if (brokerId == "1162" && shouldCallApi) {
            await updateNameInsuredApi(jobId, filterData[1]?.m || filterData[1]?.v || "");
        }
    }

    const onSaveClick = async (shouldCallApi, isAutoUpdate = false) => {
        try {
            luckysheet.exitEditMode();
            const allSheets = luckysheet.getAllSheets();
            const policyData = allSheets?.filter((f) => f?.name === "PolicyReviewChecklist");
            const formData = allSheets?.filter((f) => f?.name === "Forms Compare");
            const exclusonData = allSheets?.filter((f) => f?.name === "Exclusion");
            const redData = allSheets?.filter((f) => f?.name === "Red");
            const greenData = allSheets?.filter((f) => f?.name === "Green");
            const qacData = allSheets?.filter((f) => f?.name === "QAC not answered questions");

            const jobId = props?.selectedJob;
            const noColumnValue = policyData[0]?.data;
            const sessionQacData = sessionStorage.getItem("qacTblRange");
            const qacTableColumnDetails = JSON?.parse(sessionQacData);

            const processData = (sheetName, sheetData, tblRangeArray, filterData) => {
                if (sheetName != "QAC not answered questions") {
                    const aodDataMSet = [];
                    if (sheetData && sheetData?.length > 0 && tblRangeArray && tblRangeArray?.length > 0) {
                        tblRangeArray?.forEach((f) => {
                            const tableFilterData = filterData?.find((d) => d.TableName === f?.TableName);
                            if (tableFilterData) {
                                const filteredOriginalData = tableFilterData?.data;
                                const start = f?.TableName == "FormTable 2" ? f?.range?.start + 3 : f?.TableName == "FormTable 3" ? f?.range?.start + 3 : f?.range?.start + 2;
                                const end = f?.range?.end + 1;
                                const splitedData = sheetData.slice(start, end);
                                if (splitedData && splitedData?.length > 0) {
                                    let aodDataSet = [];
                                    splitedData?.forEach((item, itemIndex) => {
                                        let record = filteredOriginalData?.find((fod) => fod?.sheetPosition && fod?.sheetPosition === (start + itemIndex));
                                        const aod = item[f?.columnNames["Actions on Discrepancy"]] || item[f?.columnNames["ActionOnDiscrepancy"]] || item[f?.columnNames[0]?.indexOf("ActionOnDiscrepancy") - 1];
                                        const re = item[f?.columnNames["Request Endorsement"]] || item[f?.columnNames["RequestEndorsement"]] || item[f?.columnNames[0]?.indexOf("RequestEndorsement") - 1];
                                        const ns = item[f?.columnNames["Notes for Endorsement"]] || item[f?.columnNames["Notes"]] || item[f?.columnNames[0]?.indexOf("Notes") - 1];
                                        const nff = item[f?.columnNames["Notes(Free Fill)"]] || item[f?.columnNames["NotesFreeFill"]] || item[f?.columnNames[0]?.indexOf("NotesFreeFill") - 1];

                                        let aodData = aod ? getText(aod, true) : '';
                                        let reData = re ? getText(re, true) : '';
                                        let nsData = ns ? getText(ns, true) : '';
                                        let nffData = nff ? getText(nff, true) : '';

                                        aodData = aodData != undefined && aodData != null ? aodData : '';
                                        reData = reData != undefined && reData != null ? reData : '';
                                        nsData = nsData != undefined && nsData != null ? nsData : '';
                                        nffData = nffData != undefined && nffData != null ? nffData : '';

                                        if (record && !(aodData?.trim() === '' && reData?.trim() === '' && nsData?.trim() === '' && nffData?.trim() === '')) {
                                            record["ActionOnDiscrepancy"] = aodData;
                                            record["RequestEndorsement"] = reData;
                                            record["Notes"] = nsData;
                                            record["NotesFreeFill"] = nffData;
                                            if (record["ActionOnDiscrepancy"] === "Click here" && record["RequestEndorsement"] === "Click here" &&
                                                record["Notes"] === "Click here"
                                            ) {
                                                record["ActionOnDiscrepancy"] = "";
                                                record["RequestEndorsement"] = "";
                                                record["Notes"] = "";
                                            } else {
                                                ["ActionOnDiscrepancy", "RequestEndorsement", "Notes"].forEach(key => {
                                                    if (record[key] === "Click here") {
                                                        record[key] = "";
                                                    }
                                                })
                                            }
                                            aodDataSet.push(record);
                                        }

                                        if (filteredOriginalData && splitedData?.length === (itemIndex + 1) && sheetName != "Exclusion") {
                                            const keys = Object.keys(filteredOriginalData[0]);
                                            const policyKey = keys?.includes("POLICY LOB") ? "POLICY LOB" : "Policy LOB";
                                            let policyLob = filteredOriginalData?.map(f => f[policyKey]);
                                            policyLob = Array.from(new Set(policyLob));
                                            aodDataMSet.push({ "TableName": f?.TableName, "Data": aodDataSet, "PolicyLob": policyLob[0] });
                                        } else if (filteredOriginalData && splitedData?.length === (itemIndex + 1) && sheetName == "Exclusion") {
                                            aodDataMSet.push({ "Data": aodDataSet });
                                        }
                                    });
                                }
                            }
                        });
                    }
                    return aodDataMSet;
                } else if (sheetName == "QAC not answered questions") {
                    const qacDataMSet = [];
                    const sheetDataRowIndex = sheetDataRowIndexRef?.current;
                    const queRowRata = Object.keys(sheetDataRowIndex);
                    queRowRata?.forEach((Key) => {
                        const lobrow = sheetDataRowIndex[Key];
                        Object.keys(lobrow)?.forEach((lob) => {
                            const filterData = lobrow[lob];
                            let tblDataSet = [];
                            if(filterData && filterData?.length > 0) {
                                filterData?.forEach((item, index) => {
                                    const findindex = luckysheet.getcellvalue(item?.RowIndex);
                                    if (findindex?.length > 0) {
                                        const mappedRow = {
                                            "Coverage Specifications": getText(findindex[0], true),
                                            "Discrepancies to note on Intranet Form:": getText(findindex[1], true),
                                            "Checked By (Initials):": getText(findindex[2], true)
                                        };  
                                        tblDataSet.push(mappedRow);
                                    }
                                    if (filterData?.length === (index + 1)) {
                                        qacDataMSet.push({ "Table": `qacTable ${index + 0}`, "dataPairSet": tblDataSet, "tableType": Key, "policyLob": lob });
                                    }
                                });
                            }
                        });
                    });

                    const reseededTblDatSet = qacDataMSet?.map((item, index) => {
                        return {
                            ...item,
                            Table: `qacTable ${index + 0}`
                        };
                    });
                    const finalDataSet = {};

                    reseededTblDatSet?.forEach(dataSet => {
                        const { Table, tableType, dataPairSet, policyLob } = dataSet;
                        if (!finalDataSet[tableType]) {
                            finalDataSet[tableType] = {};
                        }
                        const tableDetail = qacTableColumnDetails[tableType]?.find(tbl => tbl[Table]);

                        if (tableDetail) {
                            const lobData = policyLob;

                            if (!finalDataSet[tableType][lobData]) {
                                finalDataSet[tableType][lobData] = [];
                            }
                            finalDataSet[tableType][lobData].push(...dataPairSet);
                        }
                    });
                    return finalDataSet;
                };
            };

            let checklistData = [];
            let formsCompareData = [];
            let exclusionData = [];
            let qacSheetData = [];
            let combineGradiationData = [];

            let sheetCheck = luckysheet?.getSheet()?.name;
            if (policyData && policyData?.length > 0 && sheetCheck == "PolicyReviewChecklist") {
                const policytrack = localStorage.getItem('policyDataTracked')
                const data = dataForSavePolicy == null || undefined || dataForSavePolicy.length == 0 ? JSON.parse(policytrack) : dataForSavePolicy;
                const reseededData = data?.map((item, index) => {
                    return {
                        ...item,
                        TableName: `Table ${index + 2}`
                    };
                });

                const positionpolicy = localStorage.getItem('positioningForPolicy')
                const policyPosDetails = dataForSavePolicyPosition == null || undefined || dataForSavePolicyPosition.length == 0 ? JSON.parse(positionpolicy) : dataForSavePolicyPosition;
                const reseededPolicyPosDetails = policyPosDetails?.map((item, index) => {
                    return {
                        ...item,
                        TableName: `Table ${index + 2}`
                    };
                });

                const sheetData = policyData[0]?.data;
                checklistData = processData("PreviewCheckList", sheetData, reseededPolicyPosDetails, reseededData);
            }

            if (formData && formData?.length > 0 && sheetCheck == "Forms Compare") {
                const formsCompare_DataTracked = localStorage.getItem('formsCompareDataTracked');
                const data = dataForSaveFormsCompare == null || undefined || dataForSaveFormsCompare.length == 0 ? JSON.parse(formsCompare_DataTracked) : dataForSaveFormsCompare;
                const positioningFor_FormsCompareData = localStorage.getItem('positioningForFormsCompareData');
                const formsComparePosDetails = dataForSaveFormsComparePosition == null || undefined || dataForSaveFormsComparePosition.length == 0 ? JSON.parse(positioningFor_FormsCompareData) : dataForSaveFormsComparePosition;

                const sheetData = formData[0]?.data;
                formsCompareData = processData("FormsCompare", sheetData, formsComparePosDetails, data);
            }

            if (exclusonData && exclusonData?.length > 0 && sheetCheck == "Exclusion") {
                const exclusionDataTracked = localStorage.getItem('exclusionDataTracked');
                const data = dataForSaveExclusion[0] == null || undefined || dataForSaveExclusion.length == 0 ? JSON.parse(exclusionDataTracked) : dataForSaveExclusion[0];
                const positioningForexclusionData = localStorage.getItem('positioningForexclusionData');
                const exclusionPosDetails = dataForSaveExclusionPosition == null || undefined || dataForSaveExclusionPosition.length == 0 ? JSON.parse(positioningForexclusionData) : dataForSaveExclusionPosition;

                const sheetData = exclusonData[0]?.data;
                exclusionData = processData("Exclusion", sheetData, exclusionPosDetails, data);
            }

            if (qacData && qacData?.length > 0 && sheetCheck == "QAC not answered questions") {
                const tableRangeData = qacTableColumnDetails;
                const sheetData = qacData[0]?.data;
                qacSheetData = processData("QAC not answered questions", sheetData, tableRangeData, "");
            }

            if (redData && redData?.length > 0 && (sheetCheck == "Red" || sheetCheck == "Green")) {
                let redSheetData = sessionStorage.getItem('redSheetData');
                let data = JSON.parse(redSheetData);
                let filterData = data?.filter(item => Object.keys(item.data).length !== 0);
                filterData?.forEach((item, index) => {
                    item.TableName = `Table ${index + 2}`;
                });

                let redTableRange = sessionStorage.getItem('redTableRangeDataUpdate');
                const gradiationRedData = JSON.parse(redTableRange);

                const sheetData = redData[0]?.data;
                combineGradiationData = combineGradiationData.concat(processData("GradiationRedSheet", sheetData, gradiationRedData, filterData));
            }

            if (greenData && greenData?.length > 0 && (sheetCheck == "Red" || sheetCheck == "Green")) {
                let greenSheetData = sessionStorage.getItem('greenSheetData');
                let data = JSON.parse(greenSheetData);
                let filterData = data?.filter(item => Object.keys(item.data).length !== 0);
                filterData?.forEach((item, index) => {
                    item.TableName = `Table ${index + 2}`;
                });

                let greenTableRange = sessionStorage.getItem('greenTableRangeDataUpdate');
                const gradiationGreenData = JSON.parse(greenTableRange);

                const sheetData = greenData[0]?.data;
                combineGradiationData = combineGradiationData.concat(processData("GradiationGreenSheet", sheetData, gradiationGreenData, filterData));
            }

            let filteredCombinedData = combineGradiationData?.filter(f => f.Data.length !== 0);  // filter the empty array after concatination

            // Merge the data for objects with the same TableName and PolicyLob
            let mergedLobData = filteredCombinedData.reduce((acc, item) => {
                const existingItem = acc?.find(i => i.PolicyLob === item?.PolicyLob);
                if (existingItem) {
                    existingItem.Data = existingItem?.Data.concat(item.Data);
                } else {
                    acc.push(item);
                }
                return acc;
            }, []);

            let filterNullLobData = mergedLobData?.filter(f => f?.Data?.length !== 0);

            let checklistDataForUpdate = checklistData?.filter(f => f?.Data && f.Data?.length > 0);
            let formsCompareDataForUpdate = formsCompareData
                ?.filter(f => f?.Data && f.Data?.length > 0)
                ?.map(f => ({
                    ...f,
                    Data: f.Data.map(d => {
                        const { COVERAGE_SPECIFICATIONS_MASTER, OBSERVATION, sheetPosition, ...rest } = d;
                        delete rest["Checklist Questions"];
                        delete rest["Policy LOB"];
                        delete rest["Page Number"];
                        delete rest["Current Term Policy Attached"];
                        delete rest["Prior Term Policy Attached"];
                        delete rest["Document Viewer"];
                        return rest;
                    })
                }));

            let exclusionDataForUpdate = exclusionData
                ?.filter(f => f?.Data && f.Data?.length > 0)
                ?.map(f => ({
                    ...f,
                    Data: f.Data.map(d => {
                        const { FormDescription, FormName, JobId, Exclusion, PageNumber, CreatedOn, UpdatedOn, sheetPosition, ...rest } = d;
                        return rest;
                    })
                }));

            let qacDataSet = qacSheetData;

            // checklistDataForUpdate.forEach(table => {
            //     table.Data.forEach(item => {
            //         Object.keys(item).forEach(key => {
            //            if(!keysToRemovePolicyReview.includes(key)){
            //                delete item[key];
            //            }
            //         });
            //     });
            //     delete table['PolicyLob'];
            //     delete table['TableName'];
            // });

            const gradiationDataForUpdate = filterNullLobData.filter(f => f?.Data && f.Data?.length > 0);

            // logic to send the Data to api if any of one sheet data is available for a JobId
            if (checklistDataForUpdate.length > 0 && sheetCheck == "PolicyReviewChecklist") {
                // checklistDataForUpdate.forEach ( Ary => {
                //     Ary.Data = Ary?.Data.filter( obj => {
                //         let spliceObj = Object.keys(obj).filter(key => key != "Id");
                //         return spliceObj.some(keys => obj[keys] != "");  // sending to api only the objects with non-empty values
                //     })
                // })
                if (shouldCallApi) {
                    await csrSheetPolicyGradiationDataSaveApi(0, props?.selectedJob, "PreviewCheckList", JSON.stringify(checklistDataForUpdate), isAutoUpdate);
                }
                await csrSaveJobidExportFun("PreviewCheckList", checklistDataForUpdate, shouldCallApi, isAutoUpdate);   // function to call the export for csrSave history Jobid in the xlpage Report Screen
            } else if (sheetCheck == "Forms Compare") {
                let mappingData = formsCompareDataForUpdate.map(e => e.Data);
                let arry = mappingData.flat();  // used to flatten the nested array into array of objects
                let splicedArry = arry.filter(obj => {
                    let spliceObj = Object.keys(obj).filter(key => key != "Id");
                    return spliceObj.some(keys => obj[keys] != "");   // sending to api only the objects with non-empty values
                })
                if (shouldCallApi) {
                    await csrSheetCommonSaveApi(props?.selectedJob, "FormTable", JSON.stringify(splicedArry), isAutoUpdate);
                }
                // await CsrSaveHistoryApiCall(props?.selectedJob, sessionStorage.getItem("csrAuthUserName"), JSON.stringify(splicedArry), brokerId, isAutoUpdate);
            } else if (sheetCheck == "Exclusion") {
                let mappingData = exclusionDataForUpdate.map(e => e.Data);
                let arry = mappingData.flat();  // used to flatten the nested array into array of objects
                let splicedArry = arry.filter(obj => {
                    let spliceObj = Object.keys(obj).filter(key => key != "Id");
                    return spliceObj.some(keys => obj[keys] != "");   // sending to api only the objects with non-empty values
                })
                if (shouldCallApi) {
                    csrSheetCommonSaveApi(props?.selectedJob, "ExclusionTable", JSON.stringify(splicedArry), isAutoUpdate);
                }
                // await CsrSaveHistoryApiCall(props?.selectedJob, sessionStorage.getItem("csrAuthUserName"), JSON.stringify(splicedArry), brokerId, isAutoUpdate);
            } else if ((sheetCheck == "Red" || sheetCheck == "Green")) {
                if (shouldCallApi) {
                    await csrSheetPolicyGradiationDataSaveApi(0, props?.selectedJob, "GradiationSheet", JSON.stringify(gradiationDataForUpdate), isAutoUpdate);
                }
                await CsrSaveHistoryApiCall("GradiationSheet", props?.selectedJob, sessionStorage.getItem("csrAuthUserName"), JSON.stringify(gradiationDataForUpdate), brokerId, isAutoUpdate);
            } else if (sheetCheck == "QAC not answered questions") {
                if (shouldCallApi) {
                    await qacSheetSaveCall(props?.selectedJob, JSON.stringify(qacDataSet), isAutoUpdate);
                }
            }

            if (brokerId == "1167") {
                await table1Datamap(jobId, noColumnValue, shouldCallApi);
            }
        } catch (error) {
            const message = isAutoUpdate ? "CSR-AOD-AuoUpdate-UI-function-error" : "CSR-AOD-Update-UI-function-error";
            updateGridAuditLog(jobId, message, typeof error === 'object' ? JSON.stringify(error) : error, (userName || sessionUserName));
        }
    };

    const Autoupdateclick = async (autoupdate) => {
        if (autoupdate == true && (luckysheet != undefined && luckysheet != null)) {
            try {
                await onSaveClick(true, true);
            } catch (error) {
                console.error("Error occurred:", error);
            }
        }
    }

    const csrSaveJobidExportFun = async (sheetType, checklistDataForUpdate, shouldCallApi, isAutoUpdate) => {
        let mergedData = checklistDataForUpdate.reduce((acc, table) => {
            return acc.concat(table.Data);
        }, []);

        let filteredArrayData = mergedData.filter(obj =>
            obj.ActionOnDiscrepancy !== "" ||
            obj.RequestEndorsement !== "" ||
            obj.Notes !== ""
        )

        if (filteredArrayData?.length > 0) {
            let UserName = sessionStorage.getItem("csrAuthUserName");
            let JobID = props?.selectedJob;
            let currentDate = new Date().toISOString();

            filteredArrayData.forEach(item => {
                item.UserName = UserName;  // Add UserName as a key-value pair
                item.JobID = JobID;        // Add JobID as a key-value pair
                item.CreatedOn = currentDate;
                delete item?.Id;
            });

            filteredArrayData.forEach(obj => {
                // Rename 'COVERAGE_SPECIFICATIONS_MASTER' or 'Coverage_Specifications_Master' to 'COVERAGE SPECIFICATIONS'
                if (obj.hasOwnProperty('COVERAGE_SPECIFICATIONS_MASTER')) {
                    obj['COVERAGE SPECIFICATIONS'] = obj['COVERAGE_SPECIFICATIONS_MASTER'];
                    delete obj['COVERAGE_SPECIFICATIONS_MASTER'];
                } else if (obj.hasOwnProperty('Coverage_Specifications_Master')) {
                    obj['COVERAGE SPECIFICATIONS'] = obj['Coverage_Specifications_Master'];
                    delete obj['Coverage_Specifications_Master'];
                }
            });

            let getLocalStgbrokerDatas = localStorage.getItem('brokerDatas');
            const brokerData = JSON.parse(getLocalStgbrokerDatas);
            const mappedData = brokerData.map(({ BrokerId, VchBrokerName }) => ({
                key: BrokerId.toString(),
                text: VchBrokerName,
            }));

            filteredArrayData.forEach(item => {
                const sliceJobid = item?.JobID.slice(0, 4);
                const brokerIdMap = mappedData.find(e => e.key === sliceJobid);

                if (brokerIdMap) {
                    item.BrokerId = brokerIdMap?.key;
                    item.BrokerName = brokerIdMap?.text;
                }
            })
            localStorage.setItem('endorsementRowData', JSON.stringify(filteredArrayData))
            if (shouldCallApi) {
                await CsrSaveHistoryApiCall(sheetType, props?.selectedJob, sessionStorage.getItem("csrAuthUserName"), JSON.stringify(filteredArrayData), brokerId, isAutoUpdate);
            }
        }
    }

    const UpdateJobPreviewStatusCall = async () => {
        const response = await UpdateJobPreviewStatus(jobId, token);
        await getCsrReviewData(jobId);
        container.current.showSnackbar('Status updated successfully', "info", true);
    }

    const UpdateJobSendPolicyInsuredCall = async () => {
        const response = await UpdateJobSendPolicyInsured(jobId, token);
        await getCsrReviewData(jobId);
        container.current.showSnackbar('Status updated successfully', "info", true);
    }

    const isPreviewCompleted = reviewData?.length > 0 && reviewData[0]?.IsPreviewCompleted;
    const isGridspiCompleted = sendPolicyInsuredData?.length > 0 && sendPolicyInsuredData[0]?.IsGridspiCompleted;

    useEffect(() => {
        checkUserAndShowDialog();
    }, [activeUserName, sessionUserName]);

    const checkUserAndShowDialog = () => {
        if (activeUserName !== sessionUserName && activeUserName) {
            setOpenDialog(true);
            setMsgText(
                <span className="msg-text">
                    The checklist for this job is currently opened by{' '}
                    {activeUserName === "" ? " - " : activeUserName}.
                    You will not be able to edit or save any changes to the checklist.
                </span>
            );
        }
    };

    const handleDialogClose = (shouldNavigate) => {
        setOpenDialog(false);
        if (shouldNavigate) {
            navigate(-1); // Go to the previous page
        }
    };

    return (
        <div>
            <div >
                {(activeUserName === sessionUserName || !activeUserName) ? ( // Show buttons only if users match
                    <div
                        // onClick={() => handleExportExcel()}
                        style={{ display: 'flex' }}>
                        {/* { brokerId != "1003" || (endorsementColflag != "Forms Compare" && endorsementColflag != "Exclusion") ? ( <PrimaryButton onClick={() => onSaveClick()} style={{ */}
                        {AODDBrokerIds.includes(brokerId) || (endorsementColflag != "Forms Compare" && endorsementColflag != "Exclusion") ? (<PrimaryButton onClick={() => onSaveClick(true, false)} style={{
                            backgroundColor: 'lightblue', height: '28px', borderRadius: '20px',
                            borderColor: 'black', fontSize: '12px', margin: '5px 5px 5px 5px', color: 'black'
                        }}>Save</PrimaryButton>) : null
                        }
                        <PrimaryButton
                            onClick={() => {
                                onSaveClick(false); // Call onSaveClick with false to skip API call
                                ExportClick(true);   // Then trigger ExportClick
                            }} style={{
                                backgroundColor: 'lightblue', height: '28px', borderRadius: '20px',
                                borderColor: 'black', fontSize: '12px', margin: '5px 5px 5px 5px', color: 'black'
                            }}>Export Checklist</PrimaryButton>
                        {/* { brokerId != "1003" || (endorsementColflag != "Forms Compare" && endorsementColflag != "Exclusion") ? ( <PrimaryButton onClick={() => toggleYesDialog()} style={{ */}
                        {((AODDBrokerIds.includes(brokerId) && endorsementColflag !== "QAC not answered questions") || (!AODDBrokerIds.includes(brokerId) && endorsementColflag !== "Forms Compare" && endorsementColflag !== "Exclusion" && endorsementColflag !== "QAC not answered questions")) && !reqEndorsementColHideBrokerId.includes(brokerId) ? (<PrimaryButton onClick={() => toggleYesDialog()} style={{
                            backgroundColor: 'lightblue', height: '28px', borderRadius: '20px',
                            borderColor: 'black', fontSize: '12px', margin: '5px 5px 5px 5px', color: 'black'
                        }}>Generate Endorsement Template</PrimaryButton>) : null
                        }
                        {(endorsementColflag !== "QAC not answered questions") ? <PrimaryButton onClick={() => UpdateJobPreviewStatusCall()} style={{
                            backgroundColor: isPreviewCompleted ? '#22bb33' : 'lightblue', height: '28px', borderRadius: '20px',
                            borderColor: 'black', fontSize: '12px', margin: '5px 5px 5px 5px', color: 'black'
                        }}>Review Completed</PrimaryButton> : null }
                        {['1003', '1165']?.includes(brokerId) && (endorsementColflag !== "QAC not answered questions") ? <PrimaryButton onClick={() => UpdateJobSendPolicyInsuredCall()} style={{
                            backgroundColor: isGridspiCompleted ? '#22bb33' : 'lightblue', height: '28px', borderRadius: '20px',
                            borderColor: 'black', fontSize: '12px', margin: '5px 5px 5px 5px', color: 'black'
                        }}>Send Policy To Insured</PrimaryButton> : null }

                        {/* <Icon iconName="Filter" onClick={handleIconClick} style={{ fontSize: '20px', margin: '5px', cursor: 'pointer' }}/> */}

                    </div>
                ) : null}

                {openDialog && (<DialogComponent isOpen={openDialog} onClose={handleDialogClose} message={msgText} />)}
                {yesDialog && <EndorsementDialogComponent isOpen={yesDialog} luckySheet={luckysheet} state={csrPolicyData} formState={formsComparedata} redSheetData={gradtionData} sheetState={sheetState} tableColumnDetails={tableColumnDetails} formTableColumnDetails={formTableColumnDetails} exTableColumnDetails={exTableColumnDetails} dataForXRayMapping={dataForXRayMapping} formdataForXRayMapping={formdataForXRayMapping} redSheetDataForXRayMapping={sessionStorage.getItem('redSheetData')} greenSheetDataForXRayMapping={sessionStorage.getItem('greenSheetData')} onClose={(e) => handleYesDailog(e)} />}
                {dropDialog && < DiscrepancyOptionsDialogComponent isOpen={dropDialog} luckySheet={luckysheet} state={csrPolicyData} jobId={jobId} onClose={(e) => funForDiscrepancyCol(e)} message={"Action On Descrepancy (from AMs)"} />}
                {gradiationDialog && < DiscrepancyOptionsDialogComponent isOpen={gradiationDialog} luckySheet={luckysheet} state={csrPolicyData} jobId={jobId} onClose={(e) => funForDiscrepancyCol(e)} message={"Action On Descrepancy (from AMs)"} />}
                {openFilterDialog && <FilterCsrDialogComponent isOpen={{ openFilterDialog, tableColumnDetails, props, luckysheet, filterSelectionData }} onClose={(e) => handleFilterDialogClose(e)} />}
            </div>
            <div style={{ position: 'relative' }}>
                {!reqEndorsementColHideBrokerId.includes(brokerId) && (endorsementColflag !== "QAC not answered questions") ? (<Icon iconName="Filter" onClick={handleIconClick} style={{
                    position: 'absolute',
                    top: '2px',
                    right: '-450px',
                    fontSize: '16.2px',
                    margin: '5px',
                    zIndex: 10,
                    cursor: 'pointer'
                }} />) : null}
                {!reqEndorsementColHideBrokerId.includes(brokerId) && (endorsementColflag !== "QAC not answered questions") ? (<h6 style={{
                    position: 'absolute',
                    top: '5px',
                    fontWeight: 500,
                    right: '-480px',
                    fontSize: '12.5px',
                    margin: '5px',
                    zIndex: 10,
                    cursor: 'pointer'
                }} onClick={handleIconClick} >Filter</h6>
                ) : null}
                <div className="csrSheet" id="luckysheet2" ref={luckyCss} ></div>

            </div>

            <SimpleSnackbar ref={container} />
        </div>
    );
}
