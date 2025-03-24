
import React, { useEffect, useRef, useState } from "react";
import { Disclaimer, Checklist, updateData, formCompare, } from '../Services/Constants';
import { getQACData, SetCheckListQuestionMasterData } from "../Services/CommonFunctions"
import '../App.css';

export default function Luckysheet(props) {
    const { selectChange } = props;
    const luckysheet = window.luckysheet;
    const [jobId, setJobId] = useState(props?.selectedJob);
    const [isFormApplicable, setIsFormApplicable] = useState(true);
    let token = sessionStorage.getItem('token');
    const [sheetsDropOption, setSheetDropOption] = useState([]);
    const [dropDownOption, setDropDownOption] = useState(props?.sheetOptionSet);
    const [selectedSheet, setSelectedSheet] = useState(props?.selectedSheet || dropDownOption[0]);


    let apiDataConfig = {
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
                    "2": 35,
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
                    "13": 50,
                    "14": 20,
                    "15": 20,
                    "16": 20,
                    "17": 31
                },
                columnlen: {
                    "1": 400,
                    "2": 400,
                    "3": 400,
                    "4": 400,
                    "5": 400,
                    "6": 400,
                    "7": 400,
                    "8": 400,
                    "9": 400,
                    "10": 400,
                    "11": 400,
                    "12": 400,
                    "13": 400,
                    "14": 400
                },
                "curentsheetView": "viewPage",
                "sheetViewZoom": {
                    "viewNormalZoomScale": 0.6,
                    "viewPageZoomScale": 0.6,
                },
            },
            chart: [], // Chart configuration
            status: "1", // Activation status
            order: "0", // The order of the worksheet
            hide: 0, // Whether to hide
            column: 50, // Number of columns
            row: 50, // Number of rows
            celldata: [],// Original cell data set
            ch_width: 2322, // The width of the worksheet area
            rh_height: 949, // The height of the worksheet area
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [], // Selected area
            calcChain: [], // Formula chain
            isPivotTable: false, // Whether to pivot table
            pivotTable: {}, // Pivot table settings
            filter_select: null, // Filter range
            filter: null, // Filter configuration
            luckysheet_alternateformat_save: [], // Alternate colors
            luckysheet_alternateformat_save_modelCustom: [], // Customize alternate colors
            sheets: []
        }
    };
    let FormCompare_appconfigdata = {

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
                    "3": 70,
                    "4": 50,
                    "5": 50,
                    "6": 50,
                    "7": 75,
                    "8": 50,
                    "9": 50,
                    "10": 50,
                    "11": 150,
                    "12": 50,
                    "13": 50,

                },
                columnlen: {
                    "1": 220,
                    "2": 220,
                    "3": 220,
                    "4": 220,
                    "5": 220,
                    "6": 220,
                    "7": 220,
                },
                "curentsheetView": "viewPage",//viewNormal, viewLayout, viewPage
                "sheetViewZoom": {
                    "viewNormalZoomScale": 0.6,
                    // "viewPageZoomScale": 1,
                    "viewPageZoomScale": 0.6,
                },
            },
            chart: [], // Chart configuration
            order: "0", // The order of the worksheet
            hide: 0, // Whether to hide
            column: 50, // Number of columns
            row: 50, // Number of rows
            celldata: [],
            ch_width: 2322, // The width of the worksheet area
            rh_height: 949, // The height of the worksheet area
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [], // Selected area
            calcChain: [], // Formula chain
            isPivotTable: false, // Whether to pivot table
            pivotTable: {}, // Pivot table settings
            filter_select: null, // Filter range
            filter: null, // Filter configuration
            luckysheet_alternateformat_save: [], // Alternate colors
            luckysheet_alternateformat_save_modelCustom: [], // Customize alternate colors
            sheets: [],
        },
    }
    let exclusionDatafigdata = {

        exclusion: {
            name: "Exclusion",
            config: {
                merge: {},
                borderInfo: [],
                columnlen: {
                    "0": 120,
                    "1": 120,
                    "2": 320,
                    "3": 120,
                },
                rowlen: {},
                "curentsheetView": "viewPage",
                "sheetViewZoom": {
                    "viewNormalZoomScale": 0.6,
                    "viewPageZoomScale": 0.6,
                },
            },
            status: "1",
            column: 50,
            row: 500,
            celldata: [],
            ch_width: 2322,
            rh_height: 949,
            scrollLeft: 0,
            scrollTop: 0,
            luckysheet_select_save: [],
            calcChain: [],
            isPivotTable: false,
            pivotTable: {},
            luckysheet_alternateformat_save: [],
            luckysheet_alternateformat_save_modelCustom: [],
            sheets: [],
        },
    }

    const luckyCss = {
        margin: '0px',
        padding: '0px',
        position: 'absolute',
        width: '100% !important',
        height: '40%',
        left: '0px',
        top: '0px',
    };

    useEffect(() => {
        const mainData = props.data;
        const sheetRenderConfig = props?.sheetRenderConfig;
        const formCompareData = props.formCompareData;
        SetCheckListQuestionMasterData(token, jobId);
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

        const renderTable1 = () => {
            const tableData1 = mainData.find((data) => data.Tablename === "Table 1");
            if (tableData1) {
                const table1json = tableData1.TemplateData;

                let sheetDataTable1 = [];
                let sheetDataTable2 = [];

                const rowIndexOfTable1 = 4
                const textBlockData = renderTextBlock();
                const listData = renderList();

                textBlockData.forEach((item, index) => {
                    const mergeConfig = apiDataConfig.demo.config.merge["0_1"];

                    sheetDataTable2.push({
                        r: index + mergeConfig.r,
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            bl: 1,
                            ff: item.ff,
                            fs: 12,
                            merge: mergeConfig,
                            fc: item.fc,
                        }
                    });
                });

                listData.forEach((item, index) => {
                    const mergeConfig = apiDataConfig.demo.config.merge["1_1"];

                    sheetDataTable2.push({
                        r: 1 + mergeConfig.r,
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            ff: item.ff,
                            fs: 17,
                            merge: mergeConfig,
                            fc: item.fc,
                        }
                    });
                });

                table1json.map((item, index) => {
                    if (item["Headers"] != null && item["Headers"] != undefined) {
                        if (item["Headers"] == "") {
                            sheetDataTable1.push({
                                r: rowIndexOfTable1 + index,
                                c: 1,
                                v: {
                                    ct: { fa: "@", t: "inlineStr", s: [{ v: " " }] },
                                    m: " ",
                                    v: " ",
                                    merge: null,
                                    bg: "rgb(139,173,212)",
                                    tb: '2',
                                }
                            });
                        } else {
                            sheetDataTable1.push({
                                r: rowIndexOfTable1 + index,
                                c: 1,
                                v: {
                                    ct: { fa: "@", t: "inlineStr", s: [{ v: item["Headers"], ff: "Tahoma", fs: 10 }] },
                                    m: item["Headers"],
                                    v: item["Headers"],
                                    ff: "Tahoma",
                                    merge: null,
                                    bg: "rgb(139,173,212)",
                                    tb: '2',
                                }
                            });
                        }

                        const tidleValue = item["NoColumnName"] !== null && item["NoColumnName"] != undefined ? item["NoColumnName"].replace(/~~/g, "\n") : "";

                        sheetDataTable1.push({
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

                const dummyData = [];

                const allRows = [...sheetDataTable1, ...sheetDataTable2];

                // Sort the rows by rowIndex if needed
                allRows.sort((a, b) => a.r - b.r);

                // Add the rows to the dummyData
                dummyData.push(...allRows);

                apiDataConfig.demo.celldata = dummyData;

                //table1 border info styles
                allRows.forEach((row) => {
                    if (sheetDataTable1.includes(row)) {
                        apiDataConfig.demo.config.borderInfo.push({
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

                const excludedTablenames = ["JobHeader", "JobCommonDeclaration", "JobCoverages", "Tbl_ChecklistForm1", "Tbl_ChecklistForm2", "Tbl_ChecklistForm3", "Tbl_ChecklistForm4"];
                mainData.map((e, index) => {
                    if (!excludedTablenames.includes(e?.Tablename) && e?.Tablename != 'Table 1' && e?.TemplateData?.length >= 1) {
                        let filteredData = apiDataConfig.demo.celldata.filter((f, index) => f != null || !f);
                        renderTable2([...filteredData], e?.Tablename);
                    }
                });

                renderLuckySheet(true, '', false);

            }
        };
        const renderTable2 = (combinedata1, tableName) => {
            if (!Array.isArray(combinedata1)) {
                return;
            }
            const needDocumentViewer = true;
            const DefaultColumns = ["Actions on Discrepancy (from AMs)", "Actions on Discrepancy", "Request Endorsement", "Notes", "Notes(Free Fill)"];
            const basedata = [...combinedata1];
            const inputData = mainData;

            inputData.forEach(item => {
                if (item.Tablename != 'Table 1') {
                    if (Array.isArray(item.TemplateData)) {
                        item.TemplateData.forEach(data => {
                            if (typeof data === 'object' && data !== null) {
                                if (!data.hasOwnProperty('CoverageSpecificationsMaster')) {
                                    data.CoverageSpecificationsMaster = null;
                                }
                            }
                        });
                    }
                }
            });

            const targetTablenames = ["JobCommonDeclaration", "JobCoverages", "Tbl_ChecklistForm1", "Tbl_ChecklistForm2", "Tbl_ChecklistForm3", "Tbl_ChecklistForm4"];
            inputData.forEach(data => {
                if (targetTablenames.includes(data.Tablename)) {
                    let index = data.TemplateData.indexOf("Policy LOB");
                    if (index === -1) {
                        index = data.TemplateData.indexOf("POLICY LOB");
                    }
                    if (index !== -1) {
                        data.TemplateData[index] = "PolicyLob";
                    }
                    let index1 = data.TemplateData.indexOf("COVERAGE_SPECIFICATIONS_MASTER");
                    if (index1 !== -1) {
                        data.TemplateData[index1] = "CoverageSpecificationsMaster"
                    }
                    if (data.Tablename == "Tbl_ChecklistForm2") {
                        let index2 = data.TemplateData.indexOf("Current Term Policy - Listed");
                        if (index2 !== -1) {
                            data.TemplateData[index2] = "CurrentTermPolicyListed1"
                        }
                    }
                    let index3 = data.TemplateData.indexOf("OBSERVATION");
                    if (index3 !== -1) {
                        data.TemplateData[index3] = "Observation"
                    }
                }
            });

            inputData.forEach(data => {
                if (targetTablenames.includes(data.Tablename)) {
                    data.TemplateData = data.TemplateData.map(item => item.replace(/ /g, '').replace(/_/g, '').replace(/-/g, ''));
                }
            });

            // Data mapping based on the appConfig table headings--*
            const tableDataMap = {
                'Table 4': { data: inputData.find(data => data.Tablename === 'Table 4'), appConfigTableData: 'Tbl_ChecklistForm1' },
                'Table 5': { data: inputData.find(data => data.Tablename === 'Table 5'), appConfigTableData: 'Tbl_ChecklistForm2' },
                'Table 6': { data: inputData.find(data => data.Tablename === 'Table 6'), appConfigTableData: 'Tbl_ChecklistForm3' },
                'Table 7': { data: inputData.find(data => data.Tablename === 'Table 7'), appConfigTableData: 'Tbl_ChecklistForm4' }
            };
            for (const tableName in tableDataMap) {
                const tableInfo = tableDataMap[tableName];
                const tableData = tableInfo.data;
                const appConfigTableData = inputData.find(data => data.Tablename === tableInfo.appConfigTableData);

                if (tableData && tableData.TemplateData.length > 0 && appConfigTableData) {
                    inputData.filter(data => data.Tablename === tableName).forEach(data => {
                        const appConfigTableKeys = new Set(appConfigTableData.TemplateData);
                        data.TemplateData.forEach(item => {
                            Object.keys(item).forEach(key => {
                                if ((key === 'Observation' || key === 'PageNumber' || key === 'CoverageSpecificationsMaster' || key === 'ChecklistQuestions') && item[key] === "Details not available in the document") {
                                    item[key] = '   ';
                                }
                                else if ((item[key] === null || item[key] === "") && appConfigTableKeys.has(key)) {
                                    item[key] = 'Details not available in the document';
                                }
                            });
                        });
                    });
                }
            }
            const tableData2 = inputData.find((data) => data.Tablename === tableName);

            if (!tableData2) {
                return;
            }

            const table22sonCopy = tableData2.TemplateData;
            const itemArray = ["CurrentTermPolicyListed", "PriorTermPolicyListed", "ProposalListed", "BinderListed", "ScheduleListed", "QuoteListed", "ApplicationListed", "CurrentTermPolicyListed1", "CurrentTermPolicyAttached"];
            for (let i = 0; i < table22sonCopy.length; i++) {
                const obj = table22sonCopy[i];
                let allDetailsNotAvailable = true;

                for (let j = 0; j < itemArray.length; j++) {
                    const key = itemArray[j];

                    if (obj[key] !== 'Details not available in the document') {
                        allDetailsNotAvailable = false;
                        break;
                    }
                    if (obj[key] !== 'MATCHED') {
                        allDetailsNotAvailable = false;
                        break;
                    }
                }
                if (allDetailsNotAvailable) {
                    obj.Observation = '';
                    obj.PageNumber = '';
                }
            }

            let tableColumnKeys = [];
            if (table22sonCopy && table22sonCopy?.length > 0) {
                const allKeys = Object.keys(table22sonCopy[0]);
                allKeys.map((e) => {
                    if (e) {
                        let keyHasData = table22sonCopy?.filter((f) => (f[e] != null && f[e] !== "") || (e == "Lob" && tableData2?.isMultipleLobSplit) || (e == "ChecklistQuestions" && (f[e] === null || f[e] === "")) || (e == "CoverageSpecificationsMaster" && (f[e] === null || f[e] === "")));
                        if (keyHasData?.length > 0) {
                            tableColumnKeys.push(e);
                        }
                    }
                });
                if (!tableColumnKeys?.includes('Observation')) {
                    tableColumnKeys.push('Observation');
                }
                if (!tableColumnKeys?.includes('PageNumber')) {
                    tableColumnKeys.push('PageNumber');
                }
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
                    Id,
                    JobId,
                    Jobid,
                    CreatedOn,
                    UpdatedOn,
                    Columnid,
                    IsDataForSp,
                    ...filteredItem
                } = item;
                return filteredItem;

            });

            let header = Object.keys(table2json[0]);
            header = header.filter(f => !["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"]?.includes(f));
            const value = Object.values(table2json);
            const policyLOBValues = value.map(item => item["PolicyLob"]);
            const table1Data = mainData && mainData?.length > 0 ? mainData?.find((f) => f?.Tablename == 'Table 1') : {};
            let heaterLob = '';

            if (tableName === "Table 3" && table1Data && table1Data?.TemplateData && table1Data?.TemplateData?.length > 0) {
                const headerPolicyLob = table1Data?.TemplateData.map((e) => e.PolicyLob);
                const filteredLob = Array.from(new Set(headerPolicyLob?.filter((f) => f != '' && f)));
                heaterLob = filteredLob[0];
            }

            let headerRows1 = [];
            let rowIndexForLOBStart = 0;
            let rowIndexForLOBEnd = 0;
            if (tableName === "Table 3") {
                rowIndexForLOBStart = basedata[basedata?.length - 1]?.r + 3;
                headerRows1 = [
                    {
                        r: basedata[basedata?.length - 1]?.r + 3,
                        rs: 1,
                        c: 1,
                        cs: header.length + 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [{ v: heaterLob || policyLOBValues[0], ff: "Tahoma", fs: 10 }] },
                            m: heaterLob || policyLOBValues[0],
                            v: heaterLob || policyLOBValues[0],
                            ff: "\"Tahoma\"",
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                ]

            };

            const excludedColumns = ["PolicyLob", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill", "PageNumber", "Observation"];
            let headers = Object.keys(table2json[0]).filter(headerw => !excludedColumns.includes(headerw));
            if (policyLOBValues && policyLOBValues?.length > 0 && policyLOBValues[0] === 'Are the forms and endorsements attached, listed in current term policy?') {
                const indexListed = headers.indexOf("CurrentTermPolicyListed");
                const indexAttached = headers.indexOf("CurrentTermPolicyAttached");
                if (indexListed !== -1 && indexAttached !== -1 && indexAttached > indexListed) {
                    // Swap the elements at the identified indices
                    [headers[indexListed], headers[indexAttached]] = [headers[indexAttached], headers[indexListed]];
                }
            }
            const removalCode = headers.map(item => (tableName !== "Table 3" && item === "CoverageSpecificationsMaster") ? policyLOBValues[0] : item);

            headerRows1 = [
                ...headerRows1,
                ...removalCode.map((item, index) => {
                    apiDataConfig.demo.config.merge[`${tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3}_${1 + index}`] = {
                        "r": tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3,
                        "c": 1 + index,
                        "rs": 2,
                        "cs": 1
                    }

                    return {
                        r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3,
                        rs: 2,
                        c: 1 + index,
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [{ v: item, ff: "Tahoma", fs: 10 }] },
                            m: item,
                            v: item,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }),
            ];

            //add documentviewer
            if (needDocumentViewer) {
                apiDataConfig.demo.config.merge[`${tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3}_${1 + headerRows1[headerRows1?.length - 1]?.c}`] = {
                    "r": tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3,
                    "c": 1 + headerRows1[headerRows1?.length - 1]?.c,
                    "rs": 2,
                    "cs": 1
                }

                const DocumentViewer = [{
                    r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3,
                    rs: 2,
                    c: 1 + headerRows1[headerRows1?.length - 1]?.c,
                    cs: 1,
                    v: {
                        ct: { fa: "@", t: "inlineStr", s: [{ v: 'Document Viewer', ff: "Tahoma", fs: 10 }] },
                        v: 'Document Viewer',
                        ff: "\"Tahoma\"",
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
                    rowIndexForLOBEnd = tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1;
                }
                if (index == 0) {
                    apiDataConfig.demo.config.merge[`${tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3}_${tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1}`] = {
                        "r": tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3,
                        "c": tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1,
                        "rs": 1,
                        "cs": 4,
                    }
                    return {
                        r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 4 : basedata[basedata?.length - 1]?.r + 3,
                        rs: 1,
                        c: tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1,
                        cs: 1,
                        v: {
                            ht: 0,
                            ct: { fa: "@", t: "inlineStr", s: [{ v: item, ff: "Tahoma", fs: 10 }] },
                            m: item,
                            v: item,
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
                        r: tableName === "Table 3" ? basedata[basedata?.length - 1]?.r + 5 : basedata[basedata?.length - 1]?.r + 4,
                        rs: 1,
                        c: tableName === "Table 3" ? headerRows1.length + index - 1 : headerRows1.length + index,
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [{ v: item, ff: "Tahoma", fs: 10 }] },
                            m: item,
                            v: item,
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
                apiDataConfig.demo.config.merge[`${rowIndexForLOBStart}_${1}`] = {
                    "r": rowIndexForLOBStart,
                    "c": 1,
                    "rs": 1,
                    "cs": rowIndexForLOBEnd - 1
                }
            }

            let headerRows1Values = [];
            let rowIndex = defaultHeaderRows1[defaultHeaderRows1.length - 1]?.r + 1;
            const actionColumnKeys = ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
            headers = [...headers, ...actionColumnKeys];
            table2json.map((item, cIndex) => {

                let rowHeight = 60;
                headers.map((key, rIndex) => {
                    if (item[key] !== null) {

                        rowIndex = headerRows1Values?.length == 0 ? rowIndex : headerRows1Values?.length > 0 && rIndex == 0 ? headerRows1Values[headerRows1Values.length - 1]?.r + 1 : headerRows1Values[headerRows1Values.length - 1]?.r;
                        let text = item[key]?.split('~~');
                        let ct = [];
                        let fs = 10;

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

                        if (text && text?.length > 0) {
                            const ttableData2 = inputData.find(data => data.Tablename === "Table 2");
                            const tt2 = removeNullValues(ttableData2.TemplateData[0], '');
                            const tt2keys = Object.keys(tt2);
                            const applicationIndex = tt2keys.indexOf("Application");
                            const keyBeforeApplication = tt2keys[applicationIndex - 1];

                            const ttableData3 = inputData.find(data => data.Tablename === "Table 3");
                            const tt3 = removeNullValues(ttableData3.TemplateData[0], "Lob");
                            const tt3keys = tt3 == undefined ? tt2keys : Object.keys(tt3);
                            const tb3applicationIndex = tt3keys.indexOf("Application");
                            const tb3keyBeforeApplication = tt3keys[tb3applicationIndex - 1];
                            text?.map((e, splitIndex) => {
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
                                        "v": "\r\n" + e.trim() + "\r\n"
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
                                        "v": "\r\n" + e.trim() + "\r\n"
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
                                        "v": "\r\n" + e.trim() + "\r\n"
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
                                        "v": "\r\n" + e.trim() + "\r\n"
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
                                        "v": "\r\n" + e.trim() + "\r\n"
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

                                else if (key === "PriorTermPolicyListed" && item["CurrentTermPolicyListed"] && item["PriorTermPolicyListed"]?.trim() != item["CurrentTermPolicyListed"]?.trim()
                                    && !(item["PriorTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "CurrentTermPolicyAttached" && item["CurrentTermPolicyAttached"]?.trim() != item["CurrentTermPolicyListed"]?.trim()
                                    && !(item["CurrentTermPolicyAttached"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "PriorTermPolicyListed" && item["CurrentTermPolicyListed1"] && item["PriorTermPolicyListed"]?.trim() != item["CurrentTermPolicyListed1"]?.trim()
                                    && !(item["PriorTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed1"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed1"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "QuoteListed" && item["QuoteListed"]?.trim() != item["CurrentTermPolicyListed"]?.trim()
                                    && !(item["QuoteListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "ProposalListed" && item["ProposalListed"]?.trim() != item["CurrentTermPolicyListed"]?.trim()
                                    && !(item["ProposalListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "BinderListed" && item["BinderListed"]?.trim() != item["CurrentTermPolicyListed"]?.trim()
                                    && !(item["BinderListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "ScheduleListed" && item["ScheduleListed"]?.trim() != item["CurrentTermPolicyListed"]?.trim()
                                    && !(item["ScheduleListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "ApplicationListed" && item["ApplicationListed"]?.trim() != item["CurrentTermPolicyListed"]?.trim()
                                    && !(item["ApplicationListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicyListed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicyListed"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "PriorTermPolicy" && item["PriorTermPolicy"]?.trim() != item["CurrentTermPolicy"]?.trim()
                                    && !(item["PriorTermPolicy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicy"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "Quote" && item["Quote"]?.trim() != item["CurrentTermPolicy"]?.trim()
                                    && !(item["Quote"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicy"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "Proposal" && item["Proposal"]?.trim() != item["CurrentTermPolicy"]?.trim()
                                    && !(item["Proposal"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicy"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "Binder" && item["Binder"]?.trim() != item["CurrentTermPolicy"]?.trim()
                                    && !(item["Binder"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicy"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (key === "Schedule" && item["Schedule"]?.trim() != item["CurrentTermPolicy"]?.trim()
                                    && !(item["Schedule"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                        || item["CurrentTermPolicy"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item["CurrentTermPolicy"]?.split('~~')[splitIndex]?.split(" ");

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
                                else if (


                                    key === "Application" &&

                                    (
                                        (item["Application"]?.trim() != item[keyBeforeApplication]?.trim() &&

                                            !(
                                                item["Application"]?.toLowerCase()?.replace(/\\r\\n/g, '').includes("details not available in the document") ||
                                                item[keyBeforeApplication]?.toLowerCase()?.replace(/\\r\\n/g, '').includes("details not available in the document")
                                            )
                                        )
                                        ||
                                        (item["Application"]?.trim() != item[tb3keyBeforeApplication]?.trim() &&
                                            !(
                                                item["Application"]?.toLowerCase()?.replace(/\\r\\n/g, '').includes("details not available in the document") ||
                                                item[tb3keyBeforeApplication]?.toLowerCase()?.replace(/\\r\\n/g, '').includes("details not available in the document")
                                            )
                                        )
                                    )
                                ) {

                                    let ptpSplitArray = e?.split(" ");
                                    let ctpSplitArray = item[keyBeforeApplication] ? item[keyBeforeApplication].split('~~')[splitIndex]?.split(" ") :
                                        item[tb3keyBeforeApplication] ? item[tb3keyBeforeApplication].split('~~')[splitIndex]?.split(" ") : []
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
                                        "v": e.trim() + " "
                                    });
                                }
                            })
                        }
                        if (key === "PageNumber") {
                            const textOfct = ct[ct.length - 1]?.v?.replace('\r\n', ' ');
                            if (textOfct && ct?.length > 0) {
                                ct[ct.length - 1]["v"] = textOfct;
                            }
                        }
                        headerRows1Values.push({
                            r: rowIndex,
                            c: rIndex + 1 + (actionColumnKeys?.includes(key) ? 1 : 0),
                            v: {
                                ct: { fa: "General", t: "inlineStr", s: ct },
                                ff: "Tahoma",
                                fc: "#3b3737",
                                merge: null,
                                w: 55,
                                tb: '2',
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

                        if (rowHeight != undefined && rowHeight != null) {
                            const rowHeight = parseInt(maxLength / 3 + 50);
                            apiDataConfig.demo.config.rowlen[`${rowIndex}`] = rowHeight;

                            if (rIndex == 0) {
                                apiDataConfig.demo.config.rowlen[`${rowIndex}`] = rowHeight;
                            }
                        }
                    }

                })
            });


            apiDataConfig.demo.config.borderInfo.push({
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

            headerRows1 = [...headerRows1, ...defaultHeaderRows1, ...headerRows1Values];

            const allRows2 = [...headerRows1];
            // Sort the rows by rowIndex if needed
            allRows2.sort((a, b) => a.r - b.r);

            // Add the rows to the dummyData2
            basedata.push(...allRows2);
            apiDataConfig.demo.celldata = basedata;
        };

        if (sheetRenderConfig?.PolicyReviewChecklist == 'true') {
            renderTable1();
        }


        const formTable1 = () => {
            if (isFormApplicable && isFormApplicable == true) {
                const formTableData1 = formCompareData.find((data) => data.Tablename === "FormTable 1");

                if (formTableData1) {
                    const formtable1 = formTableData1.TemplateData;

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
                                ff: item.ff,
                                fs: 24,
                                merge: mergeConfig,
                                fc: item.fc,
                            }
                        });
                    });

                    formtable1.map((item, index) => {
                        if (item["Headers"] != null && item["Headers"] != undefined) {
                            if (item["Headers"] == "") {
                                sheetDataTable4.push({
                                    r: rowIndexOfTable1 + index,
                                    c: 1,
                                    v: {
                                        ct: { fa: "@", t: "inlineStr", s: [{ v: " " }] },
                                        m: " ",
                                        v: " ",
                                        merge: null,
                                        bg: "rgb(139,173,212)",
                                        tb: '2',
                                    }
                                });
                            } else {
                                sheetDataTable4.push({
                                    r: rowIndexOfTable1 + index,
                                    c: 1,
                                    v: {
                                        ct: { fa: "@", t: "inlineStr", s: [{ v: item["Headers"], ff: "Tahoma", fs: 10 }] },
                                        m: item["Headers"],
                                        v: item["Headers"],
                                        ff: "Tahoma",
                                        merge: null,
                                        bg: "rgb(139,173,212)",
                                        tb: '2',
                                    }
                                });
                            }

                            const tidleValue = item["NoColumnName"] !== null && item["NoColumnName"] != undefined ? item["NoColumnName"].replace(/~~/g, "\n") : "";

                            sheetDataTable4.push({
                                r: rowIndexOfTable1 + index,
                                c: 2,
                                v: {
                                    ct: { fa: "@", t: "inlineStr" },
                                    m: tidleValue,
                                    v: tidleValue,
                                    ff: "Tahoma",
                                    merge: null,
                                    tb: '2',
                                }
                            });
                        }
                    });

                    const dummyData1 = [];
                    const allFormRows = [...sheetDataTable4, ...sheetDataTable3];
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
                    formCompareData.map((e, index) => {
                        if (e?.Tablename != 'FormTable 1' && (e?.TemplateData?.length >= 3 || e?.TemplateData?.length < 3)) {
                            let filteredData = FormCompare_appconfigdata.forms.celldata.filter((f, index) => f != null || !f);
                            formTable2([...filteredData], e?.Tablename);
                        }
                    });
                }
            }
            renderLuckySheet(true, '', false);
        }

        const formTable2 = (combinedata1, tableName) => {
            if (!Array.isArray(combinedata1)) {
                return;
            }

            const basedata = [...combinedata1];

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
                    Id,
                    JobId,
                    Jobid,
                    CreatedOn,
                    UpdatedOn,
                    IsMatched,
                    ...filteredItem
                } = item;
                return filteredItem;

            });
            const header = Object.keys(formtable2[0]);
            const value = Object.values(formtable2);
            const policyLOBValues = value.map(item => item["PolicyLob"]);

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
                            ct: { fa: "@", t: "inlineStr", s: [{ v: tableName === "FormTable 2" ? "Unmatched Forms" : "Matched Forms", ff: "Tahoma", fs: 10 }] },
                            m: tableName === "FormTable 2" ? "Unmatched Forms" : "Matched Forms",
                            v: tableName === "FormTable 2" ? "Unmatched Forms" : "Matched Forms",
                            ff: "\"Tahoma\"",
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                ]
            };



            const excludedColumns = ["PolicyLob", "PageNumber", "Observation"];
            let headers = Object.keys(formtable2[0]).filter(headerw => !excludedColumns.includes(headerw));

            const removalCode = headers.map(item => (item === "CoverageSpecificationsMaster") ? policyLOBValues[0] : item);

            headerRows1 = [
                ...headerRows1,
                ...removalCode.map((item, index) => {
                    if (removalCode?.length === index + 1) {
                        rowIndexForLOBEnd = headerRows1.length + index + 1;
                    }
                    return {
                        r: basedata[basedata?.length - 1]?.r + 3,
                        rs: 2,
                        c: 1 + index,
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [{ v: item, ff: "Tahoma", fs: 10 }] },
                            m: item,
                            v: item,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                })
            ];

            if (headerRows1?.length > 0)


                if (tableName === "FormTable 2" || tableName === "FormTable 3") {
                    FormCompare_appconfigdata.forms.config.merge[`${rowIndexForLOBStart}_${1}`] = {
                        "r": rowIndexForLOBStart,
                        "c": 1,
                        "rs": 1,
                        "cs": rowIndexForLOBEnd - 1
                    }
                }


            let headerRows1Values = [];
            let rowIndex = basedata[basedata?.length - 1]?.r + 4;
            formtable2.map((item, cIndex) => {
                let rowHeight = 21;
                headers.map((key, rIndex) => {
                    rowIndex = headerRows1Values?.length == 0 ? rowIndex : headerRows1Values?.length > 0 && rIndex == 0 ? headerRows1Values[headerRows1Values.length - 1]?.r + 1 : headerRows1Values[headerRows1Values.length - 1]?.r;
                    let text = item[key].toString().split('~~');
                    let ss = [];
                    let fs = 10;

                    if (text && text?.length > 0) {
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
                                    "v": e.trim() + "\r\n"
                                });
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
                                    "v": e.trim() + "\r\n"
                                });
                            }
                        });
                    }
                    headerRows1Values.push({
                        r: rowIndex,
                        c: rIndex + 1,
                        v: {
                            ct: { fa: "General", t: "inlineStr", s: ss },
                            merge: null,
                            w: 55,
                            tb: '2',
                        }
                    });
                    if (text && rowHeight < parseInt(item[key]?.length / 2 + 20)) {
                        rowHeight = parseInt(item[key]?.length / 2 + 20);
                        FormCompare_appconfigdata.forms.config.rowlen[`${rowIndex}`] = rowHeight;
                    }
                    if (rowIndex == 0) {
                        FormCompare_appconfigdata.forms.config.rowlen[`${rowIndex}`] = rowHeight;
                    }
                })
            });

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
                            headerRows1[headerRows1?.length - 1].c
                        ],
                        "row_focus": headerRows1[0]?.r,
                        "column_focus": headerRows1[0]?.c
                    }
                ]
            });

            headerRows1 = [...headerRows1, ...headerRows1Values];
            const allRows2 = [...headerRows1];
            allRows2.sort((a, b) => a.r - b.r);

            basedata.push(...allRows2);

            FormCompare_appconfigdata.forms.celldata = basedata;

        };
        if (isFormApplicable && sheetRenderConfig?.FormsCompare == 'true') {
            formTable1();
        }


        const exclusionTable = () => {
            const basedata = [];
            const data = props.exclusionRenderData;
            const dataMap = data;
            if (dataMap && dataMap?.length > 0 && !Array.isArray(dataMap[0])) {
                const exclusionjson = dataMap.map(item => {
                    const {
                        Id,
                        JobId,
                        CreatedOn,
                        UpdatedOn,
                        ...filteredItem
                    } = item;
                    return filteredItem;
                });
                const headers = Object.keys(exclusionjson[0]);
                let headerRows1 = headers.map((item, index) => {
                    return {
                        r: 0,
                        rs: 2,
                        c: index,
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [{ v: item, ff: "Tahoma", fs: 10 }] },
                            m: item,
                            v: item,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                });

                headerRows1 = [...headerRows1];

                let headerRows1Values = [];
                let rowIndex = headerRows1[headerRows1.length - 1]?.r + 1;
                let rowHeight = 60;
                let fs = 10;

                exclusionjson.map((item, indexr) => {

                    headers.map((key, rIndex) => {
                        let text = item[key].toString().split('~~');
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

                        headerRows1Values.push({
                            r: rowIndex + indexr,
                            c: rIndex,
                            v: {
                                ct: { fa: "General", t: "inlineStr", s: ss },
                                merge: null,
                                ff: "\"Tahoma\"",
                                w: 55,
                                tb: '2',
                            }
                        });

                        if (text && rowHeight < parseInt(item[key]?.length / 2 + 10)) {
                            rowHeight = parseInt(item[key]?.length / 2 + 10);
                            exclusionDatafigdata.exclusion.config.rowlen[`${rowIndex}`] = rowHeight;
                        }

                    })
                });
                exclusionDatafigdata.exclusion.config.borderInfo.push({
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
                                3
                            ],
                            "row_focus": headerRows1[0]?.r,
                            "column_focus": headerRows1[0]?.c
                        }
                    ]
                });

                headerRows1 = [...headerRows1, ...headerRows1Values];
                const allRows2 = [...headerRows1];

                allRows2.sort((a, b) => a.r - b.r);
                basedata.push(...allRows2);
                exclusionDatafigdata.exclusion.celldata = basedata;
            }
        };

        if (sheetRenderConfig?.Exclusion == 'true') {
            exclusionTable();
        }

        renderLuckySheet();
    }, [props.data, props?.sheetRenderConfig]);


    function removeNullValues(obj, key) {
        for (const prop in obj) {
            if (obj[prop] === null && prop != key) {
                delete obj[prop];
            }
        }
        return obj;
    }



    let isLuckysheetRendered = false;
    const renderLuckySheet = async (needConfigAdjustment, sheetConfig, isDelete) => {
        let data1 = FormCompare_appconfigdata.forms.celldata;
        let configData1 = FormCompare_appconfigdata.forms.config;
        if (data1 && isFormApplicable) {
            FormCompare_appconfigdata.forms.celldata = data1.map((item) => {
                if (item?.v && item?.v?.fs && item?.v?.fs > 9) {
                    item.v.fs = item.v.fs - 5;
                } else if (item?.v && item?.v?.ct && item?.v?.ct?.s?.length > 0) {
                    item?.v?.ct?.s.map((subItem, index) => {
                        if (subItem?.fs) {
                            item.v.ct.s[index].fs = 8;
                        } else {
                            item.v.ct.s[index]["fs"] = 7;
                        }
                    });
                } else {
                    if ((item?.v?.m || item?.v?.v) && !item?.v?.fs) {
                        item.v["fs"] = 8;
                    }
                }
                return item;
            });

        }
        if (configData1) {
            if (configData1?.columnlen) {
                const keys = Object.keys(configData1?.columnlen);
                if (keys?.length > 0) {
                    keys.map((key) => {
                        configData1.columnlen[key] = 220;
                    })
                }
            }
            if (configData1?.rowlen) {
                const keys = Object.keys(configData1?.rowlen);
                if (keys?.length > 0) {
                    keys.map((key) => {
                        configData1.rowlen[key] = configData1.rowlen[key] - 10 > 30 ? configData1.rowlen[key] - 10 : 30;
                    })
                }
            }
            FormCompare_appconfigdata.forms.config = configData1;
        }
        if ((!isLuckysheetRendered || isDelete) && luckysheet) { // Check if Luckysheet is not rendered and luckysheet instance exists
            isLuckysheetRendered = true;

            const qacDataSet = await getQACData(jobId, token);
            if (luckysheet) {
                if (needConfigAdjustment) {
                    let data = apiDataConfig.demo.celldata;
                    let configData = apiDataConfig.demo.config;
                    if (data) {
                        apiDataConfig.demo.celldata = data.map((item) => {
                            if (item?.v && item?.v?.fs && item?.v?.fs > 9) {
                                item.v.fs = item.v.fs - 5;
                            } else if (item?.v && item?.v?.ct && item?.v?.ct?.s?.length > 0) {
                                item?.v?.ct?.s.map((subItem, index) => {
                                    if (subItem?.fs) {
                                        item.v.ct.s[index].fs = 8;
                                    } else {
                                        item.v.ct.s[index]["fs"] = 7;
                                    }
                                });
                            } else {
                                if ((item?.v?.m || item?.v?.v) && !item?.v?.fs) {
                                    item.v["fs"] = 7;
                                }
                            }
                            return item;
                        });

                    }
                    if (configData) {
                        if (configData?.columnlen) {
                            const keys = Object.keys(configData?.columnlen);
                            if (keys?.length > 0) {
                                keys.map((key) => {
                                    configData.columnlen[key] = 250;
                                })
                            }
                        }
                        if (configData?.rowlen) {
                            const keys = Object.keys(configData?.rowlen);
                            if (keys?.length > 0) {
                                keys.map((key) => {
                                    configData.rowlen[key] = configData.rowlen[key] - 10 > 15 ? configData.rowlen[key] - 10 : 15;
                                })
                            }
                        }
                        apiDataConfig.demo.config = configData;
                    }
                } else {
                    if (isDelete) {

                        if (sheetConfig[0]?.top == undefined || sheetConfig[0]?.top == null) {
                            sheetConfig[0].top = sessionStorage.getItem("sheetConfigTop");
                        } else {
                            sessionStorage.setItem("sheetConfigTop", sheetConfig[0]?.top);
                        }
                    }

                }
                // Create options for Luckysheet
                const sheetRenderConfig = props?.sheetRenderConfig;
                let sheetDataSet = [];
                if (sheetRenderConfig?.PolicyReviewChecklist == 'true') {
                    sheetDataSet = [apiDataConfig.demo];
                } else if (sheetRenderConfig?.FormsCompare == 'true') {
                    sheetDataSet = [FormCompare_appconfigdata.forms];
                } else if (sheetRenderConfig?.Exclusion == 'true') {
                    sheetDataSet = [exclusionDatafigdata.exclusion];
                } else if (sheetRenderConfig?.QAC_not_answered_questions == 'true' && qacDataSet?.canRender) {
                    sheetDataSet = [qacDataSet?.data];
                }

                const selectedOptions = dropDownOption?.map((sheet) => ({ key: sheet, text: sheet }));
                setSheetDropOption(selectedOptions);

                const options = {
                    container: "luckysheet", // Container ID
                    showinfobar: false,
                    showsheetbar: true,
                    lang: 'en',
                    data: sheetDataSet,
                    enableAddRow: true,
                    showtoolbar: true,
                    row: 2,
                    column: 3,
                    allowUpdate: true,
                    enableAddBackTop: true,
                    sheetRightClickConfig: {
                        delete: false, //Delete
                        copy: false, //Copy
                        rename: false, //Rename
                        color: false, //Change color
                        hide: false, //Hide, unhide
                        move: false, //Move to the left, move to the right
                    },
                    showsheetbarConfig: {
                        add: false, // Hide the Add Sheet button
                        menu: false, // Hide the menu button
                    },
                    showstatisticBar: true,
                    hook: {
                        workbookCreateAfter(json) {
                            luckysheet.setSheetZoom(1);// after rendering setting the screen zoom size to 0.65 for scroll support in chrome
                        },
                    },
                    cellRightClickConfig: {
                        insertRow: false, // insert row
                        insertColumn: false, // insert column
                        deleteRow: false, // delete the selected row
                        deleteColumn: false, // delete the se
                        deleteCell: false, // delete cell
                        clear: false, // clear content
                        sort: false, // sort selection
                        filter: false, // filter selection
                        chart: false, // chart generation
                        image: false, // insert picture
                        link: false, // insert link
                    },
                    showtoolbarConfig: {
                        moreFormats: false, //'More Formats'
                        sortAndFilter: false, //'Sort and filter'
                        link: false, // insert link
                        chart: false, // chart generation
                        print: false,//  print 
                        textRotateMode: false, //'Text Rotation Mode'
                        image: false, // 'Insert picture'
                        postil: false, //'comment'
                        dataVerification: false, // 'Data Verification'
                        splitColumn: false, //'Split column'
                        screenshot: false, //'screenshot'
                        findAndReplace: false, //'Find and Replace'
                    }
                };
                luckysheet.create(options);
                FormCompare_appconfigdata = {};
                apiDataConfig = {};
                exclusionDatafigdata = {};
            }
        }
    }

    const handleSheetChange = async (e) => {
        const value = e?.target?.value;
        setTimeout(() => {
            setSelectedSheet(value);
            selectChange(value);
        }, 3000);

    };

    return (
        <div>
            <div className="p2">
                <select
                    className="dropDown1"
                    value={selectedSheet}
                    onChange={handleSheetChange}
                >
                    {sheetsDropOption.map(option => (
                        <option key={option.key} value={option.key}>
                            {option.text}
                        </option>
                    ))}
                </select>
            </div>
            <div className="App" id="luckysheet" ref={luckyCss}></div>
        </div>
    );
}
