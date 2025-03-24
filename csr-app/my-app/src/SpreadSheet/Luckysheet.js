
import React, { useEffect, useRef, useState } from "react";
import { jobData, baseUrl, Disclaimer, Checklist, Formpolicydata, updateData, formCompare, excludedColumnlist, staticExclusionData, defaultData } from '../Services/Constants';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import {
    endorsementCheck, getQACData, getTextForUpdate, getPreviewChecklistDataForUpdate,
    getEndIdex, setDocumentDetails, getDocumentDetails, validateEndorsementEntry, mapLOBColumns, SetCheckListQuestionMasterData, updateFormsPHData,
    PageHighlighterProcess, processAndUpdateToken, tableDataFormatting, formTableDataFormatting, getObervationReplacerKey, getText, getTextByRequirement, getEmptyDataSet, isARType,
    splitPageKekFromText, getPageKey, getObservationKey, getOtherApplications, getIndexForForms, getTableApplicationColumns, getExistingPageKey, autoupdate, findTableForIndex,ExportData,
    getConfidenceScoreConfigStatus, getCsRespectiveColumn, getKeyByValue, getCsRespectiveColumn_formsCompare
} from "../Services/CommonFunctions"
import axios from "axios";
import { Icon } from '@fluentui/react';
import '../App.css';
import $ from 'jquery';
import { DialogComponent, InputDialogComponent, FindDialogComponent, DiscrepancyOptionsDialogComponent, FilterDialogComponent } from '../Services/dialogComponent';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { SimpleSnackbarWithOutTimeOut } from '../Components/SnackBar';
import {reorderTemplateData} from '../Services/CommonFunctions'
import { apiCallSwitch, updateGridAuditLog, TriggerBackUp } from '../Services/PreviewChecklistDataService';
import {auditProcessNames} from '../Services/enums';

export default function Luckysheet( props ) {
    const { selectChange } = props;
    const container = useRef();
    const luckysheet = window.luckysheet;
    const [ state, setState ] = useState( props?.data );
    const [ formstate, setFormState ] = useState( props.formCompareData );
    const [ exclusionstate, setExclusionState ] = useState( props.exclusionRenderData );
    const [ jobId, setJobId ] = useState( props?.selectedJob );
    // const [ apiCallInProgress, setApiCallInProgress ] = useState( true );
    const [ autoprogress, setautoprogress ] = useState( false );
    const [ msgVisible, setMsgVisible ] = useState( false );
    const [ issavessheet, setIssavessheet ] = useState( false );
    // const sheetDatas = luckysheet.getSheetData();
    const [ sheetState, setsheetState ] = useState( [] );
    const [ openDialog, setOpenDialog ] = useState( false );
    const [ findDialog, setfindDialog ] = useState( false );
    const [ dropDialog, setDropDialog ] = useState( false );
    const [ searchResults, setSearchResults ] = useState( [] );
    // const [ isFormApplicable, setIsFormApplicable ] = useState( props?.formCompareData[0]?.isFormCompareApplicable );
    const [ isFormApplicable, setIsFormApplicable ] = useState( true );
    const [ openInputDialog, setOpenInputDialog ] = useState( false );
    const [ openFilterDialog, setOpenFilterDialog ] = useState( false );
    const [ msgClass, setMsgClass ] = useState( '' );
    const [ msgText, setMsgText ] = useState( '' );
    const [ tablenameArray, setTablenameArray ] = useState( [] );
    let token = sessionStorage.getItem( 'token' );
    let nonEditable = sessionStorage.setItem( 'nonEditable', true );
    let onUpdateClickCalled = sessionStorage.setItem( 'onUpdateClickCalled', false );
    // const [ token, setToken ] = useState( sessionStorage.getItem( 'token' ) );
    const [ setectedRowIndex, setSelectedRowIned ] = useState( '' );
    const [ selectedRowIndexRange, setSelectedRowIndexRange ] = useState( [] );
    const [ tableColumnDetails, setTableColumnDetails ] = useState( { "Table 1": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 2": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 3": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 4": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 5": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 6": { "columnNames": {}, "range": { "start": "", "end": "" } }, "Table 7": { "columnNames": {}, "range": { "start": "", "end": "" } } } );
    const [ formTableColumnDetails, setFormTableColumnDetails ] = useState( { "FormTable 1": { "columnNames": {}, "range": { "start": "", "end": "" } }, "FormTable 2": { "columnNames": {}, "range": { "start": "", "end": "" } }, "FormTable 3": { "columnNames": {}, "range": { "start": "", "end": "" } } } );
    const [ exTableColumnDetails, setExTableColumnDetails ] = useState( { "ExTable 1": { "columnNames": {}, "range": { "start": "", "end": "" } } } );
    const [ dependencyColumn, setDependencyColumn ] = useState();
    const [ hasMultipleRowsSelected, setHasMultipleRowsSelected ] = useState( false );
    const [ filterSelectionData, setFilterSelectionData ] = useState( null );
    const [ uparrowValue, setUparrowValue ] = useState( null );
    const [ uparrowsecoundValue, setUparrowsecoundValue ] = useState( null );
    const [ downarrowValue, setDownarrowValue ] = useState( false );
    const [ secoundtablerange, setSecoundtablerange ] = useState( [] );
    const [ uparrowlastValue, setUparrowlastValue ] = useState( null );
    const [sheetsDropOption, setSheetDropOption] = useState([]);
    const [dropDownOption, setDropDownOption] = useState(props?.sheetOptionSet);
    const [selectedSheet, setSelectedSheet] = useState(props?.selectedSheet || dropDownOption[0]);
    const [lockingIndex, setLockingIndex] = useState({});
    const exclusionApplicableIdx = [ 2 ];

    const brokerId = jobId.slice( 0, 4 );
    const apiDataConfig = {
        demo: {
            name: "PolicyReviewChecklist", // Worksheet name
            color: "", // Worksheet color
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
                    // "0": 300,
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
            status: "1", // Activation status
            order: "0", // The order of the worksheet
            hide: 0, // Whether to hide
            column: 50, // Number of columns
            row: 50, // Number of rows
            celldata: [],// Original cell data set
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
    };
    const FormCompare_appconfigdata = {

        forms: {
            name: "Forms Compare", // Worksheet name
            color: "", // Worksheet color
            config: {
                merge: {
                    "1_1": {
                        "rs": 1,
                        "cs": 2,
                        "r": 0,
                        "c": 1
                    },
                },
                // sheetcheck: "Formscompare",
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
                    "8": 220,
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
            //status: "1", // Activation status
            order: "0", // The order of the worksheet
            hide: 0, // Whether to hide
            column: 50, // Number of columns
            row: 50, // Number of rows
            celldata: [],

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
            sheets: [],
        },
    }
    const exclusionDatafigdata = {

        exclusion: {
            name: "Exclusion",
            config: {
                merge: {},
                borderInfo: [],
                columnlen: {
                    "0": 200,
                    "1": 600,
                    "2": 250,
                    "3": 250,
                    "4": 250,
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

    useEffect( () => {
        sessionStorage.setItem("IsDataRendering",true);
        sessionStorage.setItem("IsAutoUpdate",true); //for auto update allow while change the sheet 
        const mainData = props.data;
        const sheetRenderConfig = props?.sheetRenderConfig;
        const formCompareData = props.formCompareData;
        setTablenameArray( mainData.map( item => item.Tablename ) );
        // SetCheckListQuestionMasterData( token, jobId );
        setDocumentDetails(jobId);
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
            const tableData1 = mainData.find( ( data ) => data.Tablename === "Table 1" );
            if ( tableData1 )
            {
                const table1json = tableData1.TemplateData;

                let sheetDataTable1 = [];
                let sheetDataTable2 = [];

                const rowIndexOfTable1 = 4
                const textBlockData = renderTextBlock();
                const listData = renderList();

                textBlockData.forEach( ( item, index ) => {
                    const mergeConfig = apiDataConfig.demo.config.merge[ "0_1" ];

                    sheetDataTable2.push( {
                        r: index + mergeConfig.r, // Adjust row index based on merge configuration
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            bl: 1,
                            ff: item.ff,
                            fs: 12,
                            merge: mergeConfig, // Use the merge configuration
                            fc: item.fc,
                            // tb: '55',
                        }
                    } );
                } );

                listData.forEach( ( item, index ) => {
                    const mergeConfig = apiDataConfig.demo.config.merge[ "1_1" ];

                    sheetDataTable2.push( {
                        r: 1 + mergeConfig.r, // Adjust row index based on merge configuration
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            ff: item.ff,
                            // bl: 0,
                            fs: 17,
                            // ff: "Arial",
                            merge: mergeConfig, // Use the merge configuration
                            fc: item.fc,
                            // tb: '55',
                        }
                    } );
                } );

                table1json.map( ( item, index ) => {
                    if ( item[ "Headers" ] != null &&  item[ "Headers" ] != undefined )
                    {
                        if (item[ "Headers" ] == "") {
                            sheetDataTable1.push( {
                                r: rowIndexOfTable1 + index, // Start from row 1 for headers
                                c: 1, // Display headers in the first column
                                v: {
                                    ct: { fa: "@", t: "inlineStr", s: [ { v: " " } ] },
                                    m: " ", // Use "Headers" as the value
                                    v: " ", // Use "Headers" as the value
                                    merge: null,
                                    bg: "rgb(139,173,212)",
                                    tb: '2',
                                }
                            } );
                        } else {
                            sheetDataTable1.push( {
                                r: rowIndexOfTable1 + index, // Start from row 1 for headers
                                c: 1, // Display headers in the first column
                                v: {
                                    ct: { fa: "@", t: "inlineStr", s: [ { v: item[ "Headers" ] , ff: "Tahoma" , fs: 10} ] },
                                    m: item[ "Headers" ], // Use "Headers" as the value
                                    v: item[ "Headers" ], // Use "Headers" as the value
                                    ff: "Tahoma",
                                    merge: null,
                                    bg: "rgb(139,173,212)",
                                    tb: '2',
                                }
                            } );
                        }
                       
                        const tidleValue = item[ "NoColumnName" ] !== null && item[ "NoColumnName" ] != undefined ? item[ "NoColumnName" ].replace( /~~/g, "\n" ) : "";

                        sheetDataTable1.push( {
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
                        } );
                    }
                } );

                const dummyData = [];
                const matchedUnMatchedFilter = [
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
                                        "m": "All Variances",
                                        "v": "All Variances"
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
                        "r": 7,
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

                const allRows = [...sheetDataTable1, ...matchedUnMatchedFilter, ...sheetDataTable2 ];
                if ( sheetDataTable1 && sheetDataTable1?.length > 0 )
                {
                    const tableColumnDetails1 = tableColumnDetails;
                    tableColumnDetails1[ "Table 1" ] = { "columnNames": table1json.map( ( e ) => e?.Headers ), "range": { "start": 0, "end": sheetDataTable1[ sheetDataTable1?.length - 1 ]?.r } }
                    setTableColumnDetails( tableColumnDetails1 );
                }

                // Sort the rows by rowIndex if needed
                allRows.sort( ( a, b ) => a.r - b.r );

                // Add the rows to the dummyData
                dummyData.push( ...allRows );
                apiDataConfig.demo.config.borderInfo.push( {
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
                                4,
                                7
                            ],
                            "column": [
                                4,
                                4
                            ],
                            "row_focus": 4,
                            "column_focus": 4
                        }
                    ]
                } );
                apiDataConfig.demo.celldata = dummyData;

                //table1 border info styles
                allRows.forEach( ( row ) => {
                    if ( sheetDataTable1.includes( row ) )
                    {
                        apiDataConfig.demo.config.borderInfo.push( {
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
                        } );
                    }
                } );

                const excludedTablenames = [ "JobHeader", "JobCommonDeclaration", "JobCoverages", "Tbl_ChecklistForm1", "Tbl_ChecklistForm2", "Tbl_ChecklistForm3", "Tbl_ChecklistForm4" ];
                mainData.map( ( e, index ) => {
                    if ( !excludedTablenames.includes( e?.Tablename ) && e?.Tablename != 'Table 1' && e?.TemplateData?.length >= 1 )
                    {
                        let filteredData = apiDataConfig.demo.celldata.filter( ( f, index ) => f != null || !f );
                        renderTable2( [ ...filteredData ], e?.Tablename );
                    }
                } );

                renderLuckySheet( true, '', false );

            }
        };
        const renderTable2 = ( combinedata1, tableName ) => {
            
            
            if ( !Array.isArray( combinedata1 ) )
            {
                //console.error( "combinedata1 is not an array:", combinedata1 );
                return;
            }
            const tableColumnNamesOfValid = {};
            const needDocumentViewer = true;
            const DefaultColumns = [ "Actions on Discrepancy (from AMs)", "Actions on Discrepancy", "Request Endorsement", "Notes", "Notes(Free Fill)" ];
            const basedata = [ ...combinedata1 ];
            const inputData = mainData;

            inputData.forEach( item => {
                if ( item.Tablename != 'Table 1' )
                {
                    if ( Array.isArray( item.TemplateData ) )
                    {
                        item.TemplateData.forEach( data => {
                            if ( typeof data === 'object' && data !== null )
                            {
                                if ( !data.hasOwnProperty( 'CoverageSpecificationsMaster' ) )
                                {
                                    data.CoverageSpecificationsMaster = null;
                                }
                            }
                        } );
                    }
                }
            } );
            // reorderTemplateData(inputData, 'Table 2', 'JobCommonDeclaration');
            // reorderTemplateData(inputData, 'Table 3', 'JobCoverages');
            // reorderTemplateData(inputData, 'Table 4', 'Tbl_ChecklistForm1');
            // reorderTemplateData(inputData, 'Table 5', 'Tbl_ChecklistForm2');
            // reorderTemplateData(inputData, 'Table 6', 'Tbl_ChecklistForm3');
            // reorderTemplateData(inputData, 'Table 7', 'Tbl_ChecklistForm4');

            const targetTablenames = [ "JobCommonDeclaration", "JobCoverages", "Tbl_ChecklistForm1", "Tbl_ChecklistForm2", "Tbl_ChecklistForm3", "Tbl_ChecklistForm4" ];
            inputData.forEach( data => {
                if ( targetTablenames.includes( data.Tablename ) )
                {
                    let index = data.TemplateData.indexOf( "Policy LOB" );
                    if ( index === -1 )
                    {
                        index = data.TemplateData.indexOf( "POLICY LOB" );
                    }
                    if ( index !== -1 )
                    {
                        data.TemplateData[ index ] = "PolicyLob";
                    }
                    let index1 = data.TemplateData.indexOf( "COVERAGE_SPECIFICATIONS_MASTER" );
                    if ( index1 !== -1 )
                    {
                        data.TemplateData[ index1 ] = "CoverageSpecificationsMaster"
                    }
                    if ( data.Tablename == "Tbl_ChecklistForm2" )
                    {
                        let index2 = data.TemplateData.indexOf( "Current Term Policy - Listed" );
                        if ( index2 !== -1 )
                        {
                            data.TemplateData[ index2 ] = "CurrentTermPolicyListed1"
                        }
                    }
                    let index3 = data.TemplateData.indexOf( "OBSERVATION" );
                    if ( index3 !== -1 )
                    {
                        data.TemplateData[ index3 ] = "Observation"
                    }
                }
            } );

            inputData.forEach( data => {
                if ( targetTablenames.includes( data.Tablename ) )
                {
                    data.TemplateData = data.TemplateData.map( item => item.replace( / /g, '' ).replace( /_/g, '' ).replace( /-/g, '' ) );
                }
            } );

            // Data mapping based on the appConfig table headings--*
            const tableDataMap = {
                'Table 4': { data: inputData.find( data => data.Tablename === 'Table 4' ), appConfigTableData: 'Tbl_ChecklistForm1' },
                'Table 5': { data: inputData.find( data => data.Tablename === 'Table 5' ), appConfigTableData: 'Tbl_ChecklistForm2' },
                'Table 6': { data: inputData.find( data => data.Tablename === 'Table 6' ), appConfigTableData: 'Tbl_ChecklistForm3' },
                'Table 7': { data: inputData.find( data => data.Tablename === 'Table 7' ), appConfigTableData: 'Tbl_ChecklistForm4' }
            };
            for ( const tableName in tableDataMap )
            {
                const tableInfo = tableDataMap[ tableName ];
                const tableData = tableInfo.data;
                const appConfigTableData = inputData.find( data => data.Tablename === tableInfo.appConfigTableData );

                if ( tableData && tableData.TemplateData.length > 0 && appConfigTableData )
                {
                    inputData.filter( data => data.Tablename === tableName ).forEach( data => {
                        const appConfigTableKeys = new Set( appConfigTableData.TemplateData );
                        data.TemplateData.forEach( item => {
                            Object.keys( item ).forEach( key => {
                                if ( ( key === 'Observation' || key === 'PageNumber' || key === 'CoverageSpecificationsMaster' || key === 'ChecklistQuestions' ) && item[ key ] === "Details not available in the document" )
                                {
                                    item[ key ] = '   ';
                                }
                                else if ( ( item[ key ] === null || item[ key ] === "" ) && appConfigTableKeys.has( key ) )
                                {
                                    item[ key ] = 'Details not available in the document';
                                }
                            } );
                        } );
                    } );
                }
            }
            const tableData2 = inputData.find( ( data ) => data.Tablename === tableName );

            if ( !tableData2 )
            {
                //console.error( "Table 2 data not found" );
                return;
            }
            if ( tableData2?.TemplateData?.length > 0 )
            {
                const headersKeys = Object.keys( tableData2?.TemplateData[ 0 ] );
                headersKeys.forEach( ( column ) => {
                    if ( tableData2?.TemplateData?.filter( ( f ) => f[ column ] != null )?.length > 0 || ( column == "Lob" && tableName === "Table 3" && tableData2?.isMultipleLobSplit ) )
                    {
                        tableColumnNamesOfValid[ column ] = 0
                    }
                } );
            }

            // const table2json = tableData2.TemplateData ;
            const table22sonCopy = tableData2.TemplateData;
            const itemArray = [ "CurrentTermPolicyListed", "PriorTermPolicyListed", "ProposalListed", "BinderListed", "ScheduleListed", "QuoteListed", "ApplicationListed", "CurrentTermPolicyListed1", "CurrentTermPolicyAttached" ];
            for ( let i = 0; i < table22sonCopy.length; i++ )
            {
                const obj = table22sonCopy[ i ];
                let allDetailsNotAvailable = true;

                for ( let j = 0; j < itemArray.length; j++ )
                {
                    const key = itemArray[ j ];

                    if ( obj[ key ] !== 'Details not available in the document' )
                    {
                        allDetailsNotAvailable = false;
                        break;
                    }
                    if ( obj[ key ] !== 'MATCHED' ) 
                    {
                        allDetailsNotAvailable = false;
                        break;
                    }
                }
                if ( allDetailsNotAvailable )
                {
                    obj.Observation = '';
                    obj.PageNumber = '';
                }
            }

            let tableColumnKeys = [];
            if ( table22sonCopy && table22sonCopy?.length > 0 )
            {
                const allKeys = Object.keys( table22sonCopy[ 0 ] );
                allKeys.map( ( e ) => {
                    if ( e )
                    {
                        let keyHasData = table22sonCopy?.filter( ( f ) => ( f[ e ] != null && f[ e ] !== "" ) || ( e == "Lob" && tableData2?.isMultipleLobSplit ) || ( e == "ChecklistQuestions" && ( f[ e ] === null || f[ e ] === "" ) ) || ( e == "CoverageSpecificationsMaster" && ( f[ e ] === null || f[ e ] === "" ) ) );
                        if ( keyHasData?.length > 0 )
                        {
                            tableColumnKeys.push( e );
                        }
                    }
                } );
                if ( !tableColumnKeys?.includes( 'Observation' ) )
                {
                    tableColumnKeys.push( 'Observation' );
                }
                if ( !tableColumnKeys?.includes( 'PageNumber' ) )
                {
                    tableColumnKeys.push( 'PageNumber' );
                }
            }

            // based on the master key enable or disable the confidence score start**
            // by ***gokul*** on 11-feb-2025
            const cs_keys = ["CurrentTermPolicyCs","PriorTermPolicyCs","BinderCs","ProposalCs","QuoteCs","ApplicationCs","ScheduleCs",
                            "CurrentTermPolicyListedCs","ProposalListedCs","BinderListedCs","ScheduleListedCs","QuoteListedCs","ApplicationListedCs",
                            "PriorTermPolicyListedCs","CurrentTermPolicyAttachedCs","CurrentTermPolicyListedCs1"];
            
            const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
            const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
            const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); //variable to store the MinLockCellScore
            if(EnableConfidenceScore != "true" || !EnableConfidenceScore || !props?.enableCs){
                tableColumnKeys = tableColumnKeys.filter((f) => !cs_keys?.includes(f));
            }else{
                // check if the respective cs column has added if not add it.
                tableColumnKeys.forEach((available_cols) => {
                    if(!cs_keys?.includes(available_cols)){
                        const respective_cs_col = getCsRespectiveColumn(available_cols);
                        if(respective_cs_col && !tableColumnKeys.includes(respective_cs_col)){
                            tableColumnKeys.push(respective_cs_col);
                        }
                    }
                })
            }
            // based on the master key enable or disable the confidence score end**

            const table2JsonCopy = table22sonCopy.map( obj => {
                let newObj = {};
                tableColumnKeys.forEach( ( key ) => {
                    newObj[ key ] = obj[ key ];
                } );
                return newObj;
            } );




            const table2json = table2JsonCopy.map( item => {
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

            } );

            const filteredColumns = table2json.map( ( item ) => {
                const filteredItem = {};
                for ( const key in item )
                {
                    if ( key === "Observation" || key === "CoverageSpecificationsMaster" )
                    {
                        filteredItem[ key ] = item[ key ];
                    }
                }
                return filteredItem;
            } );


            let header = Object.keys( table2json[ 0 ] );
            header = header.filter( f => ![ "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill" ]?.includes( f ) );
            const value = Object.values( table2json );
            const policyLOBValues = value.map( item => item[ "PolicyLob" ] );
            const table1Data = mainData && mainData?.length > 0 ? mainData?.find( ( f ) => f?.Tablename == 'Table 1' ) : {};
            let heaterLob = '';

            if ( tableName === "Table 3" && table1Data && table1Data?.TemplateData && table1Data?.TemplateData?.length > 0 )
            {
                const headerPolicyLob = table1Data?.TemplateData.map( ( e ) => e.PolicyLob );
                const filteredLob = Array.from( new Set( headerPolicyLob?.filter( ( f ) => f != '' && f ) ) );
                heaterLob = filteredLob[ 0 ];
            }

            let headerRows1 = [];
            let rowIndexForLOBStart = 0;
            let rowIndexForLOBEnd = 0;
            if ( tableName === "Table 3" )
            {
                rowIndexForLOBStart = basedata[ basedata?.length - 1 ]?.r + 3;
                headerRows1 = [
                    {
                        r: basedata[ basedata?.length - 1 ]?.r + 3, // Start from row 1 for headers
                        rs: 1, // Span two rows for "POLICY LOB"
                        c: 1, // Start from column 1 for "POLICY LOB"
                        cs: header.length + 1, // Span all columns for the sub-headers
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [ { v: heaterLob || policyLOBValues[ 0 ] , ff: "Tahoma", fs: 10 } ] },
                            m: heaterLob || policyLOBValues[ 0 ], // Use "POLICY LOB" || "Policy LOB" as the value
                            v: heaterLob || policyLOBValues[ 0 ], // Use "POLICY LOB" || "Policy LOB" as the value
                            ff: "\"Tahoma\"",
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                ]

            };

            const excludedColumns = [ "PolicyLob", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill" ];
            let headers = Object.keys( table2json[ 0 ] ).filter( headerw => !excludedColumns.includes( headerw ) );
            if ( policyLOBValues && policyLOBValues?.length > 0 && policyLOBValues[ 0 ] === 'Are the forms and endorsements attached, listed in current term policy?' )
            {
                const indexListed = headers.indexOf( "CurrentTermPolicyListed" );
                const indexAttached = headers.indexOf( "CurrentTermPolicyAttached" );
                if ( indexListed !== -1 && indexAttached !== -1 && indexAttached > indexListed )
                {
                    // Swap the elements at the identified indices
                    [ headers[ indexListed ], headers[ indexAttached ] ] = [ headers[ indexAttached ], headers[ indexListed ] ];
                }
            }
            const removalCode = headers.map( item => ( tableName !== "Table 3" && item === "CoverageSpecificationsMaster" ) ? policyLOBValues[ 0 ] : item );

            headerRows1 = [
                ...headerRows1,
                ...removalCode.map( ( item, index ) => {
                    apiDataConfig.demo.config.merge[ `${ tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3 }_${ 1 + index }` ] = {
                        "r": tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3,
                        "c": 1 + index,
                        "rs": 2,
                        "cs": 1
                    }
                    let convertedItem = item;
                    if(brokerId === "1150"){
                        if((tableName === "Table 2" || tableName === "Table 3") && item?.trim()?.toLowerCase() === "application"){
                            convertedItem = "Epic";
                        }else if(tableName === "Table 4" && item?.trim()?.toLowerCase() === "applicationlisted"){
                            convertedItem = "EpicListed";
                        }
                    }
                    return {
                        r: tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3, // Start from row 1 for headers
                        rs: 2, // Start from row 1 for headers
                        c: 1 + index, // Display headers in the first column
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [ { v: convertedItem , ff: "Tahoma", fs: 10 } ] },
                            m: convertedItem, // Use the header as the value
                            v: convertedItem, // Use the header as the value
                            ff: "\"Tahoma\"",
                            merge: null, // No merging in this example
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                } ),
            ];

            //set the column location in the variable for auto populate of ct, pt and other applications
            if ( headerRows1?.length > 0 )
            {
                const columnHeaderss = Object.keys( tableColumnNamesOfValid );
                headerRows1.forEach( ( f, index ) => {
                    if ( tableName === "Table 3" && ( index === 0 || index === 1 ) )
                    {
                        tableColumnNamesOfValid[ "CoverageSpecificationsMaster" ] = f?.c;
                    } else if ( index === 0 )
                    {
                        tableColumnNamesOfValid[ "CoverageSpecificationsMaster" ] = f?.c;
                    } else
                    {
                        if ( columnHeaderss.includes( f?.v?.m ) || columnHeaderss.includes( f?.v?.v ) || 
                        (["Epic","EpicListed"]?.includes(f?.v?.m) || ["Epic","EpicListed"]?.includes(f?.v?.v)))
                        {
                            const chText = f?.v?.m;
                            if(chText?.trim()?.toLowerCase() === "epic"){
                                tableColumnNamesOfValid[ "Application" ] = f?.c;
                            }else if(chText?.trim()?.toLowerCase() === "epiclisted"){
                                tableColumnNamesOfValid[ "ApplicationListed" ] = f?.c;
                            }else{
                                tableColumnNamesOfValid[ f?.v?.v ] = f?.c;
                            }
                        }
                    }
                } );
            }
            //add documentviewer
            if ( needDocumentViewer )
            {
                apiDataConfig.demo.config.merge[ `${ tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3 }_${ 1 + headerRows1[ headerRows1?.length - 1 ]?.c }` ] = {
                    "r": tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3,
                    "c": 1 + headerRows1[ headerRows1?.length - 1 ]?.c,
                    "rs": 2,
                    "cs": 1
                }

                const DocumentViewer = [ {
                    r: tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3, // Start from row 1 for headers
                    rs: 2, // Start from row 1 for headers
                    c: 1 + headerRows1[ headerRows1?.length - 1 ]?.c, // Display headers in the first column
                    cs: 1,
                    v: {
                        ct: { fa: "@", t: "inlineStr", s: [ { v: 'Document Viewer' , ff: "Tahoma", fs: 10 } ] },
                        // m: 'Document Viewer', // Use "Headers" as the value
                        v: 'Document Viewer', // Use "HeaderChecklist Questionss" as the value
                        ff: "\"Tahoma\"",
                        merge: null, // No merging in this example
                        bg: "rgb(139,173,212)",
                        tb: '2',
                        w: 55,
                    }
                } ];

                headerRows1 = [ ...headerRows1, ...DocumentViewer ];
            }


            const defaultHeaderRows1 = DefaultColumns.map( ( item, index ) => {
                if ( DefaultColumns?.length === index + 1 )
                {
                    rowIndexForLOBEnd = tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1;
                }
                if ( index == 0 )
                {
                    apiDataConfig.demo.config.merge[ `${ tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3 }_${ tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1 }` ] = {
                        "r": tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3,
                        "c": tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1,
                        "rs": 1,
                        "cs": 4,
                    }
                    // if()
                    return {
                        r: tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 4 : basedata[ basedata?.length - 1 ]?.r + 3, // Start from row 1 for headers
                        rs: 1, // Start from row 1 for headers
                        c: tableName === "Table 3" ? headerRows1.length + index : headerRows1.length + index + 1, // Display headers in the first column
                        cs: 1,
                        v: {
                            ht: 0,
                            ct: { fa: "@", t: "inlineStr", s: [ { v: item , ff: "Tahoma", fs: 10 } ] },
                            m: item, // Use "Headers" as the value
                            v: item, // Use "Headers" as the value
                            ff: "\"Tahoma\"",
                            merge: null, // No merging in this example
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }
                else
                {
                    return {
                        r: tableName === "Table 3" ? basedata[ basedata?.length - 1 ]?.r + 5 : basedata[ basedata?.length - 1 ]?.r + 4, // Start from row 1 for headers
                        rs: 1, // Start from row 1 for headers
                        c: tableName === "Table 3" ? headerRows1.length + index - 1 : headerRows1.length + index, // Display headers in the first column
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [ { v: item , ff: "Tahoma", fs: 10 } ] },
                            m: item, // Use "Headers" as the value
                            v: item, // Use "Headers" as the value
                            ff: "\"Tahoma\"",
                            merge: null, // No merging in this example
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                }
            } );

            if ( tableName === "Table 3" )
            {
                apiDataConfig.demo.config.merge[ `${ rowIndexForLOBStart }_${ 1 }` ] = {
                    "r": rowIndexForLOBStart,
                    "c": 1,
                    "rs": 1,
                    "cs": rowIndexForLOBEnd - 1
                }
            }

            let rowLockingMaster = {}; // contaisn the source column positions. sample {"Table 1": [3,4,...]}
            let headerRows1Values = [];
            let rowIndex = defaultHeaderRows1[ defaultHeaderRows1.length - 1 ]?.r + 1;
            const actionColumnKeys = [ "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill" ];
            headers = [ ...headers, ...actionColumnKeys ];
            table2json.forEach( ( item, cIndex ) => {

                let rowHeight = 60;
                headers.forEach( ( key, rIndex ) => {
                    rowIndex = headerRows1Values?.length == 0 ? rowIndex : headerRows1Values?.length > 0 && rIndex == 0 ? headerRows1Values[ headerRows1Values.length - 1 ]?.r + 1 : headerRows1Values[ headerRows1Values.length - 1 ]?.r;
                    if ( item[ key ] !== null )
                    {

                        let text = item[ key ]?.split( '~~' );
                        let ct = [];
                        let fs = 10;

                        function escapeRegExp( str ) {
                            return str.replace( /[.*+?^${}()|[\]\\]/g, "\\$&" ); // $& means the whole matched string
                        }
                        // DNA for CS columns;
                        if(cs_keys?.includes(key) && !text){
                            text = ["Deatils not available in the document"];
                        }

                        function splitWordsWithComma( array ) {
                            if ( !array || array.length === 0 )
                            {
                                return [];
                            }
                            let newArray = [];

                            array.forEach( ( word ) => {
                                // Check if the word ends with a comma
                                word = word.trim();
                                if ( word.endsWith( ',' ) )
                                {
                                    const wordWithoutComma = word.slice( 0, -1 ).trim();
                                    // Add the word without the comma as a separate character, excluding leading spaces
                                    if ( wordWithoutComma !== '' )
                                    {
                                        newArray.push( wordWithoutComma );
                                    }
                                    newArray.push( ',' );
                                } else if ( word.includes( '(' ) && word.includes( ')' ) )
                                {
                                    // If the word contains both '(' and ')', split them into separate characters
                                    const openingParen = word.indexOf( '(' );
                                    const closingParen = word.indexOf( ')' );
                                    const beforeParen = word.slice( 0, openingParen );
                                    const insideParen = word.slice( openingParen + 1, closingParen );
                                    const afterParen = word.slice( closingParen + 1 );
                                    if ( beforeParen !== '' )
                                    {
                                        newArray.push( beforeParen );
                                    }
                                    newArray.push( '(' );
                                    if ( insideParen !== '' )
                                    {
                                        newArray.push( insideParen );
                                    }
                                    newArray.push( ')' );
                                    if ( afterParen !== '' )
                                    {
                                        newArray.push( afterParen );
                                    }
                                } else if ( word.includes( '(' ) )
                                {
                                    // If the word contains an open parenthesis, split it into separate characters
                                    const openingParen = word.indexOf( '(' );
                                    const beforeParen = word.slice( 0, openingParen );
                                    const insideParen = word.slice( openingParen + 1 );
                                    if ( beforeParen !== '' )
                                    {
                                        newArray.push( beforeParen );
                                    }
                                    newArray.push( '(' );
                                    if ( insideParen !== '' )
                                    {
                                        newArray.push( insideParen );
                                    }
                                } else if ( word.includes( ')' ) )
                                {
                                    // If the word contains a closing parenthesis, split it into separate characters
                                    let closingParen = word.indexOf( ')' );
                                    let insideParen = word.slice( 0, closingParen ).trim();
                                    const afterParen = word.slice( closingParen + 1 );
                                    if ( insideParen !== '' )
                                    {
                                        newArray.push( insideParen );
                                    }
                                    newArray.push( ')' );
                                    if ( afterParen !== '' )
                                    {
                                        newArray.push( afterParen );
                                    }
                                } else
                                {
                                    // If no comma, just add the word to the new array
                                    newArray.push( word );
                                }
                            } );

                            return newArray;
                        }

                        if ( text && text?.length > 0 )
                        {
                            const ttableData2 = inputData.find( data => data.Tablename === "Table 2" );
                            const tt2 = removeNullValues( ttableData2.TemplateData[ 0 ], '' );
                            const tt2keys = Object.keys( tt2 );
                            const applicationIndex = tt2keys.indexOf( "Application" );
                            const keysBeforeApplication = tt2keys.slice( 0, applicationIndex );
                            const keyBeforeApplication = tt2keys[ applicationIndex - 1 ];

                            const ttableData3 = inputData.find( data => data.Tablename === "Table 3" );
                            const tt3 = removeNullValues( ttableData3.TemplateData[ 0 ], "Lob" );
                            const tt3keys = tt3 == undefined ? tt2keys : Object.keys( tt3 );
                            const tb3applicationIndex = tt3keys.indexOf( "Application" );
                            const tb3keysBeforeApplication = tt3keys.slice( 0, tb3applicationIndex );
                            const tb3keyBeforeApplication = tt3keys[ tb3applicationIndex - 1 ];
                            const textLength = text.length;
                            text?.map( ( e, splitIndex ) => {
                                if ( e?.toLowerCase().includes( 'page #' ) )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": "\r\n" + e.trim() + "\r\n"
                                    } );
                                } else if ( e?.toLowerCase().includes( 'endorsement page #' ) )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": "\r\n" + e.trim() + "\r\n"
                                    } );
                                }
                                else if ( key === "PageNumber" )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "#000000",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": e.trim() + "\r\n"
                                    } );
                                }
                                else if ( e?.toLowerCase().includes( 'current policy listed' ) )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": "\r\n" + e.trim() + "\r\n"
                                    } );
                                } else if ( e?.toLowerCase().includes( 'current policy endorsement listed' ) )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": "\r\n" + e.trim() + "\r\n"
                                    } );
                                } else if ( e?.toLowerCase().includes( 'current policy attached' ) )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": "\r\n" + e.trim() + "\r\n"
                                    } );
                                } else if ( e?.toLowerCase().includes( 'current policy endorsement attached' ) )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": "\r\n" + e.trim() + "\r\n"
                                    } );
                                } else if ( e === 'MATCHED' )
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(0, 128, 0)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": e.trim()
                                    } );
                                }

                                else if ( key === "PriorTermPolicyListed" && item[ "CurrentTermPolicyListed" ] && item[ "PriorTermPolicyListed" ]?.trim() != item[ "CurrentTermPolicyListed" ]?.trim()
                                    && !( item[ "PriorTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "CurrentTermPolicyAttached" && item[ "CurrentTermPolicyAttached" ]?.trim() != item[ "CurrentTermPolicyListed" ]?.trim()
                                    && !( item[ "CurrentTermPolicyAttached" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "PriorTermPolicyListed" && item[ "CurrentTermPolicyListed1" ] && item[ "PriorTermPolicyListed" ]?.trim() != item[ "CurrentTermPolicyListed1" ]?.trim()
                                    && !( item[ "PriorTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed1" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed1" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "QuoteListed" && item[ "QuoteListed" ]?.trim() != item[ "CurrentTermPolicyListed" ]?.trim()
                                    && !( item[ "QuoteListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "ProposalListed" && item[ "ProposalListed" ]?.trim() != item[ "CurrentTermPolicyListed" ]?.trim()
                                    && !( item[ "ProposalListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "BinderListed" && item[ "BinderListed" ]?.trim() != item[ "CurrentTermPolicyListed" ]?.trim()
                                    && !( item[ "BinderListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "ScheduleListed" && item[ "ScheduleListed" ]?.trim() != item[ "CurrentTermPolicyListed" ]?.trim()
                                    && !( item[ "ScheduleListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "ApplicationListed" && item[ "ApplicationListed" ]?.trim() != item[ "CurrentTermPolicyListed" ]?.trim()
                                    && !( item[ "ApplicationListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicyListed" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicyListed" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "PriorTermPolicy" && item[ "PriorTermPolicy" ]?.trim() != item[ "CurrentTermPolicy" ]?.trim()
                                    && !( item[ "PriorTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicy" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "Quote" && item[ "Quote" ]?.trim() != item[ "CurrentTermPolicy" ]?.trim()
                                    && !( item[ "Quote" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicy" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "Proposal" && item[ "Proposal" ]?.trim() != item[ "CurrentTermPolicy" ]?.trim()
                                    && !( item[ "Proposal" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicy" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "Binder" && item[ "Binder" ]?.trim() != item[ "CurrentTermPolicy" ]?.trim()
                                    && !( item[ "Binder" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicy" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                else if ( key === "Schedule" && item[ "Schedule" ]?.trim() != item[ "CurrentTermPolicy" ]?.trim()
                                    && !( item[ "Schedule" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicy" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                
                                else if ( key === "Application" && item[ "Application" ]?.trim() != item[ "CurrentTermPolicy" ]?.trim()
                                    && !( item[ "Application" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                        || item[ "CurrentTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                {

                                    let ptpSplitArray = e?.split( " " );
                                    let ctpSplitArray = item[ "CurrentTermPolicy" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                    const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                    const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                    ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                        let css = "#000000";
                                        let ctpText = ctpFlattenedArray.join( " " );

                                        if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                        {
                                            css = "#000000";
                                        } else
                                        {
                                            let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                            const ctpWordsArray = ctpText.split( ' ' );

                                            // Check if each word in ptpe is present in ctpWordsArray
                                            ptpe.split( ' ' ).forEach( ( word ) => {
                                                if ( !ctpWordsArray.includes( word.trim() ) )
                                                {
                                                    css = "#ff0000";
                                                }
                                            } );

                                            if ( !pattern.test( ctpText ) )
                                            {
                                                css = "#ff0000";
                                            }
                                        }
                                        ct.push( {
                                            "ff": "\"Tahoma\"",
                                            "fc": css,
                                            "fs": `${fs}`,
                                            "cl": 0,
                                            "un": 0,
                                            "bl": 0,
                                            "it": 0,
                                            "v": ptpe.trim() + " "
                                        } );
                                    } );
                                }
                                
                                // else if ( key === "Application" && item[ "Application" ]?.trim() != item[ "CurrentTermPolicy" ]?.trim()
                                //     && !( item[ "Application" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" )
                                //         || item[ "CurrentTermPolicy" ]?.toLowerCase()?.replace( /\\r\\n/g, '' )?.includes( "details not available in the document" ) ) )
                                // {

                                //     let ptpSplitArray = e?.split( " " );
                                //     let ctpSplitArray = item[ "CurrentTermPolicy" ]?.split( '~~' )[ splitIndex ]?.split( " " );

                                //     const ptpFlattenedArray = splitWordsWithComma( ptpSplitArray );
                                //     const ctpFlattenedArray = splitWordsWithComma( ctpSplitArray );

                                //     ctpFlattenedArray && ctpFlattenedArray?.length > 0 && ptpFlattenedArray.forEach( ( ptpe ) => {
                                //         let css = "#000000";
                                //         let ctpText = ctpFlattenedArray.join( " " );

                                //         if ( ptpe.includes( "$||" ) || ptpe.includes( "||" ) || ptpe.includes( "(" ) || ptpe.includes( ")" ) )
                                //         {
                                //             css = "#000000";
                                //         } else
                                //         {
                                //             let pattern = new RegExp( `\\b${ escapeRegExp( ptpe.trim() ) }\\b`, 'i' );
                                //             const ctpWordsArray = ctpText.split( ' ' );

                                //             // Check if each word in ptpe is present in ctpWordsArray
                                //             ptpe.split( ' ' ).forEach( ( word ) => {
                                //                 if ( !ctpWordsArray.includes( word.trim() ) )
                                //                 {
                                //                     css = "#ff0000";
                                //                 }
                                //             } );

                                //             if ( !pattern.test( ctpText ) )
                                //             {
                                //                 css = "#ff0000";
                                //             }
                                //         }
                                //         ct.push( {
                                //             "ff": "\"Tahoma\"",
                                //             "fc": css,
                                //             "fs": `${fs}`,
                                //             "cl": 0,
                                //             "un": 0,
                                //             "bl": 0,
                                //             "it": 0,
                                //             "v": ptpe.trim() + " "
                                //         } );
                                //     } );
                                // }
                                else
                                {
                                    ct.push( {
                                        "ff": "\"Tahoma\"",
                                        "fc": "#000000",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": e.trim() + " "
                                    } );
                                }
                            } )
                        }
                        if ( key === "PageNumber" )
                        {
                            const textOfct = ct[ ct.length - 1 ]?.v?.replace( '\r\n', ' ' );
                            if ( textOfct && ct?.length > 0 )
                            {
                                ct[ ct.length - 1 ][ "v" ] = textOfct;
                            }
                        }
                        // let concatString = ct.map((cte) => cte?.v ).join("");

                        // setting background color for cs score greater that configured score start**
                        let bs_color_code = "#ffffff";
                        if(props?.enableCs && props?.enableCellLock){
                            const isStpValid = getConfidenceScoreConfigStatus(props?.data?.find((f) => f.Tablename === "JobHeader")?.StpMappings, "question check" ,item["ChecklistQuestions"] );
                            if(EnableConfidenceScore === "true" && !cs_keys?.includes(key) && isStpValid){
                                const org_col_key = getCsRespectiveColumn(key);
                                const cs_score = item[org_col_key];                     
                                if(cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score){
                                    if(parseFloat(cs_score) > parseFloat(MinLockCellScore)){
                                        bs_color_code = "#7bf87b";
                                    }
                                }
                            }
                        }
                        // end**
                        headerRows1Values.push( {
                            r: rowIndex, // Start from row 1 for headers
                            c: rIndex + 1 + ( actionColumnKeys?.includes( key ) ? 1 : 0 ), // Display headers in the first column
                            v: {
                                ct: { fa: "General", t: "inlineStr", s: ct },                              
                                ff: "Tahoma",
                                fc: "#3b3737",
                                merge: null,                              
                                w: 55,
                                tb: '2',
                                bg: bs_color_code
                            }
                        } );

                        // if ( text && rowHeight < parseInt( item[ key ]?.length / 3 + 20 ) )
                        // {
                        //     rowHeight = parseInt( item[ key ]?.length / 3 + 20 );
                        //     apiDataConfig.demo.config.rowlen[ `${ rowIndex }` ] = rowHeight;
                        // }
                        // if ( rIndex == 0 )
                        // {
                        //     apiDataConfig.demo.config.rowlen[ `${ rowIndex }` ] = rowHeight;
                        // }
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

                        if(rowHeight != undefined && rowHeight != null) {
                            const rowHeight = parseInt(maxLength / 3 + 50);
                            apiDataConfig.demo.config.rowlen[`${rowIndex}`] = rowHeight;

                            if (rIndex == 0) {
                                apiDataConfig.demo.config.rowlen[`${rowIndex}`] = rowHeight;
                            }
                        }    
                    }

                    if(cIndex === 0 && cs_keys?.includes(key)){
                        const org_col_key = getCsRespectiveColumn(key);
                        const org_col_position = tableColumnNamesOfValid[org_col_key];

                        if (rowLockingMaster[tableName] && Array.isArray(rowLockingMaster[tableName]) && rowLockingMaster[tableName].length > 0) {
                            const org_col_set = rowLockingMaster[tableName];
                            if(!org_col_set?.includes(org_col_position)){
                                rowLockingMaster[tableName] = [...rowLockingMaster[tableName], org_col_position];
                            }
                        } else {
                            rowLockingMaster[tableName] = [org_col_position];
                        }                          
                    }
                } )
            } );

            // for CS start**
            const lockingIndexCopy = lockingIndex;
            lockingIndexCopy[tableName] = rowLockingMaster[tableName];
            setLockingIndex(lockingIndexCopy);
            // end**

            apiDataConfig.demo.config.borderInfo.push( {
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
                            headerRows1[ 0 ]?.r,
                            headerRows1Values[ headerRows1Values?.length - 1 ]?.r
                        ],
                        "column": [
                            headerRows1[ 0 ]?.c,
                            defaultHeaderRows1[ defaultHeaderRows1?.length - 1 ]?.c
                        ],
                        "row_focus": headerRows1[ 0 ]?.r,
                        "column_focus": headerRows1[ 0 ]?.c
                    }
                ]
            } );
            defaultHeaderRows1.forEach( row => {
                if ( row?.v && row?.v?.m && typeof row?.c === 'number' && row.v.m !== 'Actions on Discrepancy (from AMs)' )
                {
                    tableColumnNamesOfValid[ row.v.m ] = row?.c;
                }
            } );

            headerRows1 = [ ...headerRows1, ...defaultHeaderRows1, ...headerRows1Values ];

            const allRows2 = [ ...headerRows1 ];
            // const allRows2 = [ ...headerRows1, ...emptyValueRows ];
            // Sort the rows by rowIndex if needed
            allRows2.sort( ( a, b ) => a.r - b.r );

            if ( allRows2 && allRows2?.length > 0 )
            {
                const tableColumnDetailss = tableColumnDetails;
                tableColumnDetailss[ tableData2?.Tablename ] = { "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[ 0 ]?.r, "end": allRows2[ allRows2?.length - 1 ]?.r } }
                setTableColumnDetails( tableColumnDetailss );
            }

            // Add the rows to the dummyData2
            basedata.push( ...allRows2 );
            apiDataConfig.demo.celldata = basedata;
        };

        if(sheetRenderConfig?.PolicyReviewChecklist == 'true'){
            renderTable1();
        } 


        const formTable1 = () => {
        if(isFormApplicable && isFormApplicable == true){
            const formTableData1 = formCompareData.find( ( data ) => data.Tablename === "FormTable 1" );

            if ( formTableData1 )
            {
                const formtable1 = formTableData1.TemplateData;

                let sheetDataTable3 = [];
                let sheetDataTable4 = [];
                const rowIndexOfTable1 = 3
                const formData = renderForm();
                formData.forEach( ( item, index ) => {
                    const mergeConfig = FormCompare_appconfigdata.forms.config.merge[ "1_1" ];

                    sheetDataTable3.push( {
                        r: 1 + mergeConfig.r,
                        c: mergeConfig.c,
                        v: {
                            ct: item.ct,
                            m: item.m,
                            v: item.v,
                            ff: item.ff,
                            // bl: 0,
                            fs: 24,
                            // ff: "Arial",
                            merge: mergeConfig,
                            fc: item.fc,
                            // tb: '55',
                        }
                    } );
                } );

                formtable1.map( ( item, index ) => {
                    if ( item[ "Headers" ] != null  && item[ "Headers" ] != undefined )
                    {
                        if (item[ "Headers" ] == "") {
                            sheetDataTable4.push( {
                                r: rowIndexOfTable1 + index, // Start from row 1 for headers
                                c: 1, // Display headers in the first column
                                v: {
                                    ct: { fa: "@", t: "inlineStr", s: [ { v: " " } ] },
                                    m: " ", // Use "Headers" as the value
                                    v: " ", // Use "Headers" as the value
                                    merge: null,
                                    bg: "rgb(139,173,212)",
                                    tb: '2',
                                }
                            } );
                        } else {
                            sheetDataTable4.push( {
                                r: rowIndexOfTable1 + index,
                                c: 1,
                                v: {
                                    ct: { fa: "@", t: "inlineStr", s: [ { v: item[ "Headers" ] , ff: "Tahoma", fs: 10} ] },
                                    m: item[ "Headers" ],
                                    v: item[ "Headers" ],
                                    ff: "Tahoma",
                                    merge: null,
                                    bg: "rgb(139,173,212)",
                                    tb: '2',
                                }
                            } );
                        }

                        const tidleValue = item[ "NoColumnName" ] !== null && item[ "NoColumnName" ] != undefined ? item[ "NoColumnName" ].replace( /~~/g, "\n" ) : "";

                        sheetDataTable4.push( {
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
                        } );
                    }
                } );

                const dummyData1 = [];
                const allFormRows = [ ...sheetDataTable4, ...sheetDataTable3 ];
                if ( sheetDataTable4 && sheetDataTable4?.length > 0 )
                {
                    const FormtableColumnDetails1 = formTableColumnDetails;
                    formTableColumnDetails[ "FormTable 1" ] = { "columnNames": formtable1.map( ( e ) => e?.Headers ), "range": { "start": 0, "end": sheetDataTable4[ sheetDataTable4?.length - 1 ]?.r } }
                    setFormTableColumnDetails( FormtableColumnDetails1 );
                }

                allFormRows.sort( ( a, b ) => a.r - b.r );

                dummyData1.push( ...allFormRows );
                FormCompare_appconfigdata.forms.celldata = dummyData1;

                allFormRows.forEach( ( row ) => {
                    if ( sheetDataTable4.includes( row ) )
                    {
                        FormCompare_appconfigdata.forms.config.borderInfo.push( {
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
                        } );
                    }
                } );
                formCompareData.map( ( e, index ) => {
                    if ( e?.Tablename != 'FormTable 1' && ( e?.TemplateData?.length >= 3 || e?.TemplateData?.length < 3 ) )
                    {
                        let filteredData = FormCompare_appconfigdata.forms.celldata.filter( ( f, index ) => f != null || !f );
                        formTable2( [ ...filteredData ], e?.Tablename );
                    }
                } );
            }
        }
            renderLuckySheet( true, '', false );
        }

        const formTable2 = ( combinedata1, tableName ) => {
            if ( !Array.isArray( combinedata1 ) )
            {
                // console.error("combinedata1 is not an array:", combinedata1);
                return;
            }
            const tableColumnNamesOfValid = {};
            const needDocumentViewer = true;
            const basedata = [ ...combinedata1 ];
            // console.log("basedata", basedata);
            const inputData = mainData;
            

            let defaultText = updateData[ 0 ];
            let propsUpdateData = JSON.parse( defaultText.TemplateData );
            for ( let i = 0; i < formCompareData.length; i++ )
            {
                if ( formCompareData[ i ].TemplateData.length === 0 )
                {
                    formCompareData[ i ].TemplateData = propsUpdateData;
                }
            }

            formCompareData.forEach( ( data ) => {
                if ( data.TemplateData && Array.isArray( data.TemplateData ) )
                {
                    data.TemplateData.forEach( ( template ) => {
                        Object.keys( template ).forEach( ( key ) => {
                            if ( template[ key ] === null )
                            {
                                template[ key ] = '';
                            }
                        } );
                    } );
                }
            } );

            const formTableData2 = formCompareData.find( ( data ) => data.Tablename === tableName && data.TemplateData.length > 0 );
            if (formTableData2 && formTableData2.TemplateData) {

                formTableData2.TemplateData = formTableData2.TemplateData.map(item => {
                //    If PolicyLob is empty, set it to "Attached Forms"
                    if(!item.CurrentTermPolicyAttachedCs ){
                        item["CurrentTermPolicyAttachedCs"] = "Details not available in the document";
                    }
                    if(!item.PriorTermPolicyAttachedCs){
                        item["PriorTermPolicyAttachedCs"] = "Details not available in the document";
                    }
                    if (item.PolicyLob === ""||item.policyLOB === undefined|| item.policyLOB === null ) {
                        
                        return { ...item, PolicyLob: "Attached Forms" };
                    }
                    return item;
                });
            }
        
            if ( formTableData2?.TemplateData?.length > 0 )
            {
                const headersKeys = Object.keys( formTableData2?.TemplateData[ 0 ] );
                headersKeys.forEach( ( column ) => {
                    if ( formTableData2?.TemplateData?.filter( ( f ) => f[ column ] != null )?.length > 0 || ( tableName === "FormTable 2" && formTableData2 ) || ( tableName === "FormTable 3" && formTableData2 ) )
                    {
                        tableColumnNamesOfValid[ column ] = 0
                    }
                } );
            }


            if ( !formTableData2 )
            {
                // console.error("Table 2 data not found");
                return;
            }

            const formtable2copy = formTableData2.TemplateData;

            const formDataCopy = formtable2copy.map( obj => {
                let newObj = {};
                Object.keys( obj ).forEach( key => {
                    if ( obj[ key ] !== null )
                    {
                        newObj[ key ] = obj[ key ];
                    }
                } );
                return newObj;
            } );

            const formtable2 = formDataCopy.map( item => {
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

            } );

            let header = Object.keys( formtable2[ 0 ] );
            const value = Object.values( formtable2 );
            const policyLOBValues = value.map( item => item[ "PolicyLob" ] );

            let headerRows1 = [];
            let rowIndexForLOBStart = 0;
            // let rowIndexForLOBEnd = 8;
            let rowIndexForLOBEnd = 0;

            const cs_form_keys = ["CurrentTermPolicyAttachedCs","PriorTermPolicyAttachedCs"];
            const cs_form_keys_org = ["CurrentTermPolicyAttached","PriorTermPolicyAttached"];
            const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
            const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
            const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); //variable to store the MinLockCellScore
           
            if(EnableConfidenceScore != "true" || !EnableConfidenceScore || !props?.enableCs){
                header = header.filter((f) => !cs_form_keys?.includes(f));
            }else{
                // check if the respective cs column has added if not add it.
                cs_form_keys.forEach((key) => {
                    if(!header?.includes(key)){
                        header.push(key);
                    }
                });
            }
            

            if ( tableName === "FormTable 2" || tableName === "FormTable 3" )
            {
                rowIndexForLOBStart = basedata[ basedata?.length - 1 ]?.r + 2;
                headerRows1 = [
                    {
                        r: basedata[ basedata?.length - 1 ]?.r + 2,
                        rs: 1,
                        c: 1,
                        cs: header.length + 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [ { v: tableName === "FormTable 2" ? "Unmatched Forms" : "Matched Forms" , ff: "Tahoma", fs: 10 } ] },
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



            const excludedColumns = [ "PolicyLob" ];
            let headers = Object.keys( formtable2[ 0 ] ).filter( headerw => !excludedColumns.includes( headerw ) );
          
            const removalCode = headers.map( item => ( item === "CoverageSpecificationsMaster" ) ? policyLOBValues[ 0 ] : item );

            headerRows1 = [
                ...headerRows1,
                ...removalCode.map( ( item, index ) => {
                    if ( removalCode?.length === index + 1 )
                    {
                        rowIndexForLOBEnd = headerRows1.length + index + 1;
                    }
                    return {
                        r: basedata[ basedata?.length - 1 ]?.r + 3,
                        rs: 2,
                        c: 1 + index,
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [ { v: item , ff: "Tahoma", fs: 10 } ] },
                            m: item,
                            v: item,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                } )
            ];
            
            


            if ( headerRows1?.length > 0 )
            {
                const columnHeaderss = Object.keys( tableColumnNamesOfValid );
                headerRows1.forEach( ( f, index ) => {
                    if ( tableName === "FormTable 3" && ( index === 0 || index === 1 ) )
                    {
                        tableColumnNamesOfValid[ "CoverageSpecificationsMaster" ] = f?.c;
                    } else if ( index === 0 )
                    {
                        tableColumnNamesOfValid[ "CoverageSpecificationsMaster" ] = f?.c;
                    } else
                    {
                        if ( columnHeaderss.includes( f?.v?.m ) || columnHeaderss.includes( f?.v?.v ) )
                        {
                            tableColumnNamesOfValid[ f?.v?.v ] = f?.c;
                        }
                    }
                } );
            }


            if ( needDocumentViewer )
            {
                // FormCompare_appconfigdata.forms.config.merge[`${basedata[basedata?.length - 1]?.r + 3}_${1 + headerRows1[headerRows1?.length - 1]?.c}`] = {
                //     "r": basedata[basedata?.length - 1]?.r + 3,
                //     "c": 1 + headerRows1[headerRows1?.length - 1]?.c,
                //     "rs": 2,
                //     "cs": 1
                // }

                const DocumentViewer = [ {
                    r: basedata[ basedata?.length - 1 ]?.r + 3,
                    rs: 2,
                    c: 1 + headerRows1[ headerRows1?.length - 1 ]?.c,
                    cs: 1,
                    v: {
                        ct: { fa: "@", t: "inlineStr", s: [ { v: 'Document Viewer' , ff: "Tahoma", fs: 10 } ] },
                        m: 'Document Viewer',
                        v: 'Document Viewer',
                        ff: "\"Tahoma\"",
                        merge: null,
                        bg: "rgb(139,173,212)",
                        tb: '2',
                        w: 55,
                    }
                } ];

                // headerRows1 = [...headerRows1, ...DocumentViewer];
                headerRows1 = [ ...headerRows1 ];
            }


            if ( tableName === "FormTable 2" || tableName === "FormTable 3" )
            {
                FormCompare_appconfigdata.forms.config.merge[ `${ rowIndexForLOBStart }_${ 1 }` ] = {
                    "r": rowIndexForLOBStart,
                    "c": 1,
                    "rs": 1,
                    "cs": rowIndexForLOBEnd - 1
                }
            }

            let rowLockingMaster = {}; // contaisn the source column positions. sample {"Table 1": [3,4,...]}
            let headerRows1Values = [];
            let rowIndex = basedata[ basedata?.length - 1 ]?.r + 4;
            formtable2.forEach( ( item, cIndex ) => {
                let rowHeight = 21;
                headers.forEach( ( key, rIndex ) => {
                    if(key === "CoverageSpecificationsMaster"){
                        item[key] = "Attached Forms";
                    }
                    if(key === "ChecklistQuestions"){
                        item[key] = "CA2";
                    }
                    if(cs_form_keys?.includes(key) && !item[key]){
                        item[key] = "Details not available in the document";
                    }
                    rowIndex = headerRows1Values?.length == 0 ? rowIndex : headerRows1Values?.length > 0 && rIndex == 0 ? headerRows1Values[ headerRows1Values.length - 1 ]?.r + 1 : headerRows1Values[ headerRows1Values.length - 1 ]?.r;
                    let text = item[ key ].toString().split( '~~' );
                    let ss = [];
                    let fs = 10;
                    // function escapeRegExp(str) {
                    //     return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); // $& means the whole matched string
                    // }

                    // function splitWordsWithComma(array) {
                    //     if (!array || array.length === 0) {
                    //         return [];
                    //     }
                    //     let newArray = [];

                    //     array.forEach((word) => {
                    //         // Check if the word ends with a comma
                    //         word = word.trim();
                    //         if (word.endsWith(',')) {
                    //             const wordWithoutComma = word.slice(0, -1).trim();
                    //             // Add the word without the comma as a separate character, excluding leading spaces
                    //             if (wordWithoutComma !== '') {
                    //                 newArray.push(wordWithoutComma);
                    //             }
                    //             newArray.push(',');
                    //         } else if (word.includes('(') && word.includes(')')) {
                    //             // If the word contains both '(' and ')', split them into separate characters
                    //             const openingParen = word.indexOf('(');
                    //             const closingParen = word.indexOf(')');
                    //             const beforeParen = word.slice(0, openingParen);
                    //             const insideParen = word.slice(openingParen + 1, closingParen);
                    //             const afterParen = word.slice(closingParen + 1);
                    //             if (beforeParen !== '') {
                    //                 newArray.push(beforeParen);
                    //             }
                    //             newArray.push('(');
                    //             if (insideParen !== '') {
                    //                 newArray.push(insideParen);
                    //             }
                    //             newArray.push(')');
                    //             if (afterParen !== '') {
                    //                 newArray.push(afterParen);
                    //             }
                    //         } else if (word.includes('(')) {
                    //             // If the word contains an open parenthesis, split it into separate characters
                    //             const openingParen = word.indexOf('(');
                    //             const beforeParen = word.slice(0, openingParen);
                    //             const insideParen = word.slice(openingParen + 1);
                    //             if (beforeParen !== '') {
                    //                 newArray.push(beforeParen);
                    //             }
                    //             newArray.push('(');
                    //             if (insideParen !== '') {
                    //                 newArray.push(insideParen);
                    //             }
                    //         } else if (word.includes(')')) {
                    //             // If the word contains a closing parenthesis, split it into separate characters
                    //             let closingParen = word.indexOf(')');
                    //             let insideParen = word.slice(0, closingParen).trim();
                    //             const afterParen = word.slice(closingParen + 1);
                    //             if (insideParen !== '') {
                    //                 newArray.push(insideParen);
                    //             }
                    //             newArray.push(')');
                    //             if (afterParen !== '') {
                    //                 newArray.push(afterParen);
                    //             }
                    //         } else {
                    //             // If no comma, just add the word to the new array
                    //             newArray.push(word);
                    //         }
                    //     });

                    //     return newArray;
                    // }
                    if ( text && text?.length > 0 )
                    {
                        const formCellCompare = formCompareData.find( data => data.Tablename === tableName )
                        const ctpa = removeNullValues( formCellCompare.TemplateData[ 0 ] );
                        const ctpakey = Object.keys( ctpa );
                        const ctpaIndex = ctpakey.indexOf( "PriorTermPolicyAttached" );
                        const ctpaBeforeColumn = ctpakey[ ctpaIndex - 1 ];

                        const ptpaIndex = ctpakey.indexOf( "CurrentTermPolicyAttached" );
                        const ptpaAfterColumn = ctpakey[ ptpaIndex + 1 ];
                        const datePattern = /^\d{1,2}\/\d{1,2}$/;

                        text.forEach( ( e, splitIndex ) => {
                            
                            if ( e.toLowerCase().includes( 'page' ) )
                            {
                                ss.push( {
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgb(68, 114, 196)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                    // "v": "\r\n" + e.trim() 
                                } );
                            }
                            // else if (key === 'PriorTermPolicyAttached' &&
                            //     !item["CurrentTermPolicyAttached"]?.includes("Details not available in the document")) {

                            //     if (!e.includes('Details not available in the document')) {
                            //         // Split the text and flatten each array
                            //         let ptpSplitArray = e?.split(" ");
                            //         let ctpSplitArray = item["CurrentTermPolicyAttached"]?.split('~~')[splitIndex]?.split(" ");

                            //         const ptpFlattenedArray = splitWordsWithComma(ptpSplitArray);
                            //         const ctpFlattenedArray = splitWordsWithComma(ctpSplitArray);

                            //         // Loop through each word in ptpFlattenedArray
                            //         ptpFlattenedArray.forEach((ptpe, index) => {
                            //             let css = "#000000";  // Default color for matched words

                            //             // Check if there’s a corresponding word in ctpFlattenedArray and if they match
                            //             if (ctpFlattenedArray[index] !== ptpe) {
                            //                 css = "#ff0000";  // Set to red if the word doesn't match
                            //             }

                            //             // Push each word with the appropriate color
                            //             ss.push({
                            //                 "ff": "\"Tahoma\"",
                            //                 "fc": css,
                            //                 "fs": `${fs}`,
                            //                 "cl": 0,
                            //                 "un": 0,
                            //                 "bl": 0,
                            //                 "it": 0,
                            //                 "v": ptpe.trim() + " "
                            //             });
                            //         });
                            //     }
                            // }
                            else
                            {
                                ss.push( {
                                    "ff": "\"Tahoma\"",
                                    "fc": "#000000",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": e.trim() + "\r\n"
                                    // "v": e.trim()
                                } );
                                // }
                            }
                        } )
                    }
                    
                    let bs_color_code = "#ffffff";
                    if (tableName != 'FormTable 2') {
                        if(props?.formsCompareHeaderData?.EnableCS && props?.formsCompareHeaderData?.EnableCellLock){
                            const isStpValid = getConfidenceScoreConfigStatus(props?.formsCompareHeaderData?.StpMappings, "question check" ,"CA2");
                            if(EnableConfidenceScore === "true" && cs_form_keys_org?.includes(key) && isStpValid){
                                const org_col_key = getCsRespectiveColumn(key);
                                const cs_score = item[org_col_key];                    
                                if(cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score){
                                    if(parseFloat(cs_score) > parseFloat(MinLockCellScore)){
                                        bs_color_code = "#7bf87b";
                                    }
                                }
                            }
                        }
                    }

                    headerRows1Values.push( {
                        r: rowIndex,
                        c: rIndex + 1,
                        v: {
                            ct: { fa: "General", t: "inlineStr", s: ss },
                            merge: null,
                            w: 55,
                            tb: '2',
                            bg: bs_color_code,
                        }
                    } );
                    if ( text && rowHeight < parseInt( item[ key ]?.length / 2 + 20 ) )
                    {
                        rowHeight = parseInt( item[ key ]?.length / 2 + 20 );
                        FormCompare_appconfigdata.forms.config.rowlen[ `${ rowIndex }` ] = rowHeight;
                    }
                    if ( rowIndex == 0 )
                    {
                        FormCompare_appconfigdata.forms.config.rowlen[ `${ rowIndex }` ] = rowHeight;
                    }
                    if(cIndex === 0 && cs_form_keys?.includes(key)){
                        const org_col_key = getCsRespectiveColumn(key);
                        const org_col_position = tableColumnNamesOfValid[org_col_key];
    
                        if (rowLockingMaster[tableName] && Array.isArray(rowLockingMaster[tableName]) && rowLockingMaster[tableName].length > 0) {
                            const org_col_set = rowLockingMaster[tableName];
                            if(!org_col_set?.includes(org_col_position)){
                                rowLockingMaster[tableName] = [...rowLockingMaster[tableName], org_col_position];
                            }
                        } else {
                            rowLockingMaster[tableName] = [org_col_position];
                        }                          
                    }
                } )
            } );
            // for CS start**
            const lockingIndexCopy = lockingIndex;
            lockingIndexCopy[tableName] = rowLockingMaster[tableName];
            setLockingIndex(lockingIndexCopy);
            // end**
            FormCompare_appconfigdata.forms.config.borderInfo.push( {
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
                            headerRows1[ 0 ]?.r,
                            headerRows1Values[ headerRows1Values?.length - 1 ]?.r
                        ],
                        "column": [
                            headerRows1[ 0 ]?.c,
                            headerRows1[ headerRows1?.length - 1 ].c
                        ],
                        "row_focus": headerRows1[ 0 ]?.r,
                        "column_focus": headerRows1[ 0 ]?.c
                    }
                ]
            } );

            headerRows1 = [ ...headerRows1, ...headerRows1Values ];
            const allRows2 = [ ...headerRows1 ];
            allRows2.sort( ( a, b ) => a.r - b.r );
            if ( allRows2 && allRows2?.length > 0 )
            {
                const tableColumnDetailss = formTableColumnDetails;
                formTableColumnDetails[ tableName ] = { "columnNames": tableColumnNamesOfValid, "range": { "start": allRows2[ 0 ]?.r, "end": allRows2[ allRows2?.length - 1 ]?.r } }
                setFormTableColumnDetails( tableColumnDetailss );
            }

            basedata.push( ...allRows2 );

            FormCompare_appconfigdata.forms.celldata = basedata;

        };
        if ( isFormApplicable && sheetRenderConfig?.FormsCompare == 'true' )
        {
            formTable1();
        }


        const exclusionTable = () => {
            const basedata = [];
            const data = props.exclusionRenderData;
            const dataMap = data;
            const tableColumnNamesOfValid = {};
            let headersKeys = Object.keys(dataMap[0]);

            const cs_key = ["ConfidenceScore"];
            const EnableCsForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableCsForExclusion");
            const EnableLockCellForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCellForExclusion");
            const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); 

            if (EnableCsForExclusion != "true" || !EnableCsForExclusion || !EnableLockCellForExclusion) {
                headersKeys = headersKeys?.filter((f) => !cs_key?.includes(f));
            }
            // const DefaultColumns = ["Actions on Discrepancy (from AMs)", "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
            if (dataMap?.length > 0) {
                headersKeys.forEach((column) => {
                        tableColumnNamesOfValid[column] = 0
                });
            }
            if ( dataMap && dataMap?.length > 0 && !Array.isArray( dataMap[ 0 ] ) )
            {
                const exclusionjson = dataMap.map( item => {
                    const {
                        Id,
                        JobId,
                        CreatedOn,
                        UpdatedOn,
                        ...filteredItem
                    } = item;
                    return filteredItem;
                } );

                const excludedColumns = [ "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill" ];
                const headers = Object.keys( exclusionjson[ 0 ] ).filter( headerw => !excludedColumns.includes( headerw ) );
                let headerRows1 = headers.map( ( item, index ) => {
                    return {
                        r: 0,
                        rs: 2,
                        c: index,
                        cs: 1,
                        v: {
                            ct: { fa: "@", t: "inlineStr", s: [ { v: item , ff: "Tahoma", fs: 10 } ] },
                            m: item,
                            v: item,
                            ff: "\"Tahoma\"",
                            merge: null,
                            bg: "rgb(139,173,212)",
                            tb: '2',
                            w: 55,
                        }
                    }
                } );

                headerRows1 = [ ...headerRows1 ];

                // const defaultHeaderRows1 = DefaultColumns.map((item, index) => {
                //     if (index == 0) {
                //         exclusionDatafigdata.exclusion.config.merge[`${ headerRows1[headerRows1?.length - 1]?.r }_${ headerRows1.length + index}`] = {
                //             "r": headerRows1[headerRows1?.length - 1]?.r,
                //             "c":  4,
                //             "rs": 1,
                //             "cs": 4,
                //         }
                //         return {
                //             r:  headerRows1[headerRows1?.length - 1]?.r,
                //             rs: 1,
                //             c:  headerRows1.length,
                //             cs: 1,
                //             v: {
                //                 ht: 0,
                //                 ct: { fa: "General", t: "g" },
                //                 m: item,
                //                 v: item,
                //                 fs: 11,
                //                 ff: "\"Tahoma\"",
                //                 merge: null,
                //                 bg: "rgb(139,173,212)",
                //                 tb: '2',
                //                 w: 55,
                //             }
                //         }
                //     }
                //     else {
                //         return {
                //             r:  headerRows1[headerRows1?.length - 1]?.r + 1,
                //             rs: 1,
                //             c:  headerRows1.length + index - 1,
                //             cs: 1,
                //             v: {
                //                 ct: { fa: "General", t: "g" },
                //                 m: item,
                //                 v: item,
                //                 fs: 11,
                //                 ff: "\"Tahoma\"",
                //                 merge: null,
                //                 bg: "rgb(139,173,212)",
                //                 tb: '2',
                //                 w: 55,
                //             }
                //         }
                //     }
                // });

                let headerRows1Values = [];
                let rowIndex = headerRows1[ headerRows1.length - 1 ]?.r + 1;
                let rowHeight = 40;
                let fs = 10;

                exclusionjson.map( ( item, indexr ) => {
                    // const actionColumnKeys = ["ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
                    // headers = [...headers, ...actionColumnKeys];
                    headers.map( ( key, rIndex ) => {
                        let text = item[ key ] != null ? item[ key ].toString()?.split( '~~' ) : [];
                        let ss = [];
                        if ( text && text?.length > 0 )
                        {
                            text.map( ( e ) => {
                                ss.push( {
                                    "ff": "\"Tahoma\"",
                                    "fc": "#000000",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": e?.trim()?.length > 0 ? e.trim() + "\r\n" : " "
                                } );
                            }
                            );
                        }else{
                            ss.push( {
                                "ff": "\"Tahoma\"",
                                "fc": "#000000",
                                "fs": `${fs}`,
                                "cl": 0,
                                "un": 0,
                                "bl": 0,
                                "it": 0,
                                "v": item[ key ] || " "
                            } );
                        }

                        let bs_color_code = "#ffffff";
                        if (EnableCsForExclusion === "true" && EnableLockCellForExclusion == "true" && !cs_key?.includes(key) && exclusionApplicableIdx?.includes(rIndex) && props?.enableExclusionCellLock) {
                            const cs_score = item["ConfidenceScore"];
                            if (cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score) {
                                if (parseFloat(cs_score) > MinLockCellScore) {
                                    bs_color_code = "#7bf87b";
                                }
                            }
                        }

                        headerRows1Values.push( {
                            r: rowIndex + indexr,
                            c: rIndex,
                            v: {
                                ct: { fa: "General", t: "inlineStr", s: ss },
                                merge: null,
                                ff: "\"Tahoma\"",
                                w: 55,
                                tb: '2',
                                bg: bs_color_code,
                            }
                        } );

                        exclusionDatafigdata.exclusion.config.rowlen[`${ indexr }`] = rowHeight;

                    } )
                } );
                exclusionDatafigdata.exclusion.config.borderInfo.push( {
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
                                headerRows1[ 0 ]?.r,
                                headerRows1Values[ headerRows1Values?.length - 1 ]?.r
                            ],
                            "column": [
                                headerRows1[ 0 ]?.c,
                                4
                            ],
                            "row_focus": headerRows1[ 0 ]?.r,
                            "column_focus": headerRows1[ 0 ]?.c
                        }
                    ]
                } );

                headerRows1 = [...headerRows1, ...headerRows1Values];
                // headerRows1 = [...headerRows1, ...defaultHeaderRows1, ...headerRows1Values];
                const allRows2 = [ ...headerRows1 ];
                if ( headerRows1 && headerRows1?.length > 0 )
                {

                    const ExTableColumnDetails = exTableColumnDetails;
                    if (headerRows1?.length > 0) {
                        const columnHeaderss = Object.keys(tableColumnNamesOfValid);
                        headerRows1.forEach((f, index) => {
                            
                                if (columnHeaderss.includes(f?.v?.m) || columnHeaderss.includes(f?.v?.v)) {
                                    tableColumnNamesOfValid[f?.v?.v] = f?.c;
                                }
                        });
                    }

                    ExTableColumnDetails[ "ExTable 1" ] = {
                        "columnNames": tableColumnNamesOfValid,
                        "range": {
                            "start": 0,
                            "end": headerRows1[ headerRows1.length - 1 ]?.r
                        }
                    };
                    setExTableColumnDetails( ExTableColumnDetails );
                }
                allRows2.sort( ( a, b ) => a.r - b.r );
                basedata.push( ...allRows2 );
                exclusionDatafigdata.exclusion.celldata = basedata;
            }
        };

        if(sheetRenderConfig?.Exclusion == 'true'){ 
            exclusionTable();
        }

        renderLuckySheet();
        let interval;
        setTimeout(() => {
            sessionStorage.setItem("IsDataRendering",false);
            const { UpdatePeriod } = autoupdate();
            console.log("UpdatePeriod",UpdatePeriod);
            interval = setInterval( () => {
                const { UpdateEnable } = autoupdate();
                console.log("UpdatePeriod","triggered");
                if ( UpdateEnable && issavessheet === false )
                {
                    Autoupdateclick( true );
                }
            }, UpdatePeriod );
        }, 2000);

        return () => {
            clearInterval( interval )
            const allowautoup = sessionStorage.getItem("IsAutoUpdate")
            if ( autoprogress && issavessheet == false && (allowautoup == "true" || allowautoup == true))
            {
                Autoupdateclick( true );
            }
            if(issavessheet == false && (allowautoup == "true" || allowautoup == true ) ){
            Autoupdateclick( true );
            }
        }

    }, [ autoprogress, props.data, props?.sheetRenderConfig ] );


    function removeNullValues( obj, key ) {
        for ( const prop in obj )
        {
            if ( obj[ prop ] === null && prop != key )
            {
                delete obj[ prop ];
            }
        }
        return obj;
    }

    const formCompareUpdate = async ( needLoader, iExportform, callback, valuesToPass ) => {
        
        let sheetFlag = luckysheet?.getSheet()?.name;
        if ( isFormApplicable && sheetFlag == 'Forms Compare')
        {
            const updatedDatas = {};
            let sheetData = luckysheet.getSheet()?.data;

            const keys = Object.keys( formTableColumnDetails );

            const separateArrays = [];

            if ( keys?.length > 0 && sheetData && sheetData != undefined )
            {
                keys.forEach( ( f ) => {
                    const tableData = formTableColumnDetails[ f ];
                    if ( tableData && tableData?.range && tableData?.range?.start != undefined && tableData?.range?.end != undefined && tableData?.columnNames && tableData?.range?.end != '' )
                    {
                        const slicedData = sheetData.slice( tableData?.range?.start, tableData?.range?.end + 1 );
                        separateArrays.push( slicedData );
                    }
                } );
            }

            let parentHeaderSegments = [];
            let tableDateExceptHeaders = [];
            let limitReached = false;

            separateArrays.map( ( e, index ) => {
                if ( index >= 1 && !limitReached )
                {
                    let tempTableIndex = [];
                    let tempTableColumnName = [];
                    let hasReachedLimit = false;
                    let tabledata = formTableColumnDetails;
                    for ( let tableName in tabledata )
                    {
                        if ( tableName != "FormTable 1" )
                        {
                            let tableInfo = tabledata[ tableName ];
                            for ( let columnName in tableInfo.columnNames )
                            {
                                let columnValue = tableInfo.columnNames[ columnName ];
                                if ( columnValue === 0 )
                                {
                                    delete tableInfo.columnNames[ columnName ];
                                }
                            }
                        }
                    }

                    const data = index == 1 ? e[ 1 ] : e[ 1 ];
                    const filteredData = data?.filter( item => item !== null );
                    filteredData?.forEach( ( e1, index1 ) => {
                        if ( index1 >= 0 && !hasReachedLimit )
                        {
                            tempTableIndex.push( index1 );
                            // tempTableColumnName.push(e1?.m || e1?.v);
                        }
                    } );

                    tempTableColumnName = tempTableColumnName.filter( column => column !== undefined );
                    hasReachedLimit = true;
                    const tableName = `FormTable ${ index + 1 }`;
                    if ( tabledata.hasOwnProperty( tableName ) )
                    {
                        const range = tabledata[ tableName ].columnNames;
                        tempTableColumnName.push( ...Object.keys( range ) );
                    }
                    const limitedIndex = tempTableIndex.slice( 0, tempTableColumnName.length );
                    parentHeaderSegments.push( { "Table": `FormTable ${ index + 1 }`, index: tempTableIndex, tempTableColumnName } );
                }

            } );

            let trueValues = [];
            let falseValues = [];

            separateArrays.map( ( f, index ) => {
                if ( index >= 1 )
                {
                    let keyValuePair = [];
                    f?.map( ( e, index1 ) => {
                        if ( index1 > 1 )
                        {
                            let tablePairingData = parentHeaderSegments[ 0 ];
                            if ( tablePairingData )
                            {
                                let object = {};

                                tablePairingData.index.map( ( i, index2 ) => {
                                    i = i + 1;
                                    if ( e[ i ]?.ct?.s && Array.isArray( e[ i ]?.ct?.s ) )
                                    {
                                        // Concatenate the 'v' values from the array of objects
                                        let filteredS = e[ i ]?.ct?.s.filter( ( f ) => f != null );
                                        let concatenatedValues = filteredS?.map( item => item?.v )?.join( '' );
                                        concatenatedValues = concatenatedValues?.replace( /\r\n/g, '~~' );
                                        const finalValue = e[ i ]?.m || e[ i ]?.v || concatenatedValues || "";
                                        object[ `${ tablePairingData?.tempTableColumnName[ index2 ] }` ] = finalValue;
                                    } else
                                    {
                                        // If e[i]?.ct?.s is not an array
                                        object[ `${ tablePairingData?.tempTableColumnName[ index2 ] }` ] = e[ i ]?.m || e[ i ]?.v || e[ i ]?.ct?.s || "";
                                    }
                                    // if (tablePairingData.index?.length == index2 + 1) {
                                    //     keyValuePair.push(object);
                                    // }
                                    if ( tablePairingData.index?.length == index2 + 1 )
                                    {
                                        keyValuePair.push( object );
                                    }
                                } )
                            }
                        }
                        // if (f?.length == index1 + 1) {
                        //     tableDateExceptHeaders.push({ Table: `FormTable ${index - 1}`, NewTemplateData: keyValuePair });
                        // }
                    } );
                    tableDateExceptHeaders.push( { Table: `FormTable ${ index + 1 }`, NewTemplateData: keyValuePair } );
                    // combinedTable = combinedTable.concat(keyValuePair);   //concatination of array of object matched and unmatched objects
                }
            } );

            const formCompareData = props.formCompareData;
            const tableNames = formCompareData.map( item => item.Tablename );

            const addDotonUpdateprocess = ( value ) => {
                if ( typeof value === 'string' )
                {
                    return value.replace( /\•/g, '.' );
                }
                return value;
            }

            let formdataSetToUpdate = [];
            tableNames.forEach( async ( table ) => {
                if ( separateArrays && separateArrays.length > 0 )
                {
                    if ( table === "FormTable 1" )
                    {
                        const jsonDataToUpdate = [];
                        // console.log( separateArrays );
                        const nullFiter = separateArrays[ 0 ].map( sublist => sublist.filter( item => item !== null ) );
                        const result1 = nullFiter.map( ( [ index1, index2 ] ) => ( {
                            index1,
                            index2,
                        } ) );
                        let result = result1.slice( 3 );
                        // for ( const key in valuesToPass )
                        // {
                        //     const value = valuesToPass[ key ];
                        //     for ( const item of result )
                        //     {
                        //         if ( item.index1 && item.index1 != undefined && item.index1.v === key )
                        //         {
                        //             item.index2.v = value;
                        //             item.index2.m = value;
                        //             break;
                        //         }
                        //     }
                        // }
                        const resultWithJoinedValues = result.map( ( { index1, index2 } ) => {
                            if ( index2?.ct && index2?.ct?.s !== undefined )
                            {
                                const joinedValue = index2?.ct?.s.map( ( { v } ) => v ).join( '' );
                                const { s, ...ctWithoutS } = index2.ct;
                                return {
                                    index1,
                                    index2: {
                                        ...index2,
                                        ct: ctWithoutS,
                                        m: joinedValue,
                                        v: joinedValue
                                    }
                                };
                            } else
                            {
                                return { index1, index2 };
                            }
                        } );

                        resultWithJoinedValues.forEach( item => {
                            if ( item?.index2 && item?.index2?.ct?.s !== undefined )
                            {
                                if ( Array.isArray( item.index2.ct.s ) && item.index2.ct.s.length > 0 )
                                {

                                    const extractedData = item.index2.ct.s[ 0 ];

                                    item.index2.m = extractedData.v;
                                    item.index2.v = extractedData.v;

                                    delete item.index2.ct.s;
                                } else
                                {
                                    item.index2.m = '';
                                    item.index2.v = '';

                                    delete item.index2.ct.s;
                                }
                            }
                        } );


                        for ( let rowKey in resultWithJoinedValues )
                        {
                            if ( rowKey != 'len' )
                            {
                                const cellData = resultWithJoinedValues[ rowKey ];
                                const formCompareData = props.formCompareData;
                                const tableData1 = formCompareData.find( ( data ) => data.Tablename === "FormTable 1" );
                                if ( tableData1 )
                                {
                                    const table1json = tableData1.TemplateData;
                                    const policyVal = table1json.map( item => item[ "PolicyLob" ] );
                                    const cell1Text = cellData.index1?.v || cellData.index1?.ct || '';
                                    const cell2Text = cellData.index2?.v !== undefined ? ( cellData.index2.v || cellData.index2?.ct ) : ( cellData.index2?.ct?.fa === "@" ? "" : cellData.index2?.ct?.fa );
                                    // const cell2Text = cellData.index2?.v !== undefined || cellData.index2?.ct ? (cellData.index2.v || cellData.index2?.ct) : (cellData.index2?.ct?.fa === "@" ? "" : cellData.index2?.ct?.fa);

                                    if ( cell1Text && cell1Text.s && cell1Text.s.length > 0 && cell2Text || cell1Text )
                                    {
                                        // const vValue = cell1Text.s.map(item => item.v || '').join(',') || cell1Text;
                                        const vValue = Array.isArray( cell1Text.s ) ? cell1Text.s.map( item => item.v || '' ).join( ',' ) : cell1Text;
                                        const concatenatedValues = cell2Text?.s && Array.isArray( cell2Text.s ) ? cell2Text.s.map( item => item.v ).join( '' ) : cell2Text;

                                        const policyLOB = policyVal[ 0 ];
                                        const formattedRow = {
                                            HeaderID: rowKey,
                                            JOBID: jobId,
                                            ...( policyLOB && { 'PolicyLob': policyLOB } ),
                                            Headers: vValue,
                                            '': concatenatedValues,
                                        };

                                        const addTidleonUpdateprocess = ( value ) => {
                                            if ( typeof value === 'string' )
                                            {
                                                return value.replace(/\n/g, '~~').replace(/"/g, '\\"');
                                            }
                                            return value;
                                        };

                                        const jsonString = `{${ Object.entries( formattedRow ).map( ( [ key, value ] ) => {
                                            if ( key === 'HeaderID' )
                                            {
                                                const updatedValue = Number( value ).toString();
                                                return `"${ key }":${ updatedValue }`;
                                            } else if ( key === '' )
                                            {
                                                return `"${ key }":"${ addTidleonUpdateprocess( value ) }"`;
                                            } else if (key === 'Headers') {
                                                const sValue = Array.isArray(value) ? `"${addTidleonUpdateprocess(value.join(', '))}"` : `"${addTidleonUpdateprocess(value)}"`;
                                                return `"${ key }":${ sValue }`;
                                            }
                                            return `"${key}":"${addTidleonUpdateprocess(value)}"`;
                                        } ).join( ',' ) }}`;
                                        jsonDataToUpdate.push( jsonString );
                                    }
                                }
                            }
                        }
                        const json = `[${ jsonDataToUpdate.join( ',' ) }]`;
                        updatedDatas[ "FormTable 1" ] = json;
                        formdataSetToUpdate.push( { id: jobId, tableName: "FormTable 1", data: updatedDatas[ "FormTable 1" ] } );
                        // formCompareUpdateTable1(jobId, "FormTable 1", updatedDatas["FormTable 1"]);
                    }
                 if ( table === "FormTable 2" )
                    {
                        const table2Data = formCompareData.find( ( f ) => f?.Tablename.toLowerCase() === 'formtable 2' )?.TemplateData || [];
                        
                        const table3Data = formCompareData.find( ( f ) => f?.Tablename.toLowerCase() === 'formtable 3' )?.TemplateData || [];

                        let table2copy = tableDateExceptHeaders[ 0 ]?.NewTemplateData || [];
                        let table3copy = tableDateExceptHeaders[ 1 ]?.NewTemplateData || [];


                        table2copy = table2copy.map( ( md, index ) => {
                            const dataIndex = index;
                            if ( md )
                            {
                                if ( !md[ 'PolicyLob' ] )
                                {
                                    md[ 'PolicyLob' ] = table2Data[ dataIndex ]?.[ 'PolicyLob' ];
                                }
                            }
                            md[ "IsMatched" ] = false;
                            return md;
                        } );

                        table3copy = table3copy.map( ( md, index ) => {
                            const dataIndex = index;
                            if ( md )
                            {
                                if ( !md[ 'PolicyLob' ] )
                                {
                                    md[ 'PolicyLob' ] = table3Data[ dataIndex ]?.[ 'PolicyLob' ];
                                }
                            }
                            md[ "IsMatched" ] = true;
                            return md;
                        } );

                        let bothDatas = table2copy.concat( table3copy );
                        let table2 = bothDatas;
                        table2 = table2.filter( ( md ) => Object.values( md ).some( value => value !== undefined ) );

                        for ( let i = 0; i < table2.length; i++ )
                        {
                            for ( let key in table2[ i ] )
                            {
                                if ( table2[ i ].undefined === undefined && key === "undefined" )
                                {
                                    delete table2[ i ][ key ];
                                }
                            }
                        }

                        for ( let obj of table2 )
                        {
                            for ( let key in obj )
                            {
                                if ( obj[ key ] === undefined || obj[ key ] === "~~" )
                                {
                                    obj[ key ] = '';
                                }
                            }
                        }


                        table2.forEach( ( obj ) => {
                            Object.keys( obj ).forEach( ( key ) => {
                                obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                            } );
                        } );

                        const replacer = ( key, value ) => {
                            if ( Array.isArray( value ) && value.length === 0 )
                            {
                                return "";
                            }
                            return value;
                        };

                        const f1 = JSON.stringify( table2, replacer );
                        const f = JSON.parse( f1 );

                        const reorderedF = f.map( item => {
                            if ( item[ "PolicyLob" ] === item[ "PolicyLob" ] )
                            {
                                item[ 'Attached Forms' ] = item[ 'CoverageSpecificationsMaster' ];
                                delete item[ 'Attached Forms' ];

                                return {
                                    'CoverageSpecificationsMaster': item[ 'CoverageSpecificationsMaster' ],
                                    ...item
                                };
                            }
                            return item;
                        } );
                        const json = JSON.stringify( reorderedF, replacer );
                        // console.log(json);
                        let parsedJson = JSON.parse( json );

                        function removeTilde( object ) {
                            for ( let key in object )
                            {
                                if ( typeof object[ key ] === 'string' )
                                {
                                    object[ key ] = object[ key ].replace( /~~$/, '' ); // Removing '~~' from the end of the string
                                } else if ( typeof object[ key ] === 'object' )
                                {
                                    removeTilde( object[ key ] ); // Recursive call for nested objects
                                }
                            }
                        }

                        removeTilde( parsedJson );
                        let jsonArray = parsedJson;
                        jsonArray.forEach( obj => {
                            if ( !obj.hasOwnProperty( "PolicyLob" ) )
                            {
                                obj[ "PolicyLob" ] = obj[ "CoverageSpecificationsMaster" ];
                            }
                        } );


                        parsedJson = formTableDataFormatting( parsedJson, 2 );
                        const phUpdateData = await updateFormsPHData( jobId, parsedJson, token );
                        const modifiedJson = JSON.stringify( parsedJson );
                        formdataSetToUpdate.push( { id: jobId, tableName: "FormTable 2", data: modifiedJson } );

                        if (iExportform === false ) {
                            updateFormPHData( phUpdateData );
                            
                            if ( needLoader )
                            {
                                document.body.classList.add( 'loading-indicator' );
                            }
                            let response;
                            try {
                                response = await formCompareUpdateTable2(formdataSetToUpdate[1]?.id, formdataSetToUpdate[1].tableName, formdataSetToUpdate[1]?.data);

                                
                                if (response !== "error") {
                                    setMsgVisible(true);
                                    setMsgClass('alert success');
                                    setMsgText('Data Updated');
                                    setTimeout(() => { setMsgVisible(false); setMsgText(''); }, 3000);
                                } else {
                                    console.error('Update failed for the given item');
                                }
                            } catch (error) {
                                console.error('Error:', error);
                            } finally {
                                document.body.classList.remove('loading-indicator');
                            }
                            await newUpdateApiCall( formdataSetToUpdate, false, true,"FormsCompare" );                            

                        }
                        
                        
                        if ( iExportform == true )
                        {
                            formdataSetToUpdate.forEach( item => {
                                const sanitizedData = item.data.replace( /[\u0000-\u001F\u007F-\u009F]/g, '' );   //sanitize the JSON string by removing any problematic control characters before parsing it.  so dont remove this

                                let parsedData = JSON.parse( sanitizedData );


                                parsedData.forEach( obj => {

                                    if ( obj[ "" ] !== undefined )
                                    {
                                        obj[ "NoColumnName" ] = obj[ "" ];
                                        delete obj[ "" ];
                                    }
                                } );
                                item.data = JSON.stringify( parsedData );
                            } );
                            const modifiedTabledata = formdataSetToUpdate.map( item => ( {
                                Id: item.id,
                                TableName: item.tableName,
                                Data: item.data
                            } ) );

                            let matchedTables = [];
                            let unmatchedTables = [];


                            let formTable2 = modifiedTabledata.find( table => table.TableName === "FormTable 2" );

                            if ( formTable2 )
                            {

                                let data = JSON.parse( formTable2.Data );


                                data.forEach( item => {
                                    if ( item.IsMatched )
                                    {
                                        matchedTables.push( item );
                                    } else
                                    {
                                        unmatchedTables.push( item );
                                    }
                                } );
                            }



                            let formTable2Index = modifiedTabledata.findIndex( table => table.TableName === "FormTable 2" );

                            let unmatchedTableObject = {
                                Id: modifiedTabledata[ formTable2Index ].Id,
                                TableName: "unmatchedtable",
                                Data: JSON.stringify( unmatchedTables )
                            };

                            let matchedTableObject = {
                                Id: modifiedTabledata[ formTable2Index ].Id,
                                TableName: "matchedtable",
                                Data: JSON.stringify( matchedTables )
                            };

                            if ( formTable2Index !== -1 )
                            {
                                modifiedTabledata.splice( formTable2Index, 1, unmatchedTableObject, matchedTableObject );
                            }


                            const dataFrom = modifiedTabledata;
                            if ( typeof callback === "function" )
                            {
                                callback( dataFrom );
                            }
                        }
                    }
                }
            } );


            // if ( formdataSetToUpdate?.length > 0 )
            // {
            //     if ( needLoader )
            //     {
            //         document.body.classList.add( 'loading-indicator' );
            //     }

            //     const promiseResponses = formdataSetToUpdate.map( async ( item ) => {
            //         if ( iExportform == false )
            //         {
            //             let response;
            //             if ( formdataSetToUpdate[ 0 ].tableName === 'FormTable 1' )
            //             {
            //                 response = await formCompareUpdateTable1( formdataSetToUpdate[ 0 ]?.id, formdataSetToUpdate[ 0 ]?.tableName, formdataSetToUpdate[ 0 ]?.data );
            //             }
            //             return response;
            //         }

            //     } );
                
            //     Promise.all( promiseResponses )
            //         .then( ( responses ) => {
            //             const isAllSuccess = responses.every( ( res ) => res !== "error" );
            //             if ( isAllSuccess )
            //             {
            //                 setMsgVisible( true );
            //                 setMsgClass( 'alert success' );
            //                 // if (iExportform == true) {
            //                 //     setMsgText('Downloaded successfully');
            //                 // } else {
            //                 setMsgText( 'Data Updated' );
            //                 // }
            //                 setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3000 );
            //             }
                        
            //         } 
                
            //     )
            //         .catch( ( error ) => {
            //             // console.error('Error:', error);
            //         } )
            //         .finally( () => {
            //             document.body.classList.remove( 'loading-indicator' );
            //         } );
                    
                    
            // }
           

        }
        
        document.body.classList.remove('loading-indicator');
    };

    const exclusionUpdate = async ( isExport, callback ) => {
        let sheetFlag = luckysheet?.getSheet()?.name;
        if(sheetFlag == "Exclusion" ) {
            let filterdatasheet = luckysheet?.getSheet();
            if ( filterdatasheet?.data != undefined && filterdatasheet?.data?.length > 0 )
            {
                let sheetData = filterdatasheet.data;
                let hasSeenData = false
                let sheetDataLength = sheetData?.length;
                for ( let index = sheetDataLength - 1; index < sheetDataLength && index != 0; index-- )
                {
                    let hasValue = sheetData[ index ].filter( ( f ) => f != null )?.length > 0;
                    if ( ( !hasValue ) && !hasSeenData )
                    {
                        sheetData = sheetData.slice( 0, index );
                    } else
                    {
                        hasSeenData = true;
                        break;
                    }
                }
                let parentHeaderSegments = [];
                let mergedData = [];
                let limitReached = false;

                sheetData && sheetData?.length > 0 && sheetData.map( ( e, index ) => {
                    if ( index == 0 && !limitReached )
                    {
                        let tempTableIndex = [];
                        let tempTableColumnName = [];
                        let hasReachedLimit = false;
                        const filteredData = e.filter( item => item !== null );
                        filteredData?.forEach( ( e1, index1 ) => {
                            if ( index1 >= 0 && !hasReachedLimit )
                            {
                                tempTableIndex.push( index1 );
                                tempTableColumnName.push( e1?.m || e1?.v );
                            }
                        } );

                        tempTableColumnName = tempTableColumnName.filter( column => column !== undefined );
                        hasReachedLimit = true;
                        const limitedIndex = tempTableIndex.slice( 0, tempTableColumnName.length );
                        parentHeaderSegments.push( { "Table": `ExTable ${ index }`, index: tempTableIndex, tempTableColumnName } );
                    }
                } );



                sheetData && sheetData?.length > 0 && sheetData.forEach( ( e, index ) => {
                    let keyValuePair = [];
                    if ( index > 0 )
                    {
                        let tablePairingData = parentHeaderSegments[ 0 ];
                        if ( tablePairingData )
                        {
                            let object = {};
                            tablePairingData.index.forEach( ( i, index2 ) => {
                                if ( e[ i ]?.ct?.s && Array.isArray( e[ i ]?.ct?.s ) )
                                {
                                    let filteredS = e[ i ]?.ct?.s.filter( ( f ) => f != null );
                                    let concatenatedValues = filteredS?.map( item => item?.v )?.join( '' );
                                    concatenatedValues = concatenatedValues?.replace( /\r\n/g, '~~' );
                                    const finalValue = e[ i ]?.m || e[ i ]?.v || concatenatedValues || "";
                                    object[ `${ tablePairingData?.tempTableColumnName[ index2 ] }` ] = finalValue;
                                } else
                                {
                                    object[ `${ tablePairingData?.tempTableColumnName[ index2 ] }` ] = e[ i ]?.m || e[ i ]?.v || e[ i ]?.ct?.s || "";
                                }
                            } )
                            keyValuePair.push( object );
                        }
                        mergedData.push( { Table: `ExTable${ index }`, NewTemplateData: keyValuePair } );
                    }
                } );
                let combinedNewTemplateData = [];
                let tableDateExceptHeaders;

                mergedData.forEach( item => {
                    const newTemplateData = item.NewTemplateData.filter( data => {
                        return Object.values( data ).some( value => value !== undefined );
                    } );
                    if ( newTemplateData.length > 0 )
                    {
                        combinedNewTemplateData = combinedNewTemplateData.concat( newTemplateData );
                    }
                } );

                if ( combinedNewTemplateData.length > 0 )
                {
                    tableDateExceptHeaders = {
                        "Table": "ExTable1",
                        "NewTemplateData": combinedNewTemplateData
                    };
                }

                const addDotonUpdateprocess = ( value ) => {
                    if ( typeof value === 'string' )
                    {
                        return value.replace( /\•/g, '.' );
                    }
                    return value;
                }

                if ( tableDateExceptHeaders !== undefined )
                {
                    if ( tableDateExceptHeaders.Table == 'ExTable1' )
                    {
                        let updateData = tableDateExceptHeaders.NewTemplateData

                        for ( let obj of updateData )
                        {
                            for ( let key in obj )
                            {
                                if ( obj[ key ] === undefined )
                                {
                                    obj[ key ] = '';
                                }
                            }
                        }

                        function removeTilde( object ) {
                            for ( let key in object )
                            {
                                if ( typeof object[ key ] === 'string' )
                                {
                                    object[ key ] = object[ key ].replace( /~~$/, '' );
                                } else if ( typeof object[ key ] === 'object' )
                                {
                                    removeTilde( object[ key ] );
                                }
                            }
                        }
                        removeTilde( updateData );

                        updateData.forEach( ( obj ) => {
                            Object.keys( obj ).forEach( ( key ) => {
                                obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                            } );
                        } );

                        updateData = updateData.map( ( e ) => {
                            if(!e["Exclusion"]){
                                e["Exclusion"] = " "
                            }
                            if(!e["FormDescription"]){
                                e["FormDescription"] = " "
                            }
                            if(!e["FormName"]){
                                e["FormName"] = " "
                            }
                            if(!e["PageNumber"]){
                                e["PageNumber"] = " "
                            }
                            e[ "JobId" ] = jobId;
                            return e;
                        } );
                        const updateJson = JSON.stringify( updateData );
                        if ( isExport == true )
                        {
                            const modifiedTabledata = [ {
                                TableName: "ExclusionTable",
                                Data: updateJson
                            } ];
                            const dataExclusionOnUpdateClick = modifiedTabledata;


                            if ( typeof callback === "function" )
                            {

                                callback( dataExclusionOnUpdateClick );
                            }
                        } else
                        {
                            exclusionUpdateApi( jobId, 'ExTable1', updateJson );
                            await newUpdateApiCall( updateJson, false, true,"Exclusion" );  
                        }
                    }
                } else if ( tableDateExceptHeaders == undefined )
                {
                    let setstaticExclusionData = [];
                    setstaticExclusionData.push( staticExclusionData )
                    const updateJson = JSON.stringify( setstaticExclusionData );
                    if ( isExport == true )
                    {
                        const modifiedTabledata = [ {
                            TableName: "ExclusionTable",
                            Data: updateJson
                        } ];
                        const dataExclusionOnUpdateClick = modifiedTabledata;


                        if ( typeof callback === "function" )
                        {

                            callback( dataExclusionOnUpdateClick );
                        }
                    } else
                    {
                        let setstaticExclusionData = [];
                        setstaticExclusionData.push( staticExclusionData )
                        const updateJson = JSON.stringify( setstaticExclusionData );
                        exclusionUpdateApi( jobId, 'ExTable1', updateJson );
                    }

                }
            }
        }
    };

    const exclusionUpdateApi = async ( JobId, tableName, updateJson ) => {

        document.body.classList.add( 'loading-indicator' );
        const Token = await processAndUpdateToken( token );//to validate and update the token
        token = Token;
        const headers = {
            'Authorization': `Bearer ${ token }`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${ baseUrl }/api/ProcedureData/UpdateExclustionData`;

        axios( {
            method: "POST",
            url: apiUrl,
            headers: headers,
            data: {
                JobId: JobId,
                TableName: tableName,
                NewTemplateData: updateJson
            }
        } )
            .then( response => {
                if ( response.status !== 200 )
                {
                    throw new Error( `HTTP error! Status: ${ response.status }` );
                }
                return response.data;
            } )
            .then( data => {
                if ( data?.status == 400 )
                {
                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( data?.title );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                } else
                {
                    setMsgVisible( true ); setMsgClass( 'alert success' ); setMsgText( 'Data Updated' );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                }
            } )
            .finally( () => {
                document.body.classList.remove( 'loading-indicator' );
            } );
    };

    const updateFormPHData = async ( phData ) => {
        document.body.classList.add( 'loading-indicator' );
        const Token = await processAndUpdateToken( token );//to validate and update the token
        token = Token;
        const headers = {
            'Authorization': `Bearer ${ token }`,
            "Content-Type": "application/json",
        };
        const data = {
            JobId: jobId,
            TemplateData: phData
        }
        try
        {
            const response = await axios.post( baseUrl + '/api/ProcedureData/UpdateFormPageHighlighter', data, {
                headers
            } );
            setTimeout( () => { document.body.classList.remove( 'loading-indicator' ); }, 2000 )
        } catch ( error )
        {
            const errorText = error;
        }
    }

    const formCompareUpdateTable1 = async ( jobId, tableName, json ) => {
        document.body.classList.add( 'loading-indicator' );
        const Token = await processAndUpdateToken( token );//to validate and update the token
        token = Token;
        const headers = {
            'Authorization': `Bearer ${ token }`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${ baseUrl }/api/ProcedureData/UpdateFormHeaderData?jobId=${ jobId }`;

        axios( {
            method: "POST",
            url: apiUrl,
            headers: headers,
            data: {
                JobId: jobId,
                TableName: tableName,
                NewTemplateData: json
            }
        } )
            .then( response => {
                if ( response.status !== 200 )
                {
                    throw new Error( `HTTP error! Status: ${ response.status }` );
                }
                return response.data;
            } )
            .then( data => {
                if ( data?.status == 400 )
                {
                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( data?.title );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                } else
                {
                    setMsgVisible( true ); setMsgClass( 'alert success' ); setMsgText( 'Data Updated' );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                }
            } )
            .finally( () => {
                document.body.classList.remove( 'loading-indicator' );
            } );
    };


    const formCompareUpdateTable2 = async ( jobId, tableName, modifiedJson ) => {
        document.body.classList.add( 'loading-indicator' );
        updateGridAuditLog(jobId,auditProcessNames.JobUpdateFormProcessInitiated,"");
        const Token = await processAndUpdateToken( token );//to validate and update the token
        token = Token;
        const headers = {
            'Authorization': `Bearer ${ token }`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${ baseUrl }/api/ProcedureData/UpdateChecklistFormData?jobId=${ jobId }`;

        axios( {
            method: "POST",
            url: apiUrl,
            headers: headers,
            data: {
                JobId: jobId,
                TableName: tableName,
                NewTemplateData: modifiedJson
            }
        } )
            .then( response => {
                if ( response.status !== 200 )
                {
                    updateGridAuditLog(jobId,auditProcessNames.JobUpdateFormProcessFailed,"");
                    throw new Error( `HTTP error! Status: ${ response.status }` );
                }
                updateGridAuditLog(jobId,auditProcessNames.JobUpdateFormProcessCompleted,"");
                return response.data;
            } )
            .then( data => {
                if ( data?.status == 400 )
                {
                    updateGridAuditLog(jobId,auditProcessNames.JobUpdateFormProcessFailed,"");
                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( data?.title );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                } else
                {
                    updateGridAuditLog(jobId,auditProcessNames.JobUpdateFormProcessCompleted,"");
                    setMsgVisible( true ); setMsgClass( 'alert success' ); setMsgText( 'Data Updated' );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                }
            } )
            .finally( () => {
                document.body.classList.remove( 'loading-indicator' );
            } );
    };

    const newUpdateApiCall = async ( previewChecklistDataSet, isRegenerate, needLoader ,sheetname) => {
        if(sheetname == "Policychecklist"){
            if ( previewChecklistDataSet?.length > 0 )
                {
                    const Token = await processAndUpdateToken( token );
                    token = Token;
                    if ( needLoader )
                    {
                        let JobPRTotalCount = 0
                        previewChecklistDataSet.forEach((e) => {JobPRTotalCount += (JobPRTotalCount?.NewTemplateData?.length)});
                        if(JobPRTotalCount && parseInt(JobPRTotalCount) > 2000){
                            container.current.showSnackbar( `Preview Checklist Data Update Initiated and is In-Progress. 
                                As it has ${JobPRTotalCount} lineItems, it will take some time please be Patience.`, "info", true );
                        }else{
                            container.current.showSnackbar( "Preview Checklist Data Update Initiated and is In-Progress", "info", true );
                        }
                    }
                    updateGridAuditLog(jobId,needLoader ? auditProcessNames.JobUpdateProcessInitiated : auditProcessNames.JobUpdateProcessAutoSaveInit,"");
                    const promiseResponse = Promise.all( previewChecklistDataSet.map( async ( item ) => {
                        if(needLoader){document.body.classList.add( 'loading-indicator' );}
                        const PCResponse = await apiCallSwitch( item, Token, needLoader );
                        return PCResponse === undefined ? "" : PCResponse == "Success" ? "" : ( item?.TableName + " :  " + PCResponse );
                    } ) );
                    promiseResponse.then(
                        async ( res ) => {
                            
                            updateGridAuditLog(jobId,needLoader ? auditProcessNames.JobUpdateProcessCompleted : auditProcessNames.JobUpdateProcessAutoSaveComplete,"");
                            let SheetName = "Policychecklist";
                            setTimeout(() => {
                                TriggerBackUp(jobId, SheetName);
                            }, 500);
                            document.body.classList.remove( 'loading-indicator' );
                            // const headersectionData = previewChecklistDataSet.find( ( f ) => f.TableName === "Table 1" );
                            // if ( headersectionData )
                            // {
                            //     const headerSetForForms = {};
                            //     headersectionData?.NewTemplateData.forEach( ( item ) => {
                            //         headerSetForForms[ item[ "Headers" ] ] = item[ "NoColumnName" ];
                            //     } );
                            //     try{
                            //         // await formCompareUpdate( false, false, false, headerSetForForms );
                            //     }catch(error){
                            //         const eror = error;
                            //     }
                            // }
                            const isAllSucces = res?.filter( ( f ) => f == "error" )?.length == 0;
                            const hasError = res.filter( ( f ) => f );
                            if ( hasError?.length > 0 )
                            {
                                let errorMsg = hasError.join( '  ,  ' );
                                // if ( needLoader ){
                                    container.current.showSnackbar( "Data update failed, please try again","error",true );
                                // }
                            }else{
                                if ( needLoader )
                                {
                                    container?.current?.showSnackbar( "Data Updated Successfully","success" ,true);
                                }
                            }
                            if ( isAllSucces || !isAllSucces )
                            {
                                // setMsgVisible( true );
                                // setMsgClass( 'alert success' );
                                // setMsgText('Data Updated');
                                // setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3000 );
                                // if ( needLoader && hasError?.length === 0 )
                                // {
                                //     container.current.showSnackbar( "Update pageHighlighter Initiated", "info",true );
                                // }
                                updataPHProcess( isRegenerate, needLoader );
                            }
                            // else{
                            //     document.body.classList.remove( 'loading-indicator' );
                            // }
                        },
                        ( error ) => { console.error( 'Error:', error ); }
                    );
                }
        }
        else if(sheetname == "FormsCompare"){
            if ( previewChecklistDataSet?.length > 0 )
                {
                    const Token = await processAndUpdateToken( token );
                    token = Token;
                    if ( needLoader )
                    {
                        setTimeout(() => {
                            let SheetName = "FormsCompare";
                            TriggerBackUp(jobId, SheetName);
                        }, 500);
                        
                }
            }
        }
        else if(sheetname == "Exclusion"){
            if ( previewChecklistDataSet?.length > 0 )
                {
                    const Token = await processAndUpdateToken( token );
                    token = Token;
                    if ( needLoader )
                    {
                        setTimeout(() => {
                            let SheetName = "Exclusion";
                            TriggerBackUp(jobId, SheetName);
                        }, 500);
                        
                }
            }
        }
       
    }

   


    const onUpdateClick = async ( isRegenerate, needLoader, isExport, callback ) => {
        if ( isExport == false )
        {
            GridBackupSave();
        }
        let QacFlag = luckysheet?.getSheet()?.name
        if ( QacFlag != 'QAC not answered questions' )
        {
            const Token = await processAndUpdateToken( token );//to validate and update the token
            const IsDataRendering = sessionStorage.getItem("IsDataRendering");
            token = Token;
            if ( !(IsDataRendering == true || IsDataRendering == "true")&& QacFlag == "PolicyReviewChecklist")
            {
                // console.log("table", tablenameArray)
                luckysheet.exitEditMode();
                const updatedDatas = {};
                let sheetDataa = luckysheet.getAllSheets()[ 0 ].data;
                const callUpdatedUpdateFn = true;
                if ( !isExport && callUpdatedUpdateFn )
                {
                    const tableNamesList = Object.keys( tableColumnDetails );

                    let dataToBeUpdated = [];

                    const tableJCRecord = [];

                    let headerLobForJC = "";

                    tableNamesList.forEach( async ( f, index ) => {
                        const filteredData = Object.keys( tableColumnDetails[ f ].columnNames ).filter( key => tableColumnDetails[ f ].columnNames[ key ] != 0 );
                        const rowStart = f === "Table 1" ? 4 : ( f === "Table 3" ? ( tableColumnDetails[ f ]?.range?.start + 3 ) : ( tableColumnDetails[ f ]?.range?.start + 2 ) );
                        const rowEnd = tableColumnDetails[ f ]?.range?.end;
                        const splittedData = sheetDataa.slice( rowStart, ( rowEnd + 1 ) );
                        const tableMasterData = state.find( ( item ) => item?.Tablename.toLowerCase() === f.toLowerCase() );
                        let policyLobToMap = tableMasterData?.TemplateData?.find( ( f ) => f.PolicyLob != null && f?.PolicyLob != undefined )?.PolicyLob || "";
                        if(f === "Table 1"){
                            headerLobForJC = policyLobToMap;
                        }
                        if ( f === "Table 3" && tableMasterData?.isMultipleLobSplit && headerLobForJC ){
                            policyLobToMap = headerLobForJC;
                        }
                        let dataFormfn = getPreviewChecklistDataForUpdate( jobId, splittedData, tableColumnDetails[ f ]?.columnNames, filteredData, f, policyLobToMap );
                        dataFormfn = tableDataFormatting( dataFormfn, ( index + 1 ) );
                        if ( f === "Table 3" && tableMasterData?.isMultipleLobSplit )
                        {
                            tableJCRecord.push( dataFormfn );
                        }
                        dataToBeUpdated.push( { JobId: jobId, NewTemplateData: dataFormfn, TableName: f } );
                    } );
                    if ( tableJCRecord?.length > 0 )
                    {
                        const mappedData = await mapLOBColumns( tableJCRecord[ 0 ], token, jobId );
                        dataToBeUpdated = dataToBeUpdated.map( ( e ) => {
                            if ( e?.TableName === "Table 3" )
                            {
                                e.NewTemplateData = mappedData;
                            }
                            return e;
                        } );
                        await newUpdateApiCall( dataToBeUpdated, isRegenerate, needLoader,"Policychecklist" );
                    } else
                    {
                        await newUpdateApiCall( dataToBeUpdated, isRegenerate, needLoader,"Policychecklist" );
                    }
                    console.log( dataToBeUpdated );
                } else
                {
                    function removeFontSize( sheetDataa ) {
                        for ( let i = 0; i < sheetDataa.length; i++ )
                        {
                            if ( sheetDataa[ i ] && typeof sheetDataa[ i ] === 'object' )
                            {
                                for ( let j in sheetDataa[ i ] )
                                {
                                    if ( sheetDataa[ i ][ j ] && typeof sheetDataa[ i ][ j ] === 'object' && sheetDataa[ i ][ j ][ 'v' ] === null && ( sheetDataa[ i ][ j ][ 'fs' ] === '9' || sheetDataa[ i ][ j ][ 'fs' ] === '11' ) )
                                    {
                                        delete sheetDataa[ i ][ j ][ 'fs' ];
                                        sheetDataa[ i ][ j ] = null;
                                    }
                                }
                            }
                        }
                        return sheetDataa;
                    }


                    let sheetData = removeFontSize( sheetDataa );

                    let previousLength = null;
                    for ( let i = 0; i < sheetData.length; i++ )
                    {
                        if ( sheetData[ i ] && typeof sheetData[ i ] === 'object' )
                        {
                            for ( let j in sheetData[ i ] )
                            {
                                if ( sheetData[ i ][ j ] && sheetData[ i ][ j ].ct && sheetData[ i ][ j ].ct.fa === "General" && sheetData[ i ][ j ].ct.t === "g" && sheetData[ i ][ j ].bg === null && !( 'm' in sheetData[ i ][ j ] ) &&
                                    !( 'v' in sheetData[ i ][ j ] ) )
                                {
                                    if ( previousLength !== null )
                                    {
                                        sheetData[ i ] = Array( previousLength ).fill( null );
                                    }
                                    break;
                                }
                            }
                        }
                        if ( sheetData[ i ] && Array.isArray( sheetData[ i ] ) )
                        {
                            previousLength = sheetData[ i ].length;
                        }
                    }

                    const keys = Object.keys( tableColumnDetails );

                    const separateArrays = [];

                    if ( keys?.length > 0 && sheetData && sheetData != undefined )
                    {
                        keys.forEach( ( f ) => {
                            const tableData = tableColumnDetails[ f ];
                            if ( tableData && tableData?.range && tableData?.range?.start != undefined && tableData?.range?.end != undefined && tableData?.columnNames && tableData?.range?.end != '' )
                            {
                                const slicedData = sheetData.slice( tableData?.range?.start, tableData?.range?.end + 1 );
                                separateArrays.push( slicedData );
                            }
                        } );
                    }

                    let parentHeaderSegments = [];
                    let tableDateExceptHeaders = [];

                    separateArrays.map( ( e, index ) => {
                        if ( index > 0 )
                        {
                            let tempTableIndex = [];
                            let tempTableColumnName = [];
                            let hasReachedLimit = false;

                            const data = index == 2 ? e[ 1 ] : e[ 0 ];
                            data?.map( ( e1, index1 ) => {
                                if ( index1 > 0 && !hasReachedLimit && e1?.m?.toLowerCase() != "document viewer" )
                                {
                                    tempTableIndex.push( index1 );
                                    tempTableColumnName.push( e1?.m || e1?.v );
                                }
                                else if ( index1 > 0 && e1?.m?.toLowerCase() == "document viewer" )
                                {
                                    hasReachedLimit = true;
                                    tempTableColumnName = [ ...tempTableColumnName, "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill" ];
                                    const maxIndex = Math.max( ...tempTableIndex );
                                    tempTableIndex = [ ...tempTableIndex, maxIndex + 2, maxIndex + 3, maxIndex + 4, maxIndex + 5 ];
                                    parentHeaderSegments.push( { "Table": `Table ${ index + 1 }`, index: tempTableIndex, tempTableColumnName } );
                                }
                            } );
                        }
                    } );

                    const checklistTableColumnDetails = tableColumnDetails;
                    // table dataMapping
                    separateArrays.map( ( f, index ) => {
                        if ( index > 0 )
                        {
                            let keyValuePair = [];
                            const keyValuePairUpdated = true;
                            const selectedTable = checklistTableColumnDetails[ keys[ index ] ];

                            if ( selectedTable?.columnNames && keyValuePairUpdated )
                            {
                                // Remove properties with value 0/null/undefined
                                const filteredData = Object.fromEntries(
                                    Object.entries( selectedTable?.columnNames ).filter( ( [ key, value ] ) => value !== 0 && value != null && value != undefined )
                                );
                                const filteredDataColumnKeys = Object.keys( filteredData );
                                let skipIndex = index == 2 ? 2 : 1;

                                f?.map( ( e1, index1 ) => {
                                    let object = {};
                                    if ( index1 > skipIndex )
                                    {
                                        filteredDataColumnKeys.map( ( column, columnIndex ) => {
                                            const cellData = e1[ filteredData[ column ] ];
                                            const dataOfCell = getTextForUpdate( cellData, true );
                                            object[ column ] = dataOfCell === null || dataOfCell === undefined ? '' : dataOfCell;
                                            if ( filteredDataColumnKeys?.length === ( columnIndex + 1 ) )
                                            {
                                                keyValuePair.push( object );
                                            }
                                        } );
                                    }
                                    if ( f?.length == index1 + 1 )
                                    {
                                        tableDateExceptHeaders.push( { Table: `Table ${ index + 1 }`, NewTemplateData: keyValuePair } );
                                    }
                                } );

                            } else
                            {
                                f?.map( ( e, index1 ) => {
                                    if ( index1 > 1 )
                                    {
                                        let tablePairingData = parentHeaderSegments[ index - 1 ];
                                        if ( tablePairingData )
                                        {
                                            let object = {};
                                            tablePairingData.index.map( ( i, index2 ) => {
                                                if ( e[ i ]?.ct?.s && Array.isArray( e[ i ]?.ct?.s ) )
                                                {
                                                    // Concatenate the 'v' values from the array of objects
                                                    let filteredS = e[ i ]?.ct?.s.filter( ( f ) => f != null );
                                                    let concatenatedValues = filteredS?.map( item => item?.v )?.join( '' );
                                                    concatenatedValues = concatenatedValues?.replace( /\r\n/g, '~~' );
                                                    const finalValue = e[ i ]?.m || e[ i ]?.v || concatenatedValues;
                                                    object[ `${ tablePairingData?.tempTableColumnName[ index2 ] }` ] = finalValue;
                                                } else
                                                {
                                                    // If e[i]?.ct?.s is not an array
                                                    object[ `${ tablePairingData?.tempTableColumnName[ index2 ] }` ] = e[ i ]?.m || e[ i ]?.v || e[ i ]?.ct?.s;
                                                }
                                                if ( tablePairingData.index?.length == index2 + 1 )
                                                {
                                                    keyValuePair.push( object );
                                                }
                                            } )
                                        }
                                    }
                                    if ( f?.length == index1 + 1 )
                                    {
                                        tableDateExceptHeaders.push( { Table: `Table ${ index + 1 }`, NewTemplateData: keyValuePair } );
                                    }
                                } );
                            }
                        }
                    } );

                    // console.log("parentHeaderSegments", parentHeaderSegments);
                    // console.log("tableDateExceptHeaders", tableDateExceptHeaders);
                    const inputData = props.data;
                    const tableNames = inputData.map( item => item.Tablename )

                    const addDotonUpdateprocess = ( value ) => {
                        if ( typeof value === 'string' )
                        {
                            return value.replace( /\•/g, '.' );
                        }
                        return value;
                    }

                    let dataSetToUpdate = []; //variable used for Promise calls
                    let coverageTableData = {};
                    tableNames.forEach( ( table, index ) => {
                        if ( separateArrays && separateArrays.length > 0 )
                        {
                            if ( table === "Table 1" )
                            {
                                const jsonDataToUpdate = [];
                                // console.log( separateArrays );
                                const result1 = separateArrays[ 0 ].map( ( [ _, index1, index2 ] ) => ( {
                                    index1,
                                    index2,
                                } ) );
                                let result = result1.slice( 4 );

                                // let valuesToPass = {};
                                // for ( const item of result )
                                // {
                                //     const key = item.index1 && item.index1.v;
                                //     const value = item.index2 && item.index2.v;
                                //     if ( key && value && ( key === "Named Insured" || key === "Term" || key === "LOB" || key === "Pol#" || key === "Carrier Name" ) )
                                //     {
                                //         valuesToPass[ key ] = value;
                                //     }
                                // }

                                // if ( !isExport && props?.sheetRenderConfig?.FormsCompare == 'true')
                                // {
                                //     formCompareUpdate( false, false, false, valuesToPass );
                                // }

                                const resultWithJoinedValues = result.map( ( { index1, index2 } ) => {
                                    if ( index2.ct && index2.ct.s !== undefined )
                                    {
                                        const joinedValue = index2.ct.s.map( ( { v } ) => v ).join( '' );
                                        const { s, ...ctWithoutS } = index2.ct;
                                        return {
                                            index1,
                                            index2: {
                                                ...index2,
                                                ct: ctWithoutS,
                                                m: joinedValue,
                                                v: joinedValue
                                            }
                                        };
                                    } else
                                    {
                                        return { index1, index2 };
                                    }
                                } );

                                resultWithJoinedValues.forEach( item => {
                                    if ( item?.index2 && item?.index2?.ct?.s !== undefined )
                                    {
                                        if ( Array.isArray( item.index2.ct.s ) && item.index2.ct.s.length > 0 )
                                        {

                                            const extractedData = item.index2.ct.s[ 0 ];

                                            item.index2.m = extractedData.v;
                                            item.index2.v = extractedData.v;

                                            delete item.index2.ct.s;
                                        } else
                                        {
                                            item.index2.m = '';
                                            item.index2.v = '';

                                            delete item.index2.ct.s;
                                        }
                                    }
                                } );


                                // Now jsonDataToUpdate contains the updated data

                                for ( let rowKey in resultWithJoinedValues )
                                {
                                    if ( rowKey != 'len' )
                                    {
                                        const cellData = resultWithJoinedValues[ rowKey ];
                                        const inputData = props.data;
                                        const tableData1 = inputData.find( ( data ) => data.Tablename === "Table 1" );
                                        if ( tableData1 )
                                        {
                                            const removedata1 = sessionStorage.getItem( "index1" );
                                            const removedata2 = sessionStorage.getItem( "index2" );
                                            // luckysheet.undo(removedata1);
                                            // luckysheet.undo(removedata2);
                                            const table1json = tableData1.TemplateData;
                                            const policyVal = table1json.map( item => item[ "PolicyLob" ] );
                                            const cell1Text = cellData.index1?.v || cellData.index1?.ct || '';
                                            const cell2Text = cellData.index2?.v !== undefined ? ( cellData.index2.v || cellData.index2?.ct ) : ( cellData.index2?.ct?.fa === "@" ? "" : cellData.index2?.ct?.fa );
                                            // const cell2Text = cellData.index2?.v !== undefined || cellData.index2?.ct ? (cellData.index2.v || cellData.index2?.ct) : (cellData.index2?.ct?.fa === "@" ? "" : cellData.index2?.ct?.fa);

                                            if ( cell1Text && cell1Text.s && cell1Text.s.length > 0 && cell2Text || cell1Text )
                                            {
                                                // const vValue = cell1Text.s.map(item => item.v || '').join(',') || cell1Text;
                                                const vValue = Array.isArray( cell1Text.s ) ? cell1Text.s.map( item => item.v || '' ).join( ',' ) : cell1Text;
                                                const concatenatedValues = cell2Text?.s && Array.isArray( cell2Text.s ) ? cell2Text.s.map( item => item.v ).join( '' ) : cell2Text;

                                                const policyLOB = policyVal[ 0 ];
                                                const formattedRow = {
                                                    HeaderID: rowKey,
                                                    JOBID: jobId,
                                                    ...( policyLOB && { 'PolicyLob': policyLOB } ),
                                                    Headers: vValue,
                                                    '': concatenatedValues,
                                                };

                                                const addTidleonUpdateprocess = ( value ) => {
                                                    if ( typeof value === 'string' )
                                                    {
                                                        return value.replace(/\n/g, '~~').replace(/"/g, '\\"');
                                                    }
                                                    return value;
                                                };

                                                const jsonString = `{${ Object.entries( formattedRow ).map( ( [ key, value ] ) => {
                                                    if ( key === 'HeaderID' )
                                                    {
                                                        const updatedValue = Number( value ).toString();
                                                        return `"${ key }":${ updatedValue }`;
                                                    } else if ( key === '' )
                                                    {
                                                        return `"${ key }":"${ addTidleonUpdateprocess( value ) }"`;
                                                    } else if (key === 'Headers') {
                                                        const sValue = Array.isArray(value) ? `"${addTidleonUpdateprocess(value.join(', '))}"` : `"${addTidleonUpdateprocess(value)}"`;
                                                        return `"${ key }":${ sValue }`;
                                                    }
                                                    return `"${key}":"${addTidleonUpdateprocess(value)}"`;
                                                } ).join( ',' ) }}`;
                                                jsonDataToUpdate.push( jsonString );
                                            }
                                        }
                                    }
                                }
                                const json = `[${ jsonDataToUpdate.join( ',' ) }]`;
                                updatedDatas[ "Table 1" ] = json;
                                dataSetToUpdate.push( { id: jobId, tableName: "Table 1", data: updatedDatas[ "Table 1" ] } );
                                // updateTemplateData( jobId, "Table 1", updatedDatas[ "Table 1" ] );
                            }
                            else if ( table === "Table 2" )
                            {
                                const masterData = state.filter( ( f ) => f?.Tablename.toLowerCase() === table.toLowerCase() );
                                let duplicateMdata = masterData[ 0 ].TemplateData;
                                let table2 = tableDateExceptHeaders[ 0 ]?.NewTemplateData;
                                table2 = table2.filter( obj => !Object.values( obj ).every( value => value === " " || value === undefined ) );
                                table2.forEach( entry => {
                                    Object.entries( entry ).forEach( ( [ key, value ] ) => {
                                        if ( value === undefined )
                                        {
                                            entry[ key ] = "";
                                        }
                                    } );
                                } );

                                table2.forEach( ( obj ) => {
                                    Object.keys( obj ).forEach( ( key ) => {
                                        obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                                    } );
                                } );

                                table2.forEach( item => {

                                    if ( item.hasOwnProperty( 'Common Declarations' ) )
                                    {
                                        item.CoverageSpecificationsMaster = item[ 'Common Declarations' ];
                                        delete item[ 'Common Declarations' ];
                                    }

                                    const coveragespecificationvalue = duplicateMdata.find( data => data.PolicyLob );
                                    const matchingItem = coveragespecificationvalue.PolicyLob;

                                    if ( matchingItem )
                                    {
                                        item.PolicyLob = matchingItem;
                                    }
                                } );
                                // console.log("table2data", table2);
                                table2 = tableDataFormatting( table2, 2 );
                                const table2data = JSON.stringify( table2 );
                                dataSetToUpdate.push( { id: jobId, tableName: "Table 2", data: table2data } );
                                // updateTemplateData( jobId, "Table 2", table2data );
                            }
                            else if ( table === "Table 3" )
                            {
                                const masterData = state.filter( ( f ) => f?.Tablename.toLowerCase() === table.toLowerCase() );
                                let duplicateMdata = masterData[ 0 ].TemplateData;
                                let table3 = tableDateExceptHeaders[ 1 ]?.NewTemplateData;
                                table3 = table3.filter( obj => !Object.values( obj ).every( value => value === " " || value === undefined ) );
                                table3.forEach( entry => {
                                    Object.entries( entry ).forEach( ( [ key, value ] ) => {
                                        if ( value === undefined )
                                        {
                                            entry[ key ] = "";
                                        }
                                    } );
                                } );

                                table3.forEach( ( obj ) => {
                                    Object.keys( obj ).forEach( ( key ) => {
                                        obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                                    } );
                                } );

                                table3.forEach( item => {

                                    if ( item.hasOwnProperty( 'Common Declarations' ) )
                                    {
                                        item.CoverageSpecificationsMaster = item[ 'Common Declarations' ];
                                        delete item[ 'Common Declarations' ];
                                    }

                                    const coveragespecificationvalue = duplicateMdata.find( data => data.PolicyLob );
                                    const matchingItem = coveragespecificationvalue.PolicyLob;

                                    if ( matchingItem )
                                    {
                                        item.PolicyLob = matchingItem;
                                    }
                                } );

                                // console.log("table3data", table3);
                                table3 = tableDataFormatting( table3, 3 );
                                coverageTableData = { table, data: table3 };
                                const table3data = JSON.stringify( table3 );
                                dataSetToUpdate.push( { id: jobId, tableName: "Table 3", data: table3data } );
                                // updateTemplateData( jobId, "Table 3", table3data );
                            }
                            else if ( table === "Table 4" && tableDateExceptHeaders[ 2 ]?.NewTemplateData )
                            {
                                const masterData = state.filter( ( f ) => f?.Tablename.toLowerCase() === table.toLowerCase() );
                                let duplicateMdata = masterData && masterData?.length > 0 ? masterData[ 0 ].TemplateData : [];
                                let table4 = tableDateExceptHeaders[ 2 ]?.NewTemplateData;
                                table4 = table4.filter( obj => !Object.values( obj ).every( value => value === " " || value === undefined ) );
                                table4.forEach( entry => {
                                    Object.entries( entry ).forEach( ( [ key, value ] ) => {
                                        if ( value === undefined )
                                        {
                                            entry[ key ] = "";
                                        }
                                    } );
                                } );
                                let Table4data = [];

                                table4.forEach( ( obj ) => {
                                    Object.keys( obj ).forEach( ( key ) => {
                                        obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                                    } );
                                } );

                                table4.forEach( item => {
                                    let newItem = { ...item };
                                    if ( duplicateMdata )
                                    {
                                        newItem.PolicyLob = duplicateMdata.find( data => data.PolicyLob )?.PolicyLob;
                                    } else
                                    {

                                        Object.keys( newItem ).forEach( key => {

                                            if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                            {

                                                const coverage = newItem[ key ];
                                                const policylob = key;

                                                newItem.CoverageSpecificationsMaster = coverage;
                                                newItem.PolicyLob = policylob;

                                            }
                                        } );
                                        for ( let key in newItem )
                                        {
                                            if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                            {
                                                delete newItem[ key ];
                                            }
                                        }
                                    }
                                    Table4data.push( newItem );
                                } );
                                // console.log("table4data", Table4data);
                                Table4data = tableDataFormatting( Table4data, 4 );
                                const table4datas = JSON.stringify( Table4data );
                                dataSetToUpdate.push( { id: jobId, tableName: "Table 4", data: table4datas } );
                                // updateTemplateData(jobId, "Table 4", table4datas);
                            }
                            else if ( table === "Table 5" && tableDateExceptHeaders[ 3 ]?.NewTemplateData )
                            {
                                const masterData = state.filter( ( f ) => f?.Tablename.toLowerCase() === table.toLowerCase() );
                                let duplicateMdata = masterData && masterData?.length > 0 ? masterData[ 0 ].TemplateData : [];
                                let table5 = tableDateExceptHeaders[ 3 ]?.NewTemplateData;
                                table5 = table5.filter( obj => !Object.values( obj ).every( value => value === " " || value === undefined ) );
                                table5.forEach( entry => {
                                    Object.entries( entry ).forEach( ( [ key, value ] ) => {
                                        if ( value === undefined )
                                        {
                                            entry[ key ] = "";
                                        }
                                    } );
                                } );
                                let Table5data = [];

                                table5.forEach( ( obj ) => {
                                    Object.keys( obj ).forEach( ( key ) => {
                                        obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                                    } );
                                } );

                                table5.forEach( item => {
                                    let newItem = { ...item };
                                    if ( duplicateMdata )
                                    {
                                        newItem.PolicyLob = duplicateMdata.find( data => data.PolicyLob )?.PolicyLob;
                                    } else
                                    {
                                        Object.keys( newItem ).forEach( key => {
                                            if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                            {

                                                const coverage = newItem[ key ];
                                                const policylob = key;


                                                newItem.CoverageSpecificationsMaster = coverage;
                                                newItem.PolicyLob = policylob;

                                            }
                                        } );
                                        for ( let key in newItem )
                                        {
                                            if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                            {
                                                delete newItem[ key ];
                                            }
                                        }
                                    }
                                    Table5data.push( newItem );
                                } );
                                // console.log("table5data", Table5data);
                                Table5data = tableDataFormatting( Table5data, 5 );
                                const table5datas = JSON.stringify( Table5data );
                                dataSetToUpdate.push( { id: jobId, tableName: "Table 5", data: table5datas } );
                                // updateTemplateData( jobId, "Table 5", table5datas );
                            }
                            else if ( table === "Table 6" && tableDateExceptHeaders[ 4 ]?.NewTemplateData )
                            {

                                const masterData = state.filter( ( f ) => f?.Tablename.toLowerCase() === table.toLowerCase() );
                                let duplicateMdata = masterData[ 0 ].TemplateData;
                                let table6 = tableDateExceptHeaders[ 4 ]?.NewTemplateData;
                                table6 = table6.filter( obj => !Object.values( obj ).every( value => value === " " || value === undefined ) );
                                table6.forEach( entry => {
                                    Object.entries( entry ).forEach( ( [ key, value ] ) => {
                                        if ( value === undefined )
                                        {
                                            entry[ key ] = "";
                                        }
                                    } );
                                } );
                                let Table6data = [];

                                table6.forEach( ( obj ) => {
                                    Object.keys( obj ).forEach( ( key ) => {
                                        obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                                    } );
                                } );

                                table6.forEach( item => {
                                    let newItem = { ...item };
                                    if ( duplicateMdata )
                                    {
                                        newItem.PolicyLob = duplicateMdata.find( data => data.PolicyLob )?.PolicyLob;
                                    } else
                                    {
                                        Object.keys( newItem ).forEach( key => {
                                            if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                            {
                                                const coverage = newItem[ key ];
                                                const policylob = key;
                                                newItem.CoverageSpecificationsMaster = coverage;
                                                newItem.PolicyLob = policylob;
                                            }
                                        } );
                                        for ( let key in newItem )
                                        {
                                            if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                            {
                                                delete newItem[ key ];
                                            }
                                        }
                                    }
                                    Table6data.push( newItem );
                                } );

                                // console.log("table6data", Table6data);
                                Table6data = tableDataFormatting( Table6data, 6 );
                                const table6datas = JSON.stringify( Table6data );
                                dataSetToUpdate.push( { id: jobId, tableName: "Table 6", data: table6datas } );
                                // updateTemplateData( jobId, "Table 6", table6datas );
                            }
                            else if ( table === "Table 7" && tableDateExceptHeaders[ 5 ]?.NewTemplateData )
                            {

                                const masterData = state.filter( ( f ) => f?.Tablename.toLowerCase() === table.toLowerCase() );
                                let duplicateMdata = masterData && masterData?.length > 0 ? masterData[ 0 ].TemplateData : [];
                                const checkdata = masterData.map( e => e.TemplateData );
                                if ( checkdata !== undefined && checkdata.length > 0 && checkdata[ 0 ].length > 0 )
                                {
                                    let duplicateMdata = masterData[ 0 ].TemplateData;
                                    let table7 = tableDateExceptHeaders[ 5 ]?.NewTemplateData;
                                    table7 = table7.filter( obj => !Object.values( obj ).every( value => value === " " || value === undefined ) );
                                    table7.forEach( entry => {
                                        Object.entries( entry ).forEach( ( [ key, value ] ) => {
                                            if ( value === undefined )
                                            {
                                                entry[ key ] = "";
                                            }
                                        } );
                                    } );
                                    let Table7data = [];

                                    table7.forEach( ( obj ) => {
                                        Object.keys( obj ).forEach( ( key ) => {
                                            obj[ key ] = addDotonUpdateprocess( obj[ key ] );
                                        } );
                                    } );

                                    table7.forEach( item => {
                                        let newItem = { ...item };
                                        if ( duplicateMdata )
                                        {
                                            newItem.PolicyLob = duplicateMdata.find( data => data.PolicyLob )?.PolicyLob;
                                        } else
                                        {
                                            Object.keys( newItem ).forEach( key => {
                                                if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                                {

                                                    const coverage = newItem[ key ];
                                                    const policylob = key;


                                                    newItem.CoverageSpecificationsMaster = coverage;
                                                    newItem.PolicyLob = policylob;

                                                }
                                            } );
                                            for ( let key in newItem )
                                            {
                                                if ( Formpolicydata.some( Formpolicydata => key.includes( Formpolicydata ) ) )
                                                {
                                                    delete newItem[ key ];
                                                }
                                            }
                                        }
                                        Table7data.push( newItem );
                                    } );

                                    // console.log("table7data", Table7data);
                                    Table7data = tableDataFormatting( Table7data, 7 );
                                    const table7datas = JSON.stringify( Table7data );
                                    dataSetToUpdate.push( { id: jobId, tableName: "Table 7", data: table7datas } );
                                    // updateTemplateData( jobId, "Table 7", table7datas );
                                }
                            }
                            if ( index == ( tableNames?.length - 1 ) )
                            {
                                setTimeout( () => {
                                    // updataPHProcess( isRegenerate );
                                }, 3000 );
                            }
                        }
                    } );

                    const coverageMaster = state.filter( ( f ) => f?.Tablename.toLowerCase() === coverageTableData?.table?.toLowerCase() );
                    if ( coverageMaster && coverageMaster?.length > 0 && coverageMaster[ 0 ]?.isMultipleLobSplit )
                    {
                        const convertedData = await mapLOBColumns( coverageTableData?.data, token, jobId );
                        dataSetToUpdate = dataSetToUpdate.map( ( e ) => {
                            if ( e?.tableName == coverageTableData?.table )
                            {
                                e.data = JSON.stringify( convertedData );
                            }
                            return e;
                        } );
                    }

                    if ( dataSetToUpdate?.length > 0 && props?.sheetRenderConfig?.PolicyReviewChecklist == 'true')
                    {
                        if ( needLoader )
                        {
                            document.body.classList.add( 'loading-indicator' );
                        }
                        const promiseResponse = Promise.all( dataSetToUpdate.map( async ( item ) => {

                            if ( isExport == true )
                            {
                                dataSetToUpdate.forEach( item => {
                                    const sanitizedData = item.data.replace( /[\u0000-\u001F\u007F-\u009F]/g, '' );   //sanitize the JSON string by removing any problematic control characters before parsing it.  so dont remove this
                                    let parsedData;
                                    try
                                    {
                                        parsedData = JSON.parse( sanitizedData );
                                    } catch ( error )
                                    {
                                        console.error( "parsing json error catch:", error );
                                        return;
                                    }

                                    parsedData.forEach( obj => {
                                        if ( obj[ "" ] !== undefined )
                                        {
                                            obj[ "NoColumnName" ] = obj[ "" ];
                                            delete obj[ "" ];
                                        }
                                    } );
                                    item.data = JSON.stringify( parsedData );
                                } );

                                const modifiedTabledata = dataSetToUpdate.map( item => ( {
                                    Id: item.id,
                                    TableName: item.tableName,
                                    Data: item.data
                                } ) );

                                const dataFromOnUpdateClick = modifiedTabledata;

                                if ( typeof callback === "function" )
                                {
                                    callback( dataFromOnUpdateClick );
                                }
                            } else
                            {
                                const response = await updateTemplateData( item?.id, item?.tableName, item?.data );
                                return response;
                            }

                        } ) );
                        if ( !isExport && needLoader )
                        {
                            promiseResponse.then(
                                ( res ) => {
                                    const isAllSucces = res?.filter( ( f ) => f == "error" )?.length == 0;
                                    if ( isAllSucces || !isAllSucces )
                                    {
                                        setMsgVisible( true );
                                        setMsgClass( 'alert success' );
                                        // setMsgText('Data Updated');
                                        setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3000 );
                                        updataPHProcess( isRegenerate, needLoader );
                                    }
                                    // else{
                                    //     document.body.classList.remove( 'loading-indicator' );
                                    // }
                                },
                                ( error ) => { console.error( 'Error:', error ); }
                            );
                        }


                    }
                }
            }

            if (!isExport && QacFlag == "Forms Compare") {
                
                formCompareUpdate(true, false, false);
            }

            if ( QacFlag == 'Exclusion' )
            {
                luckysheet.exitEditMode();
                exclusionUpdate();
                return;
            }
        }

    };

    const Autoupdateclick = async ( autoupdate ) => {
        if ( autoupdate == true && (luckysheet != undefined && luckysheet != null))
        {
            try
            {
                     await onUpdateClick( false, false ,false);
            } catch ( error )
            {
                console.error( "Error occurred:", error );
            }
        }
    }

    const Regenrateclick = async ( isSave ) => {
        let QacFlag = luckysheet?.getSheet()?.name
        if ( QacFlag == 'QAC not answered questions' )
        {
            return;
        }
        const parallelFlag = props?.formCompareData && props?.formCompareData?.length > 0 ? props?.formCompareData[ 0 ]?.isFormCompareApplicable : false;
        let flagCheck = luckysheet.getSheet()?.name; //formscompare
        if ( flagCheck === 'Forms Compare' || flagCheck === 'Exclusion' )
        {
            setOpenDialog( true );
            setMsgText( 'Save and Regenrate only applicable for PolicyReviewChecklist' );
            setTimeout( () => {
                setMsgVisible( false );
                setMsgText( '' );
            }, 4500 );
            // const message = "Save and Regenrate only applicable for PolicyReviewChecklist";
            { openDialog && <DialogComponent isOpen={ openDialog } onClose={ ( e ) => handleDialogClose( e ) } message={ msgText } /> }
        } else
        {
            if ( parallelFlag === true && isSave === false )
            {
                try
                {
                    // await onUpdateClick( true, true, false, );
                    // await formCompareUpdate( true, false );
                    await onUpdateClick( false, true, false )
                } catch ( error )
                {
                    console.error( "Error occurred:", error );
                }
            } else
            {
                await onUpdateClick( true, true, false );
            }
        }
    }

    const Exportclick = async (hasExport) => {
        let getFlag = luckysheet.getSheet().name;
        let sheetdatas = sheetsDropOption;
        const mergedData = [];
    
        if (getFlag !== 'QAC not answered questions') {
            if (sheetdatas.some(item => item.key === "Forms Compare") && getFlag !== 'Forms Compare') {
                const jobId = props.selectedJob;  
                const sheetType = "formscompare";  
                const token = sessionStorage.getItem("token");  
    
                try {
                    const dataFromExport = await ExportData(jobId, sheetType, token);
                    const Tabledata1 = dataFromExport;
                    mergedData.push(...Tabledata1);
                } catch (error) {
                    console.error('Error exporting data:', error);
                }
            } else {
                await formCompareUpdate(true, true, (dataFrom) => {
                    const Tabledata1 = dataFrom;
                    mergedData.push(...Tabledata1);
                });
            }
            if (sheetdatas.some(item => item.key === "PolicyReviewChecklist") && getFlag !== 'PolicyReviewChecklist') {
                const jobId = props.selectedJob;  
                const sheetType = "policyreviewchecklist";  
                const token = sessionStorage.getItem("token");  
    
                try {
                    const dataFromExport = await ExportData(jobId, sheetType, token);
                    const Tabledata2 = dataFromExport;
                    mergedData.push(...Tabledata2);
                } catch (error) {
                    console.error('Error exporting data:', error);
                }
            }else{
                await onUpdateClick(false, false, true, (dataFromOnUpdateClick) => {
                    const Tabledata2 = dataFromOnUpdateClick;
                    // console.log("tabledata2", Tabledata2);
                    mergedData.push(...Tabledata2);
                });
            }
           
            if (sheetdatas.some(item => item.key === "Exclusion") && getFlag !== 'Exclusion') {
                const jobId = props.selectedJob;  
                const sheetType = "Exclusion"; 
                const token = sessionStorage.getItem("token");  
    
                try {
                    const dataFromExport = await ExportData(jobId, sheetType, token);
                    const Tabledata3 = dataFromExport;
                    mergedData.push(...Tabledata3);
                } catch (error) {
                    console.error('Error exporting data:', error);
                }
            }
            await exclusionUpdate(true, (dataExclusionOnUpdateClick) => {
                const Tabledata3 = dataExclusionOnUpdateClick;
                mergedData.push(...Tabledata3);
            });
    
            if (brokerId === "1162" || brokerId === "1003") {
                const Tabledata4 = props.qacdataapi;
                mergedData.push(...Tabledata4);
            }
    
            const uniqueMergedData = mergedData.reduce((acc, curr) => {
                if (!acc.find(item => item.TableName === curr.TableName)) {
                    acc.push(curr);
                }
                return acc;
            }, []);
    
            const filteredData = uniqueMergedData.filter(item => {
                const data = JSON.parse(item.Data);
                return !Array.isArray(data) || data.length > 0;
            });
    
            let formTableData = [];
    
            filteredData.forEach(item => {
                if (item.TableName === "FormTable 1" || item.TableName === "Table 1" || item.TableName === "ExclusionTable" || item.TableName === "HighVolumeTable1") {
                    formTableData.push(item.TableName);
                }
            });
    
            const formTableDataJson = '"' + JSON.stringify(formTableData) + '"';
    
            if (hasExport === true) {
                let exportcheck = sessionStorage.getItem('onUpdateClickCalled');
                if (exportcheck === "false" || exportcheck === "true") {
                    const response = exportExcelData(filteredData, formTableDataJson);
                    return response;
                }
            } else {
                var gridbackupdata = "{" + '"' + "Data" + '"' + ":" + JSON.stringify(formTableDataJson) + "," + '"' + "Tabledata" + '"' + ":" + JSON.stringify(filteredData) + "}";
                var jobid = props.selectedJob;
                const response = gridBackupTemplateData(jobid, gridbackupdata);
                return response;
            }
        }
    };

    const GridBackupSave = async () => {
        let getFlag = luckysheet.getSheet().name;
        let sheetdatas = sheetsDropOption;
        const mergedData = [];
    
        if (getFlag !== 'QAC not answered questions') {
            if ( getFlag == 'Forms Compare') {
                await formCompareUpdate(false, true, (dataFrom) => {
                    const Tabledata1 = dataFrom;
                    mergedData.push(...Tabledata1);
                    if (mergedData .length !=0) {
                        const uniqueMergedData = mergedData.reduce((acc, curr) => {
                            if (!acc.find(item => item.TableName === curr.TableName)) {
                                acc.push(curr);
                            }
                            return acc;
                        }, []);
                    
                        const filteredData = uniqueMergedData.filter(item => {
                            const data = JSON.parse(item.Data);
                            return !Array.isArray(data) || data.length > 0;
                        });
                    
                        let formTableData = [];
                    
                        filteredData.forEach(item => {
                            if (item.TableName === "FormTable 1" || item.TableName === "Table 1" || item.TableName === "ExclusionTable" || item.TableName === "HighVolumeTable1") {
                                formTableData.push(item.TableName);
                            }
                        });
                    
                        const formTableDataJson = '"' + JSON.stringify(formTableData) + '"';
                    
                      
                            var gridbackupdata = "{" + '"' + "Data" + '"' + ":" + JSON.stringify(formTableDataJson) + "," + '"' + "Tabledata" + '"' + ":" + JSON.stringify(filteredData) + "}";
                            var jobid = props.selectedJob;
                            const response = gridBackupTemplateData(jobid, gridbackupdata);
                            return response;
                    
                    }
                });
               

                
            }
            if (getFlag == 'PolicyReviewChecklist') {
                
                await onUpdateClick(false, false, true, (dataFromOnUpdateClick) => {
                    const Tabledata2 = dataFromOnUpdateClick;
                    // console.log("tabledata2", Tabledata2);
                    mergedData.push(...Tabledata2);
                });
            }
           
            if ( getFlag == 'Exclusion') {
                
            await exclusionUpdate(true, (dataExclusionOnUpdateClick) => {
                const Tabledata3 = dataExclusionOnUpdateClick;
                mergedData.push(...Tabledata3);
                if (mergedData .length !=0) {
                    const uniqueMergedData = mergedData.reduce((acc, curr) => {
                        if (!acc.find(item => item.TableName === curr.TableName)) {
                            acc.push(curr);
                        }
                        return acc;
                    }, []);
                
                    const filteredData = uniqueMergedData.filter(item => {
                        const data = JSON.parse(item.Data);
                        return !Array.isArray(data) || data.length > 0;
                    });
                
                    let formTableData = [];
                
                    filteredData.forEach(item => {
                        if (item.TableName === "FormTable 1" || item.TableName === "Table 1" || item.TableName === "ExclusionTable" || item.TableName === "HighVolumeTable1") {
                            formTableData.push(item.TableName);
                        }
                    });
                
                    const formTableDataJson = '"' + JSON.stringify(formTableData) + '"';
                
                  
                        var gridbackupdata = "{" + '"' + "Data" + '"' + ":" + JSON.stringify(formTableDataJson) + "," + '"' + "Tabledata" + '"' + ":" + JSON.stringify(filteredData) + "}";
                        var jobid = props.selectedJob;
                        const response = gridBackupTemplateData(jobid, gridbackupdata);
                        return response;
                
                }
            });
        }

    if (getFlag != 'Forms Compare' && getFlag != 'Exclusion') {
        const uniqueMergedData = mergedData.reduce((acc, curr) => {
            if (!acc.find(item => item.TableName === curr.TableName)) {
                acc.push(curr);
            }
            return acc;
        }, []);

        const filteredData = uniqueMergedData.filter(item => {
            const data = JSON.parse(item.Data);
            return !Array.isArray(data) || data.length > 0;
        });

        let formTableData = [];

        filteredData.forEach(item => {
            if (item.TableName === "FormTable 1" || item.TableName === "Table 1" || item.TableName === "ExclusionTable" || item.TableName === "HighVolumeTable1") {
                formTableData.push(item.TableName);
            }
        });

        const formTableDataJson = '"' + JSON.stringify(formTableData) + '"';

      
            var gridbackupdata = "{" + '"' + "Data" + '"' + ":" + JSON.stringify(formTableDataJson) + "," + '"' + "Tabledata" + '"' + ":" + JSON.stringify(filteredData) + "}";
            var jobid = props.selectedJob;
            const response = gridBackupTemplateData(jobid, gridbackupdata);
            return response;
    }
           
            
        }
    };
    const updateTemplateData = async ( jobId, tableName, json ) => {
        document.body.classList.add( 'loading-indicator' );
        const headers = {
            'Authorization': `Bearer ${ token }`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${ baseUrl }/api/ProcedureData/Update?jobId=${ jobId }`;

        // let isApiCallPending = true;
        try
        {
            const response = await axios( {
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    JobId: jobId,
                    TableName: tableName,
                    NewTemplateData: json
                }
            } );
            if ( response.status !== 200 )
            {
                return "error";
                // throw new Error( `HTTP error! Status: ${ response.status }` );
            }

            return response.data;
        } catch ( error )
        {
            // console.error( 'Error:', error );
            return "error"; // Rethrow the error to be caught in the calling function
        } finally
        {
            document.body.classList.remove( 'loading-indicator' );
            // return "success";
        }
    };


    const gridBackupTemplateData = async ( jobId, gridbackupdata ) => {
        const headers = {
            'Authorization': `Bearer ${ token }`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${ baseUrl }/api/Excel/UpdateCheckListGridBackUpData/UpdateCheckListGridBackUpData`;

        // let isApiCallPending = true;
        try
        {
            const response = await axios( {
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    JobId: jobId,
                    NewTemplateData: gridbackupdata,
                    TableName: "griddata"
                }
            } );
            if ( response.status !== 200 )
            {
                return "error";
                // throw new Error( `HTTP error! Status: ${ response.status }` );
            }

            return response.data;
        } catch ( error )
        {
            // console.error( 'Error:', error );
            return "error"; // Rethrow the error to be caught in the calling function
        } finally
        {
            // return "success";
        }
    };

    const exportExcelData = async ( Tabledata, TableNames ) => {
        document.body.classList.add( 'loading-indicator' );
        let setupdateclicktrue = sessionStorage.setItem( 'onUpdateClickCalled', true );
        const Token = await processAndUpdateToken( token );
        const headers = {
            'Authorization': `Bearer ${ Token }`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${ baseUrl }/api/Excel/ExportExcel`;

        try
        {
            const response = await axios( {
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    Data: TableNames,
                    Tabledata: Tabledata
                },
                responseType: 'blob'
            } );
            if ( response.status !== 200 )
            {
                return "error";
            }

            const url = window.URL.createObjectURL( new Blob( [ response.data ] ) );
            const link = document.createElement( 'a' );
            link.href = url;
            link.setAttribute( 'download', `${ Tabledata[ 0 ].Id }GridExcel.xlsx` );
            document.body.appendChild( link );
            link.click();
            let setupdateclickfalse = sessionStorage.setItem( 'onUpdateClickCalled', false );
            return "success";
        } catch ( error )
        {
            return "error";
        } finally
        {
            document.body.classList.remove( 'loading-indicator' );
            // return "success";
        }
    };
    let previousCell = { row: null, col: null, text: 'string' };
    const handleCellSelection = async (range, flagCheck) => {
        if (flagCheck === 'PolicyReviewChecklist') {
            if (range && range[0]?.row && range[0]?.column && range[0]?.row[0] === range[0]?.row[1] &&
                range[0]?.column[0] === range[0]?.column[1] && [ 4, 5, 6, 7].includes(range[0]?.row[0]) &&
                range[0]?.column[0] === 4) {
                const currentRow = range[0]?.row[0];
                const currentCol = 4;

                let getRowValue = luckysheet.getcellvalue(currentRow);
                const objectAtIndex4 = getRowValue[4];
                // console.log(objectAtIndex4);
                if (objectAtIndex4?.ct?.s && Array.isArray(objectAtIndex4?.ct?.s)) {
                    let varianceText = objectAtIndex4?.ct?.s[1]?.v
                    if (varianceText === 'Matched' || varianceText === 'All Variances' || varianceText === 'Variances' || varianceText === 'Details not available in the document') {
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


    const matchedOrUnMatchedFilter = (rowIndex) => {
        if (rowIndex === 6 || rowIndex === 7 || rowIndex === 5) {
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
                const sourceColumns = Object.keys(columnKeys).filter((key) => columnKeys[key] > columnKeys["ChecklistQuestions"] &&
                    columnKeys[key] < columnKeys["Observation"]);
                if (sourceColumns && sourceColumns?.length > 0) {
                    const findData = checklistData.find((fi) => fi?.Tablename == f);
                    if (findData?.TemplateData && findData?.TemplateData?.length > 0) {
                        findData?.TemplateData.forEach((item, itemIndex) => {
                            if (rowIndex === 6) { //for unmatched(variance)
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
                            } else if (rowIndex === 5) {
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
                            } else if (rowIndex === 7) {
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
           
        }
    }
    const showOrHideRecords = (rowSet) => {
        const config = luckysheet.getConfig();
        const hiddenRows = config?.rowhidden ? Object.keys(config?.rowhidden) : [];
        if (hiddenRows && hiddenRows?.length > 0) {
            const parsedSet = hiddenRows.map((f) => parseInt(f));
            const grouppedSet = groupNumbers(parsedSet);
            grouppedSet.forEach((f) => {
                luckysheet.showRow(f[0], f[f?.length - 1]);
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
        }
    }


    let isLuckysheetRendered = false;
    const renderLuckySheet = async ( needConfigAdjustment, sheetConfig, isDelete) => {
        let data1 = FormCompare_appconfigdata.forms.celldata;
        let configData1 = FormCompare_appconfigdata.forms.config;
        if ( data1 && isFormApplicable )
        {
            FormCompare_appconfigdata.forms.celldata = data1.map( ( item ) => {
                if ( item?.v && item?.v?.fs && item?.v?.fs > 9 )
                {
                    item.v.fs = item.v.fs - 5;
                } else if ( item?.v && item?.v?.ct && item?.v?.ct?.s?.length > 0 )
                {
                    item?.v?.ct?.s.map( ( subItem, index ) => {
                        if ( subItem?.fs )
                        {
                            item.v.ct.s[ index ].fs = 8;
                        } else
                        {
                            item.v.ct.s[ index ][ "fs" ] = 7;
                        }
                    } );
                } else
                {
                    if ( ( item?.v?.m || item?.v?.v ) && !item?.v?.fs )
                    {
                        item.v[ "fs" ] = 8;
                    }
                }
                return item;
            } );

        }
        if ( configData1 )
        {
            if ( configData1?.columnlen )
            {
                const keys = Object.keys( configData1?.columnlen );
                if ( keys?.length > 0 )
                {
                    keys.map( ( key ) => {
                        configData1.columnlen[ key ] = 220;
                    } )
                }
            }
            if ( configData1?.rowlen )
            {
                const keys = Object.keys( configData1?.rowlen );
                if ( keys?.length > 0 )
                {
                    keys.map( ( key ) => {
                        configData1.rowlen[ key ] = configData1.rowlen[ key ] - 10 > 30 ? configData1.rowlen[ key ] - 10 : 30;
                    } )
                }
            }
            FormCompare_appconfigdata.forms.config = configData1;
        }
        if ( ( !isLuckysheetRendered || isDelete ) && luckysheet )
        { // Check if Luckysheet is not rendered and luckysheet instance exists
            isLuckysheetRendered = true;

            const qacDataSet = await getQACData( jobId, token );
            if ( luckysheet )
            {
                if ( needConfigAdjustment )
                {
                    let data = apiDataConfig.demo.celldata;
                    let configData = apiDataConfig.demo.config;
                    if ( data )
                    {
                        apiDataConfig.demo.celldata = data.map( ( item ) => {
                            if ( item?.v && item?.v?.fs && item?.v?.fs > 9 )
                            {
                                item.v.fs = item.v.fs - 5;
                            } else if ( item?.v && item?.v?.ct && item?.v?.ct?.s?.length > 0 )
                            {
                                item?.v?.ct?.s.map( ( subItem, index ) => {
                                    if ( subItem?.fs )
                                    {
                                        item.v.ct.s[ index ].fs = 8;
                                    } else
                                    {
                                        item.v.ct.s[ index ][ "fs" ] = 7;
                                    }
                                } );
                            } else
                            {
                                if ( ( item?.v?.m || item?.v?.v ) && !item?.v?.fs )
                                {
                                    item.v[ "fs" ] = 7;
                                }
                            }
                            return item;
                        } );

                    }
                    if ( configData )
                    {
                        if ( configData?.columnlen )
                        {
                            const keys = Object.keys( configData?.columnlen );
                            if ( keys?.length > 0 )
                            {
                                keys.map( ( key ) => {
                                    configData.columnlen[ key ] = 250;
                                } )
                            }
                        }
                        if ( configData?.rowlen )
                        {
                            const keys = Object.keys( configData?.rowlen );
                            if ( keys?.length > 0 )
                            {
                                keys.map( ( key ) => {
                                    configData.rowlen[ key ] = configData.rowlen[ key ] - 10 > 15 ? configData.rowlen[ key ] - 10 : 15;
                                } )
                            }
                        }
                        apiDataConfig.demo.config = configData;
                    }
                } else
                {
                    if ( isDelete )
                    {

                        if ( sheetConfig[ 0 ]?.top == undefined || sheetConfig[ 0 ]?.top == null )
                        {
                            sheetConfig[ 0 ].top = sessionStorage.getItem( "sheetConfigTop" );
                        } else
                        {
                            sessionStorage.setItem( "sheetConfigTop", sheetConfig[ 0 ]?.top );
                        }
                    }

                    let sheetcheck = luckysheet.getSheet();
                    // if ( sheetcheck.name == "PolicyReviewChecklist" )
                    // {
                    //     apiDataConfig.demo[ "scrollTop" ] = sheetConfig[ 0 ]?.top - 50 > 0 ? sheetConfig[ 0 ]?.top - 50 : 0;
                    // } else if ( sheetcheck.name == "Forms Compare" )
                    // {
                    //     FormCompare_appconfigdata.forms[ "scrollTop" ] = sheetConfig[ 0 ]?.top - 50 > 0 ? sheetConfig[ 0 ]?.top - 50 : 0;
                    //     FormCompare_appconfigdata.forms[ "status" ] = "1";
                    //     apiDataConfig.demo[ "status" ] = 0;
                    // } else if ( sheetcheck.name == "Exclusion" )
                    // {
                    //     FormCompare_appconfigdata.forms[ "status" ] = 0;
                    //     apiDataConfig.demo[ "status" ] = 0;
                    //     exclusionDatafigdata.exclusion[ "status" ] = "1";
                    // }
                }
                // Create options for Luckysheet
                const sheetRenderConfig = props?.sheetRenderConfig;
                let sheetDataSet = [];
                if(sheetRenderConfig?.PolicyReviewChecklist == 'true'){
                    sheetDataSet = [apiDataConfig.demo ];
                } else if(sheetRenderConfig?.FormsCompare == 'true'){
                    sheetDataSet = [FormCompare_appconfigdata.forms];
                } else  if(sheetRenderConfig?.Exclusion == 'true'){
                    sheetDataSet = [exclusionDatafigdata.exclusion];
                }  else  if(sheetRenderConfig?.QAC_not_answered_questions == 'true' &&  qacDataSet?.canRender){
                    sheetDataSet = [qacDataSet?.data];
                }
                // if ( qacDataSet?.canRender && isFormApplicable )
                // {
                //     sheetDataSet = [ apiDataConfig.demo, FormCompare_appconfigdata.forms, qacDataSet?.data ];
                // } else if ( qacDataSet?.canRender )
                // {
                //     sheetDataSet = [ apiDataConfig.demo, qacDataSet?.data ];
                // } else if ( isFormApplicable )
                // {
                //     sheetDataSet = [ apiDataConfig.demo, FormCompare_appconfigdata.forms ];
                // } else
                // {
                //     sheetDataSet = [ apiDataConfig.demo ];
                // }
                // if ( brokerId === '1003' )
                // {
                //     sheetDataSet = [ ...sheetDataSet, exclusionDatafigdata.exclusion ]
                // }
                
                const selectedOptions = dropDownOption?.map((sheet) => ({ key: sheet, text: sheet }));
                setSheetDropOption(selectedOptions);

                const options = {
                    container: "luckysheet", // Container ID
                    showinfobar: false,
                    showsheetbar: true,
                    lang: 'en',
                    // data:  filteredSheet || renderDefaultSheetData,
                    data:  sheetDataSet,
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
                    // showstatisticBarConfig: {
                    //     zoom: false,
                    // },
                    hook: {
                        workbookCreateAfter( json ) {
                            luckysheet.setSheetZoom( 1 );// after rendering setting the screen zoom size to 0.65 for scroll support in chrome
                        },
                        updated: function ( val ) {
                            // undoing the gragged value based on session
                            if(val?.type && val.type === "datachange" ){
                                sessionStorage.setItem("cs_range_select","true");
                                setTimeout(() => {
                                    sessionStorage.setItem("cs_range_select","false");
                                }, 2000);
                            }
                            // console.log(val);
                            //on undoing if insert row is reverted updating the state used for tables are updating accordingly.
                            const sheetName = luckysheet?.getSheet().name;
                            if ( val?.type && val.type === "addRC" && !val?.data && val?.curdata?.length > 0 )
                            {
                                const selectedcolumnindex = val.dataRange[ 0 ]?.row[ 0 ];
                                if ( selectedcolumnindex != undefined && selectedcolumnindex != null )
                                {
                                    onDeleteUpdateTableColumnDetails( selectedcolumnindex, 1 );
                                }
                            }
                            else if ( val?.type && val.type === "addRC" && val?.data && val?.data?.length > 0 )
                            {
                                const selectedcolumnindex = val.dataRange[ 0 ]?.row[ 0 ];
                                if ( selectedcolumnindex != undefined && selectedcolumnindex != null )
                                {
                                    onInsertUpdateTableColumnDetails( selectedcolumnindex, 1 );
                                }
                            }
                            else if ( val?.type && val.type === "delRC" && !val?.data && val?.curdata?.length > 0 )
                            {
                                const selectedcolumnindex = val.dataRange[ 0 ]?.row[ 0 ];
                                const difference = val.dataRange[ 0 ]?.row[ 1 ] - val.dataRange[ 0 ]?.row[ 0 ];
                                if ( selectedcolumnindex != undefined && selectedcolumnindex != null )
                                {
                                    onInsertUpdateTableColumnDetails( selectedcolumnindex, difference + 1 );
                                }
                            }
                            else if ( val?.type && val.type === "delRC" && val?.data && val?.data?.length > 0 )
                            {
                                const selectedcolumnindex = val.dataRange[ 0 ]?.row[ 0 ];
                                const difference = val.dataRange[ 0 ]?.row[ 1 ] - val.dataRange[ 0 ]?.row[ 0 ];
                                if ( selectedcolumnindex != undefined && selectedcolumnindex != null )
                                {
                                    onDeleteUpdateTableColumnDetails( selectedcolumnindex, difference + 1 );
                                }
                            }
                            else if( val && val != undefined && val?.type === "zoomChange")  //Zoom-in and zoom-out scroll Config set
                            {
                                if(val?.curZoomRatio > val?.zoomRatio || val?.zoomRatio === 1){
                                    let sheetScrollConfigSet = luckysheet.getluckysheet_select_save();
                                    $("#luckysheet-scrollbar-x").scrollLeft(sheetScrollConfigSet[0].left - 100);
                                    $("#luckysheet-scrollbar-y").scrollTop(sheetScrollConfigSet[0].top - 100);
                                } else if(val?.curZoomRatio < val?.zoomRatio){
                                    let sheetScrollConfigSet = luckysheet.getluckysheet_select_save();
                                    $("#luckysheet-scrollbar-x").scrollLeft(sheetScrollConfigSet[0].left - 100);
                                    $("#luckysheet-scrollbar-y").scrollTop(sheetScrollConfigSet[0].top - 100);
                                }
                            }
                        },
                        rangeSelect: function ( index, sheet ) {
                            //In this copyvalueset, we have the paste data array, and we get the targetrow to set the cell value in the sheet. After finishing this process, we can do the formatting data process in cellAllRenderBefore [HOOK]
                            let range = luckysheet.getRange();
                            let sheetchecks = luckysheet.getSheet().name;
                            let excededcolumn = range[ 0 ].column[ 1 ];
                            if ( excededcolumn != 12 && excededcolumn != 49 )
                            {
                                sessionStorage.removeItem( "ctrloptions" );
                                let ctrldata = JSON.stringify( range );
                                sessionStorage.setItem( "ctrloptions", ctrldata );
                            }

                            if (sheetchecks == "Forms Compare" || sheetchecks == "PolicyReviewChecklist" )
                            {
                                let selectedIndex = range[ 0 ].row[ 0 ];
                                let selectedColIndex = range[ 0 ].column[ 0 ];
                                let tabledata = sheetchecks == "PolicyReviewChecklist" ? tableColumnDetails : sheetchecks == "Forms Compare" ? formTableColumnDetails : sheetchecks == "Exclusion" ? exTableColumnDetails : "";
                                const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
                                const selectedTable = findTableForIndex( selectedIndex, tabledata, excludedColumns );

                                let firstindex = range[ 0 ].column[ 0 ];
                                let secoundindex = range[ 0 ].column[ 1 ];
                                if ( firstindex === 0 || secoundindex !== 6 )
                                {
                                    let formsdata = JSON.stringify( luckysheet.getRange() );
                                    localStorage.setItem( 'formrange', formsdata );
                                }
                                if ( secoundindex !== 11 )
                                {
                                    let secoundtabledata = JSON.stringify( luckysheet.getRange() );
                                    localStorage.setItem( 'secoundtabledata', secoundtabledata );
                                    // setSecoundtablerange( secoundtabledata );
                                }


                                let formsdata = JSON.stringify( luckysheet.getRange() );
                                localStorage.setItem( 'formrange', formsdata );

                                if ( firstindex !== 0 && secoundindex !== 6 )
                                {
                                    let formsdata = JSON.stringify( luckysheet.getRange() );
                                    localStorage.setItem( 'formfullrange', formsdata );
                                } else if ( firstindex !== 0 )
                                {
                                    let formsdata = JSON.stringify( luckysheet.getRange() );
                                    localStorage.setItem( 'formbackupfullrange', formsdata );
                                }

                                // if(sheetchecks == "PolicyReviewChecklist") {
                                //     setTimeout(() => {
                                //         luckysheet.exitEditMode();
                                //         // delKeyRestrictCSColumn(null, selectedTable, selectedIndex, selectedColIndex, tabledata);
                                //     }, 100);
                                // }
                                      //parvesh
                                if (range && range?.length > 0 && (sheetchecks == "PolicyReviewChecklist" || sheetchecks == "Forms Compare" )) {
                                    setTimeout(() => {
                                        luckysheet.exitEditMode();
                                        cellDragRestrict(range, selectedTable, tabledata);
                                    }, 100);
                                }
                            } 
                            let copyvalueset = JSON.parse( localStorage.getItem( 'pastevalue' ) );
                            let targetrow = range[ 0 ].row[ 0 ];
                            let targetcolumn = range[ 0 ].column;
                            let selectedIndex = range[ 0 ].row[ 0 ];
                            let tabledata = sheetchecks == "PolicyReviewChecklist" ? tableColumnDetails : sheetchecks == "Forms Compare" ? formTableColumnDetails : sheetchecks == "Exclusion" ? exTableColumnDetails : "";
                            const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
                            const selectedTable = findTableForIndex( selectedIndex, tabledata, excludedColumns );
                            let selectedtabledata = tabledata[ selectedTable ];
                            let endrange = selectedtabledata != undefined ? selectedtabledata.range.end : 0;
                            if ( range != undefined || selectedtabledata != undefined )
                            {
                                if ( copyvalueset != null )
                                {
                                    if ( range !== undefined && targetcolumn !== undefined && targetrow !== undefined )
                                    {
                                        let numRows = range[ 0 ].row[ 1 ] - range[ 0 ].row[ 0 ] + 1; // Number of rows in the range
                                        let numCols = range[ 0 ].column[ 1 ] - range[ 0 ].column[ 0 ] + 1; // Number of columns in the range
                                        if ( numRows !== undefined && numRows !== null && numCols !== undefined && numCols )
                                        {
                                            for ( let i = 0; i < numRows && ( targetrow + i ) <= endrange; i++ )
                                            {
                                                for ( let j = 0; j < numCols; j++ )
                                                {
                                                    let currentRow = targetrow + i;
                                                    let currentCol = targetcolumn[ 0 ] + j;
                                                    let index = i * numCols + j;
                                                    if ( index < copyvalueset.length )
                                                    {
                                                        luckysheet.setcellvalue( currentRow, currentCol, luckysheet.flowdata(), copyvalueset[ index ] );
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    // let length = copyvalueset.length;
                                    // for (let i = 0; i < length && (targetrow + i) <= endrange; i++) {
                                    //     let currentRow = targetrow + i;
                                    //     luckysheet.setcellvalue(currentRow, targetcolumn[0], luckysheet.flowdata(), copyvalueset[i]);
                                    // }
                                    luckysheet.jfrefreshgrid();
                                }

                            }

                            if( sheetchecks == "Exclusion") {
                                let targetRowIndexRange = range[ 0 ].row;
                                let data = luckysheet.getcellvalue();
                                if ( targetRowIndexRange != 0 )
                                    {
                                        for ( let i = targetRowIndexRange[ 0 ]; i <= targetRowIndexRange[ 1 ]; i++ )
                                        {
                                            let rowData = data[ i ];
                                            if ( rowData && rowData.length > 0 )
                                            {
                                                for ( let j = 0; j < rowData.length; j++ )
                                                {
                                                    let object = rowData[ j ];
                                                    if ( object && 'm' in object && 'v' in object && object.m === object.v )
                                                    {
                                                        const restructuredObject = {
                                                            ...object,
                                                            ct: {
                                                                fa: "General",
                                                                t: "inlineStr",
                                                                s: [
                                                                    {
                                                                        ff: "\"Tahoma\"",
                                                                        fc: "#000000",
                                                                        fs: 8,
                                                                        cl: 0,
                                                                        un: 0,
                                                                        bl: 0,
                                                                        it: 0,
                                                                        v: object.v
                                                                    }
                                                                ]
                                                            },
                                                            merge: object.merge || null,
                                                            w: object.w || 55,
                                                            tb: object.tb || "2"
                                                        };
                                                        rowData[ j ] = restructuredObject;
                                                    }
                                                }
                                            }
                                        }
                                    // luckysheet.jfrefreshgrid();
                                    }
                            }
                            //*//
                            const indexData = sheet;
                            if ( indexData && indexData?.length > 0 )
                            {
                                if ( indexData[ 0 ]?.row?.length > 0 && indexData[ 0 ]?.row[ 0 ] == indexData[ 0 ]?.row[ 1 ] )
                                {
                                    setSelectedRowIned( indexData[ 0 ]?.row[ 0 ] );
                                    setHasMultipleRowsSelected( false );
                                } else if ( indexData[ 0 ]?.row?.length > 0 && indexData[ 0 ]?.row[ 0 ] != indexData[ 0 ]?.row[ 1 ] )
                                {
                                    setSelectedRowIndexRange( indexData[ 0 ]?.row )
                                    setHasMultipleRowsSelected( true );
                                } else
                                {
                                    setSelectedRowIned( null );
                                    setHasMultipleRowsSelected( true );
                                    setSelectedRowIndexRange( [] );
                                }
                            }
                            
                             //Handled the functionality for Del key operation for Exclusion Sheet Header Section
                             if(range && range.length > 0 && range != undefined){
                                let selectedRowIndex = range[0].row[0];
                                if(sheetchecks == "Exclusion"){
                                    const tabledata = exTableColumnDetails;
                                    const excludedColumns = ["columnid"];
                                    const selectedTable = findTableForIndex(selectedRowIndex, tabledata, excludedColumns);
                                    const tblSelectedRow = tabledata[selectedTable];
                                    if(tblSelectedRow != undefined) {
                                        let values = tblSelectedRow?.range?.start;
                                        if(index && sheet && sheet.length > 0) {
                                            document.onkeyup = function (e) {
                                                if ( e.which != 40 ) {
                                                    if(e.which == 46 || e.which == 8) {
                                                        if (selectedRowIndex == values) {
                                                            luckysheet.undo()
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (range && range?.length > 0 && sheetchecks == "Exclusion") {
                                setTimeout(() => {
                                    luckysheet.exitEditMode();
                                        cellDragRestrict(range, selectedTable, tabledata);
                                }, 100);
                            }
                        },
                        cellEditBefore( range ) {
                            let flagCheck = luckysheet.getSheet().name;
                            if (range && range.length > 0 && range != undefined) {
                                if (flagCheck === 'PolicyReviewChecklist') {
                                    if (range && range[0]?.row && range[0]?.column && range[0]?.row[0] === range[0]?.row[1] &&
                                        range[0]?.column[0] === range[0]?.column[1] && [ 4, 5, 6, 7].includes(range[0].row[0]) &&
                                        range[0].column[0] === 4) {
                                        setTimeout(() => {
                                            luckysheet.exitEditMode();
                                            matchedOrUnMatchedFilter(range[0]?.row[0]);
                                            container.current.showSnackbar(range[0]?.row[0] === 5 ?
                                                "Matched Records Filtered" : range[0]?.row[0] === 6 ? "Variances Records Filtered" :
                                                    range[0]?.row[0] === 7 ? "Details not available Questions filtered" : "Filter Removed", "info", true);
                                        }, 100);


                                        // checkbox setvalue for Variances columns
                                        if (range && range.length > 0 && range != undefined) {
                                            handleCellSelection(range, flagCheck);
                                        }
                                        return;
                                    }
                                }
                            }
                            if ( flagCheck != 'Exclusion'  && flagCheck != 'QAC not answered questions' )
                            {
                                let selectedRowIndex = range[ 0 ].row[ 0 ];
                                let ranges = luckysheet.getRange();
                                let selectedcolumnindex = ranges[ 0 ].column[ 0 ];
                                let nullcolumncheck = luckysheet.getSheetData()[ selectedRowIndex ];
                                const isAllNull = nullcolumncheck.every( element => element === null );
                                if ( !isAllNull )
                                {
                                    let selectedRowIndex = range[ 0 ].row[ 0 ];
                                    let tabledata = flagCheck == 'PolicyReviewChecklist' ? tableColumnDetails : formTableColumnDetails;
                                    const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
                                    const selectedTable = findTableForIndex( selectedRowIndex, tabledata, excludedColumns );
                                    function findTableForIndex( selectedRowIndex, tableDetails, excludedColumns ) {
                                        for ( const tableName in tableDetails )
                                        {
                                            if ( tableDetails.hasOwnProperty( tableName ) )
                                            {
                                                const range = tableDetails[ tableName ].range;
                                                const columnNames = tableDetails[ tableName ].columnNames;
                                                if ( typeof range.start === 'number' && typeof range.end === 'number' )
                                                {
                                                    if ( selectedRowIndex >= range.start && selectedRowIndex <= range.end )
                                                    {
                                                        if ( columnNames && typeof columnNames === 'object' )
                                                        {
                                                            const validColumns = Object.keys( columnNames ).filter( colName => !excludedColumns.includes( colName ) );
                                                            if ( validColumns.length > 0 )
                                                            {
                                                                return tableName;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        return null;
                                    }

                                    const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
                                    const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
                                    const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore");

                                    const columnData = tabledata[selectedTable]?.columnNames;
                                    if(EnableConfidenceScore == "true" && EnableLockCell == "true" && props?.enableCs && props?.enableCellLock){

                                        if (selectedTable != "Table 1" && selectedTable != "FormTable 1" && selectedTable != "FormTable 2"){
 
                                            if(flagCheck == 'Forms Compare' ? columnData?.PageNumber < selectedcolumnindex : columnData?.PageNumber < selectedcolumnindex && columnData["Actions on Discrepancy"] - 1 > selectedcolumnindex ) {
                                                setTimeout(() => {
                                                    luckysheet.exitEditMode();
                                                }, 100);
                                                return false;
                                            }else{
                                                // for reverting back the data updation for CS functionality by gokul on (feb-11-2025) start**
                                                const document_cols = lockingIndex[selectedTable];
                                                if(Array.isArray(document_cols) && document_cols?.length > 0 && document_cols.includes(selectedcolumnindex)){
                                                    const row_data = luckysheet.getcellvalue(selectedRowIndex);
                                                    const col_key_text = getKeyByValue(columnData, selectedcolumnindex);
                                                    const question = columnData["ChecklistQuestions"] > 0 ? getText( row_data[ columnData["ChecklistQuestions"] ], false ) : "";
                                                    let formheaderdata = flagCheck == "PolicyReviewChecklist" ? props.data?.find((f) => f.Tablename === "JobHeader") : flagCheck == "Forms Compare" ? props.formsCompareHeaderData :[];
                                                    const isStpValid = getConfidenceScoreConfigStatus(formheaderdata?.StpMappings, "question check" , question);
                                                    if(col_key_text && isStpValid){
                                                        const col_cs_key_text = getCsRespectiveColumn(col_key_text);
                                                        const cs_col_value = columnData[col_cs_key_text] > 0 ? getText( row_data[ columnData[col_cs_key_text] ], false ) : "";
                                                        if(cs_col_value?.trim() !== "" && cs_col_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_col_value){
                                                            if(parseFloat(cs_col_value) > parseFloat(MinLockCellScore)){
                                                                setTimeout(() => {
                                                                    luckysheet.exitEditMode();
                                                                }, 100);
                                                                return false;
                                                            }
                                                        }
                                                    }
                                                }
                                                // end**
                                            }
                                        }
                                    }
                                    
                                    let actionColumnTable = tableColumnDetails[ selectedTable ];
                                    let values = actionColumnTable ? Object.values( actionColumnTable.columnNames ) : [];
                                    let largestIndex = Math.max( ...values );
                                    let Actioncolumnindex = largestIndex - 3;
                                    let Requestcolumnindex = largestIndex - 2;
                                    let Notescolumnindex = largestIndex - 1;
                                    if ( actionColumnTable !== undefined )
                                    {
                                        if ( selectedTable != 'Table 3' && selectedRowIndex >= actionColumnTable.range.start + 2 || selectedTable == 'Table 3' && selectedRowIndex >= actionColumnTable.range.start + 3 )
                                        {
                                            if ( ( Actioncolumnindex == selectedcolumnindex || Requestcolumnindex == selectedcolumnindex || Notescolumnindex == selectedcolumnindex ) )
                                            {   //--> for now noneed this
                                                toggleDropDialog()
                                                return false;
                                            }
                                        }
                                    }  else
                                    {
                                        return false;
                                    }
                                }
                            }
                            if (flagCheck == 'Exclusion') {
                                const EnableCsForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableCsForExclusion");
                                const EnableLockCellForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCellForExclusion");
                                const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); 

                                if (range && range.length > 0 && range != undefined) {
                                    const columnData = exTableColumnDetails["ExTable 1"]?.columnNames;
                                    let selectedRowIndex = range[0].row[0];
                                    let selectedcolumnindex = range[0].column[0];
                                    if (EnableCsForExclusion == "true" && EnableLockCellForExclusion == "true") {
                                        if (exclusionApplicableIdx?.includes(selectedcolumnindex) && selectedcolumnindex != columnData?.ConfidenceScore && props?.enableExclusionCellLock ) {
                                            const row_data = luckysheet.getcellvalue(selectedRowIndex);
                                            const cs_score = getText(row_data[columnData["ConfidenceScore"]], false);
                                            if (cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score) {
                                                if (parseFloat(cs_score) > MinLockCellScore) {
                                                    setTimeout(() => {
                                                        luckysheet.exitEditMode();
                                                    }, 100);
                                                    return false;
                                                }
                                            }
                                        } else if (selectedcolumnindex == columnData?.ConfidenceScore) {
                                            setTimeout(() => {
                                                luckysheet.exitEditMode();
                                            }, 100);
                                            return false;
                                        }
                                    }
                                }
                            }
                        },
                        cellAllRenderBefore: function ( data, sheetFile, ctx ) {
                            //In this, we can target the row index object to check if the object has both "(m)" and "(v)", and then verify if the font size (fs) is greater than 10. If it is, we can reduce the object's font size to 8.
                            let range = luckysheet.getRange();
                            let flagCheck = luckysheet?.getSheet()?.name;
                            let tableData = flagCheck === "Forms Compare" ? formTableColumnDetails : exTableColumnDetails;
                            if(flagCheck === "Forms Compare") {
                               var  tableend = tableData["FormTable 3"]?.range.end;
                               var rangestart = range[0].row[0];
                               var rangeend =range[0].row[1];
                               if(rangestart != 0 && rangeend !=0){
                                if (rangestart != rangeend) {
                                    if((rangestart > tableend ||rangeend > tableend) && rangestart != 0 && rangeend != 0) {
                                        //   luckysheet.undo();
                                          setTimeout(() => {
                                            container.current.showSnackbar(
                                                "Data can only be added / edited within the tables.", 
                                                "error", 
                                                true
                                            );
                                            setTimeout(() => {
                                                container.current.hideSnackbar(); 
                                            }, 2000); 
                                        }, 100);
                                    }
                                }
                            
                               }
                                
                            }
                            if(flagCheck == 'Exclusion') {
                              var  tableend = tableData["ExTable 1"]?.range.end;
                              var rangestart = range[0].row[0];
                              var rangeend =range[0].row[1];
                              if(rangestart != rangeend ){
                                if((rangestart > tableend ||rangeend > tableend) && rangestart != 0 && rangeend != 0) {
                                    setTimeout(() => {
                                      container.current.showSnackbar(
                                          "Data can only be added / edited within the tables.", 
                                          "error", 
                                          true
                                      );
                                      setTimeout(() => {
                                          container.current.hideSnackbar(); 
                                      }, 2000); 
                                  }, 200);
                                }
                              }
                             
                            }; 
                            let targetrow = range[ 0 ].row[ 0 ];

                            let sheetrange = range[ 0 ].row;
                            if ( targetrow != 0 )
                            {
                                for ( let i = sheetrange[ 0 ]; i <= sheetrange[ 1 ]; i++ )
                                {
                                    let rowData = data[ i ];
                                    if ( rowData && rowData.length > 2 )
                                    {
                                        for ( let j = 0; j < rowData.length; j++ )
                                        {
                                            let object = rowData[ j ];
                                            if ( object && 'm' in object && 'v' in object && object.m === object.v )
                                            {
                                                if ( object.fs && object.fs > 8 || object.fs < 8 )
                                                {
                                                    object.fs = 8;
                                                }
                                                const restructuredObject = {
                                                    ...object,
                                                    ct: {
                                                        fa: "General",
                                                        t: "inlineStr",
                                                        s: [
                                                            {
                                                                ff: "\"Tahoma\"",
                                                                fc: "#000000",
                                                                fs: 8,
                                                                cl: 0,
                                                                un: 0,
                                                                bl: 0,
                                                                it: 0,
                                                                v: object.v
                                                            }
                                                        ]
                                                    },
                                                    merge: object.merge || null,
                                                    w: object.w || 55,
                                                    tb: object.tb || "2"
                                                };
                                                rowData[ j ] = restructuredObject;
                                            }
                                        }
                                    }
                                }
                            }
                            localStorage.removeItem( 'pastevalue' );
                            //             let flagcheck = luckysheet.getSheet().name
                            //             if( flagcheck != 'Exclusion' || flagcheck != 'QAC not answered questions'){ // this for do not delete the header column names 
                            //            let range = luckysheet.getRange();
                            //            let isinitialrender = range[0].row[1] == 0 ? false : true;
                            //            let selectedRowIndex = range[0].row[0];
                            //            let nullcolumncheck = luckysheet.getSheetData()[selectedRowIndex];
                            //            const isAllNull = nullcolumncheck.every(element => element === null);
                            //          if (isinitialrender == true && !isAllNull ) {

                            //            let selectedIndex = range[0].row[0];
                            //            let sheetname = luckysheet.getSheet().name
                            //            let tabledata = sheetname == "PolicyReviewChecklist" ? tableColumnDetails : formTableColumnDetails  ;
                            //            const excludedColumns = ["Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp"];
                            //            const selectedTable = findTableForIndex(selectedIndex, tabledata, excludedColumns);
                            //            function findTableForIndex(selectedIndex, tableDetails, excludedColumns) {
                            //                for (const tableName in tableDetails) {
                            //                    if (tableDetails.hasOwnProperty(tableName)) {
                            //                        const range = tableDetails[tableName].range;
                            //                        const columnNames = tableDetails[tableName].columnNames;
                            //                        if (typeof range.start === 'number' && typeof range.end === 'number') {
                            //                            if (selectedIndex >= range.start && selectedIndex <= range.end) {
                            //                                if (columnNames && typeof columnNames === 'object') {
                            //                                    const validColumns = Object.keys(columnNames).filter(colName => !excludedColumns.includes(colName));
                            //                                    if (validColumns.length > 0) {
                            //                                        return tableName;
                            //                                    }
                            //                                }
                            //                            }
                            //                        }
                            //                    }
                            //                }
                            //                return null;
                            //            }

                            //    let tabledatas = sheetname == "PolicyReviewChecklist" ? tableColumnDetails : formTableColumnDetails ;
                            //    let selectedtabledata = tabledatas[selectedTable];
                            //    let targetrow1 = selectedtabledata != undefined ? selectedtabledata.range.start : 0;
                            //    let targetrow2 = selectedtabledata != undefined ? selectedtabledata.range.start + 1: 0;
                            //    let currentrow = range[0].row[0];

                            //    if (targetrow1 == currentrow ) {
                            //     let targetcolumn  = range[0].column[1];
                            //     luckysheet.undo([targetrow1,targetcolumn]);
                            //     return true;
                            //    }else if(targetrow2 == currentrow){
                            //     let targetcolumn  = range[0].column[1];
                            //     luckysheet.undo([targetrow2,targetcolumn]);
                            //     return true;
                            //    }
                            //    return true;
                            //         }
                            //     }
                            return true;
                        },
                        cellUpdateBefore: function ( r, c, value, isRefresh ) {
                            // not allowing the edits in the table headers section
                            const sheetDetails = luckysheet.getSheet();
                            let currentActiveSheetTableName = '';

                            if ( sheetDetails?.name === "PolicyReviewChecklist" )
                            {
                                const tblCKeys = Object.keys( tableColumnDetails );
                                tblCKeys.map( ( tblName ) => {
                                    const tblRangeData = tableColumnDetails[ tblName ];
                                    if ( tblRangeData && tblRangeData?.range && tblRangeData?.range?.start && tblRangeData?.range?.end &&
                                        tblRangeData?.range?.start <= r && tblRangeData?.range?.end >= r )
                                    {
                                        currentActiveSheetTableName = tblName;
                                    }
                                } );
                                if ( currentActiveSheetTableName )
                                {
                                    const finalTblDetail = tableColumnDetails[ currentActiveSheetTableName ];
                                    // if ( currentActiveSheetTableName === "Table 1" && c === 1 )
                                    // {
                                    //     return false;
                                    // } else 
                                    if ( finalTblDetail && finalTblDetail?.range && finalTblDetail?.range?.start && finalTblDetail?.range?.end &&
                                        ( finalTblDetail?.range?.start + ( currentActiveSheetTableName === "Table 3" ? 2 : 1 ) ) >= r )
                                    {
                                        return false;
                                    }
                                }
                            } else if ( sheetDetails?.name === "Forms Compare" )
                            {
                                const tblCKeys = Object.keys( formTableColumnDetails );
                                tblCKeys.map( ( tblName ) => {
                                    const tblRangeData = formTableColumnDetails[ tblName ];
                                    if ( tblRangeData && tblRangeData?.range && tblRangeData?.range?.start && tblRangeData?.range?.end &&
                                        tblRangeData?.range?.start <= r && tblRangeData?.range?.end >= r )
                                    {
                                        currentActiveSheetTableName = tblName;
                                    }
                                } );
                                if ( currentActiveSheetTableName )
                                {
                                    const finalTblDetail = formTableColumnDetails[ currentActiveSheetTableName ];
                                    // if ( currentActiveSheetTableName === "FormTable 1" && c === 1 )
                                    // {
                                    //     return false;
                                    // } else 
                                    if ( finalTblDetail && finalTblDetail?.range && finalTblDetail?.range?.start && finalTblDetail?.range?.end &&
                                        ( ( finalTblDetail?.range?.start + 1 ) >= r ) )
                                    {
                                        return false;
                                    }
                                }

                            }
                            return true;
                        },
                        cellUpdated: function ( r, c, oldValue, newValue, isRefresh ) {
                            const sheetName = luckysheet?.getSheet().name;
                            let rowIdx = r - 1;
                            let tableData = sheetName === "PolicyReviewChecklist" ? tableColumnDetails : sheetName === "Forms Compare" ? formTableColumnDetails : exTableColumnDetails;

                         if(r != undefined && c != undefined) {
                            if(sheetName === "PolicyReviewChecklist" || sheetName == "Forms Compare") {
                                const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
                                const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
                                const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore");

                                if(EnableConfidenceScore == "true" && EnableLockCell == "true" && props?.enableCs && props?.enableCellLock){
                                    const findTbl = findTableForIndex(r, tableData, "");
                                    if (findTbl != "Table 1" && findTbl != "FormTable 1" && findTbl != "FormTable 2"){

                                        const columnData = tableData[findTbl]?.columnNames;
                                        if(sheetName == "Forms Compare" ?columnData?.PageNumber < c : columnData?.PageNumber < c && columnData["Actions on Discrepancy"] - 1 > c ) {
                                            setCellValue(r, c, oldValue);
                                            return false;
                                        }else{
                                            // for reverting back the data updation for CS functionality by gokul on (feb-11-2025) start**
                                            const document_cols = lockingIndex[findTbl];
                                            if(Array.isArray(document_cols) && document_cols?.length > 0 && document_cols.includes(c)){
                                                const row_data = luckysheet.getcellvalue(r);
                                                const col_key_text = getKeyByValue(columnData, c);
                                                const question = columnData["ChecklistQuestions"] > 0 ? getText( row_data[ columnData["ChecklistQuestions"] ], false ) : "";
                                                let datas = sheetName == "PolicyReviewChecklist" ? props.data?.find((f) => f.Tablename === "JobHeader") : sheetName == "Forms Compare" ? props.formsCompareHeaderData :[];
                                                const isStpValid = getConfidenceScoreConfigStatus(datas?.StpMappings, "question check" , question );
                                                if(col_key_text && isStpValid){
                                                    const col_cs_key_text = getCsRespectiveColumn(col_key_text);
                                                    const cs_col_value = columnData[col_cs_key_text] > 0 ? getText( row_data[ columnData[col_cs_key_text] ], false ) : "";
                                                    if(cs_col_value?.trim() !== "" && cs_col_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_col_value){
                                                        if(parseFloat(cs_col_value) > parseFloat(MinLockCellScore)){
                                                            setCellValue(r, c, oldValue);
                                                            // return false; // if want to break here check for observation and page number auto population
                                                        }
                                                    }
                                                }
                                            }
                                            // end**
                                        }
                                    }
                                }
                            }

                             if (sheetName === "Exclusion") {
                                 const EnableCsForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableCsForExclusion");
                                 const EnableLockCellForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCellForExclusion");
                                 const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); 

                                 if (r != undefined && c != undefined) {
                                     const columnData = exTableColumnDetails["ExTable 1"]?.columnNames;
                                     if (EnableCsForExclusion == "true" && EnableLockCellForExclusion == "true") {
                                         if (exclusionApplicableIdx?.includes(c) && c != columnData?.ConfidenceScore && props?.enableExclusionCellLock) {

                                             const row_data = luckysheet.getcellvalue(r);
                                            const cs_score = getText(row_data[columnData["ConfidenceScore"]], false);
                                            if (cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score) {
                                                if (parseFloat(cs_score) > MinLockCellScore) {
                                                    setCellValue(r, c, oldValue);
                                                    return false;
                                                }
                                            }
                                         } else if (c == columnData?.ConfidenceScore) {
                                             setCellValue(r, c, oldValue);
                                             return false;
                                         }
                                     }
                                }
                            }
                            
                            if(sheetName === "Forms Compare") {
                                if(rowIdx >= tableData["FormTable 3"]?.range.end) {
                                    //   luckysheet.undo();
                                      setTimeout(() => {
                                        container.current.showSnackbar(
                                            "Data can only be added / edited within the tables.", 
                                            "error", 
                                            true
                                        );
                                        setTimeout(() => {
                                            container.current.hideSnackbar(); 
                                        }, 2000); 
                                    }, 100);
                                }
                            }

                            if(sheetName == 'Exclusion') {
                                
                                if(rowIdx >= tableData["ExTable 1"]?.range.end){
                                    setTimeout(() => {
                                      container.current.showSnackbar(
                                          "Data can only be added / edited within the tables.", 
                                          "error", 
                                          true
                                      );
                                      setTimeout(() => {
                                          container.current.hideSnackbar(); 
                                      }, 2000); 
                                  }, 200);
                                }
                            }; 
                        }
                            let rowData = luckysheet.getcellvalue( r );
                            rowData = rowData.filter( ( f ) => f != null );
                            let length = [];
                            let maxLength = 0;
                            rowData.forEach( ( f ) => {
                                if ( f?.ct?.s )
                                {
                                    if ( f?.ct?.s?.length > 1 )
                                    {
                                        var text = '';
                                        f?.ct?.s?.forEach( ( e ) => { text += e?.v } )
                                        length.push( text?.length );
                                    } else { length.push( f?.ct?.s[ 0 ]?.v?.length ) }
                                }
                            } );
                            length = Array.from( new Set( length ) );
                            length.forEach( ( f ) => {
                                if ( f > maxLength )
                                {
                                    maxLength = f;
                                }
                            } );
                            let config = luckysheet.getConfig();
                            config.rowlen[ r ] = maxLength && maxLength > 15 ? maxLength / 2 + 30 : 30;


                            luckysheet.setConfig( config );
                            const tabledats = tableColumnDetails;
                            const formtabledats = formTableColumnDetails;

                            if ( ( r > tabledats[ 'Table 1' ].range.start && r <= tabledats[ 'Table 1' ].range.end && c == 1 ) ||
                                ( r > formtabledats[ 'FormTable 1' ].range.start && r < formtabledats[ 'FormTable 1' ].range.end && c == 1 ) )
                            {
                                if (
                                    oldValue && oldValue.ct && oldValue.ct.s && oldValue.ct.s[ 0 ] && newValue &&
                                    newValue.ct && newValue.ct.s && newValue.ct.s[ 0 ] )
                                {
                                    if ( oldValue.m == newValue.ct.s[ 0 ].v )
                                    {
                                        return false;
                                    } else if ( oldValue.ct.s[ 0 ].v == newValue.ct.s[ 0 ].v )
                                    {
                                        return false;
                                    } else
                                    {
                                        const currentSheetDetails = luckysheet.getSheet();
                                        if ( currentSheetDetails?.name != 'Exclusion' )
                                        {
                                            luckysheet.setCellValue( r, c, oldValue );
                                        }
                                    }
                                }
                            }      ////Table 1 cell disable  
                            if(r != undefined && c!= undefined) {   // Handled the scroll issue on edit for the cells
                                let targetColumnIndex = c;
                                let sheetCheck = luckysheet?.getSheet();
                                let zoomCheck = sheetCheck?.zoomRatio;
                                if(zoomCheck != 1) {
                                    let rowDataa = luckysheet.getcellvalue( r ); 
                                    let filteredRow = rowDataa[targetColumnIndex];
                                    if(filteredRow && filteredRow != undefined) {
                                        let valueAtV = filteredRow?.v || filteredRow?.ct?.s[0]?.v;
                                        let transformedData = {
                                            "ct": {
                                                "fa": "@",
                                                "t": "inlineStr",
                                                "s": [
                                                    {
                                                        "v": valueAtV,
                                                        "ff": filteredRow?.ff,
                                                        "fs": 7 || filteredRow?.ct?.s[0]?.fs
                                                    }
                                                ]
                                            },
                                            "m": valueAtV,
                                            "v": valueAtV,
                                            "ff": `"${filteredRow?.ff}"`,
                                            "bg": "rgb(139,173,212)",
                                            "tb": filteredRow?.tb,
                                            "w": filteredRow?.w,
                                            "row": r,
                                            "column": c
                                        };
                                        luckysheet.scroll({
                                            targetRow: transformedData?.row - 1,
                                            targetColumn: 0
                                        });
                                    }
                                }
                            }
                            if ( sheetName != 'Exclusion' ) {
                                autoUpdateCtPt( r, c <= 2 ? 4 : c, newValue );
                             }
                        },
                        //nirshee
                        rangePasteBefore: function ( range, data ) {
                            let selectedIndex = range[ 0 ].row[ 0 ];
                            let selectedColIndex = range[ 0 ].column[ 0 ];
                            let selectedEndColIndex = range[ 0 ].column[ 1 ];
                             let flagCheck = luckysheet?.getSheet()?.name;
                            if ( flagCheck !== 'Exclusion' )
                            {
                                let tabledata = flagCheck == 'PolicyReviewChecklist' ? tableColumnDetails : formTableColumnDetails;
                                const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
                                const selectedTable = findTableForIndex( selectedIndex, tabledata, excludedColumns );
                               
                                const headercolval1 = tabledata[ selectedTable ]?.range?.start;
                                const headercolval2 = headercolval1 + 1;
                                let isHeader = false;

                                range.forEach( item => {
                                    const targetRow = item.row[ 0 ];
                                    if ( targetRow === headercolval1 || targetRow === headercolval2 )
                                    {
                                        isHeader = true;
                                    }
                                } );

                                const columnData = tabledata[selectedTable]?.columnNames;
                                if (columnData == null || columnData == undefined) {
                                    return false
                                }
                                if((flagCheck == 'PolicyReviewChecklist' || flagCheck == 'Forms Compare') && Object.keys(columnData).length > 0) {
                                    const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
                                    const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
                                    const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore");
                                    // cs source document prevent paste by gokul** on (feb-11-2025) start**
                                    if ((flagCheck == 'PolicyReviewChecklist' || flagCheck == 'Forms Compare') && (selectedTable != "Table 1" && selectedTable != "FormTable 1" && selectedTable != "FormTable 2") && EnableConfidenceScore == "true" && EnableLockCell == "true" && props?.enableCs) {
                                        if (flagCheck == 'PolicyReviewChecklist' || flagCheck == 'Forms Compare') {
                                            if((flagCheck == 'Forms Compare' ? columnData?.PageNumber < selectedColIndex : columnData?.PageNumber < selectedColIndex && columnData["Actions on Discrepancy"] - 1 > selectedColIndex ) || 
                                            flagCheck == 'Forms Compare' ? columnData?.PageNumber < selectedEndColIndex : columnData?.PageNumber < selectedEndColIndex && columnData["Actions on Discrepancy"] - 1 > selectedEndColIndex ) {
                                            return false;
                                        }
                                        }
                                       
                                        const document_cols = lockingIndex[selectedTable];
                                        if(Array.isArray(document_cols) && document_cols?.length > 0 && (document_cols.includes(selectedColIndex) || document_cols.includes(selectedEndColIndex))){
                                            let return_value = true;
                                            for (let index = range[ 0 ].row[ 0 ]; index <= range[ 0 ].row[ 1 ]; index++) {
                                                const row_data = luckysheet.getcellvalue(index);
                                                
                                                const question = columnData["ChecklistQuestions"] > 0 ? getText( row_data[ columnData["ChecklistQuestions"] ], false ) : "";
                                                let datas = flagCheck == "PolicyReviewChecklist" ? props.data?.find((f) => f.Tablename === "JobHeader") : flagCheck == "Forms Compare" ? props.formsCompareHeaderData :[];
                                                const isStpValid = getConfidenceScoreConfigStatus(datas?.StpMappings, "question check" , question );

                                                if(isStpValid){
                                                    
                                                    if(document_cols.includes(selectedColIndex)){
                                                        const col_key_text = getKeyByValue(columnData, selectedColIndex);
                                                        if(col_key_text){
                                                            const col_cs_key_text = getCsRespectiveColumn(col_key_text);
                                                            const cs_col_value = columnData[col_cs_key_text] > 0 ? getText( row_data[ columnData[col_cs_key_text] ], false ) : "";
                                                            if(cs_col_value?.trim() !== "" && cs_col_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_col_value){
                                                                if(parseFloat(cs_col_value) > parseFloat(MinLockCellScore)){
                                                                    return_value = false;
                                                                    return false; //if want to tell the use which are the restricted column remove this break and track the col and row positions
                                                                }
                                                            }
                                                        }
                                                    }
                                                    
                                                    if(selectedEndColIndex != selectedColIndex && document_cols.includes(selectedEndColIndex)){
                                                        const col_key_text = getKeyByValue(columnData, selectedEndColIndex);
                                                        if(col_key_text){
                                                            const col_cs_key_text = getCsRespectiveColumn(col_key_text);
                                                            const cs_col_value = columnData[col_cs_key_text] > 0 ? getText( row_data[ columnData[col_cs_key_text] ], false ) : "";
                                                            if(cs_col_value?.trim() !== "" && cs_col_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_col_value){
                                                                if(parseFloat(cs_col_value) > parseFloat(MinLockCellScore)){
                                                                    return_value = false;
                                                                    return false; //if want to tell the use which are the restricted column remove this break and track the col and row positions
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            if(!return_value){
                                                return false;
                                            }
                                        }
                                    }
                                    // end**
                                }

                                if ( isHeader == true )
                                {
                                    const msg = "Cannot Paste content in the header sections.";
                                    setMsgVisible( true );
                                    setMsgClass( 'alert error' );
                                    setMsgText( msg );
                                    setTimeout( () => {
                                        setMsgVisible( false );
                                        setMsgText( '' );
                                    }, 3500 );
                                    return false;
                                }
                                else
                                {
                                    /*/if there is not a headervalue below the part is running /*/
                                    // This block is for getting the copied content in 'data'. We can remove the '.' in <td> and store the data in localStorage. 
                                    // The remaining process will be done in the range select HOOK to set the cell value.
                                    const parser = new DOMParser();
                                    const htmlDoc = parser.parseFromString( data, 'text/html' );
                                    var Htmlcollections = htmlDoc.all;
                                    var tdArray = Array.from( Htmlcollections );
                                    var tdElements = tdArray.filter( element => element.tagName.toLowerCase() === "td" );

                                    const tdValues = [];
                                    const tdOriginalValues = [];
                                    // Iterate over each <td> element and push its text content into the array
                                    tdElements.forEach( tdElement => {
                                        const tdValue = tdElement.innerText.trim().replace( /\./g, '•' );
                                        tdValues.push( tdValue );
                                        tdOriginalValues.push( tdElement.innerText.trim() );
                                    } );

                                    function containsCurrency(strings) {
                                        // Regex to match numbers with decimals (e.g., 1234.56)
                                        const hasDecimalRegex = /\b\d+\.\d+\b/g;
                                        
                                        // Regex to match numbers with optional thousands separators and a dot with optional spaces before the dot
                                        const endsWithDotRegex = /\b\d+(?:,\d{3})*(?:\.\d+)?\s*\.\s*.*$/g;
                                    
                                        return strings.some(str => {
                                            
                                            // Check if the string contains numbers with decimals
                                            const hasDecimal = hasDecimalRegex.test(str);
                                            
                                            // Check if the string contains numbers that end with a dot, allowing for spaces before the dot
                                            const endsWithDot = endsWithDotRegex.test(str);
                                            
                                            return hasDecimal || endsWithDot;
                                        });
                                    }

                                    const hasDot = containsCurrency(tdOriginalValues);
                                    if(hasDot){
                                        const tdValuesJSON = JSON.stringify( tdValues );
                                        localStorage.setItem( 'pastevalue', tdValuesJSON );
                                        //*//
                                        let config = luckysheet.getConfig();
                                        setTimeout( () => {
                                            if ( flagCheck == "PolicyReviewChecklist" )
                                                {
                                                    const tabledats = tableColumnDetails;
                                            let tableKeysToRemove = [];
                                            for ( const tableName in tabledats )
                                            {
                                                if ( Object.keys( tabledats[ tableName ].columnNames ).length === 0 )
                                                {
                                                    tableKeysToRemove.push( tableName );
                                                }
                                            }
                                            tableKeysToRemove.forEach( tableName => {
                                                delete tabledats[ tableName ];
                                            } );

                                            const lastTableName = Object.keys( tabledats ).pop();

                                            let rowlen = range[ 0 ].row[ 0 ];
                                            const rangeValue = rowlen;
                                            let matchedTableName = null;
                                            let endValue = null;
                                            for ( const [ tableName, tableInfo ] of Object.entries( tabledats ) )
                                            {
                                                const { start, end } = tableInfo.range;
                                                if ( rangeValue >= start && rangeValue <= end )
                                                {
                                                    matchedTableName = tableName;
                                                    endValue = end;
                                                    break;
                                                }
                                            }
                                            const columnvalue = range[ 0 ].column_focus;
                                            const index1 = endValue + 1; // Dynamically set index1
                                            const index2 = endValue + 2; // Dynamically set index2
                                            sessionStorage.setItem( 'index1', index1 );
                                            sessionStorage.setItem( 'index2', index2 );
                                            const luckySheet = luckysheet.getSheetData()[ 1 ];
                                            let flagCheck = luckySheet[ 1 ].m;
                                            if ( flagCheck != 'FORM COMPARE' )
                                            {//blink fix
                                                if ( lastTableName != matchedTableName )
                                                {
                                                    luckysheet.clearCell( index1, columnvalue );
                                                    luckysheet.clearCell( index2, columnvalue );
                                                }
                                            }

                                            const multiplerowrange = luckysheet.getRange();
                                            range.forEach( ( item, index ) => {
                                                const multiRow = multiplerowrange[ index ]?.row; // Get the row values from multiplerowrange
                                                if ( multiRow && multiRow.length === 2 )
                                                {
                                                    range[ index ].row = multiRow; // Update row values in range
                                                    range[ index ].row_focus = multiRow[ 0 ]; // Update row_focus value as well
                                                }
                                            } );
                                            range.forEach( item => {
                                                if ( item.row[ 1 ] > endValue )
                                                {
                                                    item.row[ 1 ] = endValue;
                                                }
                                            } );

                                            if ( range && range?.length > 0 && range[ 0 ]?.row?.length > 0 )
                                            {
                                                let rowRangeLength = 0;
                                                rowRangeLength = range[ 0 ]?.row[ 1 ] - range[ 0 ]?.row[ 0 ];
                                                const rowStartIndex = range[ 0 ]?.row[ 0 ];
                                                const tableColumnDetail = Object.keys( tableColumnDetails[ 'Table 3' ].columnNames )[ '3' ];
                                                let dynamicColumnValue = tableColumnDetail == 'Lob' ? 4 : 3;
                                                if ( rowRangeLength >= 0 )
                                                {
                                                    for ( let index = 0; index <= rowRangeLength; index++ )
                                                    {
                                                        const row = rowStartIndex + index;
                                                        config.rowlen[ row ] = config.rowlen[ row ] > 50 ? config.rowlen[ row ] : 50;
                                                        autoUpdateCtPt( row, dynamicColumnValue, "value" );
                                                    }
                                                    luckysheet.setConfig( config );
                                                }
                                            }
                                        }
                                        else if ( flagCheck == "Forms Compare" )
                                        {
                                            const tabledats = formTableColumnDetails;
                                            let tableKeysToRemove = [];
                                            for ( const tableName in tabledats )
                                            {
                                                if ( Object.keys( tabledats[ tableName ].columnNames ).length === 0 )
                                                {
                                                    tableKeysToRemove.push( tableName );
                                                }
                                            }
                                            tableKeysToRemove.forEach( tableName => {
                                                delete tabledats[ tableName ];
                                            } );

                                            const lastTableName = Object.keys( tabledats ).pop();

                                            let rowlen = range[ 0 ].row[ 0 ];
                                            const rangeValue = rowlen;
                                            let matchedTableName = null;
                                            let endValue = null;
                                            for ( const [ tableName, tableInfo ] of Object.entries( tabledats ) )
                                            {
                                                const { start, end } = tableInfo.range;
                                                if ( rangeValue >= start && rangeValue <= end )
                                                {
                                                    matchedTableName = tableName;
                                                    endValue = end;
                                                    break;
                                                }
                                            }
                                            const columnvalue = range[ 0 ].column_focus;
                                            const index1 = endValue + 1;
                                            const index2 = endValue + 2;
                                            sessionStorage.setItem( 'index1', index1 );
                                            sessionStorage.setItem( 'index2', index2 );
                                            const luckySheet = luckysheet.getSheetData()[ 1 ];
                                            let flagCheck = luckySheet[ 1 ].m;
                                            if ( flagCheck != 'POLICY REVIEW CHECKLIST' )
                                            {
                                                if ( lastTableName != matchedTableName )
                                                {
                                                    luckysheet.clearCell( index1, columnvalue );
                                                    luckysheet.clearCell( index2, columnvalue );
                                                }
                                            }

                                            const multiplerowrange = luckysheet.getRange();
                                            range.forEach( ( item, index ) => {
                                                const multiRow = multiplerowrange[ index ]?.row;
                                                if ( multiRow && multiRow.length === 2 )
                                                {
                                                    range[ index ].row = multiRow;
                                                    range[ index ].row_focus = multiRow[ 0 ];
                                                }
                                            } );
                                            range.forEach( item => {
                                                if ( item.row[ 1 ] > endValue )
                                                {
                                                    item.row[ 1 ] = endValue;
                                                }
                                            } );

                                            if ( range && range?.length > 0 && range[ 0 ]?.row?.length > 0 )
                                            {
                                                let rowRangeLength = 0;
                                                rowRangeLength = range[ 0 ]?.row[ 1 ] - range[ 0 ]?.row[ 0 ];
                                                const rowStartIndex = range[ 0 ]?.row[ 0 ];

                                                if ( rowRangeLength >= 0 )
                                                {
                                                    for ( let index = 0; index <= rowRangeLength; index++ )
                                                    {
                                                        const row = rowStartIndex + index;
                                                        config.rowlen[ row ] = config.rowlen[ row ] > 50 ? config.rowlen[ row ] : 50;
                                                        autoUpdateCtPt( row, 3, "value" );
                                                    }
                                                    luckysheet.setConfig( config );
                                                }
                                            }
                                        }
                                        }, 100 );
                                    }else{
                                        range.forEach((r) => {
                                            for (let index = 0; index <= tdOriginalValues?.length; index++) {
                                                setTimeout(() => {
                                                    autoUpdateCtPt( (r.row[0] + index), r.column[0], "value" );
                                                }, 100);                                                
                                            }
                                        });
                                    }
                                    return true;

                                }
                            } else if( flagCheck == 'Exclusion' ) {
                                const EnableCsForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableCsForExclusion");
                                const EnableLockCellForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCellForExclusion");
                                const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); 
                                if (range && range.length > 0 && range != undefined) {
                                    const columnData = exTableColumnDetails["ExTable 1"]?.columnNames;
                                    let selectedRowIndex = range[0].row[0];
                                    if (EnableCsForExclusion == "true" && EnableLockCellForExclusion == "true") {
                                        if (exclusionApplicableIdx?.includes(selectedColIndex) && selectedColIndex != columnData?.ConfidenceScore && props?.enableExclusionCellLock) {

                                            const row_data = luckysheet.getcellvalue(selectedRowIndex);
                                            const cs_score = getText(row_data[columnData["ConfidenceScore"]], false);
                                            if (cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score) {
                                                if (parseFloat(cs_score) > MinLockCellScore) {
                                                    return false;
                                                }
                                            }
                                        } else if (selectedColIndex == columnData?.ConfidenceScore) {
                                            return false;
                                        }
                                    }
                                }

                                let tabledata = exTableColumnDetails;
                                const selectedTable = "ExTable 1";
                                const exclusionHeader = tabledata[ selectedTable ]?.range?.start;
                                let isHeader = false;

                                range.forEach( item => {
                                    const targetRow = item.row[ 0 ];
                                    if ( targetRow === exclusionHeader )
                                    {
                                        isHeader = true;
                                    }
                                } );

                                if ( isHeader == true )
                                {
                                    const msg = "Cannot Paste content in the header sections.";
                                    setMsgVisible( true );
                                    setMsgClass( 'alert error' );
                                    setMsgText( msg );
                                    setTimeout( () => {
                                        setMsgVisible( false );
                                        setMsgText( '' );
                                    }, 3500 );
                                    return false;
                                }
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
                        textColor: false, //'Text color'
                        findAndReplace: false, //'Find and Replace'
                    }
                };
                luckysheet.create( options );
            }
        }
    }

    $(document).ready(function () {       // right-click event trigger for confident score applied document cell values
        let flagCheck = props.selectedSheet;
        const range = luckysheet.getRange();
        const rowRange = range[0].row[0];
        const columnRange = range[ 0 ].column[ 0 ];
        let tabledata =  flagCheck == 'PolicyReviewChecklist' ? tableColumnDetails : formTableColumnDetails;
        const selectedTable = findTableForIndex( rowRange, tabledata, "" );
        const columnData = tabledata[selectedTable]?.columnNames;
        const checkRange = tabledata[selectedTable]?.range;
        if(checkRange != undefined) {
            if(rowRange > checkRange?.start && rowRange <= checkRange?.end) {

                $('body').append(`
                    <ul id="customDropdown" style="display: none; position: absolute; background: white; border: 1px solid #ccc; padding: 5px 0; list-style: none; min-width: 160px; box-shadow: 2px 2px 5px rgba(0, 0, 0, 0.2); border-radius: 3px; font-family: Arial, sans-serif; font-size: 14px;">
                        <li class="dropdown-item" data-value="MATCHED" style="padding: 8px 15px; cursor: pointer; white-space: nowrap;">MATCHED</li>
                        <li class="dropdown-item" data-value="Details not available in the document" style="padding: 8px 15px; cursor: pointer; white-space: nowrap;">Details not available in the document</li>
                    </ul>
                `);

                //----------------***-------------------//
                $("#luckysheet").off("mousedown");              
                $(document).off("click", ".dropdown-item");            // Unbind all existing events first before adding new ones.
                $(document).off("click");
                //----------------***-------------------//

                const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
                const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
                const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore");

                if(EnableConfidenceScore == "true" && EnableLockCell == "true" && props?.enableCs){
                    if(selectedTable != "Table 1" || selectedTable != "FormTable 1") {
                        const document_cols = lockingIndex[selectedTable];
                        if (Array.isArray(document_cols) && document_cols?.length > 0 && document_cols.includes(columnRange)) {
                            const row_data = luckysheet.getcellvalue(rowRange);
                            const col_key_text = getKeyByValue(columnData, columnRange);
                            const question = columnData["ChecklistQuestions"] > 0 ? getText(row_data[columnData["ChecklistQuestions"]], false) : "";
                            let formheaderdata = flagCheck == "PolicyReviewChecklist" ? props.data?.find((f) => f.Tablename === "JobHeader") : flagCheck == "Forms Compare" ? props.formsCompareHeaderData :[];
                            const isStpValid = getConfidenceScoreConfigStatus(formheaderdata?. StpMappings, "question check", question);
                            if (col_key_text && isStpValid) {
                                const col_cs_key_text = getCsRespectiveColumn(col_key_text);
                                const cs_col_value = columnData[col_cs_key_text] > 0 ? getText(row_data[columnData[col_cs_key_text]], false) : "";
                                    if ( cs_col_value?.trim() !== "" && cs_col_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_col_value && parseFloat(cs_col_value) > parseFloat(MinLockCellScore)) {
                                        $("#luckysheet").off("mousedown").on("mousedown", function (event) {
                                            let rang = luckysheet.getRange();
                                            if (event.which === 3 && document_cols?.includes( rang[ 0 ].column[ 0 ]) && (rang[ 0 ].row[ 0 ] > checkRange?.start && rang[ 0 ].row[ 0 ] <= checkRange?.end)) { // Right click detected
                                                event.preventDefault(); 
                                                    $("#customDropdown") 
                                                        .css({
                                                            top: event.pageY + "px",
                                                            left: event.pageX + "px",
                                                            display: "block"
                                                        })
                                            } else {
                                                $("#customDropdown").hide();
                                            }
                                        });

                                        //Handle the logic for dropdown selection
                                        $(document).off("click", ".dropdown-item").on("click", ".dropdown-item", function () {
                                            const selectedOption = $(this).data("value");
                                            if(selectedOption) {
                                                const sheetValueStructure = getEmptyDataSet();
                                                if(selectedOption === "MATCHED") {
                                                    for(const item of sheetValueStructure["ct"]["s"]) {
                                                        item["v"] = selectedOption
                                                        item["ff"] = "\"Tahoma\""
                                                        item["bl"] = 1
                                                        item["fc"] = "rgb(0, 128, 0)"
                                                    }
                                                    setCellValue(rowRange, columnRange,sheetValueStructure);
                                                } 
                                                if(selectedOption === "Details not available in the document") {
                                                    for(const item of sheetValueStructure["ct"]["s"]) {
                                                        item["v"] = selectedOption
                                                        item["ff"] = "\"Tahoma\""
                                                    }
                                                    setCellValue(rowRange, columnRange,sheetValueStructure);
                                                }
                                            }
                                            $("#customDropdown").hide();    // Hide dropdown after selection
                                        });
                                } else {
                                    $("#luckysheet").off("mousedown").on("mousedown", function (event) {
                                        if (event.which !== 3) { // Ignore right-clicks
                                            $("#customDropdown").hide();
                                        }
                                    });
                                }
                                //Hide dropdown if it is not a configure score valued cell
                                $(document).on("click", function (event) {
                                    if (!$(event.target).closest("#customDropdown").length) {
                                        $("#customDropdown").hide();
                                    }
                                });
                            }
                        }
                    }
                }
            }
        } 
    });

    const delKeyRestrictCSColum = (selectedTable, selectedRowIndex, selectedColIndex, tabledata) => {                // restrict the del key functionlity for the cs score applied cells
        const columnData = tabledata[selectedTable]?.columnNames;
        const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
        const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
        const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore");
        let flagCheck = luckysheet.getSheet()?.name;
        if (flagCheck != 'Exclusion') {
            if(columnData && Object.keys(columnData).length > 0) {
                // document.onkeyup = function (e) {
                //     if ( e.which != 40 ) {
                //         if(e.which == 46 || e.which == 8) {
                            if(EnableConfidenceScore == "true" && EnableLockCell == "true" && props?.enableCs){
                                if (selectedTable != "Table 1" && selectedTable != "FormTable 1" && selectedTable != "FormTable 2"){
                                    if(flagCheck == "PolicyReviewChecklist" ? columnData?.PageNumber < selectedColIndex && columnData["Actions on Discrepancy"] - 1 > selectedColIndex : columnData?.PageNumber < selectedColIndex  ) {
                                        luckysheet.undo()
                                    }
                                    else {
                                        const document_cols = lockingIndex[selectedTable];
                                        if(Array.isArray(document_cols) && document_cols?.length > 0 && document_cols.includes(selectedColIndex)){
                                            const row_data = luckysheet.getcellvalue(selectedRowIndex);
                                            const col_key_text = getKeyByValue(columnData, selectedColIndex);
                                            const question = columnData["ChecklistQuestions"] > 0 ? getText( row_data[ columnData["ChecklistQuestions"] ], false ) : "";
                                            let formheaderdata = flagCheck =="PolicyReviewChecklist" ? props.data?.find((f) => f.Tablename === "JobHeader") : flagCheck == "Forms Compare" ? props.formsCompareHeaderData :[];
                                            const isStpValid = getConfidenceScoreConfigStatus(formheaderdata?.StpMappings, "question check" , question );
                                            if(col_key_text && isStpValid){
                                                const col_cs_key_text = getCsRespectiveColumn(col_key_text);
                                                const cs_col_value = columnData[col_cs_key_text] > 0 ? getText( row_data[ columnData[col_cs_key_text] ], false ) : "";
                                                if(cs_col_value?.trim() !== "" && cs_col_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_col_value){
                                                    if(parseFloat(cs_col_value) > parseFloat(MinLockCellScore)){
                                                        // setTimeout(() => {
                                                            luckysheet.undo();
                                                        // }, 100);
                                                        return;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } 
                            }
                //         }
                //     }
                // }
                return false;
            }
        } else if (flagCheck == 'Exclusion') {
            const columnData = exTableColumnDetails["ExTable 1"]?.columnNames;
            const EnableCsForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableCsForExclusion");
            const EnableLockCellForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCellForExclusion");
            const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); 

            if (EnableCsForExclusion == "true" && EnableLockCellForExclusion == "true") {
                if (exclusionApplicableIdx?.includes(selectedColIndex) && selectedColIndex != columnData?.ConfidenceScore && props?.enableExclusionCellLock) {
                        const row_data = luckysheet.getcellvalue(selectedRowIndex);
                        const cs_score = getText(row_data[columnData["ConfidenceScore"]], false);
                        if (cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score) {
                            if (parseFloat(cs_score) > MinLockCellScore) {
                                luckysheet.undo()
                            }
                        }
                    } else if (selectedColIndex == columnData?.ConfidenceScore) {
                        luckysheet.undo()
                        return false;
                    }
            }
        }
    };

    const cellDragRestrict = (range, selectedTable, tabledata) => {            // restrict the drag function on cells which has cs scored
        const hasCangeDetected = sessionStorage.getItem("cs_range_select");
        let flagCheck = luckysheet.getSheet()?.name;
        if (flagCheck != 'Exclusion') {
            if(hasCangeDetected && hasCangeDetected === "true"){
                range?.forEach((r, index) => {
                    if(r?.row[0] === r?.row[1] && r?.column[0] === r?.column[1]){
                        return;
                    }
                    const columnData = tabledata[selectedTable]?.columnNames;
                    const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
                    const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
                    const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore");
    
                    const document_cols = lockingIndex[selectedTable];
                    if (selectedTable != "Table 1" && selectedTable != "FormTable 1" && selectedTable != "FormTable 2" && (r?.column[0] != 0)){
                        if(EnableConfidenceScore == "true" && EnableLockCell == "true" && props?.enableCs && props?.enableCellLock){
                            for (let row = r?.row[0]; row <= r?.row[1]; row++) {
                                for (let column = r.column[0]; column <= r.column[1]; column++) {
                                    // if(Array.isArray(document_cols) && document_cols?.length > 0 && document_cols.includes(column)){
                                        if(flagCheck == "PolicyReviewChecklist" ? columnData?.PageNumber < column && columnData["Actions on Discrepancy"] - 1 > column : columnData?.PageNumber < column ) {
                                            setTimeout(() => {
                                                luckysheet.undo();
                                            }, 100);
                                            return;
                                        }
                                    // }
                                    else {
                                        // for reverting back the data updation for CS functionality by gokul on (feb-11-2025) start**
                                        // const document_cols = lockingIndex[selectedTable];
                                        if(Array.isArray(document_cols) && document_cols?.length > 0 && document_cols.includes(column)){
                                            const row_data = luckysheet.getcellvalue(row);
                                            const col_key_text = getKeyByValue(columnData, column);
                                            const question = columnData["ChecklistQuestions"] > 0 ? getText( row_data[ columnData["ChecklistQuestions"] ], false ) : "";
                                            let formheaderdata = flagCheck == "PolicyReviewChecklist" ? props.data?.find((f) => f.Tablename === "JobHeader") : flagCheck == "Forms Compare" ? props.formsCompareHeaderData :[];
                                            const isStpValid = getConfidenceScoreConfigStatus(formheaderdata?.StpMappings, "question check" ,question);
                                            if(col_key_text && isStpValid){
                                                const col_cs_key_text = getCsRespectiveColumn(col_key_text);
                                                const cs_col_value = columnData[col_cs_key_text] > 0 ? getText( row_data[ columnData[col_cs_key_text] ], false ) : "";
                                                if(cs_col_value?.trim() !== "" && cs_col_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_col_value){
                                                    if(parseFloat(cs_col_value) > parseFloat(MinLockCellScore)){
                                                        const hasCangeDetected = sessionStorage.getItem("cs_range_select");
                                                        // if(hasCangeDetected && hasCangeDetected === "true"){
                                                            setTimeout(() => {
                                                                luckysheet.undo();
                                                                sessionStorage.setItem("cs_range_select","false");
                                                            }, 100);
                                                            return;
                                                        // }
                                                    }
                                                }
                                            }
                                        }
                                        // end**
                                    }
                                }
                                // For auto population
                                // if(r?.row[0] != r?.row[1] || r.column[0] != r.column[1]){
                                //     autoUpdateCtPt(row,4, null);
                                // }
                            }
                        }
                    }
                });
            }
        } else if (flagCheck == 'Exclusion') {
            const columnData = tabledata["ExTable 1"]?.columnNames;
            const EnableCsForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableCsForExclusion");
            const EnableLockCellForExclusion = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCellForExclusion");
            const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); 
            
            if (hasCangeDetected && hasCangeDetected === "true" ) {

                range?.forEach((r, index) => {
                    if (r?.row[0] === r?.row[1] && r?.column[0] === r?.column[1]) {
                        return;
                    }
                    if(r != undefined) {
                        for (let row = r?.row[0]; row <= r?.row[1]; row++) {
                            for (let column = r.column[0]; column <= r.column[1]; column++) {
                                if (EnableCsForExclusion == "true" && EnableLockCellForExclusion == "true") {
                                    if (exclusionApplicableIdx?.includes(column) && column != columnData?.ConfidenceScore && props?.enableExclusionCellLock) {
                                        const row_data = luckysheet.getcellvalue(row);
                                        const cs_score = getText(row_data[columnData["ConfidenceScore"]], false);
                                        if (cs_score?.trim() !== "" && cs_score?.trim()?.toLowerCase() !== "details not available in the document" && cs_score) {
                                            if (parseFloat(cs_score) > MinLockCellScore) {
                                                setTimeout(() => {
                                                    luckysheet.undo();
                                                    sessionStorage.setItem("cs_range_select", "false");
                                                }, 100);
                                                return;
                                            }
                                        }
                                    } else if (column == columnData?.ConfidenceScore) {
                                        luckysheet.undo()
                                        sessionStorage.setItem("cs_range_select", "false");
                                        return;
                                    }
                            } 
                        }
                    }
                    }
                })
            }
        }
    };

    document.onkeyup = function (e) {
        const range = luckysheet.getRange();
        if (!range || range.length === 0) return;
        let flagCheck = luckysheet.getSheet()?.name;
        let tabledata = flagCheck == "PolicyReviewChecklist" ? tableColumnDetails : formTableColumnDetails;
        const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
        const selectedIndex = range[0].row[0];
        const selectedRowIndex = range[0].row[0];
        const selectedColIndex = range[0].column[0];
        const selectedTable = findTableForIndex( selectedIndex, tabledata, excludedColumns );

         // Delete/Backspace Restrictions for confidence_score column cells
         if ( e.which != 40 ) {
            if (e.which === 46 || e.which === 8) {
                if (flagCheck == "PolicyReviewChecklist" || flagCheck == 'Forms Compare') {
                    if (range[0].row[0] == tabledata[selectedTable].range.start ) {
                        luckysheet.undo();
                    } 
                    else if(selectedTable != "Table 1" || selectedTable == "FormTable 1"){
                        if(range[0].row[0] == tabledata[selectedTable].range.start+1){
                            luckysheet.undo();
                        }
                    }
                }
                delKeyRestrictCSColum(selectedTable, selectedRowIndex, selectedColIndex, tabledata);
            }
        }

        // SHIFT + SPACEBAR - Select row
        if (e.shiftKey && e.key === " ") {
            let range = luckysheet.getRange();
            let selectedIndex = range[ 0 ].row[ 0 ];
            let selectedrow1 = range[ 0 ].row[ 0 ];
            let selectedrow2 = range[ 0 ].row[ 0 ];
            let tabledata = luckysheet.getSheetData()[ selectedIndex ];
            // console.log( "tabledata", tabledata );
            let count = 0;

            for ( let i = 0; i < tabledata.length; i++ )
            {
                if ( tabledata[ i ] !== null )
                {
                    count++;
                }
            }

            luckysheet.exitEditMode();
            luckysheet.setluckysheet_select_save( [ { row: [ selectedrow1, selectedrow2 ], column: [ 1, count ] } ] );
            luckysheet.selectHightlightShow();
        }

        // CTRL + D - Excel Option (Fill Down)
        if (e.ctrlKey && e.which === 68) {
            const targetrow = range[ 0 ].row[ 0 ]
            const targetcolumn = range[ 0 ].column
            const sheetdatas = luckysheet.getSheetData();
            const getrowdata = sheetdatas[ targetrow - 1 ]

            if ( range[ 0 ].column[ 0 ] == range[ 0 ].column[ 1 ] )
            {
                const getrowdata = sheetdatas[ targetrow - 1 ]
                if ( getrowdata[ targetcolumn[ 0 ] ] )
                {
                    luckysheet.setCellValue( targetrow, targetcolumn[ 0 ], getrowdata[ targetcolumn[ 0 ] ] );
                }
            } else
            {
                const startColumn = Math.min( range[ 0 ].column[ 0 ], range[ 0 ].column[ 1 ] );
                const endColumn = Math.max( range[ 0 ].column[ 0 ], range[ 0 ].column[ 1 ] );

                for ( let idx = startColumn; idx <= endColumn; idx++ )
                {
                    if ( getrowdata[ idx ] )
                    {
                        luckysheet.setCellValue( targetrow, idx, getrowdata[ idx ] );
                    }
                }
            }
        }

        // CTRL + Plus (Insert Row/Column)
        if (e.ctrlKey && (e.which === 187 || e.which === 17)) {
            singleMultipleSwitchInsert();
        }

        // CTRL + Minus (Delete Row/Column)
        if (e.ctrlKey && (e.which === 189 || e.which === 17)) {
            singleMultipleSwitchDelete();
        }

        // CTRL + S (Save)
        if (e.ctrlKey &&(e.which === 83 || e.which === 17)) {
            e.preventDefault();
            onUpdateClick(false, true, false);
        }

         // CTRL + SHIFT + F (Open Filter Dialog)  and  CTRL + F (Find Option)
        if (e.ctrlKey) {
            if(e.shiftKey && (e.key === "F" || e.key === "f" || e.keyCode === 70)) {
                if (luckysheet.getSheet()?.name === "PolicyReviewChecklist") {
                    toggleFilterDialog();
                }
            } else {
                if ((e.which === 70 || e.keyCode === 70)) {
                    toggleFindDialog();
                    setsheetState(luckysheet.getSheetData());
                }
            }
        }

        // CTRL + SHIFT + Up Arrow
        if ( e.ctrlKey && e.shiftKey && e.keyCode === 38 )
            { // CTRL + Shift + upArrow Excel Options
                let sheetcheck = luckysheet.getSheet().name;
                let range = luckysheet.getRange();
                const targetrow = sheetcheck == "Forms Compare" ? range[ 0 ].row[ 1 ] : range[ 0 ].row[ 0 ];
                let nullstartcolumncheck = luckysheet.getSheetData()[ targetrow ];
                const isAllNullstart = nullstartcolumncheck.every( element => element === null );
                if ( isAllNullstart == false )
                {
                    let sheetcheck = luckysheet.getSheet().name;
                    let range = luckysheet.getRange();
                    let selectedIndex = range[ 0 ].row[ 0 ];
                    let tabledata = sheetcheck == "PolicyReviewChecklist" ? tableColumnDetails : sheetcheck == "Forms Compare" ? formTableColumnDetails : sheetcheck == "Exclusion" ? exTableColumnDetails : "";
                    const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
                    const selectedTable = findTableForIndex( selectedIndex, tabledata, excludedColumns );
                    if ( selectedTable == 'Table 2' || selectedTable == 'FormTable 2' || selectedTable == 'Table 3' || selectedTable == 'ExTable 1' || selectedTable == 'Table 1' || selectedTable == 'FormTable 1' )
                    {
                        console.log( "tablename", selectedTable );

                        luckysheet.enterEditMode();
                        luckysheet.exitEditMode()
                        let range = luckysheet.getRange();
                        let selectedtabledata = tabledata[ selectedTable ];
                        if ( selectedTable == 'Table 3' || selectedTable == 'FormTable 2' || selectedTable == 'ExTable 1' || selectedTable == 'Table 1' )
                        {
                            let lastvalue = range[ 0 ].row[ 1 ];
                            let valueCheck = uparrowlastValue != null ? setUparrowlastValue( null ) : "";
                            setUparrowlastValue( lastvalue );
                        }
                        let row2 = range[ 0 ].row[ 0 ];
                        let row1 = selectedTable == 'Table 3' ? selectedtabledata.range.start + 3 : selectedTable == 'ExTable 1' ? selectedtabledata.range.start + 1 : selectedTable == 'Table 1' ? selectedtabledata.range.start + 4 : selectedTable == 'FormTable 1' ? selectedtabledata.range.start + 3 : selectedtabledata.range.start + 2;
                        let columns = range[ 0 ].column[ 0 ];
                        luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ columns, columns ] } ] )
                        luckysheet.selectHightlightShow()
                    } else
                    {
                        let selectedtabledata = tabledata[ selectedTable ];
                        luckysheet.enterEditMode();
                        luckysheet.exitEditMode()
                        let range = luckysheet.getRange();
                        if ( selectedTable == 'Table 4' || selectedTable == 'Table 5' || selectedTable == 'Table 6' || selectedTable == 'Table 7' || selectedTable == 'FormTable 3' )
                        {
                            let lastvalue = range[ 0 ].row[ 1 ];
                            let valueCheck = uparrowlastValue != null ? setUparrowlastValue( null ) : "";
                            setUparrowlastValue( lastvalue );
                        }
                        let row1 = range[ 0 ].row[ 0 ];
                        let row2 = selectedtabledata.range.start + 2;
                        let columns = range[ 0 ].column[ 0 ];
                        if ( row1 > row2 )
                        {
                            luckysheet.setluckysheet_select_save( [ { row: [ row2, row1 ], column: [ columns, columns ] } ] )
                            luckysheet.selectHightlightShow()
                        } else
                        {
                            luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ columns, columns ] } ] )
                            luckysheet.selectHightlightShow()
                        }
                    }
                }
        }

        // CTRL + SHIFT + Down Arrow
        if ( e.ctrlKey && e.shiftKey && e.keyCode === 40 )
        { // CTRL + Shift + Down Arrow 
            e.preventDefault();
            let secoundtablerange = localStorage.getItem('secoundtabledata');
            let sheetchecks = luckysheet.getSheet().name;
            let secoundrange = sheetchecks == 'Exclusion' ? luckysheet.getRange() : JSON.parse( secoundtablerange ) != undefined || null ? JSON.parse( secoundtablerange ) : luckysheet.getRange();
            let range = secoundrange;
            // setSecoundtablerange( [] );
            localStorage.removeItem('secoundtabledata');
            const targetrow = range[ 0 ].row[ 0 ]
            let nullstartcolumncheck = luckysheet.getSheetData()[ targetrow ];
            const isAllNullstart = nullstartcolumncheck.every( element => element === null );
            if ( isAllNullstart == false || sheetchecks == 'Exclusion' )
            {
                let sheetcheck = luckysheet.getSheet().name;
                let selectedIndex = range[ 0 ].row[ 0 ];
                let tabledata = sheetcheck == "PolicyReviewChecklist" ? tableColumnDetails : sheetcheck == "Forms Compare" ? formTableColumnDetails : sheetcheck == "Exclusion" ? exTableColumnDetails : "";
                const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
                const selectedTable = findTableForIndex( selectedIndex, tabledata, excludedColumns );
                if ( selectedTable == 'Table 2' )
                {
                    console.log( "tablename", selectedTable );
                    luckysheet.enterEditMode();
                    luckysheet.exitEditMode()
                    let range = secoundrange;
                    let selectedtabledata = tabledata[ selectedTable ];
                    let row1 = range[ 0 ].row[ 1 ];
                    let row2 = selectedtabledata.range.end
                    let columns = range[ 0 ].column[ 0 ];
                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ columns, columns ] } ] )
                    luckysheet.selectHightlightShow()
                }
                if ( selectedTable == 'Table 3' || selectedTable == 'Table 1' || selectedTable == 'FormTable 2' || selectedTable == 'ExTable 1' || selectedTable == 'Table 4' || selectedTable == 'Table 5' || selectedTable == 'Table 6' || selectedTable == 'Table 7' || selectedTable == 'FormTable 3' )
                {
                    let selectedtabledata = tabledata[ selectedTable ];
                    let row1 = uparrowlastValue == null ? range[ 0 ].row[ 0 ] : uparrowlastValue;
                    let row2 = selectedtabledata.range.end;
                    let columns = range[ 0 ].column[ 0 ];
                    setUparrowlastValue( null )
                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ columns, columns ] } ] )
                    luckysheet.selectHightlightShow()
                }
                if ( sheetchecks == 'Exclusion' || selectedTable == 'FormTable 3' )
                {
                    const sheetConfig = luckysheet.getluckysheet_select_save();
                    $( "#luckysheet-scrollbar-y" ).scrollTop( sheetConfig[ 0 ]?.top );
                }
            }
        }

        // CTRL + SHIFT + Right Arrow and CTRL + SHIFT + Left Arrow
        if ( e.ctrlKey && e.shiftKey && e?.which === 39 || e.ctrlKey && e.shiftKey && e?.which === 37 )
            {   // CTRL + Shift + rightArrow && CTRL + Shift + leftArrow Excel Options
                let ctrloptions = sessionStorage.getItem( 'ctrloptions' );
                let range = JSON.parse( ctrloptions );
                let col1 = range[ 0 ].column[ 0 ];
                let col2 = range[ 0 ].column[ 1 ];
                let row1 = range[ 0 ].row[ 0 ];
                let row2 = range[ 0 ].row[ 1 ];
                let sheetcheck = luckysheet.getSheet().name;
                let tabledata = sheetcheck == "PolicyReviewChecklist" ? tableColumnDetails : sheetcheck == "Forms Compare" ? formTableColumnDetails : sheetcheck == "Exclusion" ? exTableColumnDetails : "";
                const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
                const selectedTable = findTableForIndex( row1, tabledata, excludedColumns );
                if ( e?.which === 39 || e?.which === 37 )
                {
                    const table = tabledata[ selectedTable ];
                    let value = table ? Object.values( table.columnNames ) : [];
                    let values = value.filter( f => f !== 0 );
                    if ( selectedTable !== 'Table 1' && selectedTable !== 'FormTable 1' && selectedTable !== 'ExTable 1' )
                    {
                        let startColumn, endColumn;
                        if ( e.which === 39 )
                        {
                            endColumn = Math.max( ...values );
                            if ( selectedTable !== 'FormTable 2' && selectedTable !== 'FormTable 3' )
                            {
                                if ( col1 > 1 )
                                {
                                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ col1, endColumn ] } ] )
                                    $( "#luckysheet-scrollbar-x" ).scrollLeft( 1000 - 800 );
                                } else
                                {
                                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ col2, endColumn ] } ] )
                                    $( "#luckysheet-scrollbar-x" ).scrollLeft( 1000 - 800 );
                                }
                            }
                            if ( selectedTable == 'FormTable 2' || selectedTable == 'FormTable 3' )
                            {
                                if ( col1 > 1 )
                                {
                                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ col1, endColumn ] } ] )
                                    $( "#luckysheet-scrollbar-x" ).scrollLeft( 1000 - 800 );
                                } else
                                {
                                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ col2 - col1, endColumn ] } ] )
                                    $( "#luckysheet-scrollbar-x" ).scrollLeft( 1000 - 800 );
                                }
                            }
                        } else if ( e.which === 37 )
                        {
                            startColumn = Math.min( ...values );
                            if ( selectedTable == 'FormTable 2' || selectedTable == 'FormTable 3' )
                            {
                                if ( col1 > 1 )
                                {
                                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ startColumn, col1 ] } ] )
                                } else
                                {
                                    luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ startColumn, col2 ] } ] )
                                }
                            } else
                            {
                                luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ startColumn, col2 ] } ] )
                            }
                        }
                        luckysheet.selectHightlightShow();
                    }
                } if ( selectedTable == 'ExTable 1' )
                {
                    if ( e?.which === 39 || e?.which === 37 )
                    {
                        const table = tabledata[ selectedTable ];
                        let startColumn, endColumn;
                        if ( e.which === 39 )
                        {
                            endColumn = table.columnNames[ 0 ].indexOf( "PageNumber" );
                            luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ col1, endColumn ] } ] );
                            $( "#luckysheet-scrollbar-x" ).scrollLeft( 1000 - 800 );
                        } else if ( e.which === 37 )
                        {
                            startColumn = table.columnNames[ 0 ].indexOf( "FormName" );
                            luckysheet.setluckysheet_select_save( [ { row: [ row1, row2 ], column: [ startColumn, col2 ] } ] );
                        }
                        luckysheet.selectHightlightShow();
                    }
                }
        }
    };

    // Zooming with CTRL + Mouse Wheel
    document.addEventListener("wheel", function (e) {
        if (e.ctrlKey) {
            e.preventDefault();
            if (!e.zoomExecuted) {
                let scrollDirection = e.deltaY || e.detail || e.wheelDelta;
                let SheetZoomValue = luckysheet.getSheet().zoomRatio || 1;
                let SetZoom;

                if (scrollDirection < 100) {
                    if (SheetZoomValue < 3.8) {
                        SetZoom = SheetZoomValue + 0.15;
                    }
                } else {
                    if (SheetZoomValue > 0.25) {
                        SetZoom = SheetZoomValue - 0.15;
                    }
                }

                luckysheet.setSheetZoom(SetZoom ?? SheetZoomValue);
                e.zoomExecuted = true;
            }
        }
            return true;
    } );

    // const singleMultipleSwitchInsert = () => {
    //     const isMultiRowSelected = hasMultipleRowsSelected;
    //     if (isMultiRowSelected && selectedRowIndexRange.length > 0) {
    //         // Insert only one row at the first selected index
    //         insertRow(selectedRowIndexRange[0]);
    //     } else {
    //         insertRow(setectedRowIndex);
    //     }
    // }
    
    const singleMultipleSwitchInsert = ( isInsertBydialog ) => {
        let getFlag = luckysheet.getSheet().name;
        if ( getFlag !== "Exclusion" )
        {
            const hardCodedStartingIndex = getFlag === 'Forms Compare' ? [ 0, 1, 2 ] : [ 0, 1, 2, 3 ];
            if ( ( hasMultipleRowsSelected && hardCodedStartingIndex.includes( selectedRowIndexRange[ 0 ] ) ) || hardCodedStartingIndex.includes( setectedRowIndex ) )
            {
                const msg = `cannot add rows in the sheetname sections`;
                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                return;
            }
        } else
        {
            const hardCodedStartingIndex = getFlag === 'Exclusion' ? [ 0 ] : [ 0 ];
            if ( ( hasMultipleRowsSelected && hardCodedStartingIndex.includes( selectedRowIndexRange[ 0 ] ) ) || hardCodedStartingIndex.includes( setectedRowIndex ) )
            {
                const msg = `cannot add rows in the sheetname sections`;
                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                return;
            }
        }
        let QacFlag = luckysheet.getSheet().name
        if ( QacFlag != 'QAC not answered questions' )
        {
            const luckySheet = luckysheet.getSheetData()[ 1 ];
            let flagCheck = QacFlag;
            if ( flagCheck == 'PolicyReviewChecklist' )
            {
                const isMultiRowSelected = hasMultipleRowsSelected;
                let currentTableRecord = "";
                let TableName = "";
                let tableNameKeys = Object.keys( tableColumnDetails );
                tableNameKeys = tableNameKeys.filter( ( f ) => f != "Table 1" );
                tableNameKeys.forEach( ( columnName ) => {
                    if ( ( ( tableColumnDetails[ columnName ]?.range?.start <= setectedRowIndex && tableColumnDetails[ columnName ]?.range?.end >= setectedRowIndex ) ||
                        ( isMultiRowSelected && tableColumnDetails[ columnName ]?.range?.start <= selectedRowIndexRange[ 0 ] && tableColumnDetails[ columnName ]?.range?.end >= selectedRowIndexRange[ 1 ] ) ) && Object.keys( tableColumnDetails[ columnName ]?.columnNames )?.length > 0 )
                    {
                        currentTableRecord = tableColumnDetails[ columnName ];
                        TableName = columnName;
                    }
                } );

                if ( currentTableRecord?.range && ( ( isMultiRowSelected && currentTableRecord?.range?.end >= selectedRowIndexRange[ 0 ] ) || currentTableRecord?.range?.end >= setectedRowIndex ) )
                {
                    if ( TableName != "Table 1" )
                    {
                        if ( TableName != "Table 3" && ( ( isMultiRowSelected && !( currentTableRecord?.range?.start + 1 < selectedRowIndexRange[ 0 ] ) ) || !( currentTableRecord?.range?.start + 1 < setectedRowIndex ) ) )
                        {
                            const msg = `cannot add rows in the ${ TableName } header sections`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        }
                        if ( TableName === "Table 3" )
                        {
                            if ( ( isMultiRowSelected && !( currentTableRecord?.range?.start + 2 < selectedRowIndexRange[ 0 ] ) ) || !( currentTableRecord?.range?.start + 2 < setectedRowIndex ) )
                            {
                                const msg = `cannot add rows in the ${ TableName } header sections`;
                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                return;
                            }
                        }
                    } else
                    {
                        const msg = `cannot add rows in the ${ TableName == "Table 1" ? "Header" : TableName } header sections`;
                        setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                        setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                        return;
                    }
                }
                if ( ( ( !isMultiRowSelected && tableColumnDetails[ "Table 2" ]?.range?.start < setectedRowIndex ) ||
                    ( isMultiRowSelected && tableColumnDetails[ "Table 2" ]?.range?.start < selectedRowIndexRange[ 0 ] ) ) && !TableName )
                {
                    const msg = "Only able to insert rows insde the tables";
                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                    return;
                }
                if ( isMultiRowSelected && selectedRowIndexRange.length > 0 )
                {
                    if ( isInsertBydialog )
                    {
                        setOpenInputDialog( true );
                        return;
                    }
                    // return; //as dialog implemented this loop is not necessary by ---**gokul**------
                    //console.log( "before", tableColumnDetails );
                    const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                    const tableNameKeysBackup = tableColumnDetails;
                    // tableNameKeys.forEach((columnName) => {
                    //     if (
                    //         (isMultiRowSelected && tableNameKeysBackup[columnName]?.range?.start > selectedRowIndexRange[0] &&
                    //             tableNameKeysBackup[columnName]?.range?.end >= selectedRowIndexRange[0]) &&
                    //         Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.start += selectedDiff + 1;
                    //         tableNameKeysBackup[columnName].range.end += selectedDiff + 1;
                    //     } else if ((isMultiRowSelected && tableNameKeysBackup[columnName]?.range?.start < selectedRowIndexRange[0] &&
                    //         tableNameKeysBackup[columnName]?.range?.end >= selectedRowIndexRange[0]) &&
                    //         Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.end += selectedDiff + 1;
                    //     }
                    // });
                    // console.log( "after", tableNameKeysBackup );
                    // setTableColumnDetails(tableNameKeysBackup);
                    // insertRow(selectedRowIndexRange[0], true, selectedDiff + 1);
                    luckySheetInsert( selectedRowIndexRange[ 0 ], selectedDiff + 1, TableName, luckysheet?.getSheet() );

                } else
                {
                    if ( isInsertBydialog )
                    {
                        setOpenInputDialog( true );
                        return;
                    }
                    // setOpenInputDialog(true); --will enable when the dialog is agreed upon by exdion team --by gokul
                    // console.log( "before", tableColumnDetails );
                    const tableNameKeysBackup = tableColumnDetails;
                    // tableNameKeys.forEach((columnName) => {
                    //     if (((tableNameKeysBackup[columnName]?.range?.start > setectedRowIndex && tableNameKeysBackup[columnName]?.range?.end >= setectedRowIndex)) && Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.start += 1;
                    //         tableNameKeysBackup[columnName].range.end += 1;
                    //     } else if (((tableNameKeysBackup[columnName]?.range?.start < setectedRowIndex && tableNameKeysBackup[columnName]?.range?.end >= setectedRowIndex)) &&
                    //         Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.end += 1;
                    //     }
                    // });
                    //console.log( "after", tableNameKeysBackup );
                    // setTableColumnDetails(tableNameKeysBackup);
                    // insertRow(setectedRowIndex, false, 1);
                    luckySheetInsert( setectedRowIndex, 1, TableName, luckysheet?.getSheet() );
                }
            } else if ( flagCheck == 'Forms Compare' )
            {
                const isMultiRowSelected = hasMultipleRowsSelected;
                let currentTableRecord = "";
                let TableName = "";
                let tableNameKeys = Object.keys( formTableColumnDetails );
                tableNameKeys = tableNameKeys.filter( ( f ) => f != "FormTable 1" );
                tableNameKeys.forEach( ( columnName ) => {
                    if ( ( ( formTableColumnDetails[ columnName ]?.range?.start <= setectedRowIndex && formTableColumnDetails[ columnName ]?.range?.end >= setectedRowIndex ) ||
                        ( isMultiRowSelected && formTableColumnDetails[ columnName ]?.range?.start <= selectedRowIndexRange[ 0 ] && formTableColumnDetails[ columnName ]?.range?.end >= selectedRowIndexRange[ 1 ] ) ) && Object.keys( formTableColumnDetails[ columnName ]?.columnNames )?.length > 0 )
                    {
                        currentTableRecord = formTableColumnDetails[ columnName ];
                        TableName = columnName;
                    }
                } );
                if ( TableName == '' )
                {
                    let nullcolumncheck = luckysheet.getSheetData()[ setectedRowIndex ];
                    const isAllNull = nullcolumncheck.every( element => element === null );
                    if ( isAllNull )
                    {
                        const msg = `cannot add rows within table section`;
                        setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                        setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                        return;
                    }

                }
                if ( currentTableRecord?.range && ( ( isMultiRowSelected && currentTableRecord?.range?.end >= selectedRowIndexRange[ 0 ] ) || currentTableRecord?.range?.end >= setectedRowIndex ) )
                {
                    if ( TableName != "FormTable 1" )
                    {
                        let nullcolumncheck = luckysheet.getSheetData()[ setectedRowIndex ];
                        const headercolcheck = currentTableRecord?.range?.start;
                        const headerendcolcheck = currentTableRecord?.range?.end;
                        var range = luckysheet.getRange();
                        const targetrow = range[ 0 ].row[ 0 ]
                        const targetendrow = range[ 0 ].row[ 1 ]

                        const headercolval1 = headercolcheck;
                        const headercolval2 = headercolcheck + 1;
                        let headercolcheck1 = headercolcheck + 1;

                        if ( headercolcheck == targetrow )
                        {
                            const msg = `cannot add rows in the ${ TableName } header sections`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        } else
                        {
                            if ( headercolcheck1 == targetrow )
                            {
                                const msg = `cannot add rows in the ${ TableName } header sections`;
                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                return;
                            }
                        }
                        var range = luckysheet.getRange();
                        let sheetselectedindex = setectedRowIndex;
                        let sheetdata = luckysheet.getSheetData();
                        if ( sheetselectedindex >= 0 && sheetselectedindex < sheetdata.length )
                        {
                            sheetdata = sheetdata.slice( sheetselectedindex );
                            sheetdata = sheetdata.map( ( element, index ) => {
                                return {
                                    [ sheetselectedindex + index + 1 ]: element
                                };
                            } );
                            sheetselectedindex = sheetselectedindex + 1;
                        }
                        let valuebeforenull = [];
                        for ( let i = 0; i < sheetdata.length; i++ )
                        {
                            const obj = sheetdata[ i ];
                            const key = Object.keys( obj )[ 0 ];
                            const values = obj[ key ];
                            if ( values.every( val => val === null ) )
                            {
                                valuebeforenull = sheetdata.slice( 0, i );
                                break;
                            }
                        }

                        if ( valuebeforenull.length === 0 && sheetdata.length > 0 )
                        {
                            valuebeforenull.push( sheetdata[ sheetdata.length - 1 ] );
                        }
                        let lastIndex = valuebeforenull.length - 1;
                        const keys = Object.keys( lastIndex == 0 ? valuebeforenull[ 0 ] : sheetdata[ lastIndex ] );
                        const key = parseInt( keys[ 0 ] );
                        var range = luckysheet.getRange();
                        let columnrender2 = range[ 0 ].row[ 1 ];
                        let checktableexist = key - 1;

                        if ( targetendrow > checktableexist )
                        {
                            const msg = `cannot add rows in the ${ TableName } header sections`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        }

                        if ( targetrow == headercolval1 && targetrow == headercolval2 )
                        {
                            if ( ( isMultiRowSelected && !( currentTableRecord?.range?.start + 2 < selectedRowIndexRange[ 0 ] ) ) || !( currentTableRecord?.range?.start + 2 < setectedRowIndex ) )
                            {
                                const msg = `cannot add rows in the ${ TableName } header sections`;
                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                return;
                            }
                        }
                    } else
                    {
                        const msg = `cannot add rows in the ${ TableName == "Table 1" ? "Header" : TableName } header sections`;
                        setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                        setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                        return;
                    }

                }
                if ( isMultiRowSelected && selectedRowIndexRange.length > 0 )
                {
                    if ( isInsertBydialog )
                    {
                        setOpenInputDialog( true );
                        return;
                    }
                    // return;
                    //console.log( "before", tableColumnDetails );
                    const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                    const tableNameKeysBackup = formTableColumnDetails;
                    const policytableNameKeysBackup = tableColumnDetails;
                    // tableNameKeys.forEach((columnName) => {
                    //     if (
                    //         (isMultiRowSelected && tableNameKeysBackup[columnName]?.range?.start > selectedRowIndexRange[0] &&
                    //             tableNameKeysBackup[columnName]?.range?.end >= selectedRowIndexRange[0]) &&
                    //         Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.start += selectedDiff + 1;
                    //         tableNameKeysBackup[columnName].range.end += selectedDiff + 1;
                    //     } else if ((isMultiRowSelected && tableNameKeysBackup[columnName]?.range?.start < selectedRowIndexRange[0] &&
                    //         tableNameKeysBackup[columnName]?.range?.end >= selectedRowIndexRange[0]) &&
                    //         Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.end += selectedDiff + 1;
                    //     }
                    // });
                    // console.log( "after", tableNameKeysBackup );
                    // setFormTableColumnDetails(tableNameKeysBackup);
                    // setTableColumnDetails( policytableNameKeysBackup );
                    // insertRow(selectedRowIndexRange[0], true, selectedDiff + 1);
                    luckySheetInsert( selectedRowIndexRange[ 0 ], selectedDiff + 1, TableName, luckysheet?.getSheet() );
                } else
                {
                    if ( isInsertBydialog )
                    {
                        setOpenInputDialog( true );
                        return;
                    }
                    // setOpenInputDialog(true);
                    // console.log( "before", tableColumnDetails );   //changedss
                    const tableNameKeysBackup = formTableColumnDetails;
                    const policytableNameKeysBackup = tableColumnDetails;
                    // tableNameKeys.forEach((columnName) => {
                    //     if (((tableNameKeysBackup[columnName]?.range?.start > setectedRowIndex && tableNameKeysBackup[columnName]?.range?.end >= setectedRowIndex)) && Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.start += 1;
                    //         tableNameKeysBackup[columnName].range.end += 1;
                    //     } else if (((tableNameKeysBackup[columnName]?.range?.start < setectedRowIndex && tableNameKeysBackup[columnName]?.range?.end >= setectedRowIndex)) &&
                    //         Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.end += 1;
                    //     }
                    // });
                    //console.log( "after", tableNameKeysBackup );
                    // setFormTableColumnDetails(tableNameKeysBackup);
                    // setTableColumnDetails( policytableNameKeysBackup );
                    // insertRow(setectedRowIndex, false, 1);
                    luckySheetInsert( setectedRowIndex, 1, TableName, luckysheet?.getSheet() );
                }
            } else if ( flagCheck == 'Exclusion' )
            {
                const isMultiRowSelected = hasMultipleRowsSelected;
                let currentTableRecord = "";
                let TableName = "";
                let tableNameKeys = Object.keys( exTableColumnDetails );
                tableNameKeys.forEach( ( columnName ) => {
                    if ( ( ( exTableColumnDetails[ columnName ]?.range?.start <= setectedRowIndex && exTableColumnDetails[ columnName ]?.range?.end >= setectedRowIndex ) ||
                        ( isMultiRowSelected && exTableColumnDetails[ columnName ]?.range?.start <= selectedRowIndexRange[ 0 ] && exTableColumnDetails[ columnName ]?.range?.end >= selectedRowIndexRange[ 1 ] ) ) && Object.keys( exTableColumnDetails[ columnName ]?.columnNames )?.length > 0 )
                    {
                        currentTableRecord = exTableColumnDetails[ columnName ];
                        TableName = columnName;
                    }
                } );
                if ( TableName == '' )
                    {
                        let nullcolumncheck = luckysheet.getSheetData()[ setectedRowIndex ];
                        const isAllNull = nullcolumncheck.every( element => element === null );
                        if ( isAllNull )
                        {
                            const msg = `cannot add rows within table section`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        }
    
                    }
                if ( isMultiRowSelected && selectedRowIndexRange.length > 0 )
                {
                    if ( isInsertBydialog )
                    {
                        setOpenInputDialog( true );
                        return;
                    }
                    const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                    // tableNameKeys.forEach((columnName) => {
                    //     if (
                    //         (isMultiRowSelected && exTableColumnDetails[columnName]?.range?.start > selectedRowIndexRange[0] &&
                    //             exTableColumnDetails[columnName]?.range?.end >= selectedRowIndexRange[0]) &&
                    //         Object.keys(exTableColumnDetails[columnName]?.columnNames)?.length > 0) {
                    //         exTableColumnDetails[columnName].range.start += selectedDiff + 1;
                    //         exTableColumnDetails[columnName].range.end += selectedDiff + 1;
                    //     } else if ((isMultiRowSelected && exTableColumnDetails[columnName]?.range?.start < selectedRowIndexRange[0] &&
                    //         exTableColumnDetails[columnName]?.range?.end >= selectedRowIndexRange[0]) &&
                    //         Object.keys(exTableColumnDetails[columnName]?.columnNames)?.length > 0) {
                    //         exTableColumnDetails[columnName].range.end += selectedDiff + 1;
                    //     }
                    // });
                    // console.log( "after", tableNameKeysBackup );
                    luckySheetInsert( selectedRowIndexRange[ 0 ], selectedDiff + 1, TableName, luckysheet?.getSheet() );
                    // insertRow(selectedRowIndexRange[0], true, selectedDiff + 1);
                } else
                {
                    if ( isInsertBydialog )
                    {
                        setOpenInputDialog( true );
                        return;
                    }
                    // const tableNameKeysBackup = exTableColumnDetails;
                    // tableNameKeys.forEach((columnName) => {
                    //     if (((tableNameKeysBackup[columnName]?.range?.start > setectedRowIndex && tableNameKeysBackup[columnName]?.range?.end >= setectedRowIndex)) && Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.start += 1;
                    //         tableNameKeysBackup[columnName].range.end += 1;
                    //     } else if (((tableNameKeysBackup[columnName]?.range?.start < setectedRowIndex && tableNameKeysBackup[columnName]?.range?.end >= setectedRowIndex)) &&
                    //         Object.keys(tableNameKeysBackup[columnName]?.columnNames)?.length > 0) {
                    //         tableNameKeysBackup[columnName].range.end += 1;
                    //     }
                    // });
                    luckySheetInsert( setectedRowIndex, 1, TableName, luckysheet?.getSheet() );
                    // setExTableColumnDetails(tableNameKeysBackup);
                    // insertRow(setectedRowIndex, false, 1);
                }
            }
        }

    }
    // insert based on the input given in the inputdialog ---by Gokul---
    const insertFnByInputDialog = ( noOfRows ) => {
        const isMultiRowSelected = true; // by default it will be true always 
        const selectedDiff = noOfRows;
        let currentTablename = '';
        const currentSheetData = luckysheet?.getSheet();
        if ( currentSheetData?.name == 'PolicyReviewChecklist' )
        {
            const tableNameKeysBackup = tableColumnDetails;
            let tableNameKeys = Object.keys( tableColumnDetails );
            tableNameKeys = tableNameKeys.filter( ( f ) => f != "Table 1" );
            tableNameKeys.forEach( ( columnName ) => {
                if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start > setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) && Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    // tableNameKeysBackup[ columnName ].range.start += selectedDiff;
                    // tableNameKeysBackup[ columnName ].range.end += selectedDiff;
                } else if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start < setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) &&
                    Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    currentTablename = columnName;
                    // tableNameKeysBackup[ columnName ].range.end += selectedDiff;
                }
            } );
            // setTableColumnDetails( tableNameKeysBackup );
        }
        if ( currentSheetData?.name == 'Forms Compare' )
        {
            const tableNameKeysBackup = formTableColumnDetails;
            let tableNameKeys = Object.keys( formTableColumnDetails );
            tableNameKeys.forEach( ( columnName ) => {
                if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start > setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) && Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    // tableNameKeysBackup[ columnName ].range.start += selectedDiff;
                    // tableNameKeysBackup[ columnName ].range.end += selectedDiff;
                } else if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start < setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) &&
                    Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    currentTablename = columnName;
                    // tableNameKeysBackup[ columnName ].range.end += selectedDiff;
                }
            } );
            // setFormTableColumnDetails( tableNameKeysBackup );
        }
        if ( currentSheetData?.name == 'Exclusion' )
        {
            const tableNameKeysBackup = exTableColumnDetails;
            let tableNameKeys = Object.keys( exTableColumnDetails );
            tableNameKeys.forEach( ( columnName ) => {
                if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start > setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) && Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    console.log( "Table" );
                } else if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start < setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) &&
                    Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    currentTablename = columnName;
                }
            } );
        }

        // console.log( "after", tableNameKeysBackup );
        // insertRow(setectedRowIndex, true, selectedDiff);
        luckySheetInsert( setectedRowIndex, selectedDiff, currentTablename, currentSheetData );
    }
    /*
    selectedIndex -> start of index from where we need to insert the new rows
    isMultipleInsert -> flag to identify whether it is multiple  insert or not
    difference -> no.of rows to insert
    */
    const insertRow = ( selectedIndex, isMultipleInsert, difference ) => {
        let flagCheck = luckysheet.getSheet()?.name;
        if ( selectedIndex != null && selectedIndex != '' && selectedIndex != undefined )
        {
            const sheetData = luckysheet.getSheetData();
            // const valueddata = sheetData[selectedIndex];
            // const alltablelength = valueddata.filter(item => item !== null);
            const configData = luckysheet.getConfig();
            const emptyCell = Array( sheetData[ 0 ].length ).fill( null );
            // Update configData based on the difference
            configData.borderInfo?.forEach( e => {
                if ( e?.rangeType === 'cell' && e?.value?.row_index >= selectedIndex )
                {
                    e.value.row_index += difference;
                } else if ( e?.rangeType === 'range' )
                {
                    if ( e?.range?.length > 0 && e?.range[ 0 ]?.row?.length > 0 )
                    {
                        const [ startRow, endRow ] = e.range[ 0 ].row;
                        if ( startRow >= selectedIndex )
                        {
                            e.range[ 0 ].row[ 0 ] += difference;
                            e.range[ 0 ].row[ 1 ] += difference;
                        } else if ( startRow <= selectedIndex && endRow >= selectedIndex )
                        {
                            e.range[ 0 ].row[ 1 ] += difference;
                        }
                    }
                }
            } );
            // Update row lengths
            const updatedRowLen = { ...configData.rowlen };
            Object.keys( updatedRowLen ).forEach( f => {
                const convertedRowLen = parseInt( f );
                if ( convertedRowLen >= selectedIndex )
                {
                    updatedRowLen[ `${ convertedRowLen + difference }` ] = configData.rowlen[ f ];
                }
            } );

            // Update merges
            const merge = {};
            Object.keys( configData.merge ).forEach( k => {
                const [ findRow, findCol ] = k.split( '_' ).map( Number );
                if ( findRow >= selectedIndex )
                {
                    const existingData = configData.merge[ k ];
                    existingData.r = findRow + difference;
                    merge[ `${ findRow + difference }_${ findCol }` ] = existingData;
                } else
                {
                    merge[ k ] = configData.merge[ k ];
                }
            } );
            configData.merge = merge;

            //data insert part
            const defaultText = "Page #";
            const exText = " ";
            const defalutAttachedForms = "Attached Forms";
            const defaultQA = "CA2"
            // Adjust sheet data
            const emptyRow = Array( sheetData[ 0 ].length ).fill( null );
            const newRow = Array.from( { length: difference }, () => emptyRow );
            if ( flagCheck == "PolicyReviewChecklist" )
            {
                let tableName = '';

                const keys = Object.keys( tableColumnDetails );

                keys.forEach( ( f ) => {
                    const targetTableDetails = tableColumnDetails[ f ];
                    // console.log(targetTableDetails);
                    if ( targetTableDetails && targetTableDetails?.range && targetTableDetails?.range?.start <= selectedIndex && targetTableDetails?.range?.end >= selectedIndex )
                    {
                        tableName = f;
                    }
                } );

                if ( tableName && tableName != 'Table 1' )
                {
                    const tagetTableColumnDetails = tableColumnDetails[ tableName ];

                    const columnData = tagetTableColumnDetails?.columnNames;
                    const filteredData = {};
                    for ( const key in columnData )
                    {
                        if ( columnData.hasOwnProperty( key ) && columnData[ key ] > 0 )
                        {
                            filteredData[ key ] = columnData[ key ];
                        }
                    }
                    console.log( filteredData );
                    const filteredKeys = Object.keys( filteredData ).filter(
                        key => filteredData[ key ] > filteredData[ "ChecklistQuestions" ] && filteredData[ key ] < filteredData[ "Observation" ]
                    );

                    console.log( filteredKeys );

                    newRow.map( ( nr ) => {
                        filteredKeys.map( ( fk ) => {
                            const toBeInsertedIndex = filteredData[ fk ];
                            nr[ toBeInsertedIndex ] = { v: defaultText, fc: 'rgb(68, 114, 196)', fs: "8", };
                        } );
                        return nr;
                    } );
                }
            } else if ( flagCheck == "Forms Compare" )
            {

                let tableName = '';

                const keys = Object.keys( formTableColumnDetails );

                keys.forEach( ( f ) => {
                    const targetTableDetails = formTableColumnDetails[ f ];
                    // console.log(targetTableDetails);
                    if ( targetTableDetails && targetTableDetails?.range && targetTableDetails?.range?.start <= selectedIndex && targetTableDetails?.range?.end >= selectedIndex )
                    {
                        tableName = f;
                    }
                } );

                if ( tableName && tableName != 'FormTable 1' )
                {
                    newRow.map( ( nr ) => {
                        nr[ 1 ] = { v: defalutAttachedForms, fs: "8" };
                        nr[ 2 ] = { v: defaultQA, fs: "8" };
                        nr[ 3 ] = { v: defaultText, fc: 'rgb(68, 114, 196)', fs: "8", };
                        nr[ 4 ] = { v: defaultText, fc: 'rgb(68, 114, 196)', fs: "8", };
                        return nr;
                    } );
                }
            }

            // console.log( newRow );
            //insert new record and placing "Page #" in Document columns//
            // let newRow = [];
            // for (let index = 0; index < difference; index++) {

            //     const data = emptyCell.map((cell, cIndex) => {
            //         const excludedColumns = excludedColumnlist
            //         if (luckysheet.getSheet().name !== 'Exclusion') {
            //             const selectedTable = findTableForIndex(selectedIndex, luckysheet.getSheet().name == 'PolicyReviewChecklist' ? tableColumnDetails : formTableColumnDetails, excludedColumns);
            //             function findTableForIndex(selectedIndex, tableDetails, excludedColumns) {
            //                 for (const tableName in tableDetails) {
            //                     if (tableDetails.hasOwnProperty(tableName)) {
            //                         const range = tableDetails[tableName].range;
            //                         const columnNames = tableDetails[tableName].columnNames;
            //                         if (typeof range.start === 'number' && typeof range.end === 'number') {
            //                             if (selectedIndex >= range.start && selectedIndex <= range.end) {
            //                                 if (columnNames && typeof columnNames === 'object') {
            //                                     const validColumns = Object.keys(columnNames).filter(colName => !excludedColumns.includes(colName));
            //                                     if (validColumns.length > 0) {
            //                                         return tableName;
            //                                     }
            //                                 }
            //                             }
            //                         }
            //                     }
            //                 }

            //                 return null;
            //             }
            //             let tablemissing = "FormTable 3";
            //             let pagelistcheck = luckysheet.getSheet().name == 'PolicyReviewChecklist' ? tableColumnDetails[selectedTable].columnNames : formTableColumnDetails[selectedTable == null ? tablemissing : selectedTable].columnNames;

            //             excludedColumns.forEach(column => {
            //                 if (pagelistcheck.hasOwnProperty(column)) {
            //                     delete pagelistcheck[column];
            //                 }
            //             });

            //             for (let key in pagelistcheck) {
            //                 if (pagelistcheck[key] === 0) {
            //                     delete pagelistcheck[key];
            //                 }
            //             }
            //             const indexArray = Object.keys(pagelistcheck).map((key, index) => index.toString());
            //             const headers = Object.keys(indexArray);
            //             const header = headers[cIndex];
            //             const val = sheetData[selectedIndex] ? sheetData[selectedIndex][cIndex] : null;


            //             const luckySheet = luckysheet.getSheetData()[1];
            //             let flagCheck = luckySheet[1].m;
            //             if (flagCheck == 'FORM COMPARE') {
            //                 if (selectedTable != 'FormTable 1') {
            //                     if (header == 3 || header == 4) {
            //                         return {
            //                             v: defaultText,
            //                             fc: 'rgb(68, 114, 196)',
            //                             fs: "8",
            //                         };
            //                     } else if (header == 1) {
            //                         return {
            //                             v: defalutAttachedForms,
            //                             fs: "8",
            //                         };
            //                     }
            //                     else if (header == 2) {
            //                         return {
            //                             v: defaultQA,
            //                             fs: "8",
            //                         };
            //                     }
            //                     else {
            //                         return null;
            //                     }
            //                 }
            //             } else {
            //                 if(pagelistcheck.hasOwnProperty("Lob")){
            //                     if (
            //                         headers.length === 7 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[3] && header !== headers[6] && header !== headers[7] && val !== null)
            //                         || headers.length === 9 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[8] && header !== headers[3] && header !== headers[9] && val !== null)
            //                         || headers.length === 10 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[3] && header !== headers[9] && header !== headers[10] && val !== null)
            //                         || headers.length === 8 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[3] && header !== headers[7] && header !== headers[8] && val !== null)
            //                         || headers.length === 6 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[3] && header !== headers[5] && header !== headers[6] && val !== null)
            //                         || headers.length === 5 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[4] && header !== headers[3] && header !== headers[5] && val !== null)
            //                     ) {
            //                         return {
            //                             v: defaultText,
            //                             fc: 'rgb(68, 114, 196)',
            //                             fs: "7",
            //                             bl: true,
            //                         };
            //                     } else {
            //                         return null;
            //                     }
            //                 }

            //                 else{ 

            //                     if (
            //                         headers.length === 7 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[6] && header !== headers[7] && val !== null)
            //                         || headers.length === 9 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[8] && header !== headers[9] && val !== null)
            //                         || headers.length === 8 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[7] && header !== headers[8] && val !== null)
            //                         || headers.length === 6 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[5] && header !== headers[6] && val !== null)
            //                         || headers.length === 5 &&
            //                         (header !== headers[1] && header !== headers[2] && header !== headers[4] && header !== headers[5] && val !== null)
            //                     ) {
            //                         return {
            //                             v: defaultText,
            //                             fc: 'rgb(68, 114, 196)',
            //                             fs: "7",
            //                             bl: true,
            //                         };
            //                     } else {
            //                         return null;
            //                     }
            //                 }
            //             }
            //         }
            //         const val = sheetData[selectedIndex] ? sheetData[selectedIndex][cIndex] : null;
            //         if (val !== null) {
            //             return {
            //                 v: exText,
            //             };
            //         } else {
            //             return null;
            //         }
            //     });
            //     newRow = [...newRow, data];
            // }

            const sheetData1 = [];
            const sheetData2 = [
                ...sheetData.slice( 0, selectedIndex ),
                ...newRow,
                ...sheetData.slice( selectedIndex ),
            ];
            sheetData2.forEach( ( f, rIndex ) => {
                const data = f.filter( ( fi ) => fi != null );
                if ( data?.length > 0 )
                {
                    const iIndex = rIndex + difference;
                    f.forEach( ( val, cIndex ) => {
                        if ( val != null && ( val?.v != undefined || val?.ct?.s?.length > 0 ) )
                        {
                            const formattedVal = {
                                "r": rIndex,
                                "c": cIndex,
                                "v": {
                                    "ct": val?.ct,
                                    "m": val?.m,
                                    "v": val?.v,
                                    "fs": val?.fs,
                                    "merge": val?.merge,
                                    "fc": val?.fc,
                                    "bl": val?.bl,
                                    "bg": val?.bg,
                                    "tb": val?.tb
                                }
                            }
                            sheetData1.push( formattedVal );
                        }
                    } );
                }
            } );

            const updatedRowLength = Object.keys( updatedRowLen )?.length;
            if ( updatedRowLength > 0 )
            {
                for ( let index = 0; index <= updatedRowLength; index++ )
                {
                    if ( updatedRowLen[ index ] == undefined || updatedRowLen[ index ] == null || updatedRowLen[ index ] == 0 )
                    {
                        updatedRowLen[ index ] = 30;
                    }
                    //no need for now ******************************** by gokul********************************
                    // if ( !updatedRowLen[ index ] || updatedRowLen[ index ] < 60 )
                    // {
                    //     let rowData = luckysheet.getcellvalue( parseInt( index ) );
                    //     rowData = rowData ? rowData.filter( ( f ) => f != null ) : [];
                    //     let maxLength = 0;
                    //     let length = [];

                    //     if ( rowData?.length > 0 )
                    //     {
                    //         // let length = [];
                    //         // let maxLength = 0;
                    //         rowData.forEach( ( f ) => {
                    //             console.log( f );
                    //             if ( f?.ct?.s )
                    //             {
                    //                 if ( f?.ct?.s?.length > 1 )
                    //                 {
                    //                     var text = '';
                    //                     f?.ct?.s?.forEach( ( e ) => { text += e?.v } )
                    //                     length.push( text?.length );
                    //                 } else { length.push( f?.ct?.s[ 0 ]?.v?.length ) }
                    //             }
                    //         } );
                    //         length = Array.from( new Set( length ) );
                    //         length.forEach( ( f ) => {
                    //             if ( f > maxLength )
                    //             {
                    //                 maxLength = f;
                    //             }
                    //         } );
                    //         updatedRowLen[ index ] = maxLength > 30 ? maxLength > 100 ? maxLength / 3 + 1 : maxLength / 2 + 20 : 30;
                    //     } else
                    //     {
                    //         updatedRowLen[ index ] = 30;
                    //     }
                    // }
                }
            }
            // console.log( sheetData1 );
            configData.rowlen = updatedRowLen;

            if ( flagCheck !== 'Exclusion' )
            {
                flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'config' ] = configData : apiDataConfig.demo[ 'config' ] = configData;
                flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'celldata' ] = sheetData1 : apiDataConfig.demo[ 'celldata' ] = sheetData1;
                flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'data' ] = sheetData2 : apiDataConfig.demo[ 'data' ] = sheetData2;
                flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'rowlen' ] = updatedRowLen : apiDataConfig.demo[ 'rowlen' ] = updatedRowLen;
            }
            let sheetallconfig = luckysheet.getAllSheets();
            var configupdate = sheetallconfig.filter( f => f.name.includes( "PolicyReviewChecklist" ) );
            var formconfigupdate = sheetallconfig.filter( f => f.name.includes( "Forms Compare" ) );
            var exconfigupdate = sheetallconfig.filter( f => f.name.includes( "Exclusion" ) );

            // if (apiDataConfig.demo['config'].borderInfo == 0) {
            //     apiDataConfig.demo['config'] = configupdate.config;
            //     apiDataConfig.demo['celldata'] = configupdate.celldata;
            //     apiDataConfig.demo['data'] = configupdate.data;
            // }
           
            // if (FormCompare_appconfigdata.forms['config'].borderInfo == 0) {
            //     FormCompare_appconfigdata.forms['config'] = formconfigupdate.config;
            //     FormCompare_appconfigdata.forms['celldata'] = formconfigupdate.celldata;
            //     FormCompare_appconfigdata.forms['data'] = formconfigupdate.data;
            // }

            if ( flagCheck == 'Forms Compare' )
            {

                FormCompare_appconfigdata.forms[ 'config' ] = configData;
                FormCompare_appconfigdata.forms[ 'celldata' ] = sheetData1;
                FormCompare_appconfigdata.forms[ 'data' ] = sheetData2;
                //let formconfigupdate = luckysheet.getAllSheets()[1];


            } else if ( flagCheck == 'PolicyReviewChecklist' )
            {
                apiDataConfig.demo[ 'config' ] = configData;
                apiDataConfig.demo[ 'celldata' ] = sheetData1;
                apiDataConfig.demo[ 'data' ] = sheetData2;
                var configupdate = sheetallconfig.filter( f => f.name.includes( "PolicyReviewChecklist" ) );

                if ( apiDataConfig.demo[ 'config' ].borderInfo == 0 )
                {
                    apiDataConfig.demo[ 'config' ] = configupdate[ 0 ].config;
                    apiDataConfig.demo[ 'celldata' ] = configupdate[ 0 ].celldata;
                    apiDataConfig.demo[ 'data' ] = configupdate[ 0 ].data;
                }
            } else if ( flagCheck == 'Exclusion' )
            {
                exclusionDatafigdata.exclusion[ 'config' ] = configData;
                exclusionDatafigdata.exclusion[ 'celldata' ] = sheetData1;
                exclusionDatafigdata.exclusion[ 'data' ] = sheetData2;

                var exconfigupdate = sheetallconfig.filter( f => f.name.includes( "Exclusion" ) );
                if ( exclusionDatafigdata.exclusion[ 'config' ].borderInfo == 0 )
                {
                    exclusionDatafigdata.exclusion[ 'config' ] = exconfigupdate[ 0 ].config;
                    exclusionDatafigdata.exclusion[ 'celldata' ] = exconfigupdate[ 0 ].celldata;
                    exclusionDatafigdata.exclusion[ 'data' ] = exconfigupdate[ 0 ].data;
                }

            }

            if ( formconfigupdate.length > 0 && formconfigupdate != undefined )
            {
                if ( FormCompare_appconfigdata.forms[ 'config' ].borderInfo == 0 )
                {
                    var formconfigupdate = sheetallconfig.filter( f => f.name.includes( "Forms Compare" ) );
                    FormCompare_appconfigdata.forms[ 'config' ] = formconfigupdate[ 0 ].config;
                    FormCompare_appconfigdata.forms[ 'celldata' ] = formconfigupdate[ 0 ].celldata;
                    FormCompare_appconfigdata.forms[ 'data' ] = formconfigupdate[ 0 ].data;
                }
            }
            if ( configupdate.length > 0 && configupdate != undefined )
            {
                if ( apiDataConfig.demo[ 'config' ].borderInfo == 0 )
                {
                    var configupdate = sheetallconfig.filter( f => f.name.includes( "PolicyReviewChecklist" ) );
                    apiDataConfig.demo[ 'config' ] = configupdate[ 0 ].config;
                    apiDataConfig.demo[ 'celldata' ] = configupdate[ 0 ].celldata;
                    apiDataConfig.demo[ 'data' ] = configupdate[ 0 ].data;
                }
            }
            if ( exconfigupdate.length > 0 && exconfigupdate != undefined )
            {
                if ( exclusionDatafigdata.exclusion[ 'config' ].borderInfo == 0 )
                {
                    var exconfigupdate = sheetallconfig.filter( f => f.name.includes( "Exclusion" ) );
                    exclusionDatafigdata.exclusion[ 'config' ] = exconfigupdate[ 0 ].config;
                    exclusionDatafigdata.exclusion[ 'celldata' ] = exconfigupdate[ 0 ].celldata;
                    exclusionDatafigdata.exclusion[ 'data' ] = exconfigupdate[ 0 ].data;
                }
            }
            apiDataConfig.demo[ 'rowlen' ] = updatedRowLen;
            renderLuckySheet( false, luckysheet.getluckysheet_select_save(), false );

            // luckysheet.setConfig(configData);
            // if (apiDataConfig?.demo?.celldata?.length > 0) {
            //     const cellData = apiDataConfig?.demo?.celldata;

            //     cellData.forEach((f) => {
            //         if (f?.r !== undefined && f?.c !== undefined && Array.isArray(luckysheet.flowdata()[f?.r]) && f?.c < luckysheet.flowdata()[f?.r].length) {
            //             if (f?.v !== undefined) {
            //                 luckysheet.setcellvalue(f?.r, f?.c, luckysheet.flowdata(), f?.v);
            //             }
            //         }
            //     });

            //     // Refresh the grid after setting cell values
            //     luckysheet.jfrefreshgrid();
            // }
        } else
        {
            alert( 'Please Select only one row' );
        }
        //  }
    }

    const luckySheetInsert = ( position, noOfRecords, currentTablename, currentSheetData ) => {
        let dataToBePopulatedOnInsert = getEmptyDataSet();

        if ( currentSheetData?.name === 'PolicyReviewChecklist' )
        {
            if ( position && position > 0 && noOfRecords && noOfRecords > 0 )
            {
                const cellsToBeUpdated = [];
                for ( let index = 0; index < noOfRecords; index++ )
                {
                    luckysheet.insertRow( position );
                    if ( currentTablename != "Table 1" )
                    {
                        const currentTableData = tableColumnDetails[ currentTablename ];
                        if ( currentTableData && currentTableData?.columnNames )
                        {
                            const maxCol = Math.max( ...Object.values( currentTableData[ "columnNames" ] ) );
                            const minCol = currentTableData && currentTableData?.columnNames && currentTableData?.columnNames?.ChecklistQuestions ? currentTableData?.columnNames?.ChecklistQuestions : 2;
                            if ( maxCol && maxCol > 0 )
                            {
                                for ( let index1 = minCol; index1 < maxCol - 7; index1++ )
                                {
                                    dataToBePopulatedOnInsert[ "ct" ][ "s" ][ 0 ] = { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8", "ff": "\"Tahoma\"",
                                    };
                                    luckysheet.setcellvalue( position, index1 + 1, luckysheet.flowdata(), { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8" ,"ff": "\"Tahoma\"", } );
                                    // cellsToBeUpdated.push( { "Row": position, "Column": index1 + 1 } );          
                                }
                            }
                        }
                    }
                }
                luckysheet.jfrefreshgrid();
            }
        } else if ( currentSheetData?.name == 'Forms Compare' )
        {
            // Define datasets
            // const dataset1 = { "v": "Attached Forms", "fc": '#000000', "fs": "8" };
            // const dataset2 = { "v": "CA2", "fc": '#000000', "fs": "8" };
            // const dataset3 = { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8" };
            if ( position && position > 0 && noOfRecords && noOfRecords > 0 )
            {
                for ( let index = 0; index < noOfRecords; index++ )
                {
                    luckysheet.insertRow( position );
                    if ( currentTablename != "FormTable 1" )
                    {
                        luckysheet.setcellvalue( position, 1, luckysheet.flowdata(), { "v": "Attached Forms", "fc": '#000000', "fs": "8" , "ff": "\"Tahoma\"", } );
                        luckysheet.setcellvalue( position, 2, luckysheet.flowdata(), { "v": "CA2", "fc": '#000000', "fs": "8", "ff": "\"Tahoma\"", } );
                        luckysheet.setcellvalue( position, 3, luckysheet.flowdata(), { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8", "ff": "\"Tahoma\"", } );
                        luckysheet.setcellvalue( position, 4, luckysheet.flowdata(), { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8", "ff": "\"Tahoma\"", } );
                    }
                }
                luckysheet.jfrefreshgrid();
            }
        } else if ( currentSheetData?.name == 'Exclusion' )
        {
            if ( position && position > 0 && noOfRecords && noOfRecords > 0 )
            {
                for ( let index = 0; index < noOfRecords; index++ )
                {
                    luckysheet.insertRow( position );
                    if ( currentTablename == "ExTable 1" )
                    {
                        luckysheet.setcellvalue( position, 1, luckysheet.flowdata(), { "v": "   " } );
                        luckysheet.setcellvalue( position, 2, luckysheet.flowdata(), { "v": "   " } );
                        luckysheet.setcellvalue( position, 3, luckysheet.flowdata(), { "v": "   " } );
                        luckysheet.setcellvalue( position, 4, luckysheet.flowdata(), { "v": "   " } );
                    }
                }
                luckysheet.jfrefreshgrid();
            }
        }
    }

    // const luckySheetInsert = (position, noOfRecords, currentTablename, currentSheetData) => {
    //     const batchSize = 100; 
    //     let remainingRecords = noOfRecords;

    //     if (currentSheetData?.name === 'PolicyReviewChecklist' && position >= 0 && noOfRecords > 0  || currentSheetData?.name == 'Forms Compare' || currentSheetData?.name == 'Exclusion') {
    //         let initialPosition = position;

    //         while (remainingRecords > 0) {
    //             const currentBatchSize = Math.min(batchSize, remainingRecords);
    //             luckysheet.insertRow(position, { number: currentBatchSize });
    //             position += currentBatchSize;
    //             remainingRecords -= currentBatchSize;
    //         }
    //         luckysheet.enterEditMode();
    //        if (currentSheetData?.name != 'Exclusion') {
    //         setTimeout(() => {
    //             populatePageNumbers(initialPosition, noOfRecords, currentTablename);
    //         }, 2000);
    //        }

    //         luckysheet.jfrefreshgrid();
    //     }  
    // };

    const populatePageNumbers = (startPosition, noOfRecords, currentTablename) => {
            let sheetname = luckysheet.getSheet().name;
            if (startPosition > 0 && noOfRecords > 0 && currentTablename !== "Table 1" && sheetname == "PolicyReviewChecklist") {
                const currentTableData = tableColumnDetails[currentTablename];

                if (currentTableData && currentTableData.columnNames) {
                    const maxCol = Math.max(...Object.values(currentTableData.columnNames));
                    const minCol = currentTableData.columnNames.ChecklistQuestions || 2;

                    if (maxCol > 0) {
                        for (let rowIndex = startPosition; rowIndex < startPosition + noOfRecords; rowIndex++) {
                            for (let colIndex = minCol; colIndex < maxCol - 5; colIndex++) {

                                    luckysheet.setcellvalue( rowIndex, colIndex , luckysheet.flowdata(), { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8" } );

                            }
                        }
                    } 
                } 
                luckysheet.jfrefreshgrid();
            } 
            else if (sheetname === "Forms Compare") { 
                let startindexposition = startPosition - 1;
                for (let rowIndex = startPosition; rowIndex <= startindexposition + noOfRecords; rowIndex++) {
                        if (currentTablename !== "FormTable 1") {
                            luckysheet.setcellvalue( rowIndex, 1, luckysheet.flowdata(), { "v": "Attached Forms", "fc": '#000000', "fs": "8" } );
                        luckysheet.setcellvalue( rowIndex, 2, luckysheet.flowdata(), { "v": "CA2", "fc": '#000000', "fs": "8" } );
                        luckysheet.setcellvalue( rowIndex, 3, luckysheet.flowdata(), { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8" } );
                        luckysheet.setcellvalue( rowIndex, 4, luckysheet.flowdata(), { "v": "Page #", "fc": 'rgb(68, 114, 196)', "fs": "8" } );  
                        }
                }
            luckysheet.jfrefreshgrid();
        }
            else if (sheetname === "Exclusion") {
                let startindexposition = startPosition - 1;
                for (let rowIndex = startindexposition; rowIndex <= startindexposition + noOfRecords; rowIndex++) {
                                    if ( currentTablename == "ExTable 1" )
                                    {
                                        luckysheet.setcellvalue( rowIndex, 1, luckysheet.flowdata(), { "v": "   " } );
                                        luckysheet.setcellvalue( rowIndex, 2, luckysheet.flowdata(), { "v": "   " } );
                                        luckysheet.setcellvalue( rowIndex, 3, luckysheet.flowdata(), { "v": "   " } );
                                        luckysheet.setcellvalue( rowIndex, 4, luckysheet.flowdata(), { "v": "   " } );
                                    }

                                luckysheet.jfrefreshgrid();
                }
            luckysheet.jfrefreshgrid();
        }

    };

    const luckySheetDelete = ( selectedRow, difference ) => {
        const userSelectedRowRange = luckysheet.getRange();
        if ( userSelectedRowRange && userSelectedRowRange?.length > 0 )
        {
            const continutyArray = [];
            userSelectedRowRange.map( ( e, index ) => {
                if ( index === 0 )
                {
                    const min = e?.row[ 0 ] < e?.row[ 1 ] ? e?.row[ 0 ] : e?.row[ 1 ];
                    const max = e?.row[ 0 ] > e?.row[ 1 ] ? e?.row[ 0 ] : e?.row[ 1 ];
                    if ( min && min > 0 && max && max > 0 )
                    {
                        for ( let index = min; index <= max; index++ )
                        {
                            continutyArray.push( index );
                        }
                    }
                }
            } );
            const uniqueRow = Array.from( new Set( continutyArray ) );
            const isValidSelection = hasMissingNumbers( uniqueRow );
            if ( uniqueRow && uniqueRow?.length > 0 && isValidSelection === false )
            {
                // onDeleteUpdateTableColumnDetails( selectedRow, difference );
                // uniqueRow.forEach((f) => {
                //     luckysheet.deleteRow(uniqueRow[0],uniqueRow[0]);
                // });
                luckysheet.deleteRow( uniqueRow[ 0 ], uniqueRow[ uniqueRow?.length - 1 ] );
                luckysheet.setRangeShow( { row: [ uniqueRow[ 0 ], uniqueRow[ 0 ] ], column: [ 1, 1 ] } );
            }
        }
    }

    const hasMissingNumbers = ( arr ) => {
        arr.sort( ( a, b ) => a - b );
        for ( let i = 1; i < arr.length; i++ )
        {
            if ( arr[ i ] !== arr[ i - 1 ] + 1 )
            {
                return true;
            }
        }
        return false;
    }

    const singleMultipleSwitchDelete = () => {
        let getFlag = luckysheet.getSheet().name;
        if ( getFlag !== "Exclusion" )
        {
            const hardCodedStartingIndex = getFlag === 'Forms Compare' ? [ 0, 1, 2 ] : [ 0, 1, 2, 3 ];
            if ( ( hasMultipleRowsSelected && hardCodedStartingIndex.includes( selectedRowIndexRange[ 0 ] ) ) || hardCodedStartingIndex.includes( setectedRowIndex ) )
            {
                const msg = `cannot delete rows in the sheetname sections`;
                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                return;
            }
        } else
        {
            const hardCodedStartingIndex = getFlag === 'Exclusion' ? [ 0 ] : [ 0 ];
            if ( ( hasMultipleRowsSelected && hardCodedStartingIndex.includes( selectedRowIndexRange[ 0 ] ) ) || hardCodedStartingIndex.includes( setectedRowIndex ) )
            {
                const msg = `cannot delete rows in the sheetname sections`;
                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                return;
            }
        }

        let QacFlag = luckysheet.getSheet().name;
        if ( QacFlag != 'QAC not answered questions' )
        {
            const luckySheet = luckysheet.getSheetData()[ 1 ];
            let flagCheck = luckysheet?.getSheet()?.name;
            let sheetData = luckysheet.getSheetData();
            let configData = luckysheet.getConfig();
            if ( flagCheck == 'PolicyReviewChecklist' )
            {
                const isMultiRowSelected = hasMultipleRowsSelected;
                let currentTableRecord = "";
                let TableName = "";
                let tableNameKeys = Object.keys( tableColumnDetails );
                tableNameKeys = tableNameKeys.filter( ( f ) => f != "Table 1" );
                tableNameKeys.forEach( ( columnName ) => {
                    if ( columnName != "Table 1" && ( ( tableColumnDetails[ columnName ]?.range?.start <= setectedRowIndex && tableColumnDetails[ columnName ]?.range?.end >= setectedRowIndex ) ||
                        ( isMultiRowSelected && tableColumnDetails[ columnName ]?.range?.start <= selectedRowIndexRange[ 0 ] && tableColumnDetails[ columnName ]?.range?.end >= selectedRowIndexRange[ 1 ] ) ) && Object.keys( tableColumnDetails[ columnName ]?.columnNames )?.length > 0 )
                    {
                        currentTableRecord = tableColumnDetails[ columnName ];
                        TableName = columnName;
                    }
                } );
                if ( TableName )
                {
                    if ( TableName != "Table 1" && Object.keys( tableColumnDetails[ TableName ]?.columnNames )?.length > 0 )
                    {
                        const multipleInsertLength = TableName == "Table 3" ? 2 : 1;
                        if ( ( ( !isMultiRowSelected && tableColumnDetails[ TableName ]?.range?.end - ( tableColumnDetails[ TableName ]?.range?.start + multipleInsertLength ) ) == 1 ) ||
                            isMultiRowSelected && ( tableColumnDetails[ TableName ]?.range?.end - ( tableColumnDetails[ TableName ]?.range?.start + multipleInsertLength ) == ( ( selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ] ) + 1 ) ) )
                        {
                            const msg = `the table must have atleast one row/record's`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        }
                        if ( !( ( setectedRowIndex && !isMultiRowSelected && tableColumnDetails[ TableName ]?.range?.start + multipleInsertLength < setectedRowIndex && tableColumnDetails[ TableName ]?.range?.end >= setectedRowIndex ) ||
                            ( isMultiRowSelected && tableColumnDetails[ TableName ]?.range?.start + multipleInsertLength < selectedRowIndexRange[ 0 ] && tableColumnDetails[ TableName ]?.range?.end >= selectedRowIndexRange[ 1 ] ) ) )
                        {
                            const msg = `cannot delete rows in the ${ TableName } header sections and the table must have atleast one row/record's`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        }
                    }
                } else if ( !TableName && ( ( !isMultiRowSelected && tableColumnDetails[ "Table 1" ]?.range?.end < setectedRowIndex ) ||
                    ( isMultiRowSelected && tableColumnDetails[ "Table 1" ]?.range?.end < selectedRowIndexRange[ 1 ] ) ) )
                {
                    const msg = `Cannot delete the empty rows are used to separate the table.`;
                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                    return;
                }

                // confidence score locked cell delete prevention logic start**
                if(TableName != "Table 1" ){                    
                    const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
                    const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
                    const JobType = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "JobType");
                    const EnableARDeleteCheck = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableARDeleteCheck");
                    if(EnableConfidenceScore === "true" && JobType.toUpperCase() == "AR" && EnableLockCell == "true" && EnableARDeleteCheck == "true" && props?.enableCs && props?.enableCellLock){
                        const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); //variable to store the MinLockCellScore
                        let documentData = sessionStorage.getItem('jobDocumentData');
                        try{
                            if(typeof documentData === 'string'){
                                documentData = JSON.parse(documentData);
                            }
                            const hasEndorsementEntry = documentData?.filter((f) => f?.FileFor?.includes('Endorsement'))?.length;
                            if(hasEndorsementEntry === 0){
                                const table_col_config = tableColumnDetails[ TableName ]?.columnNames;
                                const cq_index = table_col_config["ChecklistQuestions"];
                                if(cq_index && cq_index > 0){
                                    if(isMultiRowSelected && selectedRowIndexRange.length > 0 ){
                                        const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                                        for (let row_position = setectedRowIndex; row_position <= (setectedRowIndex + (selectedDiff)); row_position++) {
                                            let isCTLocked = false;
                                            let isPTLocked = false;
                                            const row_data = luckysheet.getcellvalue(row_position);
                                            const ct_text = getText(row_data[cq_index + 1], false);
                                            const pt_text = getText(row_data[cq_index + 2], false);
                                            const cs_ct_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 1))];
                                            const cs_pt_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 2))];
                                            const cs_ct_text = getText(row_data[cs_ct_index]);
                                            const cs_pt_text = getText(row_data[cs_pt_index]);
                                            const isStpValid = getConfidenceScoreConfigStatus(props?.data?.find((f) => f.Tablename === "JobHeader")?.StpMappings, "question check" ,getText(row_data[cq_index]));
                                            if(isStpValid){
                                                if(cs_ct_text?.trim() !== "" && cs_ct_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_ct_text){
                                                    if(parseFloat(cs_ct_text) > parseFloat(MinLockCellScore)){
                                                        isCTLocked = true;
                                                    }
                                                }
                                                if(cs_pt_text?.trim() !== "" && cs_pt_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_pt_text){
                                                    if(parseFloat(cs_pt_text) > parseFloat(MinLockCellScore)){
                                                        isPTLocked = true;
                                                    }
                                                }
                                                if((isCTLocked && isPTLocked) || 
                                                (
                                                    (isCTLocked && (pt_text?.trim()?.toLowerCase() != "details not available in the document" && pt_text?.trim()?.toLowerCase() != "matched")) || 
                                                    (isPTLocked && (ct_text?.trim()?.toLowerCase() != "details not available in the document" && ct_text?.trim()?.toLowerCase() != "matched"))
                                                )){
                                                    const msg = `Cannot delete the line items at row ${row_position + 1} which CT - ${isCTLocked? "Locked" : "has variance"} and PT - ${isPTLocked ? "Locked" : "has variance"}`;
                                                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                                    return;
                                                }
                                            }
                                        }
                                    }
                                    else if(setectedRowIndex != 0){
                                        let isCTLocked = false;
                                        let isPTLocked = false;
                                        const row_data = luckysheet.getcellvalue(setectedRowIndex);
                                        const ct_text = getText(row_data[cq_index + 1], false);
                                        const pt_text = getText(row_data[cq_index + 2], false);
                                        const cs_ct_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 1))];
                                        const cs_pt_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 2))];
                                        const cs_ct_text = getText(row_data[cs_ct_index]);
                                        const cs_pt_text = getText(row_data[cs_pt_index]);
                                        const isStpValid = getConfidenceScoreConfigStatus(props?.data?.find((f) => f.Tablename === "JobHeader")?.StpMappings, "question check" ,getText(row_data[cq_index]));
                                        if(isStpValid){
                                            if(cs_ct_text?.trim() !== "" && cs_ct_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_ct_text){
                                                if(parseFloat(cs_ct_text) > parseFloat(MinLockCellScore)){
                                                    isCTLocked = true;
                                                }
                                            }
                                            if(cs_pt_text?.trim() !== "" && cs_pt_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_pt_text){
                                                if(parseFloat(cs_pt_text) > parseFloat(MinLockCellScore)){
                                                    isPTLocked = true;
                                                }
                                            }
                                            if((isCTLocked && isPTLocked) || 
                                            (
                                                (isCTLocked && (pt_text?.trim()?.toLowerCase() != "details not available in the document" && pt_text?.trim()?.toLowerCase() != "matched")) || 
                                                (isPTLocked && (ct_text?.trim()?.toLowerCase() != "details not available in the document" && ct_text?.trim()?.toLowerCase() != "matched"))
                                            )){
                                                const msg = `Cannot delete the line items at row ${setectedRowIndex + 1} which CT - ${isCTLocked? "Locked" : "has variance"} and PT - ${isPTLocked ? "Locked" : "has variance"}`;
                                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                                return;
                                            }
                                        }
                                    }
                                }
                            }
                        }catch(error){
                            console.log(error);
                        }                        
                    }
                }
                // end**

                if ( isMultiRowSelected && selectedRowIndexRange.length > 0 )
                {
                    const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                    // for (let index = 0; index <= selectedDiff; index++) {
                    //     if (selectedRowIndexRange[0] != 0 && index == selectedDiff) {
                    //         onDeleteUpdateTableColumnDetails(setectedRowIndex, selectedDiff + 1);
                    //     }
                    //     const response = deleteRow( selectedRowIndexRange[ 0 ], selectedDiff + 1, sheetData, configData );
                    //     if (response?.sheetData2) {
                    //         sheetData = response?.sheetData2;
                    //     }
                    //     if (response?.configData) {
                    //         configData = response?.configData;
                    //     }
                    // }
                    if ( selectedDiff >= 1 && sheetData && sheetData?.length > 0 )
                    {
                        // onDeleteUpdateTableColumnDetails(setectedRowIndex, selectedDiff + 1);
                        // const response = deleteRow(selectedRowIndexRange[0], selectedDiff + 1, sheetData, configData);
                        luckySheetDelete( setectedRowIndex, selectedDiff + 1 );
                        // reRenderSheetAfterDeleteLoopProcessed();
                    }
                } else
                {
                    if ( setectedRowIndex != 0 )
                    {
                        // onDeleteUpdateTableColumnDetails(setectedRowIndex, 1);
                        // const response = deleteRow(setectedRowIndex, 1, sheetData, configData);
                        luckySheetDelete( setectedRowIndex, 1 );
                    }
                    // if (response?.sheetData2) {
                    //     sheetData = response?.sheetData2;
                    // }
                    // if (response?.configData) {
                    //     configData = response?.configData;
                    // }
                    // reRenderSheetAfterDeleteLoopProcessed();
                }
            } else if ( flagCheck == 'Forms Compare' )
            {
                const isMultiRowSelected = hasMultipleRowsSelected;
                let currentTableRecord = "";
                let TableName = "";
                let tableNameKeys = Object.keys( formTableColumnDetails );
                tableNameKeys = tableNameKeys.filter( ( f ) => f != "FormTable 1" );
                tableNameKeys.forEach( ( columnName ) => {
                    if ( columnName != "FormTable 1" && ( ( formTableColumnDetails[ columnName ]?.range?.start <= setectedRowIndex && formTableColumnDetails[ columnName ]?.range?.end >= setectedRowIndex ) ||
                        ( isMultiRowSelected && formTableColumnDetails[ columnName ]?.range?.start <= selectedRowIndexRange[ 0 ] && formTableColumnDetails[ columnName ]?.range?.end >= selectedRowIndexRange[ 1 ] ) ) && Object.keys( formTableColumnDetails[ columnName ]?.columnNames )?.length > 0 )
                    {
                        currentTableRecord = formTableColumnDetails[ columnName ];
                        TableName = columnName;
                    }
                } );
                if ( TableName == "" && isMultiRowSelected == true )
                {
                    const range = luckysheet.getRange();
                    const targetrow = range[ 0 ].row[ 0 ]
                    const targetendrow = range[ 0 ].row[ 1 ]
                    let nullstartcolumncheck = luckysheet.getSheetData()[ targetrow ];
                    let nullendcolumncheck = luckysheet.getSheetData()[ targetendrow ];
                    const isAllNullstart = nullstartcolumncheck.every( element => element === null );
                    const isAllNullend = nullendcolumncheck.every( element => element === null );

                    if ( isAllNullstart == true )
                    {
                        const msg = `cannot delete rows in the ${ TableName } header sections`;
                        setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                        setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                        return;
                    } else
                    {
                        if ( isAllNullend == true )
                        {
                            const msg = `cannot delete rows in the ${ TableName } header sections`;

                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );

                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )

                            return;
                        }
                    }

                }
                if ( TableName )
                {
                    if ( TableName != "FormTable 1" && Object.keys( formTableColumnDetails[ TableName ]?.columnNames )?.length > 0 )
                    {
                        const multipleInsertLength = TableName == "FormTable 2" ? 2 : 2;
                        // if (isMultiRowSelected && (formTableColumnDetails[TableName]?.range?.end - formTableColumnDetails[TableName]?.range?.start + multipleInsertLength) == selectedRowIndexRange[1] - selectedRowIndexRange[0] + 4 || !isMultiRowSelected && (formTableColumnDetails[TableName]?.range?.end - formTableColumnDetails[TableName]?.range?.start + multipleInsertLength) == selectedRowIndexRange[1] - selectedRowIndexRange[0] + 4) {
                        //     const msg = `the table must have atleast one row/record's`;
                        //     setMsgVisible(true); setMsgClass('alert error'); setMsgText(msg);
                        //     setTimeout(() => { setMsgVisible(false); setMsgText(''); }, 3500)
                        //     return;
                        // }
                        // if (!((setectedRowIndex && !isMultiRowSelected && formTableColumnDetails[TableName]?.range?.start + multipleInsertLength < setectedRowIndex && tableColumnDetails[TableName]?.range?.end >= setectedRowIndex) ||
                        //     (isMultiRowSelected && formTableColumnDetails[TableName]?.range?.start + multipleInsertLength < selectedRowIndexRange[0] && formTableColumnDetails[TableName]?.range?.end >= selectedRowIndexRange[1]))) {
                        //     const msg = `cannot delete rows in the ${TableName} header sections and the table must have atleast one row/record's`;
                        //     setMsgVisible(true); setMsgClass('alert error'); setMsgText(msg);
                        //     setTimeout(() => { setMsgVisible(false); setMsgText(''); }, 3500)
                        //     return;
                        // }
                        var range = luckysheet.getRange();
                        let sheetselectedindex = setectedRowIndex;


                        let sheetdata = luckysheet.getSheetData();

                        if ( sheetselectedindex >= 0 && sheetselectedindex < sheetdata.length )
                        {
                            sheetdata = sheetdata.slice( sheetselectedindex );
                            sheetdata = sheetdata.map( ( element, index ) => {
                                return {
                                    [ sheetselectedindex + index + 1 ]: element
                                };
                            } );
                            sheetselectedindex = sheetselectedindex + 1;
                        }

                        // let valuebeforenull = [];

                        // for (let i = 0; i < sheetdata.length; i++) {
                        //     const obj = sheetdata[i];
                        //     const key = Object.keys(obj)[0];
                        //     const values = obj[key];
                        //     if (values.every(val => val === null)) {
                        //         valuebeforenull = sheetdata.slice(0, i);
                        //         break;
                        //     }
                        // }


                        // let lastIndex = valuebeforenull.length - 1;
                        let valuebeforenull = [];

                        for ( let i = 0; i < sheetdata.length; i++ )
                        {
                            const obj = sheetdata[ i ];
                            const key = Object.keys( obj )[ 0 ];
                            const values = obj[ key ];
                            if ( values.every( val => val === null ) )
                            {
                                valuebeforenull = sheetdata.slice( 0, i );
                                break;
                            }
                        }


                        if ( valuebeforenull.length === 0 && sheetdata.length > 0 )
                        {
                            valuebeforenull.push( sheetdata[ sheetdata.length - 1 ] );
                        }

                        let lastIndex = valuebeforenull.length - 1;

                        const keys = Object.keys( lastIndex == 0 ? valuebeforenull[ 0 ] : sheetdata[ lastIndex ] );
                        const key = parseInt( keys[ 0 ] );
                        var range = luckysheet.getRange();
                        let columnrender2 = range[ 0 ].row[ 1 ];
                        let checktableexist = key - 1;


                        if ( columnrender2 > checktableexist )
                        {
                            const msg = `cannot delete rows in the ${ TableName } header sections and the table must have atleast one row/record's`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        }

                        const headercolcheck = formTableColumnDetails[ TableName ]?.range?.start;
                        const headercolcheck1 = headercolcheck + 1;
                        const headercolcheckend = formTableColumnDetails[ TableName ]?.range?.end;
                        const headercolcheck2 = headercolcheck + 1;
                        const entiernullcheck = headercolcheck2 + 1;
                        const targetrow = range[ 0 ].row[ 0 ]
                        const targetendrow = range[ 0 ].row[ 1 ]
                        // if (targetrow == headercolcheck || targetendrow == headercolcheck2) {
                        const headervalue = headercolcheck + 2;
                        const headervalueend = headercolcheckend - 1;
                        const sheetendcheck = checktableexist;
                        if ( headercolcheck == targetrow )
                        {
                            const msg = `cannot delete rows in the ${ TableName } header sections`;

                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );

                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )

                            return;
                        } else
                        {
                            if ( headercolcheck1 == targetrow )
                            {
                                const msg = `cannot delete rows in the ${ TableName } header sections`;

                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );

                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )

                                return;
                            }
                        }

                        if ( isMultiRowSelected == true )
                        {
                            // if (headervalue == targetrow && headervalueend == targetendrow || sheetendcheck == targetendrow) {
                            if ( headervalue == targetrow && sheetendcheck == targetendrow )
                            {
                                const msg = `the table must have atleast one row/record's`;
                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                return;
                            }
                            // else{
                            //     if (sheetendcheck == targetendrow) {
                            //         const msg = `the table must have atleast one row/record's`;
                            //         setMsgVisible(true); setMsgClass('alert error'); setMsgText(msg);
                            //         setTimeout(() => { setMsgVisible(false); setMsgText(''); }, 3500)
                            //         return; 
                            //     }
                            // }
                        }


                        // }
                        if ( targetrow == targetendrow )
                        {
                            const headervalue = headercolcheck + 2;
                            const headervalueend = headercolcheckend - 1;

                            if ( headervalue == targetrow && checktableexist == targetendrow )
                            {
                                const msg = `the table must have atleast one row/record's`;
                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                return;
                            }

                        }

                        let nullstartcolumncheck = luckysheet.getSheetData()[ targetrow ];
                        let nullendcolumncheck = luckysheet.getSheetData()[ targetendrow ];
                        const isAllNullstart = nullstartcolumncheck.every( element => element === null );
                        const isAllNullend = nullendcolumncheck.every( element => element === null );

                        if ( isAllNullstart == true )
                        {
                            const msg = `cannot delete rows in the ${ TableName } header sections`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        } else
                        {
                            if ( isAllNullend == true )
                            {
                                const msg = `cannot delete rows in the ${ TableName } header sections`;

                                setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );

                                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )

                                return;
                            }
                        }




                        // if (entiernullcheck == setectedRowIndex) {
                        //     const msg = `the table must have atleast one row/record's`;
                        //     setMsgVisible(true); setMsgClass('alert error'); setMsgText(msg);
                        //     setTimeout(() => { setMsgVisible(false); setMsgText(''); }, 3500)
                        //     return;
                        // }
                        let tablestartcheck = headercolcheck + 2;
                        let tableendcheck = headercolcheckend;
                        if ( tablestartcheck == setectedRowIndex && checktableexist == targetendrow )
                        {
                            const msg = `the table must have atleast one row/record's`;
                            setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                            setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                            return;
                        }
                    }
                     // confidence score locked cell delete prevention logic start**
                    if(TableName != "FormTable 1" ){                    
                        const EnableConfidenceScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableConfidenceScore");
                        const EnableLockCell = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableLockCell");
                        const JobType = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "JobType");
                        const EnableARDeleteCheck = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "EnableARDeleteCheck");
                        if(EnableConfidenceScore === "true" && JobType.toUpperCase() == "AR" && EnableLockCell == "true" && EnableARDeleteCheck == "true" && props?.enableCs){
                            const MinLockCellScore = getConfidenceScoreConfigStatus(props?.confidenceScoreConfig, "MinLockCellScore"); //variable to store the MinLockCellScore
                            let documentData = sessionStorage.getItem('jobDocumentData');
                            try{
                                if(typeof documentData === 'string'){
                                    documentData = JSON.parse(documentData);
                                }
                                const hasEndorsementEntry = documentData?.filter((f) => f?.FileFor?.includes('Endorsement'))?.length;
                                if(hasEndorsementEntry === 0){
                                    const table_col_config = formTableColumnDetails[ TableName ]?.columnNames;
                                    const cq_index = table_col_config["ChecklistQuestions"];
                                    if(cq_index && cq_index > 0){
                                        if(isMultiRowSelected && selectedRowIndexRange.length > 0 ){
                                            const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                                            for (let row_position = setectedRowIndex; row_position <= (setectedRowIndex + (selectedDiff)); row_position++) {
                                                let isCTLocked = false;
                                                let isPTLocked = false;
                                                const row_data = luckysheet.getcellvalue(row_position);
                                                const ct_text = getText(row_data[cq_index + 1], false);
                                                const pt_text = getText(row_data[cq_index + 2], false);
                                                const cs_ct_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 1))];
                                                const cs_pt_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 2))];
                                                const cs_ct_text = getText(row_data[cs_ct_index]);
                                                const cs_pt_text = getText(row_data[cs_pt_index]);
                                                const isStpValid = getConfidenceScoreConfigStatus( props?.data?.find((f) => f.Tablename === "JobHeader")?.StpMappings, "question check" ,getText(row_data[cq_index]));
                                                if(isStpValid){
                                                    if(cs_ct_text?.trim() !== "" && cs_ct_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_ct_text){
                                                        if(parseFloat(cs_ct_text) > parseFloat(MinLockCellScore)){
                                                            isCTLocked = true;
                                                        }
                                                    }
                                                    if(cs_pt_text?.trim() !== "" && cs_pt_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_pt_text){
                                                        if(parseFloat(cs_pt_text) > parseFloat(MinLockCellScore)){
                                                            isPTLocked = true;
                                                        }
                                                    }
                                                    if((isCTLocked && isPTLocked) || 
                                                    (
                                                        (isCTLocked && (pt_text?.trim()?.toLowerCase() != "details not available in the document" && pt_text?.trim()?.toLowerCase() != "matched")) || 
                                                        (isPTLocked && (ct_text?.trim()?.toLowerCase() != "details not available in the document" && ct_text?.trim()?.toLowerCase() != "matched"))
                                                    )){
                                                        const msg = `Cannot delete the line items at row ${row_position + 1} which CT - ${isCTLocked? "Locked" : "has variance"} and PT - ${isPTLocked ? "Locked" : "has variance"}`;
                                                        setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                                        setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                                        return;
                                                    }
                                                }
                                            }
                                        }
                                        else if(setectedRowIndex != 0){
                                            let isCTLocked = false;
                                            let isPTLocked = false;
                                            const row_data = luckysheet.getcellvalue(setectedRowIndex);
                                            const ct_text = getText(row_data[cq_index + 1], false);
                                            const pt_text = getText(row_data[cq_index + 2], false);
                                            const cs_ct_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 1))];
                                            const cs_pt_index = table_col_config[getCsRespectiveColumn(getKeyByValue(table_col_config, cq_index + 2))];
                                            const cs_ct_text = getText(row_data[cs_ct_index]);
                                            const cs_pt_text = getText(row_data[cs_pt_index]);
                                            const isStpValid = getConfidenceScoreConfigStatus( props?.data?.find((f) => f.Tablename === "JobHeader")?.StpMappings, "question check" ,getText(row_data[cq_index]));
                                            if(isStpValid){
                                                if(cs_ct_text?.trim() !== "" && cs_ct_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_ct_text){
                                                    if(parseFloat(cs_ct_text) > parseFloat(MinLockCellScore)){
                                                        isCTLocked = true;
                                                    }
                                                }
                                                if(cs_pt_text?.trim() !== "" && cs_pt_text?.trim()?.toLowerCase() !== "details not available in the document" && cs_pt_text){
                                                    if(parseFloat(cs_pt_text) > parseFloat(MinLockCellScore)){
                                                        isPTLocked = true;
                                                    }
                                                }
                                                if((isCTLocked && isPTLocked) || 
                                                (
                                                    (isCTLocked && (pt_text?.trim()?.toLowerCase() != "details not available in the document" && pt_text?.trim()?.toLowerCase() != "matched")) || 
                                                    (isPTLocked && (ct_text?.trim()?.toLowerCase() != "details not available in the document" && ct_text?.trim()?.toLowerCase() != "matched"))
                                                )){
                                                    const msg = `Cannot delete the line items at row ${setectedRowIndex + 1} which CT - ${isCTLocked? "Locked" : "has variance"} and PT - ${isPTLocked ? "Locked" : "has variance"}`;
                                                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                                                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                                                    return;
                                                }
                                            }
                                        }
                                    }
                                }
                            }catch(error){
                                console.log(error);
                            }                        
                        }
                    }
                    // end**
                }
                else if ( !TableName && ( ( !isMultiRowSelected && formTableColumnDetails[ "FormTable 1" ]?.range?.end < setectedRowIndex ) ) )
                {
                    const msg = `Cannot delete the empty rows are used to separate the table.`;
                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                    return;
                }
                if ( isMultiRowSelected && selectedRowIndexRange.length > 0 )
                {
                    const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                    // for (let index = 0; index <= selectedDiff; index++) {
                    //     if (selectedRowIndexRange[0] != 0 && index == selectedDiff) {
                    //         onDeleteUpdateTableColumnDetails(setectedRowIndex, selectedDiff + 1);
                    //     }
                    //     const response = deleteRow(selectedRowIndexRange[0], index, sheetData, configData);
                    //     if (response?.sheetData2) {
                    //         sheetData = response?.sheetData2;
                    //     }
                    //     if (response?.configData) {
                    //         configData = response?.configData;
                    //     }
                    // }
                    // reRenderSheetAfterDeleteLoopProcessed();

                    if ( selectedDiff >= 1 && sheetData && sheetData?.length > 0 )
                    {
                        // onDeleteUpdateTableColumnDetails(setectedRowIndex, selectedDiff + 1);
                        // const response = deleteRow(selectedRowIndexRange[0], selectedDiff + 1, sheetData, configData);
                        luckySheetDelete( selectedRowIndexRange[ 0 ], selectedDiff + 1 );
                        // reRenderSheetAfterDeleteLoopProcessed();
                    }
                } else
                {
                    if ( setectedRowIndex != 0 )
                    {
                        // onDeleteUpdateTableColumnDetails(setectedRowIndex, 1);
                        // const response = deleteRow(setectedRowIndex, 1, sheetData, configData);
                        luckySheetDelete( setectedRowIndex, 1 );
                    }
                    // if (response?.sheetData2) {
                    //     sheetData = response?.sheetData2;
                    // }
                    // if (response?.configData) {
                    //     configData = response?.configData;
                    // }
                    // reRenderSheetAfterDeleteLoopProcessed();
                }
            } else if ( flagCheck == 'Exclusion' )
            {
                const isMultiRowSelected = hasMultipleRowsSelected;
                let currentTableRecord = "";
                let TableName = "";
                let tableNameKeys = Object.keys( exTableColumnDetails );
                let Table = tableNameKeys[0];
                tableNameKeys.forEach( ( columnName ) => {
                    if ( ( ( exTableColumnDetails[ columnName ]?.range?.start <= setectedRowIndex && exTableColumnDetails[ columnName ]?.range?.end >= setectedRowIndex ) ||
                        ( isMultiRowSelected && exTableColumnDetails[ columnName ]?.range?.start <= selectedRowIndexRange[ 0 ] && exTableColumnDetails[ columnName ]?.range?.end >= selectedRowIndexRange[ 1 ] ) ) && Object.keys( exTableColumnDetails[ columnName ]?.columnNames )?.length > 0 )
                    {   
                        currentTableRecord = exTableColumnDetails[ columnName ];
                        TableName = columnName;
                    }
                } );
           
                let range = luckysheet.getRange();

                if ( Table == "ExTable 1" && Object.keys( exTableColumnDetails[ Table ]?.columnNames )?.length > 0 ){
                    const rangeStartMatches = exTableColumnDetails[Table]?.range?.start + 1 === range[0].row[0];
                    const rangeEndMatches = exTableColumnDetails[Table]?.range?.end === range[0].row[1];
                    
                    if (rangeStartMatches && rangeEndMatches) {
                        const msg = `Not allowed to delete all the rows in the table`;
                        setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( msg );
                        setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                        return;
                    } 
                }

                if (Table == "ExTable 1" && props?.enableExclusionCellLock) {
                    let lockCellFlag = false;
                    const cs_colum = exTableColumnDetails[Table]?.columnNames;
                    const row_data = luckysheet.getcellvalue(setectedRowIndex);
                    const cs_column_value = getText(row_data[cs_colum["ConfidenceScore"]], false);
                    if (cs_column_value?.trim() !== "" && cs_column_value?.trim()?.toLowerCase() !== "details not available in the document" && cs_column_value) {
                        if (parseFloat(cs_column_value) >= 80) {
                            lockCellFlag = true;
                        }
                        if (lockCellFlag ||
                            (
                            (lockCellFlag && (cs_column_value?.trim()?.toLowerCase() != "details not available in the document" )) 
                            )) {
                            const msg = `Cannot delete the line items at row ${setectedRowIndex + 1} which ${lockCellFlag ? "is Locked" : "has variance"}`;
                            setMsgVisible(true); setMsgClass('alert error'); setMsgText(msg);
                            setTimeout(() => { setMsgVisible(false); setMsgText(''); }, 3500)
                            return;
                        }
                    }
                }

                if ( isMultiRowSelected && selectedRowIndexRange.length > 0 )
                {
                    const selectedDiff = selectedRowIndexRange[ 1 ] - selectedRowIndexRange[ 0 ];
                    // for (let index = 0; index <= selectedDiff; index++) {
                    //     if (selectedRowIndexRange[0] != 0 && index == selectedDiff) {
                    //         onDeleteUpdateTableColumnDetails(setectedRowIndex, selectedDiff + 1);
                    //     }
                    //     const response = deleteRow(selectedRowIndexRange[0], 1, sheetData, configData);
                    //     if (response?.sheetData) {
                    //         sheetData = response?.sheetData;
                    //     }
                    //     if (response?.configData) {
                    //         configData = response?.configData;
                    //     }
                    // }
                    if ( selectedDiff >= 1 && sheetData && sheetData?.length > 0 )
                    {
                        luckySheetDelete( setectedRowIndex, selectedDiff + 1 );
                    }
                } else
                {
                    if ( setectedRowIndex != 0 )
                    {
                        luckySheetDelete( setectedRowIndex, 1 );
                    }
                }
            }
        }

    }

    const onDeleteUpdateTableColumnDetails = ( setectedRowIndex, difference ) => {
        let flagCheck = luckysheet.getSheet().name; //formscompare

        if ( flagCheck !== 'Exclusion' )
        {
            const tableNameKeysBackup = flagCheck == 'Forms Compare' ? formTableColumnDetails : tableColumnDetails;
            const tableNameKeys = Object.keys( tableNameKeysBackup );
            tableNameKeys.forEach( ( columnName ) => {
                if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start > setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) && Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.start -= difference;
                    tableNameKeysBackup[ columnName ].range.end -= difference;
                } else if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start < setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) &&
                    Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.end -= difference;
                }
            } );
            //console.log( "after", tableNameKeysBackup );
            if ( flagCheck == 'Forms Compare' )
            {
                setFormTableColumnDetails( tableNameKeysBackup )
            } else
            {
                setTableColumnDetails( tableNameKeysBackup );
            }
        } else
        {
            const tableNameKeysBackup = exTableColumnDetails;
            const tableNameKeys = Object.keys( tableNameKeysBackup );
            tableNameKeys.forEach( ( columnName ) => {
                if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start > setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) && Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.start -= difference;
                    tableNameKeysBackup[ columnName ].range.end -= difference;
                } else if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start < setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) &&
                    Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.end -= difference;
                }
            } );
            setExTableColumnDetails( tableNameKeysBackup );
        }
    }

    const onInsertUpdateTableColumnDetails = ( setectedRowIndex, difference ) => { //for deleteion undo functionality
        let flagCheck = luckysheet.getSheet().name; //formscompare

        if ( flagCheck !== 'Exclusion' )
        {
            const tableNameKeysBackup = flagCheck == 'Forms Compare' ? formTableColumnDetails : tableColumnDetails;
            const tableNameKeys = Object.keys( tableNameKeysBackup );
            tableNameKeys.forEach( ( columnName ) => {
                if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start > setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) && Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.start += difference;
                    tableNameKeysBackup[ columnName ].range.end += difference;
                } else if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start < setectedRowIndex && (tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex || tableNameKeysBackup[ columnName ]?.range?.end >= (setectedRowIndex - 1)) ) ) &&
                    Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.end += difference;
                }
            } );
            //console.log( "after", tableNameKeysBackup );
            if ( flagCheck == 'Forms Compare' )
            {
                setFormTableColumnDetails( tableNameKeysBackup )
            } else
            {
                setTableColumnDetails( tableNameKeysBackup );
            }
        } else
        {
            const tableNameKeysBackup = exTableColumnDetails;
            const tableNameKeys = Object.keys( tableNameKeysBackup );
            tableNameKeys.forEach( ( columnName ) => {
                if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start > setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) && Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.start += difference;
                    tableNameKeysBackup[ columnName ].range.end += difference;
                } else if ( ( ( tableNameKeysBackup[ columnName ]?.range?.start < setectedRowIndex && tableNameKeysBackup[ columnName ]?.range?.end >= setectedRowIndex ) ) &&
                    Object.keys( tableNameKeysBackup[ columnName ]?.columnNames )?.length > 0 )
                {
                    tableNameKeysBackup[ columnName ].range.end += difference;
                }
            } );
            setExTableColumnDetails( tableNameKeysBackup );
        }
    }

    const saveReset = async () => {
        const luckySheet = luckysheet.getSheetData()[ 1 ];

        document.body.classList.add( 'loading-indicator' );
        const Token = await processAndUpdateToken( token );//to validate and update the token
        token = Token;
        var token = sessionStorage.getItem( 'token' );

        const headers = {
            'Authorization': `Bearer ${ Token }`
        };
        const jobId = props?.selectedJob;

        document.body.classList.add( 'loading-indicator' );
        axios.post( baseUrl + '/api/ProcedureData/RegeneratedChecklist', { jobId }, { headers } )
            .then( response => {
                if ( response.status !== 200 )
                {
                    throw new Error( `HTTP error! Status: ${ response.status }` );
                }
                return response.data;
            } )
            .then( data => {
                if ( data?.status == 400 )
                {
                    setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( data?.title );
                    setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
                } else
                {

                    setMsgVisible( true );
                    setMsgClass( 'alert success' );
                    setMsgText( 'Saved and Regenerated Successfully' );
                    setTimeout( () => {
                        setMsgVisible( false );
                        setMsgText( '' );
                    }, 3500 );
                }
            } )
            .finally( () => {
                setTimeout( () => {
                    document.body.classList.remove( 'loading-indicator' );
                }, 200 )
            } );

    }


    const deleteRow = ( selectedIndex, loopDIndex, sheetData, configData ) => {
        const luckySheet = luckysheet.getSheetData()[ 1 ];
        let flagCheck = luckysheet.getSheet()?.name;
        // if (flagCheck != 'FORM COMPARE') {
        if ( selectedIndex != null && selectedIndex != '' && selectedIndex != undefined )
        {
            // Add condition to ensure there are more than one row before attempting deletion
            if ( sheetData.length > 1 && loopDIndex > 0 )
            {
                // Update borderInfo
                let sheet1BackupData = sheetData;
                for ( let index = 0; index < loopDIndex; index++ )
                {
                    configData.borderInfo = configData?.borderInfo?.map( e => {
                        if ( e?.rangeType === "cell" && e?.value?.row_index >= selectedIndex )
                        {
                            e.value.row_index -= 1;
                        } else if ( e?.rangeType === "range" && e?.range?.length > 0 && e?.range[ 0 ]?.row?.length > 0 )
                        {
                            const newRange = e.range[ 0 ].row.map( row => ( row >= selectedIndex ? row - 1 : row ) );
                            e.range[ 0 ].row[ 0 ] = newRange[ 0 ];
                            e.range[ 0 ].row[ 1 ] = newRange[ 1 ];
                        }
                        return e;
                    } );

                    // Row height adjustment
                    const updatedRowLen = {};
                    const rowlen = Object.keys( configData.rowlen );
                    rowlen.forEach( ( f ) => {
                        const convertedRowLen = parseInt( f );
                        if ( convertedRowLen >= selectedIndex )
                        {
                            updatedRowLen[ `${ convertedRowLen - 1 }` ] = configData.rowlen[ f ] >= 90 ? 60 : configData.rowlen[ f ];
                        } else
                        {
                            updatedRowLen[ f ] = configData.rowlen[ f ] >= 90 ? 50 : configData.rowlen[ f ];
                        }
                    } );

                    // Merge cells adjustment
                    const merge = {};
                    const mergeKeys = Object.keys( configData?.merge );
                    if ( mergeKeys?.length > 0 )
                    {
                        mergeKeys.map( ( k ) => {
                            const findRow = parseInt( k.split( '_' )[ 0 ] );
                            const findCol = parseInt( k.split( '_' )[ 1 ] );
                            if ( findRow > selectedIndex )
                            {
                                const existingData = configData?.merge[ k ];
                                existingData.r = findRow - 1;
                                merge[ `${ findRow - 1 }` + '_' + findCol ] = existingData;
                            } else
                            {
                                merge[ k ] = configData?.merge[ k ];
                            }
                        } );
                        configData.merge = merge;
                    }
                    // const sheetData1 = [];
                    // Data deletion
                    const sheetData2 = [
                        ...sheetData.slice( 0, selectedIndex ),
                        ...sheetData.slice( selectedIndex + 1 ),
                    ];

                    const sheetData1 = sheetData2.reduce( ( acc, row, rIndex ) => {
                        const formattedRow = row?.filter( val => val != null && ( val?.v !== undefined || val?.ct?.s?.length > 0 ) )
                            .map( ( val, cIndex ) => ( {
                                r: rIndex,
                                c: cIndex,
                                v: {
                                    ct: val?.ct,
                                    m: val?.m,
                                    v: val?.v,
                                    fs: val?.fs,
                                    merge: val?.merge,
                                    fc: val?.fc,
                                    bl: val?.bl,
                                    bg: val?.bg,
                                    tb: val?.tb
                                }
                            } ) );
                        return [ ...acc, ...formattedRow ];
                    }, [] );
                    // let totalSheetRow = [];
                    // sheetData1.forEach((e) => {
                    //     if (e?.r != null || e?.r != undefined) {
                    //         if (!totalSheetRow.includes(e.r)) {
                    //             totalSheetRow.push(e?.r);
                    //         }
                    //     }
                    // });
                    // totalSheetRow = Array.from(new Set(totalSheetRow));
                    // const totalLength = Object.keys(updatedRowLen);
                    // for (let index = 0; index < totalLength?.length; index++) {
                    //     if (updatedRowLen[index] == undefined || updatedRowLen[index] == null || updatedRowLen[index] == 0 || !totalSheetRow?.includes(index)) {
                    //         updatedRowLen[index] = 15;
                    //     }
                    // }
                    configData[ 'rowlen' ] = updatedRowLen;


                    sheetData = sheetData2;
                    sheet1BackupData = sheetData1
                }
                if ( flagCheck !== 'Exclusion' )
                {
                    flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'config' ] = configData : apiDataConfig.demo[ 'config' ] = configData;
                    flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'celldata' ] = sheet1BackupData : apiDataConfig.demo[ 'celldata' ] = sheet1BackupData;
                    flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'data' ] = sheetData : apiDataConfig.demo[ 'data' ] = sheetData;
                    // flagCheck == 'Forms Compare' ? FormCompare_appconfigdata.forms[ 'rowlen' ] = updatedRowLen : apiDataConfig.demo[ 'rowlen' ] = updatedRowLen;
                }
                let sheetallconfig = luckysheet.getAllSheets();
                var configupdate = sheetallconfig.filter( f => f.name.includes( "PolicyReviewChecklist" ) );
                var formconfigupdate = sheetallconfig.filter( f => f.name.includes( "Forms Compare" ) );
                var exconfigupdate = sheetallconfig.filter( f => f.name.includes( "Exclusion" ) );
                if ( flagCheck == 'Exclusion' )
                {
                    exclusionDatafigdata.exclusion[ 'config' ] = configData;
                    exclusionDatafigdata.exclusion[ 'celldata' ] = sheet1BackupData;
                    exclusionDatafigdata.exclusion[ 'data' ] = sheetData;
                    // let exconfigupdate = luckysheet.getAllSheets()[2];

                    if ( exclusionDatafigdata?.exclusion[ 'config' ]?.borderInfo == 0 )
                    {
                        var exconfigupdate = sheetallconfig.filter( f => f.name.includes( "Exclusion" ) );
                        exclusionDatafigdata.exclusion[ 'config' ] = exconfigupdate[ 0 ].config;
                        exclusionDatafigdata.exclusion[ 'celldata' ] = exconfigupdate[ 0 ].celldata;
                        exclusionDatafigdata.exclusion[ 'data' ] = exconfigupdate[ 0 ].data;
                    }
                }


                if ( configupdate != undefined && configupdate.length > 0 )
                {
                    if ( apiDataConfig.demo[ 'config' ].borderInfo == 0 )
                    {
                        var configupdate = sheetallconfig.filter( f => f.name.includes( "PolicyReviewChecklist" ) );
                        apiDataConfig.demo[ 'config' ] = configupdate[ 0 ].config;
                        apiDataConfig.demo[ 'celldata' ] = configupdate[ 0 ].celldata;
                        apiDataConfig.demo[ 'data' ] = configupdate[ 0 ].data;
                    }
                }

                if ( formconfigupdate != undefined && formconfigupdate.length > 0 )
                {
                    if ( FormCompare_appconfigdata.forms[ 'config' ].borderInfo == 0 )
                    {
                        var formconfigupdate = sheetallconfig.filter( f => f.name.includes( "Forms Compare" ) );
                        FormCompare_appconfigdata.forms[ 'config' ] = formconfigupdate[ 0 ].config;
                        FormCompare_appconfigdata.forms[ 'celldata' ] = formconfigupdate[ 0 ].celldata;
                        FormCompare_appconfigdata.forms[ 'data' ] = formconfigupdate[ 0 ].data;
                    }
                }

                if ( exconfigupdate != undefined && exconfigupdate?.length > 0 )
                {
                    if ( exclusionDatafigdata.exclusion[ 'config' ].borderInfo == 0 )
                    {
                        var exconfigupdate = sheetallconfig.filter( f => f.name.includes( "Exclusion" ) );
                        exclusionDatafigdata.exclusion[ 'config' ] = exconfigupdate[ 0 ].config;
                        exclusionDatafigdata.exclusion[ 'celldata' ] = exconfigupdate[ 0 ].celldata;
                        exclusionDatafigdata.exclusion[ 'data' ] = exconfigupdate[ 0 ].data;
                    }
                }
                return { sheetData, configData };
            }
        } else
        {
            alert( 'Please select only one row' );
        }

    }

    //on delete multiple loop handled --by Gokul--
    const reRenderSheetAfterDeleteLoopProcessed = () => {
        renderLuckySheet( false, luckysheet.getluckysheet_select_save(), true );
    }

    const UpdateHCheck = async (needLoader) => {
        if(needLoader){document.body.classList.add( 'loading-indicator' );}
        const Token = await processAndUpdateToken( token );//to validate and update the token
        token = Token;
        const headers = {
            'Authorization': `Bearer ${ Token }`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${ baseUrl }/api/Defaultdatum/GetAllJobRData?jobId=${ jobId }&isPhNeeded=true`;

        try
        {
            const response = await axios( {
                method: "GET",
                url: apiUrl,
                headers: headers
            } );
            if ( response.status !== 200 )
            {
                throw new Error( `HTTP error! Status: ${ response.status }` );
            }

            return response.data;
        } catch ( error )
        {
            // console.error( 'Error:', error );
            throw error; // Rethrow the error to be caught in the calling function
        } finally
        {
            if(needLoader){document.body.classList.remove( 'loading-indicator' );}
        }
    }

    const updataPHProcess = async ( isRegenerate, needLoader ) => {
        // document.body.classList.add( 'loading-indicator' );
        const data = await UpdateHCheck(needLoader);
        const processedData = await PageHighlighterProcess( data, jobId );
        if ( data && processedData != null && processedData != undefined )
        {
            const userName = sessionStorage.getItem( 'userName' );
            const Token = await processAndUpdateToken( token );//to validate and update the token
            token = Token;
            const data = {
                "tblchecklistPagenumberHighlighter": processedData,
                "JobId": jobId,
                "UserName": userName
            }
            const headers = {
                'Authorization': `Bearer ${ Token }`,
                "Content-Type": "application/json",
            };
            const apiUrl = `${ baseUrl }/api/Defaultdatum/ResetPageHighlighter`;

            try
            {
                const response = await axios( {
                    method: "POST",
                    url: apiUrl,
                    headers: headers,
                    data
                } );
                if ( response.status !== 200 )
                {
                    throw new Error( `HTTP error! Status: ${ response.status }` );
                }
                if ( isRegenerate )
                {
                    saveReset();
                }
                setMsgVisible( true ); setMsgClass( 'alert success' ); setMsgText( 'Page Highlighter Data Updated' );
                setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 );
                return response.data;
            } catch ( error )
            {
                // console.error( 'Error:', error );
                throw error; // Rethrow the error to be caught in the calling function
            } finally
            {
                if ( needLoader && !isRegenerate )
                {
                    setTimeout( () => {
                        document.body.classList.remove( 'loading-indicator' );
                    }, 200 )
                }
            }

            // axios( {
            //     method: "POST",
            //     url: apiUrl,
            //     headers: headers,
            //     data
            // } )
            //     .then( response => {
            //         if ( response.status !== 200 )
            //         {
            //             throw new Error( `HTTP error! Status: ${ response.status }` );
            //         }
            //         if ( isRegenerate )
            //         {
            //             saveReset();
            //         }
            //         return response.data;
            //     } )
            //     .then( data => {
            //         if ( data?.status == 400 )
            //         {
            //             setMsgVisible( true ); setMsgClass( 'alert error' ); setMsgText( data?.title );
            //             setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
            //         } else
            //         {
            //             setMsgVisible( true ); setMsgClass( 'alert success' ); setMsgText( 'Data Updated' );
            //             setTimeout( () => { setMsgVisible( false ); setMsgText( '' ); }, 3500 )
            //         }
            //     } )
            //     .finally( () => {
            //         document.body.classList.remove( 'loading-indicator' );
            //     } );
        }
    }

    const autoUpdateCtPt = ( row, column, newValue ) => {
        // const luckySheet = luckysheet.getSheetData()[1];
        // let flagCheck = luckySheet[1].m;
        // if (flagCheck != 'FORM COMPARE') {
        const rowData = luckysheet.getcellvalue( row );
        const tabelDetails = tableColumnDetails;
        const formTabelDetails = formTableColumnDetails;
        let tableName = '';
        const Keys = Object.keys( tabelDetails );
        const formKeys = Object.keys( formTabelDetails );
        const lucky = luckysheet.getSheet()?.name
        const newlucky = lucky;
        let columnData = {};
        if ( newlucky == 'PolicyReviewChecklist' )
        {
            Keys.map( ( e ) => {
                const tableData = tabelDetails[ e ];
                if ( tableData?.range?.start <= row && tableData?.range?.end >= row ) { tableName = e }
            } );
            columnData = tabelDetails[ tableName ]?.columnNames
        } else if ( newlucky == 'Forms Compare' )
        {
            formKeys.map( ( e ) => {
                const tableData = formTabelDetails[ e ];
                if ( tableData?.range?.start <= row && tableData?.range?.end >= row ) { tableName = e }
            } );
            columnData = formTabelDetails[ tableName ]?.columnNames
        }
        //this isARCheckData model data is ****important**** ------by gokul
        let isARCheckData = { isAR: true, presentColumns: [] };
        // const columnData = tabelDetails[tableName]?.columnNames || formTabelDetails[tableName]?.columnNames;
        validateEndorsementEntry( rowData, columnData, tableName, jobId, token ).then( ( res ) => {
            if ( res )
            {
                const colMapText = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' ];
                let endorsementMsgText = 'Invalid Endorsement entry in ';
                endorsementMsgText += `Row - ${ row + 1 } `
                setMsgVisible( true ); setMsgClass( 'alert success' );
                setOpenDialog( true );
                setMsgText( endorsementMsgText );
                setTimeout( () => {
                    setMsgVisible( false );
                    setMsgText( '' );
                }, 4500 );
                const message = endorsementMsgText;
            }
        } );
        if ( tableName == "Table 2" || tableName == "Table 3" )
        {
            isARCheckData = isARType( columnData, tableName );
            isARCheckData.isAR = false;
            if ( columnData?.CurrentTermPolicy <= column && columnData?.Observation > column )
            {
                observationColumnChange( isARCheckData, tableName, row, column, rowData, columnData, "currentTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicy, columnData?.PriorTermPolicy, columnData?.Application, columnData?.Schedule, columnData?.Quote, columnData?.Proposal, columnData?.Binder );
            }
            // else if ( columnData?.PriorTermPolicy === column )
            // {
            //     observationColumnChange( isARCheckData, tableName, row, column, rowData, columnData, "priorTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicy, columnData?.PriorTermPolicy );
            // }
        } else if ( tableName == "Table 4" || tableName == "Table 5" )
        {
            const keys = getIndexForForms( columnData )
            const column1 = keys[ "column1" ];
            const column2 = keys[ "column2" ];
            isARCheckData = isARType( columnData, tableName );
            isARCheckData.isAR = false;
            if ( 3 <= column && columnData?.Observation > column )
            {
                observationColumnChange( isARCheckData, tableName, row, column, rowData, columnData, "currentTerm", tableName == "Table 2" ? 2 : 3, columnData[ column1 ], columnData[ column2 ] );
            }
            // if ( columnData?.CurrentTermPolicyListed === column )
            // {
            //     observationColumnChange( tableName , row, column, rowData, columnData, "currentTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicyListed, columnData?.PriorTermPolicyListed );
            // } else if ( columnData?.PriorTermPolicyListed === column )
            // {
            //     observationColumnChange( tableName , row, column, rowData, columnData, "priorTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicyListed, columnData?.PriorTermPolicyListed );
            // }
        }
        // else if ( tableName == "Table 5" || tableName == "Table 6"){
        //     if ( columnData?.CurrentTermPolicyListed === column )
        //     {
        //         observationColumnChange( tableName , row, column, rowData, columnData, "currentTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicyListed, columnData?.CurrentTermPolicyAttached );
        //     } else if ( columnData?.CurrentTermPolicyAttached === column )
        //     {
        //         observationColumnChange( tableName, row, column, rowData, columnData, "priorTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicyListed, columnData?.CurrentTermPolicyAttached );
        //     }
        // } 
        else if ( tableName == "FormTable 2" || tableName == "FormTable 3" )
        {
            if ( columnData?.CurrentTermPolicyAttached === column || columnData?.PriorTermPolicyAttached === column )
            {
                observationColumnChange( isARCheckData, tableName, row, column, rowData, columnData, "currentTerm", tableName == "FormTable 2" ? 2 : 3, columnData?.CurrentTermPolicyAttached
                    , columnData?.PriorTermPolicyAttached );
            }
            // else if ( columnData?.CurrentTermPolicyAttached === column )
            // {
            //     observationColumnChange( isARCheckData, tableName, row, column, rowData, columnData, "priorTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicyAttached, columnData?.CurrentTermPolicyListed );
            // }
        }
        else if ( tableName == "Table 6" || tableName == "Table 7" )
        {
            if ( columnData?.CurrentTermPolicyListed === column || columnData?.CurrentTermPolicyAttached === column )
            {
                observationColumnChange( isARCheckData, tableName, row, column, rowData, columnData, "currentTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicyAttached, columnData?.CurrentTermPolicyListed );
            }
            // else if ( columnData?.CurrentTermPolicyAttached === column )
            // {
            //     observationColumnChange( isARCheckData, tableName, row, column, rowData, columnData, "priorTerm", tableName == "Table 2" ? 2 : 3, columnData?.CurrentTermPolicyAttached, columnData?.CurrentTermPolicyListed );
            // }
        }
        // }
    }

    const observationColumnChange = ( isARCheckData, tableName, row, column, rowData, columnData, returnString, tableIndex, currentTermIndex, priorTermIndex, QuoteIndex, ApplicationIndex, ScheduleIndex, ProposalIndex, BinderIndex ) => {
        if ( !isARCheckData?.isAR )
        {
            const isAllPresentColumnsAreValid = true;
            if ( isARCheckData?.presentColumns?.length > 0 )
            {
                const matechedColumns = [];
                const dnaColumns = [];
                const needProcessColumns = [];
                const emptyDataColumns = [];
                const presentColumns = isARCheckData?.presentColumns;
                const currentEditedColumnText = getText( rowData[ column ], false );
                // if ( currentEditedColumnText?.includes( "MATCHED" ) || currentEditedColumnText?.toLowerCase()?.includes( "matched" ) )
                // {
                //     return;
                // }
                presentColumns.forEach( ( pColumn ) => {
                    let text = getText( rowData[ columnData[ pColumn ] ], false );
                    if ( text )
                    {
                        text = text?.replace( /\s+/g, ' ' );//replace more than single space with single space only
                        if ( text?.includes( "MATCHED" ) || text?.toLowerCase()?.includes( "matched" ) )
                        {
                            matechedColumns.push( pColumn );
                        } else if ( text?.toLowerCase()?.includes( "details not available" ) || text?.toLowerCase()?.includes( "details not available in the document" ) )
                        {
                            dnaColumns.push( pColumn );
                        } else
                        {
                            needProcessColumns.push( pColumn );
                        }
                    } else
                    {
                        emptyDataColumns.push( pColumn );
                    }
                } );
                if ( presentColumns?.length == ( matechedColumns?.length + dnaColumns?.length ) )
                {
                    setCellValue( row, columnData?.Observation, getEmptyDataSet() );
                    setCellValue( row, columnData?.PageNumber, getEmptyDataSet() );
                    return;
                }
                if ( presentColumns?.length == dnaColumns?.length )
                {
                    setCellValue( row, columnData?.Observation, getEmptyDataSet() );
                    setCellValue( row, columnData?.PageNumber, getEmptyDataSet() );
                    return;
                }
                if ( emptyDataColumns?.length > 0 )
                {
                    // need to handle this scenorio if needed
                }
                if ( needProcessColumns?.length > 0 || dnaColumns?.length > 0 )
                {
                    let obervationText = '';
                    let observationDataSet = getEmptyDataSet();
                    needProcessColumns.forEach( ( npColumn ) => {
                        const columnText = getText( rowData[ columnData[ npColumn ] ], false );
                        const columnText1 = getText( rowData[ columnData[ npColumn ] ], true );
                        const columnTextDuplicate = getText( rowData[ columnData[ npColumn ] ], true );
                        const observationKey = getObservationKey( npColumn, tableName );
                        let columnSplitText = '';
                        if ( tableName == "Table 6" || tableName == "Table 7" )
                        {
                            columnSplitText = splitPageKekFromText( columnText1, npColumn );
                        } else if ( columnTextDuplicate?.toLowerCase()?.includes( 'endorsement page' ) || columnTextDuplicate?.toLowerCase()?.includes( 'current policy endorsement attached' ) ||
                            columnTextDuplicate?.toLowerCase()?.includes( 'current policy endorsement listed' ) )
                        {
                            columnSplitText = splitPageKekFromText( columnText1, 'all' );
                        } else
                        {
                            columnSplitText = splitPageKekFromText( columnText1, '' );
                        }
                        obervationText += ( observationKey + columnSplitText + " " );
                    } );
                    dnaColumns.forEach( ( dnaColumn ) => {
                        const observationKey = getObservationKey( dnaColumn, tableName );
                        obervationText += ( observationKey + "NO RECORDS" + " " );
                    } );
                    observationDataSet.ct.s[ 0 ].v = obervationText;
                    setCellValue( row, columnData?.Observation, observationDataSet );
                }
                let index1 = 0;
                let index2 = 0;
                let splitKeyForCp = '';
                let splitKeyForPp = '';
                let currentTermText = '';
                let priorTermText = '';
                const applicationKeys = getTableApplicationColumns( tableName );
                for ( let index = 0; index < applicationKeys.length; index++ )
                {
                    if ( index1 === 0 )
                    {
                        const key = getKeysForMRType( needProcessColumns, dnaColumns, emptyDataColumns, matechedColumns, applicationKeys[ index ] );
                        if ( key )
                        {
                            index1 = columnData[ key ];
                            splitKeyForCp = key;
                            currentTermText = getText( rowData[ index1 ], true );
                        }
                    } else if ( index2 === 0 )
                    {
                        const key = getKeysForMRType( needProcessColumns, dnaColumns, emptyDataColumns, matechedColumns, applicationKeys[ index ] );
                        if ( key )
                        {
                            index2 = columnData[ key ];
                            splitKeyForPp = key;
                            priorTermText = getText( rowData[ index2 ], true );
                        }
                    } else
                    {
                        break;
                    }
                }
                updatePageColummn( tableName, row, column, rowData, columnData, tableIndex, index1, index2, splitKeyForCp, splitKeyForPp, currentTermText, priorTermText );
                // console.log("process done");
            }
        } else
        {
            let text = '';
            const cellData = rowData[ column ];
            //valid field data check
            let currentTermText = getText( rowData[ currentTermIndex ] );
            let applicationTermText = getText( rowData[ ApplicationIndex ] );
            let quoteTermText = getText( rowData[ QuoteIndex ] );
            let scheduleTermText = getText( rowData[ ScheduleIndex ] );
            let BinderTermText = getText( rowData[ BinderIndex ] );
            let ProposalTermText = getText( rowData[ ProposalIndex ] );
            let hasCPNorecords = false;
            let priorTermText = getText( rowData[ priorTermIndex ] );
            let haPPNorecords = false;
            let currentTermText1 = getText( rowData[ currentTermIndex ], true );
            let priorTermText1 = getText( rowData[ priorTermIndex ], true );
            const otherApplications = getOtherApplications( columnData );
            const hasOtherApplications = otherApplications?.length > 0;
            if ( currentTermText && ( currentTermText.includes( "MATCHED" ) || currentTermText.toLowerCase().includes( "details not available" ) || currentTermText?.toLowerCase()?.includes( "matched" ) ) || applicationTermText && ( applicationTermText.includes( "MATCHED" ) || applicationTermText.toLowerCase().includes( "details not available" ) || applicationTermText?.toLowerCase()?.includes( "matched" ) ) )
            {
                hasCPNorecords = true;
            }
            if ( priorTermText && ( priorTermText.includes( "MATCHED" ) || priorTermText.toLowerCase().includes( "details not available" ) || priorTermText?.toLowerCase()?.includes( "matched" ) ) || quoteTermText && ( quoteTermText.includes( "MATCHED" ) || quoteTermText.toLowerCase().includes( "details not available" ) || quoteTermText?.toLowerCase()?.includes( "matched" ) ) || scheduleTermText && ( scheduleTermText.includes( "MATCHED" ) || scheduleTermText.toLowerCase().includes( "details not available" ) || scheduleTermText?.toLowerCase()?.includes( "matched" ) ) )
            {
                haPPNorecords = true;
            }
            if ( hasCPNorecords && haPPNorecords )
            {
                if ( ( ( currentTermText.toLowerCase().includes( "details not available" ) || currentTermText.toLowerCase().includes( "details not available in the document" ) ) ||
                    ( priorTermText.toLowerCase().includes( "details not available" ) || priorTermText.toLowerCase().includes( "details not available in the document" ) ) ||
                    ( scheduleTermText.toLowerCase().includes( "details not available" ) || scheduleTermText.toLowerCase().includes( "details not available in the document" ) ) ||
                    ( applicationTermText.toLowerCase().includes( "details not available" ) || applicationTermText.toLowerCase().includes( "details not available in the document" ) ) ||
                    ( quoteTermText.toLowerCase().includes( "details not available" ) || quoteTermText.toLowerCase().includes( "details not available in the document" ) ) ||
                    ( BinderTermText.toLowerCase().includes( "details not available" ) || BinderTermText.toLowerCase().includes( "details not available in the document" ) ) ||
                    ( ProposalTermText.toLowerCase().includes( "details not available" ) || ProposalTermText.toLowerCase().includes( "details not available in the document" ) )
                ) ||
                    ( ( currentTermText.toLowerCase().includes( "MATCHED" ) || currentTermText.toLowerCase().includes( "matched" ) ) ||
                        ( priorTermText.toLowerCase().includes( "MATCHED" ) || priorTermText.toLowerCase().includes( "matched" ) ) ||
                        ( scheduleTermText.toLowerCase().includes( "MATCHED" ) || scheduleTermText.toLowerCase().includes( "matched" ) ) ||
                        ( applicationTermText.toLowerCase().includes( "MATCHED" ) || applicationTermText.toLowerCase().includes( "matched" ) ) ||
                        ( quoteTermText.toLowerCase().includes( "MATCHED" ) || quoteTermText.toLowerCase().includes( "matched" ) ) ||
                        ( BinderTermText.toLowerCase().includes( "MATCHED" ) || BinderTermText.toLowerCase().includes( "matched" ) ) ||
                        ( ProposalTermText.toLowerCase().includes( "MATCHED" ) || ProposalTermText.toLowerCase().includes( "matched" ) )
                    ) )
                {
                    setCellValue( row, columnData?.Observation, getEmptyDataSet() );
                    setCellValue( row, columnData?.PageNumber, getEmptyDataSet() );
                }
                return;
            }
            // if ( ( returnString == "currentTerm" && hasCPNorecords ) || ( returnString == "priorTerm" && haPPNorecords ) ){
            //     return;
            // }
            if ( cellData && cellData?.ct?.s?.length > 0 )
            {
                text = getText( cellData );
            } else if ( cellData?.m || cellData?.v )
            {
                text = cellData?.m || cellData?.v;
            }
            const key = text.includes( "Page" ) ? "Page" : text.includes( "page" ) ? "page" : "";
            if ( key )
            {
                text = text.split( key )[ 0 ].replace( /\r\n/g, '' );
            }
            let splitKeyForCp = '';
            let splitKeyForPp = '';
            let splitKeyForCpl = '';
            let splitKeyForCpa = '';
            const columnDataKeys = Object.keys( columnData );
            columnDataKeys.forEach( ( key ) => {
                if ( columnData[ key ] === currentTermIndex )
                {
                    splitKeyForCp = key;
                } else if ( columnData[ key ] === priorTermIndex )
                {
                    splitKeyForPp = key;
                }
            } );
            let isCpaAtInitial = false;
            if ( text )
            {
                let currentTermContent = hasCPNorecords ? "NO RECORDS" : splitPageKekFromText( currentTermText1, splitKeyForCp );
                let priorTermContent = haPPNorecords ? "NO RECORDS" : splitPageKekFromText( priorTermText1, splitKeyForPp );
                if ( !currentTermContent )
                {
                    currentTermContent = "NO RECORDS"
                }
                if ( !priorTermContent )
                {
                    priorTermContent = "NO RECORDS"
                }
                let trimmedText = '';
                if ( tableName == "Table 6" || tableName == "Table 7" )
                {
                    const dataSet = state;
                    const filteredData = dataSet.filter( ( f ) => f?.Tablename === tableName );
                    let lobData = Array.from( new Set( filteredData[ 0 ]?.TemplateData.filter( ( f ) => f?.PolicyLob ).map( ( e ) => e?.PolicyLob ) ) );
                    lobData = lobData.filter( ( f ) => f != undefined && f != null );
                    if ( lobData?.length > 0 )
                    {
                        trimmedText = lobData[ 0 ]?.replace( /\s+/g, '' );
                    }
                }
                let ObservationData = rowData[ columnData?.Observation ];
                if ( ObservationData && Object.keys( ObservationData )?.length > 0 && ObservationData?.ct?.s && ObservationData?.ct?.s?.length > 0 )
                {
                    if ( ObservationData?.ct?.s?.length > 0 || !ObservationData?.ct?.s && ( ObservationData?.v || ObservationData?.m ) )
                    {
                        //replacerSection **gokul**
                        if ( ( tableName == "Table 6" || tableName == "Table 7" ) && trimmedText )
                        {
                            ObservationData.ct.s = [ ObservationData.ct.s[ 0 ] ];
                            if ( splitKeyForPp == "CurrentTermPolicyAttached" && trimmedText?.toLocaleLowerCase()?.includes( 'attached,listed' ) || trimmedText?.toLocaleLowerCase()?.includes( 'attachedlisted' ) )
                            {
                                isCpaAtInitial = true;
                                ObservationData.ct.s[ 0 ].v = getObservationKey( splitKeyForPp, tableName ) + priorTermContent + " " + getObservationKey( splitKeyForCp, tableName ) + currentTermContent;
                            } else if ( splitKeyForCp == "CurrentTermPolicyAttached" && trimmedText?.toLocaleLowerCase()?.includes( 'attached,listed' ) || trimmedText?.toLocaleLowerCase()?.includes( 'attachedlisted' ) )
                            {
                                isCpaAtInitial = true;
                                ObservationData.ct.s[ 0 ].v = getObservationKey( splitKeyForCp, tableName ) + currentTermContent + " " + getObservationKey( splitKeyForPp, tableName ) + priorTermContent;
                            } else if ( splitKeyForCp == "CurrentTermPolicyListed" && trimmedText?.toLocaleLowerCase()?.includes( 'listed,attached' ) || trimmedText?.toLocaleLowerCase()?.includes( 'listedattached' ) )
                            {
                                ObservationData.ct.s[ 0 ].v = getObservationKey( splitKeyForCp, tableName ) + currentTermContent + " " + getObservationKey( splitKeyForPp, tableName ) + priorTermContent;
                            } else if ( splitKeyForPp == "CurrentTermPolicyListed" && trimmedText?.toLocaleLowerCase()?.includes( 'listed,attached' ) || trimmedText?.toLocaleLowerCase()?.includes( 'listedattached' ) )
                            {
                                ObservationData.ct.s[ 0 ].v = getObservationKey( splitKeyForPp, tableName ) + priorTermContent + " " + getObservationKey( splitKeyForCp, tableName ) + currentTermContent;
                            } else
                            {
                                ObservationData.ct.s[ 0 ].v = getObservationKey( splitKeyForCp, tableName ) + currentTermContent + " " + getObservationKey( splitKeyForPp, tableName ) + priorTermContent;
                            }
                        } else
                        {
                            ObservationData.ct.s[ 0 ].v = getObservationKey( splitKeyForCp, tableName ) + currentTermContent + " " + getObservationKey( splitKeyForPp, tableName ) + priorTermContent;
                        }
                        // console.log(ObservationData);
                        setCellValue( row, columnData?.Observation, ObservationData )
                    } else
                    {
                        //more than 1 array need to handle
                    }
                } else
                {
                    //if the obervation has no ct or s[] is empty
                    let observationDataSet = getEmptyDataSet(); //if the cell has no set structure
                    if ( ( tableName == "Table 6" || tableName == "Table 7" ) && splitKeyForPp == "CurrentTermPolicyAttached" )
                    {
                        observationDataSet.ct.s[ 0 ].v = getObservationKey( splitKeyForPp, tableName ) + priorTermContent + " " + getObservationKey( splitKeyForCp, tableName ) + currentTermContent;
                    } else
                    {
                        observationDataSet.ct.s[ 0 ].v = getObservationKey( splitKeyForCp, tableName ) + currentTermContent + " " + getObservationKey( splitKeyForPp, tableName ) + priorTermContent;
                    }
                    setCellValue( row, columnData?.Observation, observationDataSet )
                }
            }
            updatePageColummn( tableName, row, column, rowData, columnData, tableIndex, currentTermIndex, priorTermIndex, splitKeyForCp, splitKeyForPp, currentTermText1, priorTermText1, isCpaAtInitial );
        }
    }

    const getKeysForMRType = ( needProcessColumns, dnaColumns, emptyDataColumns, matechedColumns, key ) => {
        if ( needProcessColumns?.length > 0 && needProcessColumns?.includes( key ) )
        {
            return key;
        }
        if ( dnaColumns?.length > 0 && dnaColumns?.includes( key ) )
        {
            return key;
        }
        if ( emptyDataColumns?.length > 0 && emptyDataColumns?.includes( key ) )
        {
            return key;
        }
        if ( matechedColumns?.length > 0 && matechedColumns?.includes( key ) )
        {
            return key;
        }

    }

    const updatePageColummn = ( tableName, row, column, rowData, columnData, tableIndex, currentTermIndex, priorTermIndex, splitKeyForCp, splitKeyForPp, currentTermText, priorTermText, isCpaAtInitial ) => {
        let pageData = rowData[ columnData?.PageNumber ];
        if ( !pageData || Object.keys( pageData )?.length == 0 )
        {
            pageData = getEmptyDataSet();
        }
        if ( !pageData?.ct?.s || pageData?.ct?.s?.length == 0 )
        {
            pageData = getEmptyDataSet();
        }
        if ( tableIndex == 2 || tableIndex == 3 )
        {
            gatherAndUpdatePageColumn( tableName, row, pageData, rowData, columnData, currentTermIndex, priorTermIndex, splitKeyForCp, splitKeyForPp, currentTermText, priorTermText, isCpaAtInitial );
        }
    }

    const gatherAndUpdatePageColumn = ( tableName, row, pageData, rowData, columnData, currentTermIndex, priorTermIndex, splitKeyForCp, splitKeyForPp, currentTermText, priorTermText, isCpaAtInitial ) => {
        const questionCode = getTextByRequirement( getText( rowData[ columnData?.ChecklistQuestions ] ), "question" );
        let sText = '';
        if ( splitKeyForCp && currentTermIndex > 0 && tableName != "Table 2" && tableName != "Table 3" && ( questionCode?.toLowerCase()?.includes( "ca" ) || questionCode?.toLowerCase()?.includes( "cl" ) ) )
        {
            sText = getText( rowData[ columnData[ splitKeyForCp ] ], true );
        }
        let hasCp = currentTermIndex > 0 ? getPageKey( splitKeyForCp, tableName, sText ) : null;
        let sText1 = '';
        if ( splitKeyForPp && priorTermIndex > 0 && tableName != "Table 2" && tableName != "Table 3" && ( questionCode?.toLowerCase()?.includes( "ca" ) || questionCode?.toLowerCase()?.includes( "cl" ) ) )
        {
            sText1 = getText( rowData[ columnData[ splitKeyForPp ] ], true );
        }
        let hasPp = priorTermIndex > 0 ? getPageKey( splitKeyForPp, tableName, sText1 ) : null;
        let endrosementList = getTableApplicationColumns( "endorsement" );
        const CpPageNo = currentTermText?.toLowerCase()?.includes( "matched" ) ? getExistingPageKey( pageData, hasCp.trim() + questionCode.trim() ) : getTextByRequirement( getText( rowData[ currentTermIndex ] ), "getPage", splitKeyForCp );
        const PpPageNo = priorTermText?.toLowerCase()?.includes( "matched" ) ? getExistingPageKey( pageData, hasPp.trim() + questionCode.trim() ) : getTextByRequirement( getText( rowData[ priorTermIndex ] ), "getPage", splitKeyForPp );
        // const CpPageNo = currentTermText?.toLowerCase()?.includes( "matched" ) ? existingPageCode( pageData, hasCp.trim() + questionCode.trim(), 0, hasPp.trim() + questionCode.trim()  ) :getTextByRequirement( getText( rowData[ currentTermIndex ] ), "getPage", splitKeyForCp );
        // const PpPageNo = priorTermText?.toLowerCase()?.includes( "matched" ) ? existingPageCode( pageData, hasPp.trim() + questionCode.trim(),1,'' ) : getTextByRequirement( getText( rowData[ priorTermIndex ] ), "getPage", splitKeyForPp );
        if ( pageData.ct.s?.length > 1 )
        {
            const firstIndex = pageData.ct.s.slice( 0, 1 );
            pageData.ct.s = firstIndex;
        }
        if ( tableName == "Table 2" || tableName == "Table 3" || tableName == "Table 4" || tableName == "Table 5" )
        {
            let hasSeenEDfCP = false;
            let hasSeenEDfPP = false;
            endrosementList = endrosementList?.filter( ( f ) => f?.toLocaleLowerCase() != "page" );
            endrosementList.forEach( ( f ) => {
                if ( !hasSeenEDfCP && currentTermText && currentTermText?.includes( f ) )
                {
                    hasCp += 'E';
                    hasSeenEDfCP = true;
                }
                if ( !hasSeenEDfPP && priorTermText && priorTermText?.includes( f ) )
                {
                    hasPp += 'E';
                    hasSeenEDfPP = true;
                }
            } )
        }
        let pageDataOfV = '';
        if ( ( tableName == "Table 6" || tableName == "Table 7" ) && ( questionCode?.toLowerCase()?.includes( "ca" ) || isCpaAtInitial ) )
        {
            if ( hasCp == "Cpa" || hasCp == "CpEa" )
            {
                pageDataOfV = hasCp + questionCode + ":" + CpPageNo + " \r\n" + hasPp + questionCode + ":" + PpPageNo;
            } else
            {
                pageDataOfV = hasPp + questionCode + ":" + PpPageNo + " \r\n" + hasCp + questionCode + ":" + CpPageNo;
            }
        } else if ( ( tableName == "Table 6" || tableName == "Table 7" ) && ( !questionCode?.toLowerCase()?.includes( "ca" ) || !isCpaAtInitial ) )
        {
            if ( hasCp != "Cpa" && hasCp != "CpEa" )
            {
                pageDataOfV = hasCp + questionCode + ":" + CpPageNo + " \r\n" + hasPp + questionCode + ":" + PpPageNo;
            } else
            {
                pageDataOfV = hasPp + questionCode + ":" + PpPageNo + " \r\n" + hasCp + questionCode + ":" + CpPageNo;
            }
        } else
        {
            pageDataOfV = hasCp + questionCode + ":" + CpPageNo + " \r\n" + hasPp + questionCode + ":" + PpPageNo;
        }
        const noOfApplications = columnData?.Observation - 3;
        if ( noOfApplications > 2 )
        {
            const keys = Object.keys( columnData );
            const presentedColumns = [];
            const allApplicationColumns = getTableApplicationColumns( tableName );
            keys.map( ( key ) => {
                if ( columnData[ key ] && columnData[ key ] >= 3 && allApplicationColumns.includes( key ) && key != splitKeyForCp && key != splitKeyForPp )
                {
                    presentedColumns.push( key );
                }
            } );
            if ( presentedColumns?.length > 0 )
            {
                presentedColumns.forEach( ( key ) => {
                    let code = getPageKey( key, tableName, '' );
                    let pageNo = '';
                    const text = getText( rowData[ columnData[ key ] ], true );
                    const lowerCasedText = text?.toLowerCase();
                    if ( lowerCasedText && lowerCasedText?.includes( 'matched' ) )
                    {
                        pageNo = getExistingPageKey( pageData, code.trim() + questionCode.trim() );
                    } else
                    {
                        pageNo = getTextByRequirement( text, "getPage", key );
                    }
                    let hasSeenEDfC = false;
                    endrosementList.forEach( ( f ) => {
                        if ( !hasSeenEDfC && text && text?.includes( f ) )
                        {
                            code += 'E';
                            hasSeenEDfC = true;
                        }
                    } )
                    pageDataOfV += " \r\n" + code + questionCode + ":" + pageNo;
                } );
            }
        }
        pageData.ct.s[ 0 ].v = pageDataOfV;
        setCellValue( row, columnData?.PageNumber, pageData );
    }

    const existingPageCode = ( data, key, index, secondKey ) => {
        const text = getText( data );
        let pageData = key ? text.split( key ) : null;
        if ( index == 0 && secondKey )
        {
            pageData = pageData[ 1 ]?.replace( /\s/g, '' )?.split( secondKey )[ 0 ];
            const pageNumber = pageData?.match( /\d+/g );
            if ( Array.isArray( pageNumber ) && pageNumber?.length == 1 )
            {
                return pageNumber[ 0 ];
            }
        }
        if ( index != 0 && pageData )
        {
            const pageNumber = pageData[ pageData?.length - 1 ].match( /\d+/g );
            if ( Array.isArray( pageNumber ) && pageNumber?.length == 1 )
            {
                return pageNumber[ 0 ];
            }
        }
        return "NO RECORDS";
    }

    const setCellValue = ( row, column, data ) => {
        // luckysheet.setCellValue(row,column,data);
        luckysheet.setcellvalue( row, column, luckysheet.flowdata(), data );
        luckysheet.jfrefreshgrid();
    }

    const handleDialogClose = ( e ) => {
        if ( e == false )
        {
            setOpenDialog( e );
        }
    }

    const toggleFindDialog = () => {
        setfindDialog( !findDialog );
    };

    const toggleDropDialog = () => {
        setDropDialog( !dropDialog );
    };

    const handleIconClick = () => {
        const currentSheetData = luckysheet.getSheet();
        if (currentSheetData?.name === 'PolicyReviewChecklist' || currentSheetData?.name === 'Red' || currentSheetData?.name === 'Green') {
            toggleFilterDialog();
        }
    }
    const toggleFilterDialog = () => {
        setOpenFilterDialog( !openFilterDialog );
    };


    const dropDialogClose = ( e ) => {
        let range = luckysheet.getRange();
        let selectedIndex = range[ 0 ].row[ 0 ];
        let tabledata = tableColumnDetails;
        const excludedColumns = [ "Id", "JobId", "CreatedOn", "UpdatedOn", "Columnid", "IsDataForSp" ];
        const selectedTable = findTableForIndex( selectedIndex, tabledata, excludedColumns );

        let actionColumnTable = tableColumnDetails[ selectedTable ];
        let values = Object.values( actionColumnTable.columnNames );
        let largestIndex = Math.max( ...values );
        let Actioncolumnindex = largestIndex - 3;
        let Requestcolumnindex = largestIndex - 2;
        let Notescolumnindex = largestIndex - 1;
        let r = range[ 0 ].row[ 0 ];
        let c = range[ 0 ].column[ 0 ];

        if ( e?.hasData )
        {
            if ( e.selectedOption1 != null && e.selectedOption2 != null && e.selectedOption3 != null )
            {
                if ( c == Actioncolumnindex )
                {
                    setCellValue( r, c, e.selectedOption1.text );
                    setCellValue( r, c + 1, e.selectedOption2.text );
                    setCellValue( r, c + 2, e.selectedOption3.text );
                } else if ( c == Requestcolumnindex )
                {
                    setCellValue( r, c - 1, e.selectedOption1.text );
                    setCellValue( r, c, e.selectedOption2.text );
                    setCellValue( r, c + 1, e.selectedOption3.text );
                } else if ( c == Notescolumnindex )
                {
                    setCellValue( r, c - 2, e.selectedOption1.text );
                    setCellValue( r, c - 1, e.selectedOption2.text );
                    setCellValue( r, c, e.selectedOption3.text );
                }
            } else
            {
                // Handle each option separately if not all are selected
                if ( e && e.selectedOption1 && e.selectedOption1 != null )
                {
                    let actionText = e.selectedOption1.text;
                    // let foundMatch = false;
                    // let matchText = autoActionTxt.includes(actionText);
                    // if (matchText) {
                    //         foundMatch = true;
                    setCellValue( r, Actioncolumnindex, actionText );
                    // }
                    luckysheet.exitEditMode();
                }
                if ( e && e.selectedOption2 && e.selectedOption2 != null )
                {
                    let requestText = e.selectedOption2.text;
                    // let foundMatch = false;
                    // let matchText = autoRequestTxt.includes(requestText);
                    //     if (matchText) {
                    //         foundMatch = true;
                    setCellValue( r, Requestcolumnindex, requestText );
                    // }
                    luckysheet.exitEditMode();
                }
                if ( e && e.selectedOption3 && e.selectedOption3 != null )
                {
                    let notesText = e.selectedOption3.text;
                    // let foundMatch = false;
                    // let matchText = autoNotesTxt.includes(notesText);
                    //     if (matchText) {
                    //         foundMatch = true;
                    setCellValue( r, Notescolumnindex, notesText );
                    // }
                    luckysheet.exitEditMode();
                }
            }
        }
        setDropDialog( false );
    };
    luckysheet.exitEditMode();

    const handleInputDialogClose = ( e ) => {
        if ( e?.input && e?.input > 0 )
        {
            insertFnByInputDialog( e?.input );
            setOpenInputDialog( false );
        } else
        {
            setOpenInputDialog( false );
        }
    }

    const findDialogClose = ( e ) => {
        // if ( e == true )
        // {
            setfindDialog( false );
        // }
    }

    const handleFilterDialogClose = ( e ) => {
        console.log( e );
        setOpenFilterDialog( false );
        if ( e?.filterData )
        {
            setFilterSelectionData( e?.filterData );
        } else
        {
            setFilterSelectionData( null );
        }
    }

    const clearDuplicateRecords = () => {
        document.body.classList.add( 'loading-indicator' );
        setTimeout( () => {
            const sheetDetails = luckysheet.getSheet();
            if ( sheetDetails?.name === 'PolicyReviewChecklist' )
            {
                const tableDetails = tableColumnDetails;
                if ( tableDetails )
                {
                    const keys = Object.keys( tableDetails );
                    if ( keys && keys?.length > 0 )
                    {
                        const questionIndexTobeRemoved = [];

                        keys.forEach( ( key ) => {
                            const keyData = tableDetails[ key ];
                            const rangeDetails = keyData?.range;
                            if ( keyData && rangeDetails?.start && rangeDetails?.end && key != 'Table 1' )
                            {
                                const QuestionSet = [];
                                for ( let index = rangeDetails?.start + ( key == 'Table 3' ? 3 : 2 ); index <= rangeDetails?.end; index++ )
                                {
                                    const sheetData = luckysheet.getSheetData();
                                    const data = sheetData[ index ];
                                    if ( data && data?.length > 0 && keyData?.columnNames?.ChecklistQuestions && data[ keyData?.columnNames?.ChecklistQuestions ] )
                                    {
                                        const questionData = data[ keyData?.columnNames?.ChecklistQuestions ];
                                        const textedData = getText( questionData );
                                        if ( textedData && typeof textedData === 'string' )
                                        {
                                            const questionCode = textedData?.trim()?.slice( 0, 3 );

                                            if ( !QuestionSet?.includes( questionCode.toUpperCase() ) )
                                            {
                                                QuestionSet.push( questionCode.toUpperCase() );
                                            } else if ( QuestionSet?.includes( questionCode.toUpperCase() ) )
                                            {
                                                questionIndexTobeRemoved.push( index );
                                                // luckysheet.deleteRow( index, index );
                                            }
                                        }
                                    }
                                }
                            }
                        } );
                        if ( questionIndexTobeRemoved?.length > 0 )
                        {
                            let grouppedNumbersSet = groupNumbers( questionIndexTobeRemoved );
                            if ( grouppedNumbersSet && grouppedNumbersSet.length > 0 )
                            {
                                grouppedNumbersSet = grouppedNumbersSet.reverse();
                                grouppedNumbersSet.forEach( groupset => {
                                    if ( groupset && groupset?.length > 0 )
                                    {
                                        luckysheet.setRangeShow( {
                                            "row": [
                                                groupset[ 0 ] - 1,
                                                groupset[ groupset?.length - 1 ] - 1
                                            ],
                                            "column": [
                                                1,
                                                1
                                            ]
                                        } );
                                        luckysheet.deleteRow( groupset[ 0 ], groupset[ groupset?.length - 1 ] );
                                    }
                                } );
                            }
                            // descOrderedRIndex.forEach((rIndex) => {
                            //     luckysheet.setRangeShow( {
                            //         "row": [
                            //             rIndex,
                            //             rIndex
                            //         ],
                            //         "column": [
                            //             1,
                            //             1
                            //         ]
                            //     } );
                            //         luckysheet.deleteRow( rIndex, rIndex );
                            //     });
                        }
                    }
                }
            } else if ( sheetDetails?.name === 'Forms Compare' )
            {
                const tableDetails = formTableColumnDetails;
                if ( tableDetails )
                {
                    const keys = Object.keys( tableDetails );
                    if ( keys && keys?.length > 0 )
                    {
                        const questionIndexTobeRemoved = [];

                        keys.forEach( ( key ) => {
                            const keyData = tableDetails[ key ];
                            const rangeDetails = keyData?.range;
                            if ( keyData && rangeDetails?.start && rangeDetails?.end && key != 'FormTable 1' )
                            {
                                const QuestionSet = [];
                                for ( let index = rangeDetails?.start + 2; index <= rangeDetails?.end; index++ )
                                {
                                    const sheetData = luckysheet.getSheetData();
                                    const data = sheetData[ index ];
                                    if ( data && data?.length > 0 && data[ 2 ] )
                                    {
                                        const questionData = data[ 2 ];
                                        const textedData = getText( questionData );
                                        if ( textedData && typeof textedData === 'string' )
                                        {
                                            const questionCode = textedData?.trim()?.slice( 0, 3 );

                                            if ( !QuestionSet?.includes( questionCode.toUpperCase() ) )
                                            {
                                                QuestionSet.push( questionCode.toUpperCase() );
                                            } else if ( QuestionSet?.includes( questionCode.toUpperCase() ) )
                                            {
                                                questionIndexTobeRemoved.push( index );
                                                // luckysheet.deleteRow( index, index );
                                            }
                                        }
                                    }
                                }
                            }
                        } );
                        if ( questionIndexTobeRemoved?.length > 0 )
                        {
                            let grouppedNumbersSet = groupNumbers( questionIndexTobeRemoved );
                            if ( grouppedNumbersSet && grouppedNumbersSet.length > 0 )
                            {
                                grouppedNumbersSet = grouppedNumbersSet.reverse();
                                grouppedNumbersSet.forEach( groupset => {
                                    if ( groupset && groupset?.length > 0 )
                                    {
                                        luckysheet.setRangeShow( {
                                            "row": [
                                                groupset[ 0 ] - 1,
                                                groupset[ groupset?.length - 1 ] - 1
                                            ],
                                            "column": [
                                                1,
                                                1
                                            ]
                                        } );
                                        luckysheet.deleteRow( groupset[ 0 ], groupset[ groupset?.length - 1 ] );
                                    }
                                } );
                            }
                        }
                    }
                }
            }
            document.body.classList.remove( 'loading-indicator' );
        }, 100 );
    }

    const groupNumbers = ( data ) => {
        data = data.sort( ( a, b ) => a - b )
        const groupedData = [];

        if ( data.length === 0 )
        {
            return;
        }

        let currentGroup = [ data[ 0 ] ];

        for ( let i = 1; i < data.length; i++ )
        {
            if ( data[ i ] === data[ i - 1 ] || data[ i ] === data[ i - 1 ] + 1 )
            {
                currentGroup.push( data[ i ] );
            } else
            {
                groupedData.push( currentGroup );
                currentGroup = [ data[ i ] ];
            }
        }
        groupedData.push( currentGroup );
        return groupedData;
    };

   
    const handleSheetChange = async (e) => {
        setIssavessheet(true);
        sessionStorage.setItem("IsAutoUpdate",false);
        const value = e?.target?.value;
        const sheets = luckysheet.getSheet().name;
        if(sheets == 'PolicyReviewChecklist'){
         await onUpdateClick( false, true,false );
        }
        else if(sheets == 'Forms Compare'){
            await formCompareUpdate(true, false, false);
            GridBackupSave();
            // await onUpdateClick( false, true,false );
        }else if(sheets == 'Exclusion'){
            await onUpdateClick( false, true );
        }
        setTimeout( () => { 
            setSelectedSheet(value);
            selectChange(value);
        }, 3000 );
       
    }; 
      
    return (
        <div>
            { msgVisible &&
                <div className="alert-container">
                    <div className={ msgClass }>{ msgText }</div>
                </div>
            }
            {/* <div className="toggle">
                <Toggle
                    label="AutoSave "
                    onText="On"
                    offText="Off"
                    styles={{
                        root: {
                            selectors: {
                                '.ms-Toggle-thumb': {
                                    width: 7,
                                    height: 8,
                                },
                                '.ms-Toggle-stateText': {
                                    display: 'none',
                                    color: 'green',
                                },
                            },
                        },
                        label: {
                            fontSize: '7.5px',
                            marginTop: '5px',
                        },
                        description: {
                            fontSize: '14px',
                            color: 'green',
                        },
                    }}
                // onChange={(e, checked) => {
                //     if (checked) {
                //         setApiCallStatus(false);
                //     }
                // }}
                />
            </div> */}
            <div className="p2">
                <PrimaryButton className="luckySheet_header_button" onClick={ () => onUpdateClick( false, true, false ) }>Save</PrimaryButton>
                <PrimaryButton className="luckySheet_header_button" onClick={ () => singleMultipleSwitchInsert( true ) }>Insert Row</PrimaryButton>
                {/* <PrimaryButton className="luckySheet_header_button" onClick={() => exclusionUpdate(false, true)}>Exclusion Update</PrimaryButton> */ }
                <PrimaryButton className="luckySheet_header_button" onClick={ () => { singleMultipleSwitchDelete(); } }>Delete Row</PrimaryButton>
                <PrimaryButton className="luckySheet_header_button" onClick={ () => Regenrateclick( true ) } >Save and Regenerate</PrimaryButton>
                <PrimaryButton className="luckySheet_header_button" onClick={ () => Exportclick( true ) } >Export</PrimaryButton>
                { ( sessionStorage.getItem( 'userName' ) == 'ramu_s@exdion.com' || sessionStorage.getItem( 'userName' ) == 'ganesh_sriramu@exdion.com' || sessionStorage.getItem( 'userName' ) == 'sandeep_kumar@exdion.com' ) && <PrimaryButton className="luckySheet_header_button" onClick={ () => clearDuplicateRecords() } >Clean</PrimaryButton> }
                {/* <Dropdown placeholder="Select a sheet" options={sheetsDropOption} selectedKey={selectedSheet} onChange={handleSheetChange} 
                 styles={customStyles}
                /> */}
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
                { openDialog && <DialogComponent isOpen={ openDialog } onClose={ ( e ) => handleDialogClose( e ) } message={ msgText } /> }
                { findDialog && <FindDialogComponent isOpen={ {findDialog} } luckySheet={ luckysheet } sheetState={ sheetState } onClose={ ( e ) => findDialogClose( e ) } message={ msgText } /> }
                { dropDialog && <DiscrepancyOptionsDialogComponent isOpen={ dropDialog } luckySheet={ luckysheet } state={ [ { JobId: jobId } ]} onClose={ ( e ) => dropDialogClose( e ) } message={ "Action On Descrepancy (from AMs)" } /> }
                { openInputDialog && <InputDialogComponent isOpen={ openInputDialog } onClose={ ( e ) => handleInputDialogClose( e ) } /> }
                { openFilterDialog && <FilterDialogComponent isOpen={ { openFilterDialog, tableColumnDetails, luckysheet, filterSelectionData } } onClose={ ( e ) => handleFilterDialogClose( e ) } /> }
            </div>
            <div style={{ position: 'relative' }}>
                <Icon iconName="Filter"
                onClick={handleIconClick}
                  style={{
                    position: 'absolute',
                    top: '2px',
                    right: '-350px',
                    fontSize: '16.2px',
                    margin: '5px',
                    zIndex: 10,
                    cursor: 'pointer'
                }} /> 
                
                <h6 style={{
                    position: 'absolute',
                    top: '5px',
                    fontWeight: 500,
                    right: '-380px',
                    fontSize: '12.5px',
                    margin: '5px',
                    zIndex: 10,
                    cursor: 'pointer'
                }}
                  onClick={handleIconClick}
                 >Filter</h6>
                <div className="App" id="luckysheet" ref={luckyCss}></div>

            </div>
           
            <SimpleSnackbarWithOutTimeOut ref={ container } />
        </div>
    );
}
