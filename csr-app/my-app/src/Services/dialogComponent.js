import React, { useState, useEffect, useRef } from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { DetailsList, DetailsListLayoutMode, SelectionMode, CheckboxVisibility } from '@fluentui/react/lib/DetailsList';
import { SpinButton, Position, ComboBox } from '@fluentui/react';
import axios from "axios";
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { processAndUpdateToken, updateAllSheetsResultArray, CsrSaveHistoryApiCall, CsrPendingReport, brokerIdsGetData } from '../Services/CommonFunctions';
import '../dialog.css'; // You can define your CSS for styling
import { baseUrl } from '../Services/Constants';
import { TooltipHost } from '@fluentui/react';
import { getTextWithoutAnyChnages, findTblRowAllIndex } from './CommonFunctions';
import { SimpleSnackbar } from '../Components/SnackBar';
import { updateGridAuditLog } from '../Services/PreviewChecklistDataService';


const dialogContentProps = {
    type: DialogType.normal,
    title: 'Notification',
    subText: 'Something went wrong ... !',
};

let token = sessionStorage.getItem('token');

export const DialogComponent = (props) => {
    const [isOpenState, setIsOpenState] = useState(props.isOpen);
    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px)': {
                        height: 200,
                        maxHeight: 200,
                        maxWidth: 650,
                        width: 650
                    }
                }
            }
        },
    };

    useEffect(() => {
        if (props?.message) {
            dialogContentProps.subText = props?.message;
        }
        setIsOpenState(props.isOpen);
    }, [props.isOpen]);

    const handleClose = () => {
        setIsOpenState(false);
        props.onClose(false);
    };

    return (
        <>
            <Dialog
                hidden={!isOpenState} // Negate the isOpenState to properly handle visibility
                onDismiss={handleClose} // Remove the parentheses from handleClose
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
                <DialogFooter>
                    <PrimaryButton onClick={handleClose} text="Ok" /> {/* Remove the parentheses from handleClose */}
                    {/* <DefaultButton onClick={ handleClose } text="Cancel" />  */}
                </DialogFooter>
            </Dialog>
        </>
    );
};

export const TransferDialogComponent = ({ isOpen, message, onClose }) => {
    const [isOpenState, setIsOpenState] = useState(isOpen);

    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px)': {
                        height: 100,
                        maxHeight: 100,
                        maxWidth: 650,
                        width: 650,
                    }
                }
            }
        }
    };

    const dialogContentProps = {
        subText: message || "Are You Sure You Want to Transfer the Data",
    };

    useEffect(() => {
        setIsOpenState(isOpen);
    }, [isOpen]);

    const handleClose = (shouldClose) => {
        setIsOpenState(false);
        onClose(shouldClose);
    };

    return (
        <Dialog
            hidden={!isOpenState}
            onDismiss={() => handleClose(false)}
            dialogContentProps={dialogContentProps}
            modalProps={modalProps}
        >
            <DialogFooter>
                <PrimaryButton onClick={() => handleClose(true)} text="Proceed" />

                <DefaultButton onClick={() => handleClose(false)} text="Close" />
            </DialogFooter>
        </Dialog>
    );
};

export const DiscrepancyOptionsDialogComponent = (props) => {
    const [isOpenState, setIsOpenState] = useState(props.isOpen);
    const [selectedOption1, setSelectedOption1] = useState(null);
    const [selectedOption2, setSelectedOption2] = useState(null);
    const [selectedOption3, setSelectedOption3] = useState(null);
    const [inputValue1, setInputValue1] = useState('');
    const [inputValue2, setInputValue2] = useState('');
    const [inputValue3, setInputValue3] = useState('');
    // const [ areAllOptionsSelected, setAreAllOptionsSelected ] = useState( false );

    dialogContentProps.subText = "";
    dialogContentProps.title = "Discrepancy Dropdown";

    let jobid = props?.jobId || props?.state[0]?.JobId;
    const brokerId = jobid?.slice(0, 4) || 0;
    const reqEndorsementColHideBrokerId = ["1162"];

    let modalProps;
    if (reqEndorsementColHideBrokerId.includes(brokerId)) {
        modalProps = {
            isBlocking: true,
            styles: {
                main: {
                    selectors: {
                        '@media (min-width: 0px)': {
                            height: 260,
                            maxHeight: 260,
                            maxWidth: 450,
                            width: 450
                        }
                    }
                }
            },
        };
    } else {
        modalProps = {
            isBlocking: true,
            styles: {
                main: {
                    selectors: {
                        '@media (min-width: 0px)': {
                            height: 320,
                            maxHeight: 320,
                            maxWidth: 450,
                            width: 450
                        }
                    }
                }
            },
        };
    }

    useEffect(() => {
        setIsOpenState(props.isOpen);
    }, [props.isOpen]);

    // useEffect( () => {
    //     setAreAllOptionsSelected( selectedOption1 && selectedOption2 && selectedOption3 );
    // }, [ selectedOption1, selectedOption2, selectedOption3 ] );

    const handleClose = () => {
        setIsOpenState(false);
        // setAreAllOptionsSelected( false );
        props.onClose({ hasData: false });
    };

    const handleOk = () => {
        setIsOpenState(false);
        props.onClose({ hasData: true, selectedOption1, selectedOption2, selectedOption3 });
    };


    let Actions_on_Discrepancy, Request_Endorsement, Notes_for_Endorsement;
    if (brokerId === "1165") {
        Actions_on_Discrepancy = [
            { key: 'Yes - Internal colleague to request', text: 'Yes - Internal colleague to request' },
            { key: 'Yes - Exdion to update EPIC', text: 'Yes - Exdion to update EPIC' },
            { key: 'OK as is', text: 'OK as is' },
            { key: 'Email Policy to Client', text: 'Email Policy to Client' },
        ];

        Request_Endorsement = [
            { key: 'Yes', text: 'Yes' },
            { key: 'No', text: 'No' },
        ];

        Notes_for_Endorsement = [
            { key: 'OK', text: 'OK' },
            { key: 'Current matches Expiring policy', text: 'Current matches Expiring policy' },
            { key: 'Add of Proposal', text: 'Add of Proposal' },
            { key: 'Clarification Needed', text: 'Clarification Needed' },
            { key: 'Change Request to Carrier', text: 'Change Request to Carrier' },
            { key: 'Confirmed OK w/Servicer', text: 'Confirmed OK w/Servicer' },
        ];
    } else if (brokerId === "1162") {
        Actions_on_Discrepancy = [
            { key: 'Not a Discrepancy - Select Notes for Endt', text: 'Not a Discrepancy - Select Notes for Endt' },
            { key: 'Send Change Request', text: 'Send Change Request' },
            { key: 'Updated EPIC', text: 'Updated EPIC' },
            { key: 'Updated Exposure Spreadsheet / IRIS', text: 'Updated Exposure Spreadsheet / IRIS' },
            { key: 'Updated Proposal', text: 'Updated Proposal' },
            { key: 'Updated Proposal / EPIC', text: 'Updated Proposal / EPIC' },
            { key: 'FOR CSR: Updated ISUB Application', text: 'FOR CSR: Updated ISUB Application' },
        ];

        Request_Endorsement = [
            { key: 'Yes', text: 'Yes' },
            { key: 'No', text: 'No' },
        ];

        Notes_for_Endorsement = [
            { key: 'Ok as is - enter reason in notes field', text: 'Ok as is - enter reason in notes field' },
            { key: 'AI Limitation - go to QAC not answered tab', text: 'AI Limitation - go to QAC not answered tab' },
            // { key: 'OK', text: 'OK' },
            // { key: 'Current matches Expiring policy', text: 'Current matches Expiring policy' },
            // { key: 'Add of Proposal', text: 'Add of Proposal' },
            // { key: 'Clarification Needed', text: 'Clarification Needed' },
            // { key: 'Change Request to Carrier', text: 'Change Request to Carrier' },
            // { key: 'Confirmed OK w/Servicer', text: 'Confirmed OK w/Servicer' },
        ];
    } else {
        Actions_on_Discrepancy = [
            { key: 'Yes - Lead Servicer to request', text: 'Yes - Lead Servicer to request' },
            { key: 'Yes - Lead Servicer directed colleague or outsourcing to request', text: 'Yes - Lead Servicer directed colleague or outsourcing to request' },
            { key: 'Yes - Update AMS', text: 'Yes - Update AMS' },
            { key: 'OK as Is', text: 'OK as Is' },
        ];

        Request_Endorsement = [
            { key: 'Yes', text: 'Yes' },
            { key: 'No', text: 'No' },
        ];

        Notes_for_Endorsement = [
            { key: 'OK', text: 'OK' },
            { key: 'Current matches Expiring policy', text: 'Current matches Expiring policy' },
            { key: 'Add of Proposal', text: 'Add of Proposal' },
            { key: 'Clarification Needed', text: 'Clarification Needed' },
            { key: 'Change Request to Carrier', text: 'Change Request to Carrier' },
            { key: 'Confirmed OK w/Servicer', text: 'Confirmed OK w/Servicer' },
        ];
    }


    return (
        <Dialog
            hidden={!isOpenState}
            onDismiss={handleClose}
            dialogContentProps={dialogContentProps}
            modalProps={modalProps}
        >
            <ComboBox
                label="Actions on Discrepancy"
                options={Actions_on_Discrepancy}
                selectedKey={selectedOption1 ? selectedOption1.key : null}
                onChange={(event, option) => {
                    if (option) {
                        setSelectedOption1(option);
                        setInputValue1(option.text);
                    } else {
                        setInputValue1(event.target.value || '');
                        setSelectedOption1({ key: event.target.value, text: event.target.value });
                    }
                }}
                text={inputValue1}
                allowFreeform
                autoComplete="on"
                placeholder="Select an Actions on Discrepancy option"
            />
            {!reqEndorsementColHideBrokerId.includes(brokerId) && (<ComboBox
                label="Request Endorsement"
                options={Request_Endorsement}
                selectedKey={selectedOption2 ? selectedOption2.key : null}
                onChange={(event, option) => {
                    if (option) {
                        setSelectedOption2(option);
                        setInputValue2(option.text);
                    } else {
                        setInputValue2(event.target.value || '');
                        setSelectedOption2({ key: event.target.value, text: event.target.value });
                    }
                }}
                text={inputValue2}
                allowFreeform
                autoComplete="on"
                placeholder="Select an Request Endorsement option"
            />)}
            <ComboBox
                label="Notes for Endorsement"
                options={Notes_for_Endorsement}
                selectedKey={selectedOption3 ? selectedOption3.key : null}
                onChange={(event, option) => {
                    if (option) {
                        setSelectedOption3(option);
                        setInputValue3(option.text);
                    } else {
                        setInputValue3(event.target.value || '');
                        setSelectedOption3({ key: event.target.value, text: event.target.value });
                    }
                }}
                text={inputValue3}
                allowFreeform
                autoComplete="on"
                placeholder="Select an Notes for Endorsement option"
            />
            <DialogFooter>
                <DefaultButton onClick={handleClose} text="Cancel" />
                <PrimaryButton onClick={handleOk} text="Ok" />
            </DialogFooter>
        </Dialog>
    );
};

export const FindDialogComponent = (props) => {
    const container = useRef();
    const [isOpenState, setIsOpenState] = useState(props.isOpen);
    const [isSheet, setSheetState] = useState(props.sheetState);
    const [isConfig, setIsConfig] = useState(props.luckySheet);
    const [searchText, setSearchText] = useState();
    const [searchResults, setSearchResults] = useState([]);
    dialogContentProps.subText = "";
    const [result, setResult] = useState([]);
    dialogContentProps.title = "Find";

    const [columns, setColumns] = useState([
        { key: 'column1', name: 'Found Text', fieldName: 'foundText', minWidth: 100, maxWidth: 125, isResizable: true, onRender: renderFoundTextColumn },
        { key: 'column2', name: 'Count', fieldName: 'count', minWidth: 100, maxWidth: 125, isResizable: true },
        { key: 'column3', name: 'Position', fieldName: 'positions', minWidth: 100, maxWidth: 126, isResizable: true, onRender: renderPositionColumn },
    ]);

    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px)': {
                        height: 250,
                        maxHeight: 250,
                        maxWidth: 500,
                        width: 500,
                        minHeight: 335
                    }
                }
            }
        },
    };

    useEffect(() => {
        setIsOpenState(props.isOpen?.findDialog);
        setSheetState(isSheet);
        findColumn(isSheet);
    }, [props.isOpen?.findDialog, isSheet]);

    const handleClose = () => {
        setIsOpenState(false);
        props.onClose({state: false});
    };

    const handleSearch = () => {
        if (searchText != undefined) {
            if (searchText?.toLowerCase() === "details not available in the document" && (result?.find(item => item?.originalText?.toLowerCase().trim() === "details not available in the document") || result?.find(item => item?.originalText.trim() === "Details not available in the document"))) {
                setSearchResults([{ foundText: searchText, count: 0, positions: '' }]);
            } else {
                const matchingResult = result?.filter(f => f?.originalText?.toLowerCase()?.includes(searchText?.toLowerCase()));

                if (matchingResult?.length > 0) {
                    const searchResults = matchingResult?.flatMap(matchingResult => {
                        return matchingResult?.positions?.map(pos => ({
                            foundText: matchingResult?.originalText,
                            count: matchingResult?.count,
                            positions: `${pos.row}_${pos.column}`
                        }));
                    });
                    setSearchResults(searchResults);
                } else {
                    setSearchResults([{ foundText: searchText, count: 0, positions: '' }]);
                }
            }
        } else {
            setTimeout(() => {
                container.current.showSnackbar('Enter any text to search...', "info", true);
            }, 300)
        }
    };

    const handleRowClick = (item) => {
        const [rowIndex, columnIndex] = item.positions.split('_').map(pos => {
            if (!isNaN(pos)) {
                return parseInt(pos);
            } else {
                return pos.charCodeAt(0) - 65;
            }
        });
        isConfig.setluckysheet_select_save([{ row: [rowIndex - 1, rowIndex - 1], column: [columnIndex, columnIndex] }]);
        isConfig.selectHightlightShow();
        isConfig.scroll({
            targetRow: rowIndex - 1,
            targetColumn: columnIndex - 1
        });
        // let configData = isConfig.getConfig()
        // let data = configData;
        // if (data) {
        //     if (data?.columnlen) {
        //         const keys = Object.keys(data?.columnlen);
        //         if (keys?.length > 0) {
        //             keys.map((key) => {
        //                 data.columnlen[key] = 250;
        //             })
        //         }
        //     }
        //     if (data?.rowlen) {
        //         const rowLenForNavigation = data.rowlen;
        //         const keys = Object.keys(rowLenForNavigation);
        //         let val = 0;
        //         if (keys.length > 0) {
        //             keys.forEach((f) => { if (parseInt(f) < rowIndex) { val += rowLenForNavigation[f] } });
        //             if (val > 0) {
        //                 $("#luckysheet-scrollbar-x").scrollLeft(val - 300);
        //                 $("#luckysheet-scrollbar-y").scrollTop(val - 150);
        //             }
        //         }
        //     }
        //     configData = data;
        // }
    };

    function removeNulls(arr) {
        if (Array.isArray(arr)) {
            return arr?.filter(item => item !== null).map(removeNulls);
        } else {
            return arr;
        }
    }

    const findColumn = (isSheet) => {
        let sheetData = isSheet;
        let sheetDataLength = sheetData?.length;
        for (let index = sheetDataLength - 1; index < sheetDataLength && index != 0; index--) {
            let hasValue = sheetData[index].filter((f) => f != null)?.length > 0;
            if ((!hasValue)) {
                sheetData = sheetData?.slice(0, index);
            } else {
                break;
            }
        }
        let dataWithoutNulls = removeNulls(sheetData);
        const duplicateWords = {};
        dataWithoutNulls?.forEach((row, rowIndex) => {
            if (row?.length > 0) {
                row?.forEach((cell, colIndex) => {
                    const ctObj = cell?.ct || {};
                    const sArray = ctObj?.s || [];
                    let concatenatedValue = ''; // Variable to store concatenated values
                    let originalValue = ''; // Variable to store original values
                    if (ctObj?.s && Array.isArray(ctObj?.s)) {
                        sArray?.forEach((s) => {
                            const value = s?.v || '';
                            concatenatedValue += value?.trim();
                            originalValue += value;
                        });
                        concatenatedValue = concatenatedValue.replace(/\s+/g, '');
                    }
                    else {
                        const value = cell?.m || cell?.v || {};
                        concatenatedValue = value;
                        originalValue = value;
                    }
                    if (typeof concatenatedValue === 'string' && concatenatedValue.trim() !== '') {
                        if (duplicateWords[concatenatedValue]) {
                            duplicateWords[concatenatedValue].count++;
                            duplicateWords[concatenatedValue].positions.push({ row: rowIndex, column: colIndex });
                        } else {
                            duplicateWords[concatenatedValue] = {
                                count: 1,
                                positions: [{ row: rowIndex, column: colIndex }],
                                originalText: originalValue.trim()
                            };
                        }
                    }
                });
            }
        });

        const result = Object.entries(duplicateWords).reduce((acc, [word, data]) => {
            if (data.count >= 1) {
                const positions = data?.positions?.map(position => ({
                    row: position.row + 1,
                    column: String.fromCharCode(66 + position.column) // Convert numeric column index to ASCII value
                }));
                acc.push({
                    count: data.count,
                    positions,
                    originalText: data.originalText
                });
            }
            return acc;
        }, []);
        setResult(result)
    }

    function renderPositionColumn(item) {
        if (item && item?.positions) {
            if (item?.positions?.length > 10) {
                return (
                    <div style={{ whiteSpace: 'pre-wrap' }}>{item?.positions}</div>
                );
            } else {
                return <div>{item?.positions}</div>;
            }
        }
    }

    function renderFoundTextColumn(item) {
        if (item && item?.foundText) {
            if (item?.foundText?.length > 10) {
                return (
                    <div style={{ whiteSpace: 'pre-wrap' }}>{item?.foundText}</div>
                );
            } else {
                return <div>{item?.foundText}</div>;
            }
        }
    }

    return (
        <>
            <Dialog
                hidden={!isOpenState}
                onDismiss={handleClose}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
                className="dialogContainer"
            >
                <div>
                    <input
                        type="text"
                        value={searchText}
                        onChange={(e) => setSearchText(e.target.value)}
                        style={{ height: '25px' }}
                    />
                    <PrimaryButton
                        onClick={handleSearch}
                        text="Search"
                        style={{ marginLeft: '10px' }}
                    />
                </div>
                <div className="detailsListContainer">
                    <DetailsList
                        items={searchResults}
                        columns={columns}
                        onActiveItemChanged={(item, index, ev) => handleRowClick(item)}
                        selectionMode={SelectionMode.none}
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                    />
                </div>
                <DialogFooter>
                    <DefaultButton onClick={handleClose} text="Close" />
                </DialogFooter>
            </Dialog>
            <SimpleSnackbar ref={container} />
        </>
    );
};

export const EndorsementDialogComponent = (props) => {
    const [isOpenState, setIsOpenState] = useState(props.isOpen);
    const [isSheet, setSheetState] = useState(props.sheetState);
    const [luckysheet, setLuckysheet] = useState(props.luckySheet);
    const [tableset, setTableSet] = useState(props.tableColumnDetails);
    const [tableFormset, setTableFormSet] = useState(props.formTableColumnDetails);
    const [tableExclusionset, setTableExclusionSet] = useState(props.exTableColumnDetails);
    const [policyData, setPolicyData] = useState(props.state);
    const [formsCompareData, setFormsCompareData] = useState(props.formState);
    const [redSheetData, setRedSheetData] = useState(props.redSheetData);
    const [xRayData, setxRAyData] = useState(props?.dataForXRayMapping);
    const [xRayFormData, setxRAyFormData] = useState(props?.formdataForXRayMapping);
    const [xRayRedSheetData, setxRAyRedSheetData] = useState(props?.redSheetDataForXRayMapping);
    const [xRayGreenSheetData, setxRAyGreenSheetData] = useState(props?.greenSheetDataForXRayMapping);
    const [selctedTableData, setSelctedTableData] = useState([]);
    const [loBResult, setLobResult] = useState([]);
    dialogContentProps.title = "Request Endorsement Row Item";
    dialogContentProps.subText = "";

    const redDailogData = sessionStorage.getItem('redDailogData');
    const greenDailogData = sessionStorage.getItem('greenDailogData');

    let lobRedDailogData = JSON.parse(redDailogData);
    let lobGreenDailogData = JSON.parse(greenDailogData);

    // Retrieve the data for both sheets
    const redTableData = sessionStorage.getItem('redTableRangeData');
    const greenTableData = sessionStorage.getItem('greenTableRangeData');

    const parsedRedTableData = JSON.parse(redTableData);
    const parsedGreenTableData = JSON.parse(greenTableData);

    useEffect(() => {
        setIsOpenState(props.isOpen);
        setSheetState(isSheet);
        yesValueMapping(isSheet);
    }, [props.isOpen, isSheet, policyData, formsCompareData]);

    const handleClose = () => {
        setIsOpenState(false);
        props.onClose({ excelDataFlag: false });
    };

    const handleOk = async () => {
        const dataToPush = [];
        const csrHistoryData = [];
        Object.keys(selctedTableData).forEach(tableName => {
            const RowDataSet = [];
            selctedTableData[tableName].rows.forEach(row => {
                const rowData = {};
                selctedTableData[tableName].headers.forEach((header, index) => {
                    rowData[header] = row[index] ? row[index].value : '';
                    // const key = JSON.stringify(header);
                    // const value = JSON.stringify(row[index] ? row[index].value : '');
                    // rowData[key] = value;
                });
                RowDataSet.push(rowData);
            });
            dataToPush.push({ TableName: getPolicyLOB(tableName), Data: RowDataSet });
            csrHistoryData.push({ RowDataSet });
        });

        let arrayOfObjs = csrHistoryData.reduce((acc, table) => {
            return acc.concat(table.RowDataSet);          // Api Data for CsrSave History for jobId's 
        }, []);

        if(arrayOfObjs?.length > 0){
            let UserName = sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName");
            let JobID = props?.state[0]?.JobId;
            let currentDate = new Date().toISOString();

            arrayOfObjs.forEach(item => {
                item.UserName = UserName;  // Add UserName as a key-value pair
                item.JobID = JobID;        // Add JobID as a key-value pair
                item.CreatedOn = currentDate;
            });
        }

        let brokerData = await brokerIdsGetData();
        const mappedData = brokerData.map(({ BrokerId, VchBrokerName }) => ({
            key: BrokerId.toString(),
            text: VchBrokerName,
        }));

        arrayOfObjs.forEach(item => {
            const sliceJobid = item?.JobID.slice(0, 4);
            const brokerIdMap = mappedData.find(e => e.key === sliceJobid);

            if (brokerIdMap) {
                item.BrokerId = brokerIdMap?.key;
                item.BrokerName = brokerIdMap?.text;
            }
        });

        const finalDataSetForExport = [];
        // dataToPush.forEach((f, index) => {
        //     // if(AvailableLobs && AvailableLobs?.length > 0 && AvailableLobs?.includes(f?.TableName)){
        //     //     const TableName = " EndorsementTable " + (index + 1);
        //     // }else{
        //     //     const TableName = " EndorsementTable " + (index + 1);
        //     // }
        //     const TableName = "EndorsementTable" + (index + 1);
        //     const data =f?.Data;
        //     const updatedData = [];
        //     if (data && data?.length > 0) {
        //         data.forEach((item) => {
        //             item["policyLob"] = f?.TableName;
        //             updatedData.push(item);
        //         });
        //     }

        //     // finalDataSetForExport.push({ TableName, 
        //     // "Data": updatedData?.length > 0 ? JSON.stringify(updatedData) : updatedData });
        // });
        // const formattedData = {
        //     Tabledata: finalDataSetForExport
        // };
        dataToPush.forEach((f, index) => {
            const TableName = "EndorsementTable " + (index + 1);
            let data = f?.Data;
            const Id = props?.state[0]?.JobId;

            if (data && data?.length > 0) {
                data = JSON.stringify(data.map(item => {
                    item["policyLob"] = f?.TableName;
                    return item;
                }));
            }
            finalDataSetForExport.push({ TableName, Data: data, Id });
        });

        const sheetNameArray = ["DiscripencyTable"];
        const table = '"' + JSON.stringify(sheetNameArray) + '"';
        var dialogData = "{" + '"' + "Data" + '"' + ":" + JSON.stringify(table) + "," + '"' + "Tabledata" + '"' + ":" + JSON.stringify(finalDataSetForExport) + "}";
        var jobid = props?.state[0]?.JobId;
        const response = await endrolmentDialogSaveApi(jobid, dialogData);
        await CsrSaveHistoryApiCall(jobid, sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName"), JSON.stringify(arrayOfObjs));
        await CsrPendingReport(jobid, sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName"), JSON.stringify(arrayOfObjs), 'CSrExport-Save');
        return response;
    };

    const endrolmentDialogSaveApi = async (jobId, dialogData) => {
        document.body.classList.add('loading-indicator');
        const Token = await processAndUpdateToken(token);
        updateGridAuditLog(jobId, "CSR-EndorsementDailogSave", "EndorsementDailogSave", (sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName")));
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${baseUrl}/api/ProcedureData/UpdateActionOnDiscrepancyData`;
        try {
            const response = await axios({
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    JobId: jobId,
                    Data: dialogData,
                }
            });
            if (response.status !== 200) {
                return "error";
            }

            return response.data;
        } catch (error) {
            // console.error( 'Error:', error );
            return "error"; // Rethrow the error to be caught in the calling function
        } finally {
            updateGridAuditLog(jobId, "CSR-EndorsementDailogSave-Success", "EndorsementDailogSave-Success", (sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName")));
            document.body.classList.remove('loading-indicator');
            return "success";
        }
    };

    const ExportDiscripancyData = async (jobId) => {
        document.body.classList.add('loading-indicator');
        const Token = await processAndUpdateToken(token);
        updateGridAuditLog(jobId, "CSR-Dialog-ExportData", "CSR-Dailog-ExportExcel", (sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName")));
        const JobId = props?.state[0]?.JobId;
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const apiUrl = `${baseUrl}/api/Excel/ExportExcelDiscripancyTable`;

        try {
            const response = await axios({
                method: "POST",
                url: apiUrl,
                headers: headers,
                data: {
                    JobId: JobId,
                },
                responseType: 'blob'
            });
            if (response.status !== 200) {
                console.error('Failed to download file, status:', response.status);
                return "error";
            }

            const url = window.URL.createObjectURL(new Blob([response.data]));
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', `${JobId}GridExcel.xlsx`);
            document.body.appendChild(link);
            link.click();
        } catch (error) {
            console.error('Error downloading file:', error);
            return "error";
        } finally {
            updateGridAuditLog(jobId, "CSR-Dailog-ExportData-Success", "Dailog-ExportExcel-Success", (sessionStorage.getItem("csrAuthUserName") || sessionStorage.getItem("userName")));
            document.body.classList.remove('loading-indicator');
            return "success";
        }
    };
    
    const handleSaveAndExport = async (jobId) => {
        const saveResult = await handleOk();
        if (saveResult !== "error") {
            handleClose();
            const exportResult = await ExportDiscripancyData(jobId);
            if (exportResult === "success") {
                console.log("Export successful");
            }
        } else {
            console.error("Error saving data");
        }
    };

    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px) and (max-width: 767px)': {
                        height: '200px',
                        width: '300px',
                    },
                    '@media (min-width: 768px) and (max-width: 1023px)': {
                        height: '300px',
                        width: '450px',
                    },
                    '@media (min-width: 1024px)': {
                        minWidth: '980px',
                        height: '480px'
                    },
                    '@media (min-width: 480px)': {
                        maxWidth: '660px',
                        height: '300px',
                    }
                }
            }
        },
    };

    function extractCharacterValues(sheetValues) {
        const sheetValue = [];
        for (const row of sheetValues) {
            const rowIndex = row.rowIndex;
            const rowValue = row.rowValues;
            const nullFilterRowValues = rowValue.filter(f => f != null);
            for (const cell of nullFilterRowValues) {
                if (cell && cell.ct && cell.ct.s && cell.ct.s.length > 0 && cell.ct.s[0].hasOwnProperty('v')) {
                    if (cell?.ct?.s && cell?.ct?.s.length >= 2 && Array.isArray(cell?.ct?.s)) {
                        let concatenatedValues = cell?.ct?.s?.map(item => item?.v)?.join('');
                        concatenatedValues = concatenatedValues.replace(/\r\n/g, '');
                        let modifiedText = concatenatedValues.replace('Page #', '~~Page #');
                        sheetValue.push({ rowIndex: rowIndex, value: modifiedText });
                    } else if (cell?.ct?.s && cell?.ct?.s.length > 0 && cell?.ct?.s.length <= 1) {
                        cell.ct.s[0].v = cell.ct.s[0].v.replace(/\r\n/g, '')
                        let modifiedText = cell.ct.s[0].v.replace('Page #', '~~Page #');
                        sheetValue.push({ rowIndex: rowIndex, value: modifiedText });
                    }
                } else if (cell.hasOwnProperty('v')) {
                    cell.v = cell.v.replace(/\r\n/g, '')
                    sheetValue.push({ rowIndex: rowIndex, value: cell.v });
                }
            }
        }
        return sheetValue;
    }

    const yesValueMapping = (isSheet) => {
        let flagCheck = luckysheet.getSheet().name;
        if (flagCheck != 'Red' && flagCheck != 'Green' && flagCheck != 'Exclusion' && flagCheck != 'Forms Compare') {
            let propsData = flagCheck == "PolicyReviewChecklist" ? policyData : formsCompareData;
            let lobResult = [];
            let uniqueLobSet = new Set();
            propsData.forEach(table => {
                if (table.Tablename !== 'Table 1' && !uniqueLobSet.has(table.Tablename)) {
                    let policyLOB = table.TemplateData[0]["POLICY LOB"] || table.TemplateData[0]["Policy LOB"];
                    lobResult.push({
                        tableName: table.Tablename,
                        "POLICY LOB": policyLOB
                    });
                    uniqueLobSet.add(table.Tablename);
                }
            });
            setLobResult(lobResult);
        } else if (flagCheck == 'Red' || flagCheck == 'Green') {
            let propsData = redSheetData;
            let lobResult = [];
            let uniqueLobSet = new Set();

            propsData.forEach(table => {
                if (table.Tablename !== 'Table 1' && !uniqueLobSet.has(table.Tablename)) {
                    let policyLOB = table.TemplateData[0]["POLICY LOB"] || table.TemplateData[0]["Policy LOB"];
                    lobResult.push({
                        tableName: table.Tablename,
                        "POLICY LOB": policyLOB
                    });
                    uniqueLobSet.add(table.Tablename);
                }
            });
            sessionStorage.setItem('gradiationLobSet', JSON.stringify(lobResult));
            setLobResult(lobResult);
        } else if (flagCheck == 'Exclusion') {
            let lobResult = [];
            lobResult.push({
                tableName: "ExTable 1",
                "POLICY LOB": "Exclusion"
            });
            setLobResult(lobResult);
        } else if (flagCheck == 'Forms Compare') {
            let lobResult = [];
            lobResult.push({
                tableName: "FormTable 2",
                "POLICY LOB": "Unmatched Forms"
            }, {
                tableName: "FormTable 3",
                "POLICY LOB": "Matched Forms"
            });
            setLobResult(lobResult);
        }

        let sheetData = isSheet;
        let rowsWithYes = [];
        let redRowsWithYes = [];
        let greenRowsWithYes = [];
        if (flagCheck != 'Red' && flagCheck != 'Green') {
            for (let i = 0; i < sheetData.length; i++) {
                let row = sheetData[i];
                for (let j = 0; j < row.length; j++) {
                    let cell = row[j];
                    if (cell && ((cell.m && cell.m.toLowerCase() === "yes") || (cell.v && cell.v.toLowerCase() === "yes"))) {
                        rowsWithYes.push({ rowIndex: i, rowValues: row });
                        break;
                    } else if (cell && cell.ct && cell.ct.s) {
                        for (let k = 0; k < cell.ct.s.length; k++) {
                            let text = cell.ct.s[k].v.trim()
                            if (text.toLowerCase() === "yes") {
                                rowsWithYes.push({ rowIndex: i, rowValues: row });
                            }
                            break;
                        }
                    }
                }
            }
        }
        if (flagCheck == 'Red' || flagCheck == 'Green') {
            let redSheetData = sheetData.redData;
            let greenSheetData = sheetData.greenData;
            for (let i = 0; i < redSheetData.length; i++) {
                let row = redSheetData[i];
                for (let j = 0; j < row.length; j++) {
                    let cell = row[j];
                    if (cell && ((cell.m && cell.m.toLowerCase() === "yes") || (cell.v && cell.v.toLowerCase() === "yes"))) {
                        redRowsWithYes.push({ rowIndex: i, rowValues: row });
                        break;
                    } else if (cell && cell.ct && cell.ct.s) {
                        for (let k = 0; k < cell.ct.s.length; k++) {
                            let text = cell.ct.s[k].v.trim()
                            if (text.toLowerCase() === "yes") {
                                redRowsWithYes.push({ rowIndex: i, rowValues: row });
                            }
                            break;
                        }
                    }
                }
            }

            for (let i = 0; i < greenSheetData.length; i++) {
                let row = greenSheetData[i];
                for (let j = 0; j < row.length; j++) {
                    let cell = row[j];
                    if (cell && ((cell.m && cell.m.toLowerCase() === "yes") || (cell.v && cell.v.toLowerCase() === "yes"))) {
                        greenRowsWithYes.push({ rowIndex: i, rowValues: row });
                        break;
                    } else if (cell && cell.ct && cell.ct.s) {
                        for (let k = 0; k < cell.ct.s.length; k++) {
                            let text = cell.ct.s[k].v.trim()
                            if (text.toLowerCase() === "yes") {
                                greenRowsWithYes.push({ rowIndex: i, rowValues: row });
                            }
                            break;
                        }
                    }
                }
            }
        }
        redRowsWithYes = redRowsWithYes.map(row => {
            return {
                ...row,
                rowValues: row.rowValues.filter(cell => cell !== null)
            };
        });
        greenRowsWithYes = greenRowsWithYes.map(row => {
            return {
                ...row,
                rowValues: row.rowValues.filter(cell => cell !== null)
            };
        });

        let vValues = extractCharacterValues(rowsWithYes);
        vValues.forEach(f => {
            if (f.value == 'Click here') {
                f.value = ""
            }
        });
        let redValues = extractCharacterValues(redRowsWithYes);
        redValues.forEach(f => {
            if (f.value == 'Click here') {
                f.value = ""
            }
        });
        let greenValues = extractCharacterValues(greenRowsWithYes);
        greenValues.forEach(f => {
            if (f.value == 'Click here') {
                f.value = ""
            }
        });
        const dataSetValues = [];
        const redDataSetValues = [];
        const greenDataSetValues = [];
        let tempObject = [];
        let redTempObject = [];
        let greenTempObject = [];
        let result = [];
        let redResult = [];
        let greenResult = [];
        vValues.forEach(item => {              // foreach Data for policyChecklistSheet, formCompareSheet, exclusionSheet
            if (item.value !== "No") {
                let value = item.value.trim();
                if (!tempObject) {
                    tempObject = [];
                }
                tempObject.push({ rowIndex: item.rowIndex, value: value });
            }
        });
        if (tempObject.length > 0) {
            dataSetValues.push(tempObject);
        }
        let policyRowIndexes = tempObject.map((e) => e.rowIndex);
        policyRowIndexes = Array.from(new Set(policyRowIndexes));

        policyRowIndexes.forEach((f) => {
            let filteredData = tempObject.filter(item => item.rowIndex === f);
            // filteredData.forEach(f => {
            //     if(f.value == 'X-Ray'){
            //         f.value = "" 
            //     }
            // });
            let removedData = filteredData.filter(symbolFilter => symbolFilter.value != "+" && symbolFilter.value != "-")
            result.push(removedData)
        });

        redValues.forEach(item => {           // foreach Data for gradiationSheetRed
            if (item.value !== "No") {
                let value = item.value.trim();
                if (!redTempObject) {
                    redTempObject = [];
                }
                redTempObject.push({ rowIndex: item.rowIndex, value: value });
            }
        });
        if (redTempObject.length > 0) {
            redDataSetValues.push(redTempObject);
        }
        let redRowIndexes = redTempObject.map((e) => e.rowIndex);
        redRowIndexes = Array.from(new Set(redRowIndexes));

        redRowIndexes.forEach((f) => {
            let filteredData = redTempObject.filter(item => item.rowIndex === f);
            // filteredData.forEach(f => {
            //     if(f.value == 'X-Ray'){
            //         f.value = "" 
            //     }
            // });
            redResult.push(filteredData);
        });

        greenValues.forEach(item => {           // foreach Data for gradiationSheetGreen
            if (item.value !== "No") {
                let value = item.value.trim();
                if (!greenTempObject) {
                    greenTempObject = [];
                }
                greenTempObject.push({ rowIndex: item.rowIndex, value: value });
            }
        });
        if (greenTempObject.length > 0) {
            greenDataSetValues.push(greenTempObject);
        }
        let greenRowIndexes = greenTempObject.map((e) => e.rowIndex);
        greenRowIndexes = Array.from(new Set(greenRowIndexes));

        greenRowIndexes.forEach((f) => {
            let filteredData = greenTempObject.filter(item => item.rowIndex === f);
            // filteredData.forEach(f => {
            //     if(f.value == 'X-Ray'){
            //         f.value = "" 
            //     }
            // });
            greenResult.push(filteredData)
        });

        if (flagCheck == "PolicyReviewChecklist") {
            const mappingXRayDataFun = updateAllSheetsResultArray(result, xRayData, flagCheck);
        } else if (flagCheck == "Forms Compare") {
            const mappingXRayDataFun = updateAllSheetsResultArray(result, xRayFormData, flagCheck);
        } else if (flagCheck == "Red") {
            let redSheetParseData = JSON.parse(xRayRedSheetData);
            const mappingXRayDataFun = updateAllSheetsResultArray(redResult, redSheetParseData, flagCheck);
        } else if (flagCheck == "Green") {
            let greenSheetParseData = JSON.parse(xRayGreenSheetData);
            const mappingXRayDataFun = updateAllSheetsResultArray(greenResult, greenSheetParseData, flagCheck);
        }

        const selctedTableData = {};
        const redSelectedTableData = {};
        const greenSelectedTableData = {};
        result.forEach(row => {
            if (row.length > 0) {
                let selectedIndex = row[0].rowIndex;
                if (flagCheck == 'PolicyReviewChecklist' || flagCheck == 'Forms Compare') {
                    let tableRangeData = flagCheck == "PolicyReviewChecklist" ? tableset : tableFormset;
                    let selectedTable = findTblRowAllIndex(selectedIndex, tableRangeData);
                    if (selectedTable) {
                        if (!selctedTableData[selectedTable]) {
                            selctedTableData[selectedTable] = { headers: [], rows: [] };
                        }
                        let table = flagCheck == "PolicyReviewChecklist" ? tableset[selectedTable] : tableFormset[selectedTable];
                        // let docIndexFilter = flagCheck == "PolicyReviewChecklist" ? table.columnNames['Document Viewer'] - 1 : table.columnNames['Document Viewer'] - 1;
                        let headerValues = Object.keys(table.columnNames);
                        if (selectedTable == "FormTable 2" || selectedTable == "FormTable 3") {
                            headerValues = headerValues.filter(value => value !== "Id" && value !== "Checklist Questions" && value !== "OBSERVATION"
                                && value !== "Policy LOB" && value !== "Page Number" && value !== "IsMatched" && value !== "columnid" && value !== "Attached Forms");
                        } else {
                            // headerValues = headerValues.filter(value => value !== "Document Viewer");
                        }
                        headerValues[0] = "COVERAGE SPECIFICATIONS"
                        selctedTableData[selectedTable].headers = headerValues;
                        // if(flagCheck == "PolicyReviewChecklist" || selectedTable == "FormTable 2") {
                        //     row.splice(docIndexFilter, 1);
                        // }
                        selctedTableData[selectedTable].rows.push(row);
                    }
                } else if (flagCheck == 'Exclusion') {
                    let tableRangeData = tableExclusionset;
                    let selectedTable = findTblRowAllIndex(selectedIndex, tableRangeData);
                    if (selectedTable) {
                        if (!selctedTableData[selectedTable]) {
                            selctedTableData[selectedTable] = { headers: [], rows: [] };
                        }
                        let table = tableExclusionset[selectedTable];
                        let headerValues = Object.values(table.columnNames[0]);
                        headerValues = headerValues.filter(value => value !== "Id" && value !== "sheetPosition");
                        selctedTableData[selectedTable].headers = headerValues;
                        selctedTableData[selectedTable].rows.push(row);
                    }
                }
            }
        });

        redResult.forEach(row => {
            if (row.length > 0) {
                let selectedIndex = row[0].rowIndex;
                if (flagCheck == 'Red' || flagCheck == 'Green') {
                    let selectedTable = findTblRowAllIndex(selectedIndex, parsedRedTableData); // both headings is common for Red and Green sheets
                    if (selectedTable) {
                        if (!redSelectedTableData[selectedTable]) {
                            redSelectedTableData[selectedTable] = { headers: [], rows: [] };
                        }
                        let table = parsedRedTableData[selectedTable];
                        let docIndexFilter = table.columnNames['Document Viewer'] - 1;
                        let headerValues = Object.keys(table.columnNames);
                        headerValues = headerValues.filter(value => value !== "Document Viewer");
                        redSelectedTableData[selectedTable].headers = headerValues;
                        row.splice(docIndexFilter, 1);
                        redSelectedTableData[selectedTable].rows.push(row);
                    }
                }
            }
        });

        greenResult.forEach(row => {
            if (row.length > 0) {
                let selectedIndex = row[0].rowIndex;
                if (flagCheck == 'Red' || flagCheck == 'Green') {
                    let selectedTable = findTblRowAllIndex(selectedIndex, parsedGreenTableData); // both headings is common for Red and Green sheets
                    if (selectedTable) {
                        if (!greenSelectedTableData[selectedTable]) {
                            greenSelectedTableData[selectedTable] = { headers: [], rows: [] };
                        }
                        let table = parsedGreenTableData[selectedTable];
                        let docIndexFilter = table.columnNames['Document Viewer'] - 1;
                        let headerValues = Object.keys(table.columnNames);
                        headerValues = headerValues.filter(value => value !== "Document Viewer");
                        greenSelectedTableData[selectedTable].headers = headerValues;
                        row.splice(docIndexFilter, 1);
                        greenSelectedTableData[selectedTable].rows.push(row);
                    }
                }
            }
        });


        if (flagCheck == 'Red' || flagCheck == 'Green') {
            function updateRedSelectedTableData(lobData, redData) {
                lobData.forEach(lobEntry => {
                    const tableName = lobEntry.TableName;
                    const policyLob = lobEntry['Policy LOB'];

                    if (redData[tableName]) {
                        redData[tableName].policyLob = policyLob;
                    }
                });
            }
            updateRedSelectedTableData(lobRedDailogData, redSelectedTableData);

            function updateGreenSelectedTableData(lobData, greenData) {
                lobData.forEach(lobEntry => {
                    const tableName = lobEntry.TableName;
                    const policyLob = lobEntry['Policy LOB'];

                    if (greenData[tableName]) {
                        greenData[tableName].policyLob = policyLob;
                    }
                });
            }
            updateGreenSelectedTableData(lobGreenDailogData, greenSelectedTableData);

            const combineData = (redData, greenData) => {
                const combinedData = {};
                // Function to add table data to combinedData
                const addToCombinedData = (data, source) => {
                    for (const table in data) {
                        if (data.hasOwnProperty(table)) {
                            const { policyLob, headers, rows } = data[table];

                            // Create a unique key based on table and policyLob to handle separate objects for different policyLob
                            const key = `${policyLob}`;

                            if (combinedData[key]) {
                                // Merge rows if the same table and policyLob already exist
                                combinedData[key].rows.push(...rows);
                            } else {
                                // Add new entry if table and policyLob combination does not exist
                                combinedData[key] = { headers, rows: rows.slice(), policyLob };
                            }
                        }
                    }
                };
                addToCombinedData(redData, 'red');
                addToCombinedData(greenData, 'green');

                // Convert combinedData to the final desired format
                const finalData = {};
                for (const key in combinedData) {
                    if (combinedData.hasOwnProperty(key)) {
                        const { headers, rows, policyLob } = combinedData[key];
                        // const table = key.split('_')[0];
                        const table = key;
                        if (!finalData[table]) {
                            finalData[table] = { headers, rows, policyLob };
                        } else if (finalData[table].policyLob === policyLob) {
                            finalData[table].rows.push(...rows);
                        } else {
                            finalData[`${table}_${policyLob}`] = { headers, rows, policyLob };
                        }
                    }
                }

                return finalData;
            };

            let combinedGradiationTableData = combineData(redSelectedTableData, greenSelectedTableData);
            const sessionLobSet = sessionStorage.getItem('gradiationLobSet');
            let tblLObSet = JSON.parse(sessionLobSet);

            let lobToTableMapping = {};
            let correctGradiationTableData = {};

            for (let entry of tblLObSet) {
                lobToTableMapping[entry["POLICY LOB"]] = entry["tableName"];
            }

            for (let tableName in combinedGradiationTableData) {
                let data = combinedGradiationTableData[tableName];
                let correctTableName = lobToTableMapping[data["policyLob"]] || tableName;
                correctGradiationTableData[correctTableName] = data;
            }
            setSelctedTableData(correctGradiationTableData);
        }

        if (flagCheck == 'PolicyReviewChecklist' || flagCheck == 'Forms Compare' || flagCheck == 'Exclusion') {
            setSelctedTableData(selctedTableData);
        }
    }

    const getPolicyLOB = (tableName) => {
        const table = loBResult.find(item => item.tableName === tableName);
        return table ? (table["POLICY LOB"] || table["Policy LOB"]) : '';
    };

    function renderPositionColumn(result, index) {

        if (result.length > index) {
            let selectedIndexMap = result.map(e => e.rowIndex);
            Array.from(new Set(selectedIndexMap));
            let selectedIndex = selectedIndexMap[0];
            let flagCheck = luckysheet.getSheet().name;
            let tableRangeData = flagCheck == "PolicyReviewChecklist" ? tableset : tableFormset;
            let selectedTable = findTblRowAllIndex(selectedIndex, tableRangeData);
            if (selectedTable) {
                let table = flagCheck == "PolicyReviewChecklist" ? tableset[selectedTable] : tableFormset[selectedTable];
                let headerValues = Object.keys(table.columnNames);
                const value = result[index].value;
                let docIndexCheckForXRay = headerValues.indexOf("Document Viewer");
                if (docIndexCheckForXRay === index) {
                    if (value != "") {
                        return (
                            <a href={value} target="_blank" rel="noopener noreferrer" style={{ display: 'block', margin: '15px 0px 0px 42px' }}>X-Ray</a>
                        );
                    }
                } else {
                    return (
                        <div>{value}</div>
                    );
                }
            }
        }
        return null;
    }

    return (
        <>
            <Dialog
                hidden={!isOpenState}
                onDismiss={handleClose}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
                <div style={{ marginTop: '-25px', padding: '0px' }}>
                    {
                        Object.keys(selctedTableData).length > 0 ? (
                            Object.keys(selctedTableData).map((tableName, idx) => (
                                <div key={idx} >
                                    <h3>Table - {getPolicyLOB(tableName)}</h3>
                                    <DetailsList
                                        className="custom-details-list  endroslment-tables"
                                        items={selctedTableData[tableName].rows}
                                        columns={selctedTableData[tableName].headers.map((name, index) => ({
                                            name: name,
                                            maxWidth: 240.5,
                                            minWidth: 240.5,
                                            isResizable: true,
                                            onRender: item => renderPositionColumn(item, index),
                                        }))}
                                        checkboxVisibility={CheckboxVisibility.hidden}
                                    />
                                </div>
                            ))
                        ) : (
                            <div style={{ marginTop: '190px', textAlign: 'center' }}>
                                <span style={{ fontSize: 'x-large' }}>No Records Found For Request Endorsement!</span>
                            </div>
                        )
                    }
                </div>
                <DialogFooter>
                    <DefaultButton onClick={handleClose} style={{ margin: '135px 10px 0px 0px' }} text="Close" />
                    <PrimaryButton onClick={handleOk} style={{ margin: '135px 10px 0px 0px' }} text="Save" disabled={Object.keys(selctedTableData).length === 0} />
                    <PrimaryButton onClick={handleSaveAndExport} style={{ margin: '135px 10px 0px 0px' }} text="Save & Export" disabled={Object.keys(selctedTableData).length === 0} />
                </DialogFooter>
            </Dialog>
        </>
    );
};

export const InputDialogComponent = (props) => {
    const [isOpenState, setIsOpenState] = useState(props.isOpen);
    const [value, setValue] = React.useState("0");
    const styles = { spinButtonWrapper: { width: 150 } };
    dialogContentProps.subText = "";
    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px)': {
                        height: 250,
                        maxHeight: 300,
                        maxWidth: 650,
                        width: 650
                    }
                }
            }
        },
    };

    useEffect(() => {
        setIsOpenState(props.isOpen);
    }, [props.isOpen]);

    const handleClose = (e) => {
        setIsOpenState(false);
        props.onClose({ state: false, input: 0 });
    };
    const handleOk = (e) => {
        setIsOpenState(false);
        const inputValue = parseInt(value)
        props.onClose({ state: false, input: inputValue });
    };

    const onChange = (newValue) => {
        // Handle validation if needed
        setValue(newValue);
    };

    return (
        <>
            <Dialog
                hidden={!isOpenState} // Negate the isOpenState to properly handle visibility
                onDismiss={handleClose} // Remove the parentheses from handleClose
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
                <SpinButton
                    label="Please enter input for which how many rows you wish to be inserted"
                    labelPosition={Position.top}
                    onChange={(event, newValue) => onChange(newValue)}
                    defaultValue={value}
                    min={0}
                    max={200}
                    step={1}
                    incrementButtonAriaLabel="Increase value by 1"
                    decrementButtonAriaLabel="Decrease value by 1"
                    styles={styles}
                />
                <DialogFooter>
                    <DefaultButton onClick={handleClose} text="Cancel" /> {/* Remove the parentheses from handleClose */}
                    <PrimaryButton onClick={handleOk} text="Ok" /> {/* Remove the parentheses from handleClose */}
                    {/* <DefaultButton onClick={ handleClose } text="Cancel" />  */}
                </DialogFooter>
            </Dialog>
        </>
    );
}

export const FilterDialogComponent = (props) => {
    const [isOpenState, setIsOpenState] = useState(props?.isOpen?.openFilterDialog);
    const [selectedOption1, setSelectedOption1] = useState(props?.isOpen?.filterSelectionData?.selectedOption1);
    const [selectedOption2, setSelectedOption2] = useState(props?.isOpen?.filterSelectionData?.selectedOption2);
    const [questionMaster, setQuestionMaster] = useState([]);
    const [cdSectionData, setCDSectionData] = useState([]);
    const [csSectionData, setCSSectionData] = useState([]);
    const [hiddenRow, setHiddenRow] = useState([]);
    const [tableColumnDetails, setTableColumnDetails] = useState(props?.isOpen?.tableColumnDetails);
    const [questionSegreation, setQuestionSegreation] = useState({ "TableCS": [], "TableCD": [] });

    const luckySheet = props?.isOpen?.luckysheet;
    dialogContentProps.subText = "";
    dialogContentProps.title = "Filter";
    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px) and (max-width: 767px)': {
                        height: '200px',
                        width: '300px',
                    },
                    '@media (min-width: 768px) and (max-width: 1023px)': {
                        height: '300px',
                        width: '450px',
                    },
                    '@media (min-width: 1024px)': {
                        // height: '350px', 
                        width: '650px',
                    },
                    '@media (min-width: 480px)': {
                        width: 'auto',
                        maxWidth: '345px',
                        minWidth: '288px',
                    }
                }
            }
        },
    };

    const TableMaster = [
        { key: 'Common Declaration Table', text: 'Common Declaration Table' },
        { key: 'Coverage Specification Table', text: 'Coverage Specification Table' }
    ];

    useEffect(() => {
        setIsOpenState(props?.isOpen?.openFilterDialog);
        processAndUpdateCSCDdata(selectedOption1);
    }, [props?.isOpen?.openFilterDialog]);

    const handleClear = (e) => {
        // setIsOpenState( false );
        // if ( hiddenRow && hiddenRow?.length == 2 )
        // {
        //     luckySheet.showRow( hiddenRow[ 0 ], hiddenRow[ 1 ] )
        // } else
        // {
        const configData = luckySheet.getconfig();
        if (configData && configData?.rowhidden) {
            let hiddenRows = Object.keys(configData?.rowhidden);
            hiddenRows = hiddenRows.map((e) => parseInt(e));
            if (hiddenRows && hiddenRows.length > 0) {
                hiddenRows = groupNumbers(hiddenRows);
                hiddenRows.forEach((f) => {
                    // hiddenRows.forEach( ( f ) => { luckySheet.showRow( f[0], f[f?.length - 1 ] ) } );
                    luckySheet.showRow(f[0], f[f?.length - 1]);
                });
            }
        }
        // }
        // props.onClose( { state: false } );
        setSelectedOption2(null);
        setSelectedOption1(null);
    };

    const handleOk = (e) => {
        getSelectedQuestionRange();
        setIsOpenState(false);
        props.onClose({ state: false, filterData: { selectedOption1, selectedOption2 } });
    };

    const handleClose = (e) => {
        // getSelectedQuestionRange();
        setIsOpenState(false);
        props.onClose({ state: false, filterData: { selectedOption1, selectedOption2 } });
    };

    const onTableSelectionChange = (e, data) => {
        const questionSegreationData = data ? data : questionSegreation;
        if (e && e?.key) {
            if (e?.key === "Common Declaration Table") {
                setQuestionMaster(questionSegreationData["TableCD"]);
            } else if (e?.key === "Coverage Specification Table") {
                setQuestionMaster(questionSegreationData["TableCS"]);
            } else {
                setQuestionMaster([]);
            }
            if (e?.key != selectedOption1?.key) {
                setSelectedOption2(null);
            }
            setSelectedOption1(e);
        }
    };

    const processAndUpdateCSCDdata = (e) => {
        if (luckySheet && tableColumnDetails) {
            const sheetData = luckySheet.getSheetData();
            const tableCDdetails = tableColumnDetails["Table 2"];
            const tableCSdetails = tableColumnDetails["Table 3"];
            if (sheetData && sheetData?.length > 0 && tableCDdetails && tableCSdetails) {
                let dataForState = { "TableCS": [], "TableCD": [] };

                const cdDataSection = sheetData.slice(tableCDdetails?.range?.start + 2, tableCDdetails?.range?.end + 1);
                const csDataSection = sheetData.slice(tableCSdetails?.range?.start + 3, tableCSdetails?.range?.end + 1);
                if (cdDataSection && cdDataSection?.length > 0) {
                    const data = cdDataSection?.map((e) => { return e[2] });
                    const coverageData = cdDataSection?.map((e) => { return e[1] });
                    if (coverageData?.length > 0) {
                        let textArrayOfColumn1 = coverageData.map((e) => {
                            const text1 = getTextWithoutAnyChnages(e);
                            if (text1) {
                                return text1?.trim()?.toLowerCase();
                            }
                            return "";
                        });
                        if (data?.length > 0) {
                            let textArrayOfCD = data.map((e) => {
                                const text = getTextWithoutAnyChnages(e);
                                if (text) {
                                    return text?.trim()?.toLowerCase()?.slice(0, 3);
                                }
                                return "";
                            });
                            const cdDataPosition = textArrayOfCD.map((questionCode, index) => {
                                return {
                                    code: questionCode,
                                    position: tableCDdetails?.range?.start + 2 + index,
                                    text: `${questionCode}  _  ${textArrayOfColumn1[index]}`
                                };
                            });
                            setCDSectionData(cdDataPosition);
                            let uniqueOptionArySet = Array.from(new Set(cdDataPosition.map(option => option.text)));
                            uniqueOptionArySet = uniqueOptionArySet.map((e) => { return e.toUpperCase() });
                            if (uniqueOptionArySet && uniqueOptionArySet?.length > 0) {
                                uniqueOptionArySet = uniqueOptionArySet.map((e) => { return { key: e, text: e } });
                                dataForState["TableCD"] = uniqueOptionArySet;
                            }
                        }
                    }
                }
                if (csDataSection && csDataSection?.length > 0) {
                    const csChecklistQuestionIndex = tableCSdetails?.columnNames?.ChecklistQuestions ? tableCSdetails?.columnNames?.ChecklistQuestions : 2;
                    const covergeColIndex = tableCSdetails?.columnNames?.CoverageSpecificationsMaster ? tableCSdetails?.columnNames?.CoverageSpecificationsMaster : 2;
                    const data = csDataSection?.map((e) => { return e[csChecklistQuestionIndex] });
                    const coverageData = csDataSection?.map((e) => { return e[covergeColIndex] });
                    if (coverageData?.length > 0) {
                        let textArrayOfColumn1 = coverageData.map((e) => {
                            const text = getTextWithoutAnyChnages(e);
                            if (text) {
                                return text?.trim()?.toLowerCase();
                            }
                            return "";
                        });
                        if (data?.length > 0) {
                            let textArrayOfCS = data.map((e) => {
                                const text = getTextWithoutAnyChnages(e);
                                if (text) {
                                    return text?.trim()?.toLowerCase()?.slice(0, 3);
                                }
                                return "";
                            });
                            const csDataPosition = textArrayOfCS.map((questionCode, index) => {
                                return {
                                    code: questionCode,
                                    position: tableCSdetails?.range?.start + 3 + index,
                                    text: `${questionCode}  _  ${textArrayOfColumn1[index]}`
                                };
                            });
                            setCSSectionData(csDataPosition);
                            let uniqueOptionArySet = Array.from(new Set(csDataPosition.map(option => option.text)));
                            uniqueOptionArySet = uniqueOptionArySet.map((e) => { return e.toUpperCase() });
                            if (uniqueOptionArySet && uniqueOptionArySet?.length > 0) {
                                uniqueOptionArySet = uniqueOptionArySet.map((e) => { return { key: e, text: e } });
                                dataForState["TableCS"] = uniqueOptionArySet;
                            }
                        }
                    }
                }

                setQuestionSegreation(dataForState);
                if (e) {
                    onTableSelectionChange(e, dataForState);
                    // setTimeout( () => {
                    //     setSelectedOption2( props?.isOpen?.filterSelectionData?.selectedOption2 );
                    // }, 200 );
                }
            }
        }
    }

    const getSelectedQuestionRange = () => {
        if (selectedOption1 && selectedOption1?.key && selectedOption2 && selectedOption2.key && selectedOption2.text) {
            if (selectedOption1?.key === 'Common Declaration Table') {
                if (cdSectionData && cdSectionData?.length > 0) {
                    const selectedKey = selectedOption2?.text;
                    const getCheckListRowCode = selectedKey.split('_')[0]
                    const keyFilter = cdSectionData.filter((f) => f?.code && f?.code.toUpperCase() == getCheckListRowCode?.trim());
                    const positions = keyFilter?.length > 0 ? keyFilter.map((e) => e?.position) : [];
                    const uniquePositions = Array.from(new Set(positions));
                    if (uniquePositions && uniquePositions?.length > 0) {
                        const isContinutiyMissing = hasMissingNumbers(uniquePositions);
                        if (!isContinutiyMissing) {
                            // luckySheet.hideRow( uniquePositions[ 0 ], uniquePositions[ uniquePositions?.length - 1 ] );
                            let allPositionIndex = cdSectionData.map((e) => e?.position);
                            allPositionIndex = allPositionIndex.filter((f) => !uniquePositions?.includes(f))
                            allPositionIndex = allPositionIndex.sort((a, b) => a - b);
                            allPositionIndex = Array.from(new Set(allPositionIndex));
                            const grouppedPositions = groupNumbers(allPositionIndex);
                            if (grouppedPositions && grouppedPositions?.length > 0) {
                                handleClear();
                                grouppedPositions.forEach((gp) => {
                                    luckySheet.hideRow(gp[0], gp[gp?.length - 1]);
                                });
                            }
                            setHiddenRow([uniquePositions[0], uniquePositions[uniquePositions?.length - 1]]);
                            const sheetdatas = luckySheet.find("Common Declarations");
                            luckySheet.scroll({
                                targetRow: sheetdatas[0].row,
                                targetColumn: 0
                            });
                        }
                    }
                }
            } else if (selectedOption1?.key === 'Coverage Specification Table') {
                if (csSectionData && csSectionData?.length > 0) {
                    const selectedKey = selectedOption2?.text;
                    const getCheckListRowCode = selectedKey.split('_')[0]
                    const keyFilter = csSectionData.filter((f) => f?.code && f?.code.toUpperCase() == getCheckListRowCode?.trim());
                    const positions = keyFilter?.length > 0 ? keyFilter.map((e) => e?.position) : [];
                    const uniquePositions = Array.from(new Set(positions));
                    if (uniquePositions && uniquePositions?.length > 0) {
                        const isContinutiyMissing = hasMissingNumbers(uniquePositions);
                        if (!isContinutiyMissing) {
                            let allPositionIndex = csSectionData.map((e) => e?.position);
                            allPositionIndex = allPositionIndex.filter((f) => !uniquePositions?.includes(f))
                            allPositionIndex = allPositionIndex.sort((a, b) => a - b);
                            allPositionIndex = Array.from(new Set(allPositionIndex));
                            const grouppedPositions = groupNumbers(allPositionIndex);
                            if (grouppedPositions && grouppedPositions?.length > 0) {
                                handleClear();
                                grouppedPositions.forEach((gp) => {
                                    luckySheet.hideRow(gp[0], gp[gp?.length - 1]);
                                });
                            }
                            setHiddenRow([uniquePositions[0], uniquePositions[uniquePositions?.length - 1]]);
                            const sheetdatas = luckySheet.find("CoverageSpecificationsMaster");
                            luckySheet.scroll({
                                targetRow: sheetdatas[0].row,
                                targetColumn: 0
                            });
                        }
                    }
                }
            }
        }
        return { "range": [] };
    }

    const hasMissingNumbers = (arr) => {
        arr.sort((a, b) => a - b);
        for (let i = 1; i < arr.length; i++) {
            if (arr[i] !== arr[i - 1] + 1) {
                return true;
            }
        }
        return false;
    }

    //get the data index apart from the selection

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


    const renderOptionList = (option) => (
        <TooltipHost content={option.tooltip} calloutProps={{ gapSpace: 0 }} >
            <div style={{ fontSize: 'smaller' }}>{option.text}</div>
        </TooltipHost>
    );

    return (
        <>
            <Dialog
                hidden={!isOpenState}
                onDismiss={handleClose}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
                <div className="dropdownContainer">
                    <Dropdown
                        label="Table"
                        options={TableMaster}
                        selectedKey={selectedOption1 ? selectedOption1.key : null}
                        onChange={(event, option) => onTableSelectionChange(option)}
                        placeholder="Select an option"
                        required
                        errorMessage={!selectedOption1 ? "Mandatory field" : ""}
                        onRenderOption={renderOptionList}
                    />
                    <Dropdown
                        label="Questions"
                        options={questionMaster}
                        selectedKey={selectedOption2 ? selectedOption2?.key : null}
                        onChange={(event, option) => setSelectedOption2(option)}
                        placeholder="Select an option"
                        required
                        errorMessage={!selectedOption2 ? "Mandatory field" : ""}
                        onRenderOption={renderOptionList}
                    />
                </div>
                <DialogFooter>
                    <DefaultButton onClick={handleClose} text="Cancel" />
                    <DefaultButton onClick={handleClear} text="Clear Filter" />
                    <PrimaryButton onClick={handleOk} text="Ok" disabled={!selectedOption2?.text} />
                </DialogFooter>
            </Dialog>
        </>
    );
}

export const FilterCsrDialogComponent = (props) => {
    const [isOpenState, setIsOpenState] = useState(props?.isOpen?.openFilterDialog);
    const [selectedOption1, setSelectedOption1] = useState(props?.isOpen?.filterSelectionData?.selectedOption1);
    const [selectedOption2, setSelectedOption2] = useState(props?.isOpen?.filterSelectionData?.selectedOption2);
    const [coverageColmData, setCoverageColmData] = useState(null);
    const [questionMaster, setQuestionMaster] = useState([]);
    const [cdSectionData, setCDSectionData] = useState([]);
    const [csSectionData, setCSSectionData] = useState([]);
    const [hiddenRow, setHiddenRow] = useState([]);
    const [tableColumnDetails, setTableColumnDetails] = useState(props?.isOpen?.tableColumnDetails);
    const [questionSegreation, setQuestionSegreation] = useState({ "TableCS": [], "TableCD": [] });

    const luckySheet = props?.isOpen?.luckysheet;
    let checklist = props?.isOpen?.props?.data;
    dialogContentProps.subText = "";
    dialogContentProps.title = "Filter";
    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px) and (max-width: 767px)': {
                        height: '200px',
                        width: '300px',
                    },
                    '@media (min-width: 768px) and (max-width: 1023px)': {
                        height: '300px',
                        width: '450px',
                    },
                    '@media (min-width: 1024px)': {
                        // height: '350px', 
                        width: '650px',
                    },
                    '@media (min-width: 480px)': {
                        width: 'auto',
                        maxWidth: '345px',
                        minWidth: '288px',
                    }
                }
            }
        },
    };

    const TableMaster = [
        { key: 'Common Declaration Table', text: 'Common Declaration Table' },
        { key: 'Coverage Specification Table', text: 'Coverage Specification Table' }
    ];

    useEffect(() => {
        setIsOpenState(props?.isOpen?.openFilterDialog);
        checkListColFilterFun(selectedOption1);
    }, [props?.isOpen?.openFilterDialog]);

    const handleClear = (e) => {
        const configData = luckySheet.getconfig();
        if (configData && configData?.rowhidden) {
            let hiddenRows = Object.keys(configData?.rowhidden);
            hiddenRows = hiddenRows.map((e) => parseInt(e));
            if (hiddenRows && hiddenRows.length > 0) {
                hiddenRows = groupNumbers(hiddenRows);
                hiddenRows.forEach((f) => {
                    luckySheet.showRow(f[0], f[f?.length - 1]);
                });
            }
        }

        setSelectedOption2(null);
        setSelectedOption1(null);
    };

    const handleOk = (e) => {
        getSelectedQuestionRange();
        setIsOpenState(false);
        props.onClose({ state: false, filterData: { selectedOption1, selectedOption2 } });
    };

    const handleClose = (e) => {
        setIsOpenState(false);
        props.onClose({ state: false, filterData: { selectedOption1, selectedOption2 } });
    };

    const onTableSelectionChange = (e, data) => {
        const questionSegreationData = data ? data : coverageColmData;
        if (e && e?.key) {
            if (e?.key === "Common Declaration Table") {
                setQuestionMaster(questionSegreationData["TableCD"]);
            } else if (e?.key === "Coverage Specification Table") {
                setQuestionMaster(questionSegreationData["TableCS"]);
            } else {
                setQuestionMaster([]);
            }
            if (e?.key != selectedOption1?.key) {
                setSelectedOption2(null);
            }
            setSelectedOption1(e);
        }
    };

    const checkListColFilterFun = (e) => {
        if (checklist && tableColumnDetails) {
            let checkData = checklist;
            let dataPositionForState = { "Table2Ary": [], "Table3Ary": [] };
            let coverageDataValue = { "Table2Ary": [], "Table3Ary": [] };
            let dataForState = { "TableCD": [], "TableCS": [] };
            let dataForCoverageSpecColumn = { "TableCD": [], "TableCS": [] };

            checkData.forEach(table => {
                const tableName = table.Tablename;
                if (tableName != "Table 1" && (tableName === "Table 2" || tableName === "Table 3")) {
                    const tableDetails = tableColumnDetails[tableName];
                    const tableKey = tableName === "Table 2" ? "Table2Ary" : "Table3Ary";

                    if (tableDetails) {
                        table.TemplateData.forEach((item, index) => {
                            const questionCode = item["Checklist Questions"].substring(0, 3);
                            if (questionCode === item["Checklist Questions"].substring(0, 3)) {
                                let coverageText = item["COVERAGE_SPECIFICATIONS_MASTER"];
                                const position = tableName == "Table 2" ? tableDetails.range.start + 2 + index : tableDetails.range.start + 3 + index;
                                dataPositionForState[tableKey].push({ "code": coverageText, "position": position });
                                coverageDataValue[tableKey].push({ "keyVariance": coverageText, "position": position });
                            }
                        });
                    }
                    setCDSectionData(dataPositionForState?.Table2Ary);
                    setCSSectionData(dataPositionForState?.Table3Ary);
                }
            });

            let textArrayOfCD = dataPositionForState.Table2Ary.map(e => e.code);
            let coverageTextTable2 = coverageDataValue.Table2Ary.map(e => e.keyVariance);
            let coverageTextTable3 = coverageDataValue.Table3Ary.map(e => e.keyVariance);
            textArrayOfCD = Array.from(new Set(textArrayOfCD));
            if (textArrayOfCD && textArrayOfCD?.length > 0) {
                textArrayOfCD = textArrayOfCD.map((e) => { return { key: e, text: e } });
                dataForState["TableCD"] = textArrayOfCD;
            }
            coverageTextTable2 = Array.from(new Set(coverageTextTable2));
            if (coverageTextTable2 && coverageTextTable2?.length > 0) {
                coverageTextTable2 = coverageTextTable2.map((e) => { return { key: e, text: e } });
                dataForCoverageSpecColumn["TableCD"] = coverageTextTable2;
            }
            let textArrayOfCS = dataPositionForState.Table3Ary.map(e => e.code);
            textArrayOfCS = Array.from(new Set(textArrayOfCS));
            if (textArrayOfCS && textArrayOfCS?.length > 0) {
                textArrayOfCS = textArrayOfCS.map((e) => { return { key: e, text: e } });
                dataForState["TableCS"] = textArrayOfCS;
            }
            coverageTextTable3 = Array.from(new Set(coverageTextTable3));
            if (coverageTextTable3 && coverageTextTable3?.length > 0) {
                coverageTextTable3 = coverageTextTable3.map((e) => { return { key: e, text: e } });
                dataForCoverageSpecColumn["TableCS"] = coverageTextTable3;
            }

            setQuestionSegreation(dataForState);
            setCoverageColmData(dataForCoverageSpecColumn);
            if (e) {
                onTableSelectionChange(e, dataForCoverageSpecColumn);
            }
        }
    }


    const getSelectedQuestionRange = () => {
        if (selectedOption1 && selectedOption1?.key && selectedOption2 && selectedOption2.key) {
            if (selectedOption1?.key === 'Common Declaration Table') {
                if (cdSectionData && cdSectionData?.length > 0) {
                    const selectedKey = selectedOption2?.key;
                    const keyFilter = cdSectionData.filter((f) => f?.code == selectedKey);
                    const positions = keyFilter?.length > 0 ? keyFilter.map((e) => e?.position) : [];
                    const uniquePositions = Array.from(new Set(positions));
                    if (uniquePositions && uniquePositions?.length > 0) {
                        const isContinutiyMissing = hasMissingNumbers(uniquePositions);
                        if (!isContinutiyMissing) {
                            // luckySheet.hideRow( uniquePositions[ 0 ], uniquePositions[ uniquePositions?.length - 1 ] );
                            let allPositionIndex = cdSectionData.map((e) => e?.position);
                            allPositionIndex = allPositionIndex.filter((f) => !uniquePositions?.includes(f))
                            allPositionIndex = allPositionIndex.sort((a, b) => a - b);
                            allPositionIndex = Array.from(new Set(allPositionIndex));
                            const grouppedPositions = groupNumbers(allPositionIndex);
                            if (grouppedPositions && grouppedPositions?.length > 0) {
                                handleClear();
                                grouppedPositions.forEach((gp) => {
                                    luckySheet.hideRow(gp[0], gp[gp?.length - 1]);
                                });
                            }
                            setHiddenRow([uniquePositions[0], uniquePositions[uniquePositions?.length - 1]]);
                            const sheetdatas = luckySheet.find("Common Declarations");
                            luckySheet.scroll({
                                targetRow: sheetdatas[0].row,
                                targetColumn: 0
                            });
                        }
                    }
                }
            } else if (selectedOption1?.key === 'Coverage Specification Table') {
                if (csSectionData && csSectionData?.length > 0) {
                    const selectedKey = selectedOption2?.key;
                    const keyFilter = csSectionData.filter((f) => f?.code == selectedKey);
                    const positions = keyFilter?.length > 0 ? keyFilter.map((e) => e?.position) : [];
                    const uniquePositions = Array.from(new Set(positions));
                    if (uniquePositions && uniquePositions?.length > 0) {
                        const isContinutiyMissing = hasMissingNumbers(uniquePositions);
                        if (!isContinutiyMissing) {
                            let allPositionIndex = csSectionData.map((e) => e?.position);
                            allPositionIndex = allPositionIndex.filter((f) => !uniquePositions?.includes(f))
                            allPositionIndex = allPositionIndex.sort((a, b) => a - b);
                            allPositionIndex = Array.from(new Set(allPositionIndex));
                            const grouppedPositions = groupNumbers(allPositionIndex);
                            if (grouppedPositions && grouppedPositions?.length > 0) {
                                handleClear();
                                grouppedPositions.forEach((gp) => {
                                    luckySheet.hideRow(gp[0], gp[gp?.length - 1]);
                                });
                            }
                            setHiddenRow([uniquePositions[0], uniquePositions[uniquePositions?.length - 1]]);
                            const sheetdatas = luckySheet.find("COVERAGE_SPECIFICATIONS_MASTER");
                            luckySheet.scroll({
                                targetRow: sheetdatas[0].row,
                                targetColumn: 0
                            });
                        }
                    }
                }
            }
        }
        return { "range": [] };
    }

    const hasMissingNumbers = (arr) => {
        arr.sort((a, b) => a - b);
        for (let i = 1; i < arr.length; i++) {
            if (arr[i] !== arr[i - 1] + 1) {
                return true;
            }
        }
        return false;
    }

    //get the data index apart from the selection

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

    return (
        <>
            <Dialog
                hidden={!isOpenState} // Negate the isOpenState to properly handle visibility
                onDismiss={handleClose} // Remove the parentheses from handleClose
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
                <div className="dropdownContainer">
                    <Dropdown
                        label="Table"
                        options={TableMaster}
                        selectedKey={selectedOption1 ? selectedOption1.key : null}
                        onChange={(event, option) => onTableSelectionChange(option)}
                        placeholder="Select an option"
                        required
                        errorMessage={!selectedOption1 ? "Mandatory field" : ""}
                    />
                    <Dropdown
                        label="Questions"
                        options={questionMaster}
                        selectedKey={selectedOption2 ? selectedOption2?.key : null}
                        onChange={(event, option) => setSelectedOption2(option)}
                        placeholder="Select an option"
                        required
                        errorMessage={!selectedOption2 ? "Mandatory field" : ""}
                    />
                </div>
                <DialogFooter>
                    <DefaultButton onClick={handleClose} text="Cancel" />
                    <DefaultButton onClick={handleClear} text="Clear Filter" />
                    <PrimaryButton onClick={handleOk} text="Ok" disabled={!selectedOption2?.key} />
                </DialogFooter>
            </Dialog>
        </>
    );
}

export const FilterQcDialogComponent = (props) => {
    const [isOpenState, setIsOpenState] = useState(props?.isOpen?.openFilterDialog);
    const [selectedOption1, setSelectedOption1] = useState(props?.isOpen?.filterSelectionData?.selectedOption1);
    const [selectedOption2, setSelectedOption2] = useState(props?.isOpen?.filterSelectionData?.selectedOption2);
    const [coverageColmData, setCoverageColmData] = useState(null);
    const [questionMaster, setQuestionMaster] = useState([]);
    const [cdSectionData, setCDSectionData] = useState([]);
    const [csSectionData, setCSSectionData] = useState([]);
    const [hiddenRow, setHiddenRow] = useState([]);
    const [quoteTableColumnDetails, setQuoteTableColumnDetails] = useState(props?.isOpen?.quoteTableColumnDetails);
    const [questionSegreation, setQuestionSegreation] = useState({ "TableCS": [], "TableCD": [] });

    const luckySheet = props?.isOpen?.luckysheet;
    let checklist = props?.isOpen?.dataSet;
    dialogContentProps.subText = "";
    dialogContentProps.title = "Filter";
    const modalProps = {
        isBlocking: true,
        styles: {
            main: {
                selectors: {
                    '@media (min-width: 0px) and (max-width: 767px)': {
                        height: '200px',
                        width: '300px',
                    },
                    '@media (min-width: 768px) and (max-width: 1023px)': {
                        height: '300px',
                        width: '450px',
                    },
                    '@media (min-width: 1024px)': {
                        width: '650px',
                    },
                    '@media (min-width: 480px)': {
                        width: 'auto',
                        maxWidth: '345px',
                        minWidth: '288px',
                    }
                }
            }
        },
    };

    const TableMaster = [
        { key: 'Common Declaration Table', text: 'Common Declaration Table' },
        { key: 'Coverage Specification Table', text: 'Coverage Specification Table' }
    ];

    useEffect(() => {
        setIsOpenState(props?.isOpen?.openFilterDialog);
        checkListColFilterFun(selectedOption1);
    }, [props?.isOpen?.openFilterDialog]);

    const handleClear = (e) => {
        const configData = luckySheet.getconfig();
        if (configData && configData?.rowhidden) {
            let hiddenRows = Object.keys(configData?.rowhidden);
            hiddenRows = hiddenRows.map((e) => parseInt(e));
            if (hiddenRows && hiddenRows.length > 0) {
                hiddenRows = groupNumbers(hiddenRows);
                hiddenRows.forEach((f) => {
                    luckySheet.showRow(f[0], f[f?.length - 1]);
                });
            }
        }

        setSelectedOption2(null);
        setSelectedOption1(null);
    };

    const handleOk = (e) => {
        getSelectedQuestionRange();
        setIsOpenState(false);
        props.onClose({ state: false, filterData: { selectedOption1, selectedOption2 } });
    };

    const handleClose = (e) => {
        setIsOpenState(false);
        props.onClose({ state: false, filterData: { selectedOption1, selectedOption2 } });
    };

    const onTableSelectionChange = (e, data) => {
        const questionSegreationData = data ? data : coverageColmData;
        if (e && e?.key) {
            if (e?.key === "Common Declaration Table") {
                setQuestionMaster(questionSegreationData["TableCD"]);
            } else if (e?.key === "Coverage Specification Table") {
                setQuestionMaster(questionSegreationData["TableCS"]);
            } else {
                setQuestionMaster([]);
            }
            if (e?.key != selectedOption1?.key) {
                setSelectedOption2(null);
            }
            setSelectedOption1(e);
        }
    };

    const checkListColFilterFun = (e) => {
        if (checklist && quoteTableColumnDetails) {
            let checkData = checklist;
            let dataPositionForState = { "Table2Ary": [], "Table3Ary": [] };
            let coverageDataValue = { "Table2Ary": [], "Table3Ary": [] };
            let dataForState = { "TableCD": [], "TableCS": [] };
            let dataForCoverageSpecColumn = { "TableCD": [], "TableCS": [] };

            checkData.forEach(table => {
                const tableName = table.Tablename;
                if (tableName != "Table 1" && (tableName === "Table 2" || tableName === "Table 3")) {
                    const tableDetails = quoteTableColumnDetails[tableName];
                    const tableKey = tableName === "Table 2" ? "Table2Ary" : "Table3Ary";

                    if (tableDetails) {
                        table.TemplateData.forEach((item, index) => {
                            const questionCode = item["ChecklistQuestion"].substring(0, 3);
                            if (questionCode === item["ChecklistQuestion"].substring(0, 3)) {
                                let coverageText = item["CoverageSpecification"];
                                const position = tableName == "Table 2" ? tableDetails.range.start + 2 + index : tableDetails.range.start + 3 + index;
                                dataPositionForState[tableKey].push({ "code": coverageText, "position": position });
                                coverageDataValue[tableKey].push({ "keyVariance": coverageText, "position": position });
                            }
                        });
                    }
                    setCDSectionData(dataPositionForState?.Table2Ary);
                    setCSSectionData(dataPositionForState?.Table3Ary);
                }
            });

            let textArrayOfCD = dataPositionForState.Table2Ary.map(e => e.code);
            let coverageTextTable2 = coverageDataValue.Table2Ary.map(e => e.keyVariance);
            let coverageTextTable3 = coverageDataValue.Table3Ary.map(e => e.keyVariance);
            textArrayOfCD = Array.from(new Set(textArrayOfCD));
            if (textArrayOfCD && textArrayOfCD?.length > 0) {
                textArrayOfCD = textArrayOfCD.map((e) => { return { key: e, text: e } });
                dataForState["TableCD"] = textArrayOfCD;
            }
            coverageTextTable2 = Array.from(new Set(coverageTextTable2));
            if (coverageTextTable2 && coverageTextTable2?.length > 0) {
                coverageTextTable2 = coverageTextTable2.map((e) => { return { key: e, text: e } });
                dataForCoverageSpecColumn["TableCD"] = coverageTextTable2;
            }
            let textArrayOfCS = dataPositionForState.Table3Ary.map(e => e.code);
            textArrayOfCS = Array.from(new Set(textArrayOfCS));
            if (textArrayOfCS && textArrayOfCS?.length > 0) {
                textArrayOfCS = textArrayOfCS.map((e) => { return { key: e, text: e } });
                dataForState["TableCS"] = textArrayOfCS;
            }
            coverageTextTable3 = Array.from(new Set(coverageTextTable3));
            if (coverageTextTable3 && coverageTextTable3?.length > 0) {
                coverageTextTable3 = coverageTextTable3.map((e) => { return { key: e, text: e } });
                dataForCoverageSpecColumn["TableCS"] = coverageTextTable3;
            }

            setQuestionSegreation(dataForState);
            setCoverageColmData(dataForCoverageSpecColumn);
            if (e) {
                onTableSelectionChange(e, dataForCoverageSpecColumn);
            }
        }
    }


    const getSelectedQuestionRange = () => {
        if (selectedOption1 && selectedOption1?.key && selectedOption2 && selectedOption2.key) {
            if (selectedOption1?.key === 'Common Declaration Table') {
                if (cdSectionData && cdSectionData?.length > 0) {
                    const selectedKey = selectedOption2?.key;
                    const keyFilter = cdSectionData.filter((f) => f?.code == selectedKey);
                    const positions = keyFilter?.length > 0 ? keyFilter.map((e) => e?.position) : [];
                    const uniquePositions = Array.from(new Set(positions));
                    if (uniquePositions && uniquePositions?.length > 0) {
                        const isContinutiyMissing = hasMissingNumbers(uniquePositions);
                        if (!isContinutiyMissing) {
                            // luckySheet.hideRow( uniquePositions[ 0 ], uniquePositions[ uniquePositions?.length - 1 ] );
                            let allPositionIndex = cdSectionData.map((e) => e?.position);
                            allPositionIndex = allPositionIndex.filter((f) => !uniquePositions?.includes(f))
                            allPositionIndex = allPositionIndex.sort((a, b) => a - b);
                            allPositionIndex = Array.from(new Set(allPositionIndex));
                            const grouppedPositions = groupNumbers(allPositionIndex);
                            if (grouppedPositions && grouppedPositions?.length > 0) {
                                handleClear();
                                grouppedPositions.forEach((gp) => {
                                    luckySheet.hideRow(gp[0], gp[gp?.length - 1]);
                                });
                            }
                            setHiddenRow([uniquePositions[0], uniquePositions[uniquePositions?.length - 1]]);
                            const sheetdatas = luckySheet.find("Common Declarations");
                            luckySheet.scroll({
                                targetRow: sheetdatas[0].row,
                                targetColumn: 0
                            });
                        }
                    }
                }
            } else if (selectedOption1?.key === 'Coverage Specification Table') {
                if (csSectionData && csSectionData?.length > 0) {
                    const selectedKey = selectedOption2?.key;
                    const keyFilter = csSectionData.filter((f) => f?.code == selectedKey);
                    const positions = keyFilter?.length > 0 ? keyFilter.map((e) => e?.position) : [];
                    const uniquePositions = Array.from(new Set(positions));
                    if (uniquePositions && uniquePositions?.length > 0) {
                        const isContinutiyMissing = hasMissingNumbers(uniquePositions);
                        if (!isContinutiyMissing) {
                            let allPositionIndex = csSectionData.map((e) => e?.position);
                            allPositionIndex = allPositionIndex.filter((f) => !uniquePositions?.includes(f))
                            allPositionIndex = allPositionIndex.sort((a, b) => a - b);
                            allPositionIndex = Array.from(new Set(allPositionIndex));
                            const grouppedPositions = groupNumbers(allPositionIndex);
                            if (grouppedPositions && grouppedPositions?.length > 0) {
                                handleClear();
                                grouppedPositions.forEach((gp) => {
                                    luckySheet.hideRow(gp[0], gp[gp?.length - 1]);
                                });
                            }
                            setHiddenRow([uniquePositions[0], uniquePositions[uniquePositions?.length - 1]]);
                            const sheetdatas = luckySheet.find("COVERAGE SPECIFICATIONS");
                            luckySheet.scroll({
                                targetRow: sheetdatas[0].row,
                                targetColumn: 0
                            });
                        }
                    }
                }
            }
        }
        return { "range": [] };
    }

    const hasMissingNumbers = (arr) => {
        arr.sort((a, b) => a - b);
        for (let i = 1; i < arr.length; i++) {
            if (arr[i] !== arr[i - 1] + 1) {
                return true;
            }
        }
        return false;
    }

    //get the data index apart from the selection

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

    return (
        <>
            <Dialog
                hidden={!isOpenState} // Negate the isOpenState to properly handle visibility
                onDismiss={handleClose} // Remove the parentheses from handleClose
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
            >
                <div className="dropdownContainer">
                    <Dropdown
                        label="Table"
                        options={TableMaster}
                        selectedKey={selectedOption1 ? selectedOption1.key : null}
                        onChange={(event, option) => onTableSelectionChange(option)}
                        placeholder="Select an option"
                        required
                        errorMessage={!selectedOption1 ? "Mandatory field" : ""}
                    />
                    <Dropdown
                        label="Questions"
                        options={questionMaster}
                        selectedKey={selectedOption2 ? selectedOption2?.key : null}
                        onChange={(event, option) => setSelectedOption2(option)}
                        placeholder="Select an option"
                        required
                        errorMessage={!selectedOption2 ? "Mandatory field" : ""}
                    />
                </div>
                <DialogFooter>
                    <DefaultButton onClick={handleClose} text="Cancel" />
                    <DefaultButton onClick={handleClear} text="Clear Filter" />
                    <PrimaryButton onClick={handleOk} text="Ok" disabled={!selectedOption2?.key} />
                </DialogFooter>
            </Dialog>
        </>
    );
}