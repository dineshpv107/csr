import { getTableApplicationColumns } from './CommonFunctions';

export const gradiationConverter = async (data, JobId, forCsr) => {
    if (forCsr) {
        try {
            data = data.map((e) => {
                const templateData = e["TemplateData"];
                if (templateData && typeof templateData === 'string') {
                    e["TemplateData"] = JSON.parse(templateData);
                }
                return e;
            });
            const dataSetRed = await processDataForRedGreenSheetCSR(data, 0, forCsr);
            const response1 = dataSetRed?.sheetDataResponse;
            const response1Range = groupNumbers(Array.from(new Set(response1.celldata.map(e => e?.r))));
            const dataSetGreen = await processDataForRedGreenSheetCSR(data, 1, forCsr);
            const response2 = dataSetGreen?.sheetDataResponse;
            const response2Range = groupNumbers(Array.from(new Set(response2.celldata.map(e => e?.r))));
            const redTableDetails = {};
            const updateDataRed = [];
            const tb1Red = dataSetRed?.tb1ColumnDetailSet[0];
            const tb1Green = dataSetGreen?.tb1ColumnDetailSet[0];
            let tableRedCounter = 2;
            if (dataSetRed?.columnDetailSet && dataSetRed?.columnDetailSet?.length > 0) {
                const cdSet = dataSetRed?.columnDetailSet.filter((f) => f && Object?.keys(f)?.length > 0);
                cdSet.forEach((f, index) => {
                    const rangeDetails = response1Range[index + 1];
                    // let startValue = rangeDetails[0];
                    // if ("Table " + tableRedCounter === "Table 3") {
                    //     startValue += 1;
                    // }
                    redTableDetails["Table " + tableRedCounter] = {
                        columnNames: f,
                        range: { start: rangeDetails[0], end: rangeDetails[rangeDetails?.length - 1] }
                    };
                    tableRedCounter++;
                    updateDataRed.push({ "TableName": "Table " + (index + 2), "columnNames": f, "range": { "start": rangeDetails[0], "end": rangeDetails[rangeDetails?.length - 1] } });
                });
                const mergedRedSheet = Object.assign({}, tb1Red, redTableDetails);
                sessionStorage.setItem('redTableRangeData', JSON.stringify(mergedRedSheet));
                sessionStorage.setItem('redTableRangeDataUpdate', JSON.stringify(updateDataRed));
            } else {
                const mergedRedSheet = Object.assign({}, tb1Red);
                sessionStorage.setItem('redTableRangeData', JSON.stringify(mergedRedSheet));
            }

            if (dataSetRed?.tableDataForUpdate && dataSetRed.tableDataForUpdate.length > 0) {
                const redTblData = dataSetRed?.tableDataForUpdate;
                const redSheetUpdateProcess = redTblData.map((tableData, index) => {
                    return {
                        "TableName": "Table " + (index + 2),
                        "data": tableData
                    };
                });
                sessionStorage.setItem('redSheetData', JSON.stringify(redSheetUpdateProcess));
            }

            if (dataSetRed?.tableDataForUpdate && dataSetRed.tableDataForUpdate.length > 0) {
                const redTblData = dataSetRed?.tableDataForUpdate;
                const redSheetUpdateProcess = redTblData.map((tableData, index) => {
                    if (Array.isArray(tableData)) {
                        const policyLOBs = tableData.map(item => item["POLICY LOB"] || item["Policy LOB"]).filter(Boolean);
                        return {
                            "TableName": "Table " + (index + 2),
                            "data": tableData,
                            "Policy LOB": policyLOBs[0]
                        };
                    }
                });
                sessionStorage.setItem('redDailogData', JSON.stringify(redSheetUpdateProcess));
            }

            const greenTableDetails = {};
            const updateDataGreen = [];
            let tableGreenCounter = 2;
            if (dataSetGreen?.columnDetailSet && dataSetGreen?.columnDetailSet?.length > 0) {
                const cdSet = dataSetGreen?.columnDetailSet.filter((f) => f && Object?.keys(f)?.length > 0);
                cdSet.forEach((f, index) => {
                    const rangeDetails = response2Range[index + 1];
                    // let startValue = rangeDetails[0];
                    // if ("Table " + tableGreenCounter === "Table 3") {
                    //     startValue += 1;
                    // }
                    greenTableDetails["Table " + tableGreenCounter] = {
                        columnNames: f,
                        range: { start: rangeDetails[0], end: rangeDetails[rangeDetails?.length - 1] }
                    };
                    tableGreenCounter++;
                    updateDataGreen.push({ "TableName": "Table " + (index + 2), "columnNames": f, "range": { "start": rangeDetails[0], "end": rangeDetails[rangeDetails?.length - 1] } });
                });
                const mergedGreenSheet = Object.assign({}, tb1Green, greenTableDetails);
                sessionStorage.setItem('greenTableRangeData', JSON.stringify(mergedGreenSheet));
                sessionStorage.setItem('greenTableRangeDataUpdate', JSON.stringify(updateDataGreen));
            } else {
                const mergedGreenSheet = Object.assign({}, tb1Green);
                sessionStorage.setItem('greenTableRangeData', JSON.stringify(mergedGreenSheet));
            }

            if (dataSetGreen?.tableDataForUpdate && dataSetGreen.tableDataForUpdate.length > 0) {
                const greenTblData = dataSetGreen?.tableDataForUpdate;
                const greenSheetUpdateProcess = greenTblData.map((tableData, index) => {
                    return {
                        "TableName": "Table " + (index + 2),
                        "data": tableData
                    };
                });
                sessionStorage.setItem('greenSheetData', JSON.stringify(greenSheetUpdateProcess));
            }

            if (dataSetGreen?.tableDataForUpdate && dataSetGreen.tableDataForUpdate.length > 0) {
                const greenTblData = dataSetGreen?.tableDataForUpdate;
                const greenSheetUpdateProcess = greenTblData.map((tableData, index) => {
                    if (Array.isArray(tableData)) {
                        const policyLOBs = tableData.map(item => item["POLICY LOB"] || item["Policy LOB"]).filter(Boolean);
                        return {
                            "TableName": "Table " + (index + 2),
                            "data": tableData,
                            "Policy LOB": policyLOBs[0]
                        };
                    }
                });
                sessionStorage.setItem('greenDailogData', JSON.stringify(greenSheetUpdateProcess));
            }

            let redData = JSON.parse(sessionStorage.getItem('redDailogData')) || [];
            let nullFilterRedData = redData.filter(f => f != null);
            nullFilterRedData.forEach((item, index) => {
                item.TableName = `Table ${index + 2}`;
            });

            let greenData = JSON.parse(sessionStorage.getItem('greenDailogData')) || [];
            let nullFilterGreenData = greenData.filter(f => f != null);
            nullFilterGreenData.forEach((item, index) => {
                item.TableName = `Table ${index + 2}`;
            });

            sessionStorage.setItem('redDailogData', JSON.stringify(nullFilterRedData));
            sessionStorage.setItem('greenDailogData', JSON.stringify(nullFilterGreenData));

            // const response2 = await processDataForGreenSheet( data );
            return { response1, response2 };
        } catch (error) {
            // console.log(error);
        }
    } else {
        const response1 = await processDataForRedGreenSheet(data, 0);
        const response2 = await processDataForRedGreenSheet(data, 1);
        // const response2 = await processDataForGreenSheet( data );
        return { response1, response2 };
    }
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

const processDataForRedGreenSheetCSR = async (data, sheetnum, forCsr) => {
    let sheetData = staticStructure();
    let sheetDataResponse = {};
    const keys = data.map((e) => e?.Tablename);
    let isMultipleSplitLob = false;
    let AvailableLobs = [];
    const columnDetailSet = [];
    const tableDataForUpdate = [];
    const tb1ColumnDetailSet = [];
    keys.forEach((key) => {
        if (key === "Table 1") {
            const headerData = data.find(f => f.Tablename === key)?.TemplateData;
            let headerResponse = processHeaderData(sheetData, headerData, forCsr);
            sheetDataResponse = headerResponse?.sheetDataForHeader;
            let columnDetailsOfHeader = headerResponse?.columnDetailss;
            tb1ColumnDetailSet.push(columnDetailsOfHeader);
        }
        else if (key === "Table 2") {
            let commonDeclarationData = data.find(f => f.Tablename === key)?.TemplateData;
            commonDeclarationData = commonDeclarationData.filter((f) => f["SHEETNUM"] == sheetnum);
            const dataCD = processCommonDeclarationData(sheetDataResponse, commonDeclarationData, forCsr);
            columnDetailSet.push(dataCD?.columnDetails);
            tableDataForUpdate.push(dataCD?.tableDataForUpdate);
            sheetDataResponse = dataCD?.sheetDataForCD
        }
        if (key === "Table 3") {
            const dataOfJC = data.find(f => f.Tablename === key);
            let JobCoverageData = dataOfJC?.TemplateData;
            isMultipleSplitLob = dataOfJC?.IsMultipleLobSplit;
            AvailableLobs = dataOfJC?.AvailableLobs;
            JobCoverageData = JobCoverageData.filter((f) => f["SHEETNUM"] == sheetnum);
            const dataJC = processCoverageData(sheetDataResponse, JobCoverageData, forCsr, true);
            columnDetailSet.push(dataJC?.columnDetails);
            tableDataForUpdate.push(dataJC?.tableDataForUpdate);
            sheetDataResponse = dataJC?.sheetDataForCD
        } else if (key != "Table 3" && isMultipleSplitLob && AvailableLobs && AvailableLobs?.length > 1) {
            let JobCoverageData = data.find(f => f.Tablename === key)?.TemplateData;
            const policyLob = JobCoverageData.filter(f => AvailableLobs.includes(f["Policy LOB"]) || AvailableLobs.includes(f["POLICY LOB"]));
            if (policyLob && policyLob?.length > 0) {
                JobCoverageData = JobCoverageData.filter((f) => f["SHEETNUM"] == sheetnum);
                const dataTF = processCoverageData(sheetDataResponse, JobCoverageData, forCsr, true);
                columnDetailSet.push(dataTF?.columnDetails);
                tableDataForUpdate.push(dataTF?.tableDataForUpdate);
                sheetDataResponse = dataTF?.sheetDataForCD
            } else {
                let JobForm1Data = data.find(f => f.Tablename === key)?.TemplateData;
                JobForm1Data = JobForm1Data.filter((f) => f["SHEETNUM"] == sheetnum);
                if (JobForm1Data && JobForm1Data?.length > 0) {
                    const dataTF = processForm1Data(sheetDataResponse, JobForm1Data, forCsr);
                    columnDetailSet.push(dataTF?.columnDetails);
                    tableDataForUpdate.push(dataTF?.tableDataForUpdate);
                    sheetDataResponse = dataTF?.sheetDataForCD
                }
            }

        }
        else if (key === "Table 4") {
            let JobForm1Data = data.find(f => f.Tablename === key)?.TemplateData;
            JobForm1Data = JobForm1Data.filter((f) => f["SHEETNUM"] == sheetnum);
            if (JobForm1Data && JobForm1Data?.length > 0) {
                const dataTF = processForm1Data(sheetDataResponse, JobForm1Data, forCsr);
                columnDetailSet.push(dataTF?.columnDetails);
                tableDataForUpdate.push(dataTF?.tableDataForUpdate);
                sheetDataResponse = dataTF?.sheetDataForCD
            }
        }
        else if (key === "Table 5") {
            let JobForm1Data = data.find(f => f.Tablename === key)?.TemplateData;
            JobForm1Data = JobForm1Data.filter((f) => f["SHEETNUM"] == sheetnum);
            if (JobForm1Data && JobForm1Data?.length > 0) {
                const dataTF = processForm2Data(sheetDataResponse, JobForm1Data, forCsr);
                columnDetailSet.push(dataTF?.columnDetails);
                tableDataForUpdate.push(dataTF?.tableDataForUpdate);
                sheetDataResponse = dataTF?.sheetDataForCD
            }
        }
        else if (key === "Table 6") {
            let JobForm1Data = data.find(f => f.Tablename === key)?.TemplateData;
            JobForm1Data = JobForm1Data.filter((f) => f["SHEETNUM"] == sheetnum);
            if (JobForm1Data && JobForm1Data?.length > 0) {
                const dataTF = processForm3Data(sheetDataResponse, JobForm1Data, forCsr);
                columnDetailSet.push(dataTF?.columnDetails);
                tableDataForUpdate.push(dataTF?.tableDataForUpdate);
                sheetDataResponse = dataTF?.sheetDataForCD
            }
        }
        else if (key === "Table 7") {
            let JobForm1Data = data.find(f => f.Tablename === key)?.TemplateData;
            JobForm1Data = JobForm1Data.filter((f) => f["SHEETNUM"] == sheetnum);
            if (JobForm1Data && JobForm1Data?.length > 0) {
                const dataTF = processForm4Data(sheetDataResponse, JobForm1Data, forCsr);
                columnDetailSet.push(dataTF?.columnDetails);
                tableDataForUpdate.push(dataTF?.tableDataForUpdate);
                sheetDataResponse = dataTF?.sheetDataForCD
            }
        }

    });

    if (sheetnum == 0) {
        sheetDataResponse["name"] = "Red";
    } else {
        sheetDataResponse["name"] = "Green";
    }

    //adding missin rowlen in config
    const rowLenKeys = sheetDataResponse?.config?.rowlen ? Object.keys(sheetDataResponse?.config?.rowlen) : [];
    if (rowLenKeys?.length > 0) {
        const rowLenData = sheetDataResponse?.config?.rowlen;
        for (let index = 0; index < rowLenKeys; index++) {
            if (rowLenData[`${index}`] == undefined || rowLenData[`${index}`] == null) {
                sheetDataResponse.config.rowlen[`${index}`] = 15;
            } else {
                // round of the length calculated values
                sheetDataResponse.config.rowlen[`${index}`] = Math.round(sheetDataResponse.config.rowlen[`${index}`]);
            }
        }
    }

    // return headerResponse;
    return { sheetDataResponse, columnDetailSet, tb1ColumnDetailSet, tableDataForUpdate };
}

const processDataForRedGreenSheet = async (data, sheetnum) => {
    let sheetData = staticStructure();
    const headerData = data["JobHeader"];
    let sheetDataResponse = processHeaderData(sheetData, headerData);
    let commonDeclarationData = data["CommonDeclaration"];
    commonDeclarationData = commonDeclarationData.filter((f) => f["SHEETNUM"] == sheetnum);
    sheetDataResponse = processCommonDeclarationData(sheetDataResponse, commonDeclarationData);
    let JobCoverageData = data["JobCoverage"];
    JobCoverageData = JobCoverageData.filter((f) => f["SHEETNUM"] == sheetnum);
    sheetDataResponse = processCoverageData(sheetDataResponse, JobCoverageData);
    let JobForm1Data = data["JobForm1"];
    let JobForm2Data = data["JobForm2"];
    let JobForm3Data = data["JobForm3"];
    let JobForm4Data = data["JobForm4"];
    JobForm1Data = JobForm1Data.filter((f) => f["SHEETNUM"] == sheetnum);
    JobForm2Data = JobForm2Data.filter((f) => f["SHEETNUM"] == sheetnum);
    JobForm3Data = JobForm3Data.filter((f) => f["SHEETNUM"] == sheetnum);
    JobForm4Data = JobForm4Data.filter((f) => f["SHEETNUM"] == sheetnum);
    if (JobForm1Data && JobForm1Data?.length > 0) {
        sheetDataResponse = processForm1Data(sheetDataResponse, JobForm1Data);
    }
    if (JobForm2Data && JobForm2Data?.length > 0) {
        sheetDataResponse = processForm2Data(sheetDataResponse, JobForm2Data);
    }
    if (JobForm3Data && JobForm3Data?.length > 0) {
        sheetDataResponse = processForm3Data(sheetDataResponse, JobForm3Data);
    }
    if (JobForm4Data && JobForm4Data?.length > 0) {
        sheetDataResponse = processForm4Data(sheetDataResponse, JobForm4Data);
    }

    if (sheetnum == 0) {
        sheetDataResponse["name"] = "Red";
    } else {
        sheetDataResponse["name"] = "Green";
    }

    //adding missin rowlen in config
    const rowLenKeys = sheetDataResponse?.config?.rowlen ? Object.keys(sheetDataResponse?.config?.rowlen) : [];
    if (rowLenKeys?.length > 0) {
        const rowLenData = sheetDataResponse?.config?.rowlen;
        for (let index = 0; index < rowLenKeys; index++) {
            if (rowLenData[`${index}`] == undefined || rowLenData[`${index}`] == null) {
                sheetDataResponse.config.rowlen[`${index}`] = 15;
            } else {
                // round of the length calculated values
                sheetDataResponse.config.rowlen[`${index}`] = Math.round(sheetDataResponse.config.rowlen[`${index}`]);
            }
        }
    }

    // return headerResponse;
    return sheetDataResponse;
}

const processDataForGreenSheet = async (data) => {
    let sheetData = staticStructure();
}

const processHeaderData = (sheetData, data, forCsr) => {
    let sheetDataForHeader = sheetData;
    let columnDetailss = {};
    let fs = 9;
    if (data && Array.isArray(data) && data?.length > 0) {
        let initialStartRow = 2;
        let start = initialStartRow;
        let end = 1;
        let columnNames = {};
        const headerDataSet = [];
        data.forEach((item, index) => {
            if (data?.length === (index + 1)) {
                end = initialStartRow + index;
            }
            columnNames[item["Headers"]] = initialStartRow + index;
            headerDataSet.push({
                "r": initialStartRow + index,
                "c": 1,
                "v": {
                    "ct": {
                        "fa": "@",
                        "t": "inlineStr"
                    },
                    "m": item["Headers"],
                    "v": item["Headers"],
                    "merge": null,
                    "tb": "2",
                    "ff": "\"Tahoma\"",
                    "fs": `${fs}`,
                    "bg": "rgb(139,173,212)",
                }
            });
            headerDataSet.push({
                "r": initialStartRow + index,
                "c": 2,
                "v": {
                    "ct": {
                        "fa": "@",
                        "t": "inlineStr"
                    },
                    "m": forCsr ? item["(No column name)"] : item["NoColumnName"],
                    "v": forCsr ? item["(No column name)"] : item["NoColumnName"],
                    "merge": null,
                    "ff": "\"Tahoma\"",
                    "tb": "2",
                    "fs": `${fs}`
                }
            });
            const text = item["NoColumnName"];
            sheetDataForHeader["config"]["rowlen"][initialStartRow + index] = text == null || text == undefined ? 40 : text?.length > 50 ? (text?.length) / 3 + 5 : 30;
        });
        columnDetailss["Table 1"] = { columnNames, "range": { start, end } };
        sheetDataForHeader["config"]["borderInfo"] = [{
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
                        2,
                        data?.length + 1
                    ],
                    "column": [
                        1,
                        2
                    ],
                    "row_focus": 2,
                    "column_focus": 1
                }
            ]
        }];

        sheetDataForHeader["celldata"] = headerDataSet;
    }
    if (forCsr) {
        return { sheetDataForHeader, columnDetailss };
    } else {
        return sheetDataForHeader;
    }
}

const processCommonDeclarationData = (sheetData, data, forCsr) => {
    let sheetDataForCD = sheetData;
    let columnDetails = {};
    let tableDataForUpdate = {};
    if (data && Array.isArray(data) && data?.length > 0) {
        const policyData = data.map((e) => e["Policy LOB"] || e["POLICY LOB"]);
        const applicableSourceColumn = forCsr ? Object.keys(data[0]).filter(f => !["SHEETNUM", "Checklist Questions", "columnid", "Notes", "Page Number", "Observation", "OBSERVATION", "POLICY LOB", "ActionOnDiscrepancy", "RequestEndorsement", "NotesforEndorsement", "NotesFreeFill", "sheetPosition"].includes(f)) : getTableApplicationColumns("Table 2");
        const columnsWithData = getDataKeys(data);
        let columnsToBeShown = [];

        if (!forCsr) {
            columnsToBeShown = [...["CoverageSpecificationsMaster"], ...columnsToBeShown];
            applicableSourceColumn.forEach(column => {
                if (columnsWithData && Array.isArray(columnsWithData) && columnsWithData?.includes(column)) {
                    columnsToBeShown.push(column);
                }
            });
            columnsToBeShown.push("Observation");
            columnsToBeShown.push("PageNumber");
        } else {
            columnsToBeShown = applicableSourceColumn;
        }
        if (!forCsr) {
            columnsToBeShown.push("DocumentViewer");
        }

        if (columnsToBeShown?.length > 0) {
            if (forCsr) {
                const dataS = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
                sheetDataForCD = dataS?.sheetDataToReturn;
                columnDetails = dataS?.columnDetails;
                tableDataForUpdate = dataS?.data;
            } else {
                sheetDataForCD = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
            }
        }
    }
    if (forCsr) {
        return { sheetDataForCD, columnDetails, tableDataForUpdate };
    } else {
        return sheetDataForCD;
    }
}

const processCoverageData = (sheetData, data, forCsr, needHeader) => {

    let sheetDataForCD = sheetData;
    let columnDetails = {};
    let tableDataForUpdate = {};
    if (data && Array.isArray(data) && data?.length > 0) {
        const policyData = data.map((e) => e["Policy LOB"] || e["POLICY LOB"] || e["PolicyLob"]);
        const applicableSourceColumn = forCsr ? Object.keys(data[0]).filter(f => !["SHEETNUM", "Checklist Questions", "columnid", "Notes", "Page Number", "Observation", "OBSERVATION", "POLICY LOB", "ActionOnDiscrepancy", "RequestEndorsement", "NotesforEndorsement", "NotesFreeFill", "sheetPosition"].includes(f)) : getTableApplicationColumns("Table 2");
        const columnsWithData = getDataKeys(data);
        let columnsToBeShown = [];

        if (!forCsr) {
            columnsToBeShown = [...["CoverageSpecificationsMaster"], ...columnsToBeShown];
            applicableSourceColumn.forEach(column => {
                if (columnsWithData && Array.isArray(columnsWithData) && columnsWithData?.includes(column)) {
                    columnsToBeShown.push(column);
                }
            });
            columnsToBeShown.push("Observation");
            columnsToBeShown.push("PageNumber");
        } else {
            columnsToBeShown = applicableSourceColumn;
        }

        if (!forCsr) {
            columnsToBeShown.push("DocumentViewer");
        }

        if (columnsToBeShown?.length > 0) {
            if (forCsr) {
                const dataS = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, needHeader, policyData);
                sheetDataForCD = dataS?.sheetDataToReturn;
                columnDetails = dataS?.columnDetails;
                tableDataForUpdate = dataS?.data;
            } else {
                sheetDataForCD = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, true, policyData);
            }
        }
    }
    if (forCsr) {
        return { sheetDataForCD, columnDetails, tableDataForUpdate };
    } else {
        return sheetDataForCD;
    }
}

const processForm1Data = (sheetData, data, forCsr) => {
    let sheetDataForCD = sheetData;
    let columnDetails = {};
    let tableDataForUpdate = {};
    if (data && Array.isArray(data) && data?.length > 0) {
        const policyData = data.map((e) => e["Policy LOB"] || e["POLICY LOB"]);
        const applicableSourceColumn = forCsr ? Object.keys(data[0]).filter(f => !["SHEETNUM", "Checklist Questions", "columnid", "Notes", "Page Number", "Observation", "OBSERVATION", "Policy LOB", "POLICY LOB", "ActionOnDiscrepancy", "RequestEndorsement", "NotesforEndorsement", "NotesFreeFill", "sheetPosition"].includes(f)) : getTableApplicationColumns("Table 2");
        const columnsWithData = getDataKeys(data);
        let columnsToBeShown = [];

        if (!forCsr) {
            columnsToBeShown = [...["CoverageSpecificationsMaster"], ...columnsToBeShown];
            applicableSourceColumn.forEach(column => {
                if (columnsWithData && Array.isArray(columnsWithData) && columnsWithData?.includes(column)) {
                    columnsToBeShown.push(column);
                }
            });
            columnsToBeShown.push("Observation");
            columnsToBeShown.push("PageNumber");
        } else {
            columnsToBeShown = applicableSourceColumn;
        }
        if (!forCsr) {
            columnsToBeShown.push("DocumentViewer");
        }

        if (columnsToBeShown?.length > 0) {
            if (forCsr) {
                const dataS = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
                sheetDataForCD = dataS?.sheetDataToReturn;
                columnDetails = dataS?.columnDetails;
                tableDataForUpdate = dataS?.data;
            } else {
                sheetDataForCD = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
            }
        }
    }
    if (forCsr) {
        return { sheetDataForCD, columnDetails, tableDataForUpdate };
    } else {
        return sheetDataForCD;
    }
}

const processForm2Data = (sheetData, data, forCsr) => {
    let sheetDataForCD = sheetData;
    let columnDetails = {};
    let tableDataForUpdate = {};
    if (data && Array.isArray(data) && data?.length > 0) {
        const policyData = data.map((e) => e["Policy LOB"] || e["POLICY LOB"]);
        const applicableSourceColumn = forCsr ? Object.keys(data[0]).filter(f => !["SHEETNUM", "Checklist Questions", "columnid", "Notes", "Page Number", "Observation", "OBSERVATION", "Policy LOB", "POLICY LOB", "ActionOnDiscrepancy", "RequestEndorsement", "NotesforEndorsement", "NotesFreeFill", "sheetPosition"].includes(f)) : getTableApplicationColumns("Table 5");
        const columnsWithData = getDataKeys(data);
        let columnsToBeShown = [];

        if (!forCsr) {
            columnsToBeShown = [...["CoverageSpecificationsMaster"], ...columnsToBeShown];
            applicableSourceColumn.forEach(column => {
                if (columnsWithData && Array.isArray(columnsWithData) && columnsWithData?.includes(column)) {
                    columnsToBeShown.push(column);
                }
            });
            columnsToBeShown.push("Observation");
            columnsToBeShown.push("PageNumber");
        } else {
            columnsToBeShown = applicableSourceColumn;
        }
        if (!forCsr) {
            columnsToBeShown.push("DocumentViewer");
        }
        if (columnsToBeShown?.length > 0) {
            if (forCsr) {
                const dataS = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
                sheetDataForCD = dataS?.sheetDataToReturn;
                columnDetails = dataS?.columnDetails;
                tableDataForUpdate = dataS?.data;
            } else {
                sheetDataForCD = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
            }
        }
    }
    if (forCsr) {
        return { sheetDataForCD, columnDetails, tableDataForUpdate };
    } else {
        return sheetDataForCD;
    }
}

const processForm3Data = (sheetData, data, forCsr) => {
    let sheetDataForCD = sheetData;
    let columnDetails = {};
    let tableDataForUpdate = {};
    if (data && Array.isArray(data) && data?.length > 0) {
        const policyData = data.map((e) => e["Policy LOB"] || e["POLICY LOB"]);
        const applicableSourceColumn = forCsr ? Object.keys(data[0]).filter(f => !["SHEETNUM", "Checklist Questions", "Notes", "columnid", "Page Number", "Observation", "OBSERVATION", "Policy LOB", "POLICY LOB", "ActionOnDiscrepancy", "RequestEndorsement", "NotesforEndorsement", "NotesFreeFill", "sheetPosition"].includes(f)) : getTableApplicationColumns("Table 6");
        const columnsWithData = getDataKeys(data);
        let columnsToBeShown = [];

        if (!forCsr) {
            columnsToBeShown = [...["CoverageSpecificationsMaster"], ...columnsToBeShown];
            applicableSourceColumn.forEach(column => {
                if (columnsWithData && Array.isArray(columnsWithData) && columnsWithData?.includes(column)) {
                    columnsToBeShown.push(column);
                }
            });
            columnsToBeShown.push("Observation");
            columnsToBeShown.push("PageNumber");
        } else {
            columnsToBeShown = applicableSourceColumn;
        }
        if (!forCsr) {
            columnsToBeShown.push("DocumentViewer");
        }

        if (columnsToBeShown?.length > 0) {
            if (forCsr) {
                const dataS = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
                sheetDataForCD = dataS?.sheetDataToReturn;
                columnDetails = dataS?.columnDetails;
                tableDataForUpdate = dataS?.data;
            } else {
                sheetDataForCD = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
            }
        }
    }
    if (forCsr) {
        return { sheetDataForCD, columnDetails, tableDataForUpdate };
    } else {
        return sheetDataForCD;
    }
}

const processForm4Data = (sheetData, data, forCsr) => {
    let sheetDataForCD = sheetData;
    let columnDetails = {};
    let tableDataForUpdate = {};
    if (data && Array.isArray(data) && data?.length > 0) {
        const policyData = data.map((e) => e["Policy LOB"] || e["POLICY LOB"]);
        const applicableSourceColumn = forCsr ? Object.keys(data[0]).filter(f => !["SHEETNUM", "Checklist Questions", "columnid", "Page Number", "Notes", "Observation", "OBSERVATION", "Policy LOB", "POLICY LOB", "ActionOnDiscrepancy", "RequestEndorsement", "NotesforEndorsement", "NotesFreeFill", "sheetPosition"].includes(f)) : getTableApplicationColumns("Table 7");
        const columnsWithData = getDataKeys(data);
        let columnsToBeShown = [];

        if (!forCsr) {
            columnsToBeShown = [...["CoverageSpecificationsMaster"], ...columnsToBeShown];
            applicableSourceColumn.forEach(column => {
                if (columnsWithData && Array.isArray(columnsWithData) && columnsWithData?.includes(column)) {
                    columnsToBeShown.push(column);
                }
            });
            columnsToBeShown.push("Observation");
            columnsToBeShown.push("PageNumber");
        } else {
            columnsToBeShown = applicableSourceColumn;
        }
        if (!forCsr) {
            columnsToBeShown.push("DocumentViewer");
        }

        if (columnsToBeShown?.length > 0) {
            if (forCsr) {
                const dataS = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
                sheetDataForCD = dataS?.sheetDataToReturn;
                columnDetails = dataS?.columnDetails;
                tableDataForUpdate = dataS?.data;
            } else {
                sheetDataForCD = processColumnHeaders(sheetDataForCD, columnsToBeShown, data, applicableSourceColumn, forCsr, false, policyData);
            }
        }
    }
    if (forCsr) {
        return { sheetDataForCD, columnDetails, tableDataForUpdate };
    } else {
        return sheetDataForCD;
    }
}

const getDataKeys = (data) => {
    const keys = [];
    const valuesKeys = Object.keys(data[0]);

    valuesKeys.forEach((f) => {
        const hasData = data.filter((item) => item[f] != null);
        if (hasData?.length > 0) {
            keys.push(f);
        }
    });
    return keys;
}

const processColumnHeaders = (sheetData, columnData, data, applicableSourceColumn, forCsr, needHeader, policyData) => {

    let sheetDataToReturn = sheetData;
    const cellData = sheetData["celldata"];
    const merge = sheetData?.config?.merge ? sheetData?.config?.merge : {};
    const borderInfo = sheetData?.config?.borderInfo ? sheetData?.config?.borderInfo : [];
    const initialRow = cellData ? cellData[cellData?.length - 1]["r"] + (needHeader ? 1 : 0) : 2;
    const headerDataSet = [];
    let columnDetails = {};
    if (columnData && columnData?.length > 0) {
        columnData.forEach((column, index) => {
            if (index == 0 && needHeader === true) {
                headerDataSet.push({
                    "r": initialRow + 2,
                    "c": index + 1,
                    "v": {
                        "ct": {
                            "fa": "@",
                            "t": "inlineStr"
                        },
                        "m": policyData[0],
                        "v": policyData[0],
                        "merge": null,
                        "tb": "2",
                        "fs": 10,
                        "ff": "\"Tahoma\"",
                        "bg": "rgb(139,173,212)"
                    }
                });

                merge[`${initialRow + 2}_${index + 1}`] = {
                    "r": initialRow + 2,
                    "c": index + 1,
                    "rs": 1,
                    "cs": columnData?.length + 4
                }
                borderInfo.push({
                    "rangeType": "range",
                    "borderType": "border-all",
                    "color": "#000",
                    "style": "1",
                    "range": [
                        {
                            "left": 74,
                            "width": 272,
                            "top": 411,
                            "height": 40,
                            "left_move": 74,
                            "width_move": 272,
                            "top_move": 412,
                            "height_move": 39,
                            "row": [
                                initialRow + 2,
                                initialRow + 2
                            ],
                            "column": [
                                1,
                                columnData?.length + 4
                            ],
                            "row_focus": 14,
                            "column_focus": 1
                        }
                    ]
                });
            }
            const policyToBeshown = index == 0 && !needHeader && policyData && policyData[0] ? policyData[0] : column;
            headerDataSet.push({
                "r": initialRow + 3,
                "c": index + 1,
                "v": {
                    "ct": {
                        "fa": "@",
                        "t": "inlineStr"
                    },
                    "m": policyToBeshown,
                    "v": policyToBeshown,
                    "merge": null,
                    "tb": "2",
                    "fs": 10,
                    "ff": "\"Tahoma\"",
                    "bg": "rgb(139,173,212)"
                }
            });
            columnDetails[column] = index + 1;
            borderInfo.push({
                "rangeType": "range",
                "borderType": "border-all",
                "color": "#000",
                "style": "1",
                "range": [
                    {
                        "left": 74,
                        "width": 272,
                        "top": 411,
                        "height": 40,
                        "left_move": 74,
                        "width_move": 272,
                        "top_move": 412,
                        "height_move": 39,
                        "row": [
                            initialRow + 3,
                            initialRow + 4
                        ],
                        "column": [
                            index + 1,
                            index + 1
                        ],
                        "row_focus": 14,
                        "column_focus": 1
                    }
                ]
            });
            merge[`${initialRow + 3}_${index + 1}`] = {
                "r": initialRow + 3,
                "c": index + 1,
                "rs": 2,
                "cs": 1
            }
        });
        //action on descripancy
        headerDataSet.push({
            "r": initialRow + 3,
            "c": columnData?.length + 1,
            "v": {
                "ct": {
                    "fa": "General",
                    "t": "g"
                },
                "m": "Actions on Discrepancy (from AMs)",
                "v": "Actions on Discrepancy (from AMs)",
                "merge": null,
                "ht": 0,
                "tb": "2",
                "ff": "\"Tahoma\"",
                "fs": 10,
                "bg": "rgb(139,173,212)"
            }
        });
        headerDataSet.push({
            "r": initialRow + 4,
            "c": columnData?.length + 1,
            "v": {
                "ct": {
                    "fa": "@",
                    "t": "inlineStr"
                },
                "m": "Actions on Discrepancy",
                "v": "Actions on Discrepancy",
                "merge": null,
                "tb": "2",
                "ff": "\"Tahoma\"",
                "fs": 10,
                "bg": "rgb(139,173,212)"
            }
        });
        columnDetails["Actions on Discrepancy"] = columnData?.length + 1;
        headerDataSet.push({
            "r": initialRow + 4,
            "c": columnData?.length + 2,
            "v": {
                "ct": {
                    "fa": "@",
                    "t": "inlineStr"
                },
                "m": "Request Endorsement",
                "v": "Request Endorsement",
                "merge": null,
                "tb": "2",
                "ff": "\"Tahoma\"",
                "fs": 10,
                "bg": "rgb(139,173,212)"
            }
        });
        columnDetails["Request Endorsement"] = columnData?.length + 2;
        headerDataSet.push({
            "r": initialRow + 4,
            "c": columnData?.length + 3,
            "v": {
                "ct": {
                    "fa": "@",
                    "t": "inlineStr"
                },
                "m": "Notes for Endorsement",
                "v": "Notes for Endorsement",
                "merge": null,
                "tb": "2",
                "ff": "\"Tahoma\"",
                "fs": 10,
                "bg": "rgb(139,173,212)"
            }
        });
        columnDetails["Notes for Endorsement"] = columnData?.length + 3;
        headerDataSet.push({
            "r": initialRow + 4,
            "c": columnData?.length + 4,
            "v": {
                "ct": {
                    "fa": "@",
                    "t": "inlineStr"
                },
                "m": "Notes(Free Fill)",
                "v": "Notes(Free Fill)",
                "merge": null,
                "tb": "2",
                "ff": "\"Tahoma\"",
                "fs": 10,
                "bg": "rgb(139,173,212)"
            }
        });
        columnDetails["Notes(Free Fill)"] = columnData?.length + 4;
        merge[`${initialRow + 3}_${columnData?.length + 1}`] = {
            "r": initialRow + 3,
            "c": columnData?.length + 1,
            "rs": 1,
            "cs": 4
        }
        borderInfo.push({
            "rangeType": "range",
            "borderType": "border-all",
            "color": "#000",
            "style": "1",
            "range": [
                {
                    "left": 74,
                    "width": 272,
                    "top": 411,
                    "height": 40,
                    "left_move": 74,
                    "width_move": 272,
                    "top_move": 412,
                    "height_move": 39,
                    "row": [
                        initialRow + 3,
                        initialRow + 3
                    ],
                    "column": [
                        columnData?.length + 1,
                        columnData?.length + 4
                    ],
                    "row_focus": 14,
                    "column_focus": 1
                }
            ]
        });
        borderInfo.push({
            "rangeType": "range",
            "borderType": "border-all",
            "color": "#000",
            "style": "1",
            "range": [
                {
                    "left": 74,
                    "width": 272,
                    "top": 411,
                    "height": 40,
                    "left_move": 74,
                    "width_move": 272,
                    "top_move": 412,
                    "height_move": 39,
                    "row": [
                        initialRow + 4,
                        initialRow + 4
                    ],
                    "column": [
                        columnData?.length + 1,
                        columnData?.length + 4
                    ],
                    "row_focus": 14,
                    "column_focus": 1
                }
            ]
        });
    }
    sheetDataToReturn["celldata"] = [...cellData, ...headerDataSet];
    sheetDataToReturn["config"]["merge"] = merge;
    sheetDataToReturn["config"]["borderInfo"] = borderInfo;
    sheetDataToReturn = processColumnData(sheetDataToReturn, columnData, data, applicableSourceColumn);
    if (forCsr) {
        return { sheetDataToReturn, columnDetails, data };
    }
    return sheetDataToReturn;
}

const processColumnData = (sheetData, columnData, data, applicableSourceColumn) => {
    const dataProcessedSheetData = sheetData;
    const cellData = sheetData["celldata"];
    const borderInfo = sheetData?.config?.borderInfo ? sheetData?.config?.borderInfo : [];
    const initialRow = cellData ? cellData[cellData?.length - 1]["r"] : 2;
    const dataSet = [];
    let fs = 9;
    let rowIndexForStateData = initialRow + 1; // finding sheetposition for each object by sanjay

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
    if (data?.length > 0) {
        columnData = [...columnData, "ActionOnDiscrepancy", "RequestEndorsement", "Notes", "NotesFreeFill"];
        data.forEach((item, index) => {
            columnData.forEach((column, cIndex) => {
                const dataForSplitCheck = item[column];
                if (dataForSplitCheck || (dataForSplitCheck && !(column === 'ActionOnDiscrepancy' || column === 'RequestEndorsement' || column === 'Notes' || column === 'NotesFreeFill'
                ))) {

                    const splittedData = dataForSplitCheck ? dataForSplitCheck.split('~~') : [];
                    const nonsplittedData = item[column];
                    const sInctData = [];
                    if (column != 'Document Viewer') {
                        if (column === "Prior Term Policy" && item["Prior Term Policy"]?.trim() != item["Current Term Policy"]?.trim()
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
                                sInctData.push({
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
                        if (column === "Binder" && item["Binder"]?.trim() != item["Current Term Policy"]?.trim()
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
                                sInctData.push({
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
                        if (column === "Proposal" && item["Proposal"]?.trim() != item["Current Term Policy"]?.trim()
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
                                sInctData.push({
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
                        if (column === "Quote" && item["Quote"]?.trim() != item["Current Term Policy"]?.trim()
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
                                sInctData.push({
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
                        if (column === "Schedule" && item["Schedule"]?.trim() != item["Current Term Policy"]?.trim()
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
                                sInctData.push({
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

                        else if (column === "Current Term Policy Listed" && item["Current Term Policy Listed"]?.trim() != item["Current Term Policy Attached"]?.trim()
                            && !(item["Current Term Policy Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                || item["Current Term Policy Attached"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                            let ptpSplitArray = item["Current Term Policy Listed"]?.split('~~')[0]?.split(" ");
                            let ctpSplitArray = item["Current Term Policy Attached"]?.split('~~')[0]?.split(" ");

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
                                sInctData.push({
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
                        else if (column === "Current Term Policy Attached" && item["Current Term Policy Attached"]?.trim() != item["Current Term Policy Listed"]?.trim()
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
                                sInctData.push({
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
                        else if (column === "Current Term Policy - Listed" && item["Current Term Policy - Listed"]?.trim() != item["Prior Term Policy - Listed"]?.trim()
                            && !(item["Current Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document")
                                || item["Prior Term Policy - Listed"]?.toLowerCase()?.replace(/\\r\\n/g, '')?.includes("details not available in the document"))) {

                            let ptpSplitArray = item["Current Term Policy - Listed"]?.split('~~')[0]?.split(" ");
                            let ctpSplitArray = item["Prior Term Policy - Listed"]?.split('~~')[0]?.split(" ");

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
                                sInctData.push({
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
                        else if (column === "Application" && item["Application"]?.trim() != item["Current Term Policy"]?.trim()
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
                                sInctData.push({
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
                        else if (column === "Application - Listed" && item["Application - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
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
                                sInctData.push({
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
                        else if (column === "Quote - Listed" && item["Quote - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
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
                                sInctData.push({
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
                        else if (column === "Proposal - Listed" && item["Proposal - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
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
                                sInctData.push({
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
                        else if (column === "Binder - Listed" && item["Binder - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
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
                                sInctData.push({
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
                        else if (column === "Schedule - Listed" && item["Schedule - Listed"]?.trim() != item["Current Term Policy - Listed"]?.trim()
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
                                sInctData.push({
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

                        if (nonsplittedData === 'MATCHED') {
                            sInctData.push({
                                "ff": "\"Tahoma\"",
                                "fc": "rgb(0, 128, 0)",
                                "fs": `${fs}`,
                                "cl": 0,
                                "un": 0,
                                "bl": 0,
                                "it": 0,
                                "v": nonsplittedData.trim() + "\r\n"
                            })
                        }
                        if (column != 'DocumentViewer' && splittedData) {
                            splittedData.forEach((f) => {
                                if (f && f?.length > 0 && (f?.toLowerCase()?.includes('page #') || f?.toLowerCase()?.includes('endorsement page #') ||
                                    f?.toLowerCase()?.includes('current policy attached') || f?.toLowerCase()?.includes('current policy listed') ||
                                    f?.toLowerCase()?.includes('current policy endorsement attached') || f?.toLowerCase()?.includes('current policy endorsement listed'))) {
                                    sInctData.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": "rgb(68, 114, 196)",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 1,
                                        "it": 0,
                                        "v": "\r\n" + f.trim() + "\r\n"
                                    })
                                } else if (f && f?.length > 0 && column === 'PageNumber') {
                                    sInctData.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": "#000000",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": f.trim() + "\r\n"
                                    })
                                }
                                else if (f && f?.length > 0 && sInctData?.length === 0) {
                                    sInctData.push({
                                        "ff": "\"Tahoma\"",
                                        "fc": "#000000",
                                        "fs": `${fs}`,
                                        "cl": 0,
                                        "un": 0,
                                        "bl": 0,
                                        "it": 0,
                                        "v": f.trim() + " "
                                    });
                                }
                            });
                        }
                    } else if (column === 'Document Viewer') {
                        const dvData = item[column];
                        if (dvData != undefined && dvData != null && dvData?.trim() != '') {
                            sInctData.push({
                                "ff": "\"Tahoma\"",
                                "fc": "rgb(61, 133, 198)",
                                "fs": `${fs}`,
                                "cl": 0,
                                "un": 1,
                                "bl": 0,
                                "it": 0,
                                "ht": "0",
                                "v": "X-Ray"
                            });
                        } else {
                            sInctData.push({
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
                    }

                    dataSet.push({
                        "r": rowIndexForStateData + index,  // for sheetposition
                        "c": cIndex + 1,
                        "v": {
                            "ct": {
                                "fa": "General",
                                "t": "inlineStr",
                                "s": sInctData
                            },
                            // "m": item[ column ],
                            // "v": item[ column ],
                            "merge": null,
                            "tb": "2",
                            "ff": "\"Tahoma\"",
                            "fs": `${fs}`,
                            "ht": column === 'Document Viewer' ? "0" : null
                        }
                    });

                    item["sheetPosition"] = rowIndexForStateData;  //finding sheetposition for each object by sanjay
                    // return item;
                } else if (column === 'ActionOnDiscrepancy' || column === 'RequestEndorsement' || column === 'Notes') {
                    dataSet.push({
                        "r": initialRow + 1 + index,
                        "c": cIndex + 1,
                        "v": {
                            "ct": {
                                "fa": "@",
                                "t": "inlineStr",
                                "s": [{
                                    "ff": "\"Tahoma\"",
                                    "fc": "rgba(171, 160, 160, 0.957)",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "ht": "0",
                                    "v": item[column] == '' || item[column] == undefined || item[column] == null ? "Click here" : item[column]
                                }]
                            },
                            // "m": item[ column ],
                            // "v": item[ column ],
                            "merge": null,
                            "ht": column === 'ActionOnDiscrepancy' ? "0" : column === 'RequestEndorsement' ? "0" :
                                column === 'Notes' ? "0" : null,
                            "tb": "2",
                            "fs": `${fs}`
                        }
                    });
                    // const dvData = item[column];
                    // if (dvData == "" || dvData == undefined || dvData == null) {
                    //     sInctData.push({
                    //         "ff": "\"Tahoma\"",
                    //         "fc": "rgba(171, 160, 160, 0.957)",
                    //         "fs": "7",
                    //         "cl": 0,
                    //         "bl": 0,
                    //         "it": 0,
                    //         "ht": "0",
                    //         "v": "Click here"
                    //     });
                    // }
                }
                else {
                    dataSet.push({
                        "r": initialRow + 1 + index,
                        "c": cIndex + 1,
                        "v": {
                            "ct": {
                                "fa": "@",
                                "t": "inlineStr",
                                "s": [{
                                    "ff": "\"Tahoma\"",
                                    "fc": "#000000",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": column == 'Document Viewer' ? " " : item[column] == '' || item[column] == undefined || item[column] == null ? "  " : item[column]
                                }]
                            },
                            // "m": item[ column ],
                            // "v": item[ column ],
                            "merge": null,
                            "tb": "2",
                            "fs": `${fs}`
                        }
                    });
                }

                item["sheetPosition"] = rowIndexForStateData + index;  //finding sheetposition for each object by sanjay

                const existingRowLen = dataProcessedSheetData?.config?.rowlen;
                const text = item[column];
                const rowMaxValue = 30;
                const currentTextLen = text && text?.length > 0 && rowMaxValue < (text?.length / 2) - 50 ? (text?.length / 2) - 30 : rowMaxValue;
                if (existingRowLen != undefined && existingRowLen != null && existingRowLen[`${initialRow + 1 + index}`] < currentTextLen) {
                    dataProcessedSheetData.config.rowlen[`${initialRow + 1 + index}`] = currentTextLen;
                } else {
                    if (existingRowLen[`${initialRow + 1 + index}`] === undefined) {
                        dataProcessedSheetData.config.rowlen[`${initialRow + 1 + index}`] = currentTextLen;
                    }
                    if (existingRowLen == undefined || existingRowLen == null) {
                        dataProcessedSheetData.config["rowlen"][`${initialRow + 1 + index}`] = currentTextLen;
                    }
                }
            });

            if (columnData && columnData?.length > 0) {
                const staticColumn = [1, 2, 3];
                staticColumn.forEach((item, indexOfS) => {
                    dataSet.push({
                        "r": initialRow + 1 + index,
                        "c": columnData?.length + indexOfS + 1,
                        "v": {
                            "ct": {
                                "fa": "@",
                                "t": "inlineStr",
                                "s": [{
                                    "ff": "\"Tahoma\"",
                                    "fc": "#000000",
                                    "fs": `${fs}`,
                                    "cl": 0,
                                    "un": 0,
                                    "bl": 0,
                                    "it": 0,
                                    "v": "  "
                                }]
                            },
                            // "m": item[ column ],
                            // "v": item[ column ],
                            "merge": null,
                            "tb": "2",
                            "fs": `${fs}`
                        }
                    });
                });
            }
            borderInfo.push({
                "rangeType": "range",
                "borderType": "border-all",
                "color": "#000",
                "style": "1",
                "range": [
                    {
                        "left": 74,
                        "width": 272,
                        "top": 411,
                        "height": 40,
                        "left_move": 74,
                        "width_move": 272,
                        "top_move": 412,
                        "height_move": 39,
                        "row": [
                            initialRow + 1 + index,
                            initialRow + 1 + index
                        ],
                        "column": [
                            1,
                            columnData?.length
                        ],
                        "row_focus": 14,
                        "column_focus": 1
                    }
                ]
            });
            return item;
        });
    }

    dataProcessedSheetData["config"]["borderInfo"] = borderInfo;
    dataProcessedSheetData["celldata"] = [...cellData, ...dataSet];
    return dataProcessedSheetData;
}

const staticStructure = () => {
    return {
        name: "", // Worksheet name
        color: "", // Worksheet color
        config: {
            merge: {},
            borderInfo: [],
            rowlen: {},
            columnlen: {
                // "0": 300,
                "1": 272,
                "2": 272,
                "3": 272,
                "4": 272,
                "5": 272,
                "6": 272,
                "7": 272,
                "8": 272,
                "9": 272,
                "10": 272,
                "11": 272,
                "12": 272,
                "13": 272,
                "14": 272
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
    };
}