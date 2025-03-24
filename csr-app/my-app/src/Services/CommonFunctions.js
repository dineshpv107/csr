import { baseUrl, isAdLogin } from '../Services/Constants';
import axios from "axios";
import { qacMasterSet } from '../Services/Constants';
import { updateGridAuditLog } from '../Services/PreviewChecklistDataService';

export const autoupdate = () => {
    const UpdatePeriod = parseInt(sessionStorage.getItem("autoUpdatePeriod"), 10);
    const UpdateEnable = sessionStorage.getItem("autoUpdateEnable") === "true"; // Convert string to boolean

    return { UpdatePeriod, UpdateEnable };
}

// export const csr_autoupdate = () => {
//     const UpdatePeriod = parseInt( sessionStorage.getItem( "Csr_autoUpdatePeriod" ), 10 );
//     const UpdateEnable = sessionStorage.getItem( "Csr_autoUpdatePeriod" ) === "true"; // Convert string to boolean

//     return { UpdatePeriod, UpdateEnable };
// }

export const InsertRowDisabling = (menuItems, clickedRowIndex, sheet, state) => {
    menuItems.forEach((menuItem) => {
        const menuItemText = menuItem?.el?.outerText;
        if (menuItemText === "Insert row") {
            if (clickedRowIndex !== null) {
                const keysAsNumbers = Object.keys(sheet.datas[0].rows._).map(Number);
                for (let i = 0; i <= Math.max(...keysAsNumbers); i++) {
                    if (i === clickedRowIndex) {
                        if (state.rows[i] !== undefined) {
                            if (state.rows[i].cells[1] !== undefined) {
                                if (state.rows[i - 1] !== undefined) {
                                    if (state.rows[i - 1].cells[1] !== undefined) {
                                        menuItem?.el?.classList.add("disabled");
                                        break;
                                    } else {
                                        menuItem?.el?.classList.remove("disabled");
                                        break;
                                    }
                                } else {
                                    menuItem?.el?.classList.remove("disabled");
                                    break;
                                }
                            } else {
                                menuItem?.el?.classList.remove("disabled");
                                break;
                            }
                        } else {
                            menuItem?.el?.classList.remove("disabled");
                            break;
                        }
                    }
                }
            }
        }
    })
}

export const DeleteRowDisabling = (menuItems, clickedRowIndex, sheet, state) => {
    menuItems.forEach((menuItem) => {
        const menuItemText = menuItem?.el?.outerText;
        if (menuItemText === "Delete row") {
            if (clickedRowIndex !== null) {
                const keysAsNumbers = Object.keys(sheet.datas[0].rows._).map(Number);
                for (let i = 0; i <= Math.max(...keysAsNumbers); i++) {
                    if (i === clickedRowIndex) {
                        if (state.rows[i] !== undefined) {
                            if (state.rows[i].cells[1] !== undefined) {
                                menuItem?.el?.classList.add("disabled");
                                break;
                            } else {
                                menuItem?.el?.classList.remove("disabled");
                                break;
                            }
                        } else {
                            menuItem?.el?.classList.remove("disabled");
                            break;
                        }
                    }
                }
            }
        }
    })
}

export const PageHighlighterProcess = async (data, jobId) => {
    if (data) {
        const columnsForPH = {
            "commonDeclaration": [],
            "jobCoverage": [],
            "checkListForm1": [],
            "checkListForm2": [],
            "checkListForm3": [],
            "checkListForm4": []
        };
        if (data["JobCommonDeclaration"]?.length > 0) {
            //find specific columns
            const columnsArray = ["Application", "Binder", "CurrentTermPolicy", "PriorTermPolicy", "Proposal", "Quote", "Schedule"];
            columnsArray.forEach((column) => {
                const hasData = data["JobCommonDeclaration"].filter((f) => f[column]);
                if (hasData.length > 0) {
                    columnsForPH["commonDeclaration"].push(column);
                }
            });
        }
        if (data["JobCoverage"]?.length > 0) {
            //find specific columns
            const columnsArray = ["Application", "Binder", "CurrentTermPolicy", "PriorTermPolicy", "Proposal", "Quote", "Schedule"];
            columnsArray.forEach((column) => {
                const hasData = data["JobCoverage"].filter((f) => f[column]);
                if (hasData.length > 0) {
                    columnsForPH["jobCoverage"].push(column);
                }
            });
        }
        if (data["TblChecklistForm1"]?.length > 0) {
            //find specific columns
            const columnsArray = ["CurrentTermPolicyListed", "PriorTermPolicyListed", "ProposalListed", "BinderListed", "ScheduleListed", "QuoteListed", "ApplicationListed"];
            columnsArray.forEach((column) => {
                const hasData = data["TblChecklistForm1"].filter((f) => f[column]);
                if (hasData.length > 0) {
                    columnsForPH["checkListForm1"].push(column);
                }
            });
        }
        if (data["TblChecklistForm2"]?.length > 0) {
            //find specific columns
            const columnsArray = ["CurrentTermPolicyListed", "PriorTermPolicyListed", "CurrentTermPolicyAttached", "CurrentTermPolicyListed1"];
            columnsArray.forEach((column) => {
                const hasData = data["TblChecklistForm2"].filter((f) => f[column]);
                if (hasData.length > 0) {
                    columnsForPH["checkListForm2"].push(column);
                }
            });
        }
        if (data["TblChecklistForm3"]?.length > 0) {
            //find specific columns
            const columnsArray = ["CurrentTermPolicyAttached", "CurrentTermPolicyListed"];
            columnsArray.forEach((column) => {
                const hasData = data["TblChecklistForm3"].filter((f) => f[column]);
                if (hasData.length > 0) {
                    columnsForPH["checkListForm3"].push(column);
                }
            });
        }
        if (data["TblChecklistForm4"]?.length > 0) {
            //find specific columns
            const columnsArray = ["CurrentTermPolicyAttached", "CurrentTermPolicyListed"];
            columnsArray.forEach((column) => {
                const hasData = data["TblChecklistForm4"].filter((f) => f[column]);
                if (hasData.length > 0) {
                    columnsForPH["checkListForm4"].push(column);
                }
            });
        }
        const jobData = data;

        let valuesToUpdatePH = [];//variable to store the PH vales to be replaced

        if (columnsForPH["commonDeclaration"]?.length > 0) {
            const cdData = CommonFnSplitter(columnsForPH["commonDeclaration"], data, "JobCommonDeclaration", jobId);
            if (cdData.length > 0)
                valuesToUpdatePH = [...valuesToUpdatePH, ...cdData];
        }
        if (columnsForPH["jobCoverage"]?.length > 0) {
            const cdData = CommonFnSplitter(columnsForPH["jobCoverage"], data, "JobCoverage", jobId);
            if (cdData.length > 0)
                valuesToUpdatePH = [...valuesToUpdatePH, ...cdData];
        }
        if (columnsForPH["checkListForm1"]?.length > 0) {
            const cdData = CommonFnSplitter(columnsForPH["checkListForm1"], data, "TblChecklistForm1", jobId);
            if (cdData.length > 0)
                valuesToUpdatePH = [...valuesToUpdatePH, ...cdData];
        }
        if (columnsForPH["checkListForm2"]?.length > 0) {
            const cdData = CommonFnSplitter(columnsForPH["checkListForm2"], data, "TblChecklistForm2", jobId);
            if (cdData.length > 0)
                valuesToUpdatePH = [...valuesToUpdatePH, ...cdData];
        }
        if (columnsForPH["checkListForm3"]?.length > 0) {
            const cdData = CommonFnSplitter(columnsForPH["checkListForm3"], data, "TblChecklistForm3", jobId);
            if (cdData.length > 0)
                valuesToUpdatePH = [...valuesToUpdatePH, ...cdData];
        }
        if (columnsForPH["checkListForm4"]?.length > 0) {
            const cdData = CommonFnSplitter(columnsForPH["checkListForm4"], data, "TblChecklistForm4", jobId);
            if (cdData.length > 0)
                valuesToUpdatePH = [...valuesToUpdatePH, ...cdData];
        }
        //console.log( valuesToUpdatePH );
        return valuesToUpdatePH;
    }
    return [];
}

//based on the column and the config data and keys will change here
const CommonFnSplitter = (data, jobData, tableName, jobId) => {
    let constructedData = [];
    data.forEach((column) => {
        if (column == "CurrentTermPolicy") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Cp", "Current Term Policy", "Page")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "PriorTermPolicy") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Pp", "Prior Term Policy", "Page")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "CurrentTermPolicyListed") {
            const label = tableName == "TblChecklistForm1" ? "Cp" : "Cpl";
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, label, "Current Term Policy", "Current policy listed")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "CurrentTermPolicyListed1") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Cp", "Current Term Policy", "Current policy listed")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "CurrentTermPolicyAttached") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Cpa", "Current Term Policy", "Current policy attached")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "PriorTermPolicyListed") {
            if (tableName == "TblChecklistForm2" || tableName == "TblChecklistForm1") {
                const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Pp", "Prior Term Policy", "Current policy listed")
                constructedData = [...constructedData, ...columnMappingData];
            } else {
                const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Ppl", "Prior Term Policy", "Current policy listed")
                constructedData = [...constructedData, ...columnMappingData];
            }
        }
        if (column == "Quote" || column == "QuoteListed") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Quo", "Carrier Quote", "Page")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "Proposal" || column == "ProposalListed") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Prop", "Proposal", "Page")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "Binder" || column == "BinderListed") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Bin", "Binder", "Page")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "Application" || column == "ApplicationListed") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "App", "Acord Application", "Page")
            constructedData = [...constructedData, ...columnMappingData];
        }
        if (column == "Schedule" || column == "ScheduleListed") {
            const columnMappingData = getPageNumber(jobId, jobData, tableName, column, "Sch", "Schedule", "Page")
            constructedData = [...constructedData, ...columnMappingData];
        }
    });
    return constructedData;
}

//common fn for commonDeclaration
/*
jobData -> total jobDataSet
tableName -> specific name of the table to be picked up
column -> specific name of the column to be picked up
key -> PrefixDoc name ex:Cc,Pp,..
key2 -> name of the column with space for which we are generating the pages for(file for value)
key3 -> keyword to get the pageNumber by split()
*/
const getPageNumber = (jobId, jobData, tableName, column, key, key2, key3) => {
    const constructedData = [];
    jobData[tableName]?.forEach((e) => {
        let backupKey2 = key2;
        const text = e[column];
        let headerKey = key;
        if (text?.toLowerCase()?.trim() != "~~matched~~" && text?.toLowerCase()?.trim() != "matched" &&
            text?.toLowerCase()?.trim() != "details not available in the document" && text?.toLowerCase()?.trim() != "details not available in") {
            let splitKey = '';
            const tags = getTableApplicationColumns("endorsement");
            // if ( key3 != "Page" && text?.includes( key3 ) )
            // {
            //     splitKey = key3;
            // }
            if (key3 === "Page" && !splitKey &&
                (column == "CurrentTermPolicy" || column == "PriorTermPolicy")) {
                tags?.forEach((f) => {
                    if (text && text?.includes(f) && !splitKey) {
                        splitKey = f;
                    }
                });
            }
            const tags1 = getTableApplicationColumns("endorsementCPLA");
            if (key3 != "Page" && !splitKey &&
                (column == "CurrentTermPolicyListed" || column == "CurrentTermPolicyAttached" || column == "CurrentTermPolicyListed1" || column == "PriorTermPolicyListed")) {
                tags1?.forEach((f) => {
                    if (text && text?.includes(f) && !splitKey) {
                        splitKey = f;
                    }
                });
            }
            if (!splitKey) {
                splitKey = 'Page';
            }
            const splittedText = text?.split(splitKey);
            const pageNumber = splittedText[splittedText?.length - 1]?.match(/\d+/g);
            if (Array.isArray(pageNumber) && pageNumber?.length == 1) {
                // const keyObjectData = jobData[ "TblPHData" ].find( ( f ) => f?.FileFor == key2 );
                if (backupKey2) {
                    if (splitKey?.toLowerCase()?.includes('endorsement')) {
                        if (backupKey2 == "Prior Term Policy") {
                            backupKey2 = "Prior Endorsement";
                        } else if (backupKey2 == "Current Term Policy") {
                            backupKey2 = "Current Endorsement";
                        }
                    }
                }
                const keyObjectData = jobData["JobFileInfo"]?.find((f) => f?.JobId === jobId && f?.FileFor === backupKey2);
                let constructedPageNumber = e["ChecklistQuestions"]?.trim().slice(0, 2);
                // constructedPageNumber = key + constructedPageNumber + ":" + pageNumber;
                if (headerKey && splitKey && (splitKey == "Endorsement Page" || splitKey?.toLowerCase()?.includes('endorsement listed') ||
                    splitKey?.toLowerCase()?.includes('endorsement attached'))) {
                    if (headerKey == "Cpa") {
                        headerKey = "CpEa"
                    } else if (headerKey == "Cpl") {
                        headerKey = "CpEl"
                    } else {
                        headerKey += 'E';
                    }
                }
                constructedPageNumber = headerKey + constructedPageNumber;
                const valueText = splittedText[0];
                constructedData.push({
                    "Valuetext": valueText ? valueText?.replace(/~/g, '')?.trim() : '',
                    "PageNumber": pageNumber[0],
                    "FileFor": backupKey2,
                    "JobId": jobId,
                    "FileName": keyObjectData?.FileName,
                    "Filepath": keyObjectData?.FilePath,
                    "Doc": constructedPageNumber
                });
            }
        }
    });
    return constructedData;
}

export const processAndUpdateToken = async (token, needAuthorizationCheck, userName) => {
    let sessionToken = sessionStorage.getItem("token");
    try {
        const decodedToken = token ? JSON.parse(atob(token?.split(".")[1])) : null;
        const decodedSessionToken = sessionToken ? JSON.parse(atob(sessionToken?.split(".")[1])) : null;
        if (needAuthorizationCheck) {
            if (isAdLogin) {                
                const adUser = sessionStorage.getItem('userName');
                const response = await axios.get(baseUrl + '/api/Authentication/CheckUserRoles?userName=' + adUser, {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                    }
                });
                if (response.status !== 200) {
                    throw new Error(`HTTP error! Status: ${response.status}`);
                }
                if (response?.data?.length == 0) {
                    sessionStorage.removeItem('userName');
                    window.location.href = '/AccessDenied';
                    return;
                }
                sessionStorage.setItem('userName', adUser);
            } else {
                const response = await axios.get(baseUrl + '/api/Authentication/GetUserSession?userName=' + userName, {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                    }
                });
                if (response.status !== 200) {
                    throw new Error(`HTTP error! Status: ${response.status}`);
                }
                if (!response?.data?.isUserValid) {
                    sessionStorage.removeItem('userName');
                    window.location.href = '/UnAuthorizedUser';
                    return;
                }
                sessionStorage.setItem('userName', userName);
            }
        }
        if (decodedToken && decodedToken?.exp && new Date((decodedToken?.exp * 1000) - 100000) >= new Date()) {
            return token;
        } else if (decodedSessionToken && decodedSessionToken?.exp && new Date((decodedSessionToken?.exp * 1000) - 100000) >= new Date()) {
            return sessionToken;
        } else {
            try {
                const loginData = {
                    UserId: 0,
                    UserName: "Exdion",
                    Password: "Exdion@123"
                };
                const response = await axios.post(baseUrl + '/api/Authentication/login', loginData, {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                    }
                });
                if (response.status !== 200) {
                    sessionStorage.setItem("LoginTokenError", response?.data);
                    throw new Error(`HTTP error! Status: ${response.status}`);
                }
                if (response?.data) {
                    sessionStorage.setItem("token", response?.data);
                }
                return response.data;
            } catch (error) {
                console.log(error);
                sessionStorage.setItem("LoginTokenError", (error?.response?.data || error?.data));
                throw error; // Rethrow the error to be caught in the calling function
            } finally {
                // document.body.classList.remove( 'loading-indicator' );
            }
        }
    } catch (error) {
        return token;
    }

}

export const autoPopulateCTPT = (row, column, newValue, oldValue, jobData, currentSheetData) => {
    if (row && column) {
        const TableColumnData = []; // variable to store the table and the respective column details
        const tableColumnHeaderIndex = []; //variable to store the position of the column header
        jobData.forEach((element) => {
            if (element?.Tablename != "Table 1") {
                const parsedData = JSON.parse(element?.TemplateData);
                if (parsedData && parsedData?.length > 0) {
                    const objectKeys = Object.keys(parsedData[0]);
                    if (objectKeys.length > 0) {
                        const tableName = element?.Tablename
                        TableColumnData.push({ [tableName]: objectKeys });
                    }
                }
            }
        });
        // console.log( TableColumnData );
    }
    return false;
}

export const tableDataFormatting = (data, tableIndex) => {
    let formattedData = data.map((e) => {
        const keys = Object.keys(e);
        keys.map((k) => {
            if (e[k] && typeof e[k] === 'string') {
                e[k] = e[k]?.replace(/\s+/g, ' ');
                e[k] = e[k]?.replace(/~{3,}/, '~~');
                e[k] = e[k]?.trim()?.replace(/~~$/, '');
                if (tableIndex == 2 || tableIndex == 3) {
                    if (k == "CurrentTermPolicy" || k == "PriorTermPolicy") {
                        e[k] = endKeyCheck(e[k]);
                    } else {
                        e[k] = pageKeyCheck(e[k]);
                    }
                } else if (tableIndex > 3 && tableIndex <= 7) {
                    const formPageKey = k === "CurrentTermPolicyListed" ? "Current policy listed" : k === "CurrentTermPolicyAttached" ? "Current policy attached" : "Page";
                    const textE = e[k];
                    if ((k === "CurrentTermPolicyListed" || k === "CurrentTermPolicyListed1" || k == "PriorTermPolicyListed" || k == "CurrentTermPolicyAttached") && textE && textE?.toLowerCase()?.includes('endorsement')) {
                        e[k] = endKeyCheck(e[k]);
                    }
                    else if (k === "CurrentTermPolicyListed") {
                        if (e[k]?.includes("page") || e[k]?.includes("Page")) {
                            e[k] = pageKeyCheck(e[k]);
                        } else {
                            e[k] = cplKeyCheck(e[k]);
                        }
                    } else if (k === "CurrentTermPolicyAttached") {
                        if (e[k]?.includes("page") || e[k]?.includes("Page")) {
                            e[k] = pageKeyCheck(e[k]);
                        } else {
                            e[k] = cpaKeyCheck(e[k]);
                        }
                    } else {
                        e[k] = pageKeyCheck(e[k]);
                    }
                }
            }
        });
        return e;
    });
    formattedData = formattedData.map((e, index) => {
        e["Columnid"] = index;
        return e;
    });

    formattedData.forEach(obj => {
        Object.keys(obj).forEach(key => {
            if (typeof obj[key] === 'string' && obj[key].startsWith('~')) {
                // Remove one or more "~" symbols at the start of the value
                obj[key] = obj[key].replace(/^~+/, '');
            }
        });
    });
    return formattedData;
}

export const formTableDataFormatting = (data, tableIndex) => {
    let formattedData = data.map((e) => {
        const keys = Object.keys(e);
        keys.map((k) => {
            if (typeof e[k] === 'string') {
                e[k] = e[k]?.replace(/\s+/g, ' ');
                e[k] = e[k]?.replace(/~{3,}/, '~~');
                e[k] = e[k]?.trim()?.replace(/~~$/, '');
                if (tableIndex == 2) {
                    if (k == "CurrentTermPolicyAttached" || k == "PriorTermPolicyAttached") {
                        e[k] = pageKeyCheck(e[k]);
                    }
                }
            }
        });
        return e;
    });
    formattedData.forEach(obj => {
        Object.keys(obj).forEach(key => {
            if (typeof obj[key] === 'string' && obj[key].startsWith('~')) {
                // Remove one or more "~" symbols at the start of the value
                obj[key] = obj[key].replace(/^~+/, '');
            }
        });
    });
    return formattedData;
}

// Function to remove extra colons after a specified key in a string for cpl/cpa
function removeExtraColons(inputString, key) {
    // Create a regular expression that matches the key followed by spaces and multiple colons
    const regex = new RegExp(`(${key}\\s*):(?:\\s*:+)+`, 'g');

    // Replace the matched key followed by multiple colons and spaces with the key followed by a single colon
    const result = inputString.replace(regex, `$1:`);

    return result;
}

const pageKeyCheck = (keyData) => {
    const pageKey = keyData?.includes("page") ? "page" : keyData?.includes("Page") ? "Page" : "";
    if (pageKey) {
        const matchedKey = pageKey == "page" ? keyData?.match(/(~~\s*page\s*#*)/) : keyData?.match(/(~~\s*Page\s*#*)/);
        if (matchedKey != null && matchedKey?.length > 0) {
            //has page key
        } else {
            const hasAshSymbol = pageKey == "page" ? keyData.match(/page\s*#\s*\d+/) : keyData.match(/Page\s*#\s*\d+/);
            if (hasAshSymbol && hasAshSymbol[0]) {
                keyData = keyData.replace(pageKey, "~~Page")
            } else {
                keyData = keyData.replace(pageKey, "~~Page")
            }
        }
    }
    return keyData;
}
const cplKeyCheck = (keyData) => {
    const cplKeys = ["Current policy listed", "current policy listed", "Current policy endorsement listed", "current policy endorsement listed"];
    let pageKey = '';
    cplKeys.forEach((k) => {
        if (!pageKey && keyData?.includes(k)) {
            pageKey = k;
        }
    });
    // const pageKey = keyData?.includes( "Current policy listed" ) ? "Current policy listed" : keyData?.includes( "current policy listed" ) ? "current policy listed" : "";
    if (pageKey) {
        let matchedKey = '';
        let hasEndorsement = false;
        if (pageKey == "Current policy listed") {
            matchedKey = keyData?.match(/(~~\s*Current policy listed\s*:*)/);
        } else if (pageKey == "current policy listed") {
            matchedKey = keyData?.match(/(~~\s*current policy listed\s*:*)/);
        } else if (pageKey == "Current policy endorsement listed") {
            hasEndorsement = true;
            matchedKey = keyData?.match(/(~~\s*Current policy endorsement listed\s*:*)/);
        } else {
            hasEndorsement = true;
            matchedKey = keyData?.match(/(~~\s*current policy endorsement listed\s*:*)/);
        }
        // pageKey == "Current policy listed" ? keyData?.match( /(~~\s*Current policy listed\s*:*)/ ) : keyData?.match( /(~~\s*current policy listed\s*:*)/ );
        if (matchedKey != null && matchedKey?.length > 0) {
            //has page key
        } else {
            let hasAshSymbol = '';
            if (pageKey == "Current policy listed") {
                hasAshSymbol = keyData.match(/current policy listed\s*:\s*\d+/);
            } else if (pageKey == "current policy listed") {
                hasAshSymbol = keyData.match(/current policy listed\s*:\s*\d+/);
            } else if (pageKey == "Current policy endorsement listed") {
                hasAshSymbol = keyData.match(/current policy endorsement listed\s*:\s*\d+/);
            } else {
                hasAshSymbol = keyData.match(/current policy endorsement listed\s*:\s*\d+/);
            }
            //  pageKey == "current policy listed" ? keyData.match( /current policy listed\s*:\s*\d+/ ) : keyData.match( /Current policy listed\s*:\s*\d+/ );
            if (hasAshSymbol && hasAshSymbol[0]) {
                keyData = hasEndorsement ? keyData.replace(pageKey, "~~Current policy endorsement listed") : keyData.replace(pageKey, "~~Current policy listed")
            } else {
                keyData = hasEndorsement ? keyData.replace(pageKey, "~~Current policy endorsement listed: ") : keyData.replace(pageKey, "~~Current policy listed: ")
            }
        }
    }
    try {
        const cleanedString = removeExtraColons(keyData, pageKey);
        return cleanedString;
    } catch (error) {
        return keyData;
    }
}
const cpaKeyCheck = (keyData) => {
    const cplKeys = ["Current policy attached", "current policy attached", "Current policy endorsement attached", "current policy endorsement attached"];
    let pageKey = '';
    cplKeys.forEach((k) => {
        if (!pageKey && keyData?.includes(k)) {
            pageKey = k;
        }
    });
    // const pageKey = keyData?.includes( "Current policy attached" ) ? "Current policy attached" : keyData?.includes( "current policy attached" ) ? "current policy attached" : "";
    if (pageKey) {
        let matchedKey = '';
        let hasEndorsement = false;
        if (pageKey == "Current policy attached") {
            matchedKey = keyData?.match(/(~~\s*Current policy attached\s*:*)/);
        } else if (pageKey == "current policy attached") {
            matchedKey = keyData?.match(/(~~\s*current policy attached\s*:*)/);
        } else if (pageKey == "Current policy endorsement attached") {
            hasEndorsement = true;
            matchedKey = keyData?.match(/(~~\s*Current policy endorsement attached\s*:*)/);
        } else {
            hasEndorsement = true;
            matchedKey = keyData?.match(/(~~\s*current policy endorsement attached\s*:*)/);
        }
        // const matchedKey = pageKey == "Current policy attached" ? keyData?.match( /(~~\s*Current policy attached\s*:*)/ ) : keyData?.match( /(~~\s*current policy attached\s*:*)/ );
        if (matchedKey != null && matchedKey?.length > 0) {
            //has page key
        } else {
            let hasAshSymbol = '';
            if (pageKey == "Current policy attached") {
                hasAshSymbol = keyData.match(/current policy attached\s*:\s*\d+/);
            } else if (pageKey == "current policy attached") {
                hasAshSymbol = keyData.match(/current policy attached\s*:\s*\d+/);
            } else if (pageKey == "Current policy endorsement attached") {
                hasAshSymbol = keyData.match(/current policy endorsement attached\s*:\s*\d+/);
            } else {
                hasAshSymbol = keyData.match(/current policy endorsement attached\s*:\s*\d+/);
            }
            // const hasAshSymbol = pageKey == "current policy attached" ? keyData.match( /current policy attached\s*:\s*\d+/ ) : keyData.match( /Current policy attached\s*:\s*\d+/ );
            if (hasAshSymbol && hasAshSymbol[0]) {
                keyData = hasEndorsement ? keyData.replace(pageKey, "~~Current policy endorsement attached") : keyData.replace(pageKey, "~~Current policy attached")
            } else {
                keyData = hasEndorsement ? keyData.replace(pageKey, "~~Current policy endorsement attached: ") : keyData.replace(pageKey, "~~Current policy attached: ")
            }
        }
    }
    try {
        const cleanedString = removeExtraColons(keyData, pageKey);
        return cleanedString;
    } catch (error) {
        return keyData;
    }
}
const endKeyCheck = (keyData) => {
    const pageKey = keyData?.includes("Endorsement Page") ? "Endorsement Page" : keyData?.includes("Endorsement page") ?
        "Endorsement page" : keyData?.includes("endorsement page") ? "endorsement page" : keyData?.includes("endorsement Page") ? "endorsement Page" : "";
    if (pageKey) {
        let matchedKey = '';
        if (pageKey == "Endorsement Page") {
            matchedKey = keyData?.match(/(~~\s*Endorsement Page\s*#*)/)
        } else if (pageKey == "Endorsement page") {
            matchedKey = keyData?.match(/(~~\s*Endorsement page\s*#*)/)
        } else if (pageKey == "endorsement Page") {
            matchedKey = keyData?.match(/(~~\s*endorsement Page\s*#*)/)
        } else {
            matchedKey = keyData?.match(/(~~\s*endorsement page\s*#*)/)
        }
        // if ( matchedKey && matchedKey[ 0 ] )
        // {
        //     keyData = keyData.replace( pageKey, "~~Endorsement Page" )
        // } else
        // {
        //     keyData = keyData.replace( pageKey, "~~Endorsement Page # " )
        // }
        if (matchedKey != null && matchedKey?.length > 0) {
            keyData = keyData.replace(pageKey, "~~Endorsement Page")
            //has page key
        } else {
            let hasAshSymbol = pageKey == "current policy attached" ? keyData.match(/current policy attached\s*:\s*\d+/) : keyData.match(/Current policy attached\s*:\s*\d+/);
            if (pageKey == "Endorsement Page") {
                hasAshSymbol = keyData.match(/Endorsement Page\s*#\s*\d+/);
            } else if (pageKey == "Endorsement page") {
                hasAshSymbol = keyData.match(/Endorsement page\s*#\s*\d+/);
            } else if (pageKey == "endorsement Page") {
                hasAshSymbol = keyData.match(/endorsement Page\s*#\s*\d+/);
            } else {
                hasAshSymbol = keyData.match(/endorsement page\s*#\s*\d+/);
            }
            if (hasAshSymbol && hasAshSymbol[0]) {
                keyData = keyData.replace(pageKey, "~~Endorsement Page")
            } else {
                keyData = keyData.replace(pageKey, "~~Endorsement Page # ")
            }
        }
    } else {
        return pageKeyCheck(keyData);
    }
    keyData = keyData ? keyData?.replace(/~{3,}/g, '~~') : keyData;
    return keyData;
}

export const getObervationReplacerKey = (text, key) => {
    const inputString = text;

    // Define regular expressions to capture prior and current term contents
    const currentTermRegex = /Current term content\(s\):(.*?)Prior term content\(s\):/;
    const priorTermRegex = /Prior term content\(s\):(.*?)$/;

    // Extract current term content
    const currentTermMatch = inputString.match(currentTermRegex);
    const currentTermContent = currentTermMatch ? currentTermMatch[1].trim() : null;

    // Extract prior term content
    const priorTermMatch = inputString.match(priorTermRegex);
    const priorTermContent = priorTermMatch ? priorTermMatch[1].trim() : null;

    // console.log( currentTermContent );
    // console.log( priorTermContent );
    if (key == "currentTerm") {
        return currentTermContent;
    }
    if (key == "priorTerm") {
        return priorTermContent?.trim();
    }
}

export const getText = (data, noNeedReplacer) => {
    if (data?.ct?.s && data?.ct?.s?.length > 0) {
        let text = '';
        data?.ct?.s?.forEach((f) => {
            text += f?.v;
        });
        text = !noNeedReplacer ? text?.replace(/\r\n/g, '')?.trim() : text?.trim();
        return text;
    } else if (!data?.ct?.s && (data?.m || data?.v)) {
        return data?.m;
    }
    return '';
}

export const getTextForUpdate = (data, noNeedReplacer) => {
    if (data?.ct?.s && data?.ct?.s?.length > 0) {
        let text = '';
        data?.ct?.s?.forEach((f) => {
            text += f?.v;
        });
        text = noNeedReplacer ? text?.replace(/\r\n/g, '~~')?.trim() : text?.trim();
        return text;
    } else if (!data?.ct?.s && (data?.m || data?.v)) {
        return data?.m;
    }
    return '';
}

export const getTextByRequirement = (text, key, splitKey) => {
    if (key == "question") {
        if (text && text?.length > 0) {
            const questionKey = text?.trim()?.slice(0, 2);
            return questionKey != undefined || questionKey != null ? questionKey?.replace(/[^a-zA-Z0-9 ]/g, '')?.trim() : '';
        } else {
            return '';
        }
    } else if (key == "getPage") {
        if (text && text?.length > 0) {
            let keyForSplit = "";
            if (splitKey == "CurrentTermPolicyListed" || splitKey == "CurrentTermPolicyAttached") {
                const keysPossibleList = ["Current policy listed", "current policy listed", "Current policy endorsement listed", "current policy endorsement listed",
                    "Current policy attached", "current policy attached", "Current policy endorsement attached", "current policy endorsement attached"];;
                // keyForSplit = text.includes( "Current policy listed" ) ? "Current policy listed" : text.includes( "current policy listed" ) ? "current policy listed" : "";
                keysPossibleList.forEach((f) => {
                    if (text && text?.includes(f) && !keyForSplit) {
                        keyForSplit = f;
                    }
                });
            } else if (splitKey == "CurrentTermPolicyAttached") {
                keyForSplit = text.includes("Current policy attached") ? "Current policy attached" : text.includes("current policy attached") ? "current policy attached" : "";
            }
            if (!keyForSplit) {
                keyForSplit = text.includes("Page") ? "Page" : text.includes("page") ? "page" : "";
            }
            const pageData = keyForSplit ? text.split(keyForSplit) : null;
            if (pageData) {
                const pageDataArray = pageData[pageData?.length - 1];
                if (pageDataArray === ' # JSON file' || pageDataArray === ' # Json file' || pageDataArray === ' # json file' ||
                    pageDataArray === ' # JSONfile' || pageDataArray === ' # Jsonfile' || pageDataArray === ' # jsonfile') {
                    return "JSON file";
                } else {
                    const pageNumber = pageDataArray.match(/\d+/g);
                    if (Array.isArray(pageNumber) && pageNumber?.length == 1) {
                        return pageNumber[0];
                    }
                }
                // const pageNumber = pageData[ pageData?.length - 1 ].match( /\d+/g );
                // if ( Array.isArray( pageNumber ) && pageNumber?.length == 1 )
                // {
                //     return pageNumber[ 0 ];
                // }
            }
            return "NO RECORDS";
        } else {
            return "NO RECORDS";
        }
    }
}

/*NOTE:
1. in v object space is needed otherewise script error will happen.
2. function to get the default structure of cell data.
3. v is the object where the values will be appended/stored.
*/
export const getEmptyDataSet = () => {
    return {
        "ct": {
            "fa": "General",
            "t": "inlineStr",
            "s": [
                {
                    "ff": "\"times new roman\"",
                    "fc": "#000000",
                    "fs": 8,
                    "cl": 0,
                    "un": 0,
                    "bl": 0,
                    "it": 0,
                    "v": " "
                }
            ]
        },
        "merge": null,
        "w": 55,
        "tb": "2"
    }
}

export const splitPageKekFromText = (text, splitKey) => {
    let data = text;
    let key = "";
    const ctplKeysMSet = ["Current policy listed", "current policy listed", "Current policy endorsement listed", "current policy endorsement listed",
        "Current policy attached", "current policy attached", "Current policy endorsement attached", "current policy endorsement attached", "Endorsement Page", "Endorsement page",
        "endorsement page", "endorsement Page"];
    if (splitKey == "all") {
        ctplKeysMSet.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
    }
    if (splitKey != "all" && splitKey == "CurrentTermPolicyListed") {
        const ctplKeys = ["Current policy listed", "current policy listed", "Current policy endorsement listed", "current policy endorsement listed"];
        ctplKeys.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
        // if ( data.includes( "Current policy listed" ) )
        // {
        //     key = "Current policy listed";
        // } else if ( data.includes( "current policy listed" ) )
        // {
        //     key = "current policy listed";
        // } else if ( data.includes( "endorsement page" ) )
        // {
        //     key = "endorsement page";
        // } else
        // {
        //     key = "";
        // }
        // key = data.includes( "Current policy listed" ) ? "Current policy listed" : data.includes( "current policy listed" ) ? "current policy listed" : "";
    } else if (splitKey == "PriorTermPolicy" || splitKey == "CurrentTermPolicy") {
        const ptplKeys = getTableApplicationColumns("endorsement");
        ptplKeys.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
    } else if (splitKey == "CurrentTermPolicyAttached") {
        if (data?.includes("Current policy attached")) {
            key = "Current policy attached";
        } else if (data?.includes("current policy attached")) {
            key = "current policy attached";
        } else if (data?.includes("Current policy endorsement attached")) {
            key = "Current policy endorsement attached";
        } else if (data?.includes("current policy endorsement attached")) {
            key = "current policy endorsement attached";
        } else {
            key = "";
        }
        // key = data?.includes( "Current policy attached" ) ? "Current policy attached" : data?.includes( "current policy attached" ) ? "current policy attached" : "";
    }
    if (!key) {
        ctplKeysMSet.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
        if (!key) {
            key = data.includes("Page") ? "Page" : data.includes("page") ? "page" : "";
        }
    }
    if (key) {
        data = data.split(key)[0].replace(/\r\n/g, '');
    }
    return data;
}

export const splitPageKekFromTextForDataRendering = (text, splitKey) => {
    let data = text;
    let key = "";
    const ctplKeysMSet = ["Current policy listed", "current policy listed", "Current policy endorsement listed", "current policy endorsement listed",
        "Current policy attached", "current policy attached", "Current policy endorsement attached", "current policy endorsement attached", "Endorsement Page", "Endorsement page",
        "endorsement page", "endorsement Page"];
    if (splitKey == "all") {
        ctplKeysMSet.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
    }
    if (splitKey != "all" && splitKey == "CurrentTermPolicyListed") {
        const ctplKeys = ["Current policy listed", "current policy listed", "Current policy endorsement listed", "current policy endorsement listed"];
        ctplKeys.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
        // if ( data.includes( "Current policy listed" ) )
        // {
        //     key = "Current policy listed";
        // } else if ( data.includes( "current policy listed" ) )
        // {
        //     key = "current policy listed";
        // } else if ( data.includes( "endorsement page" ) )
        // {
        //     key = "endorsement page";
        // } else
        // {
        //     key = "";
        // }
        // key = data.includes( "Current policy listed" ) ? "Current policy listed" : data.includes( "current policy listed" ) ? "current policy listed" : "";
    } else if (splitKey == "PriorTermPolicy" || splitKey == "CurrentTermPolicy") {
        const ptplKeys = getTableApplicationColumns("endorsement");
        ptplKeys.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
    } else if (splitKey == "CurrentTermPolicyAttached") {
        if (data?.includes("Current policy attached")) {
            key = "Current policy attached";
        } else if (data?.includes("current policy attached")) {
            key = "current policy attached";
        } else if (data?.includes("Current policy endorsement attached")) {
            key = "Current policy endorsement attached";
        } else if (data?.includes("current policy endorsement attached")) {
            key = "current policy endorsement attached";
        } else {
            key = "";
        }
        // key = data?.includes( "Current policy attached" ) ? "Current policy attached" : data?.includes( "current policy attached" ) ? "current policy attached" : "";
    }
    if (!key) {
        ctplKeysMSet.forEach((f) => {
            if (data && data?.includes(f) && !key) {
                key = f;
            }
        });
        if (!key) {
            key = data.includes("Page") ? "Page" : data.includes("page") ? "page" : "";
        }
    }
    if (key) {
        data = data.split(key)[0].replace(/~~/g, '');
    }
    return data;
}

export const getPageKey = (key, tableName, text) => {
    if (key == "CurrentTermPolicyAttached") {
        if (text && text?.toLowerCase()?.includes('current policy endorsement attached')) {
            return "CpEa";
        }
        if (tableName?.toLowerCase()?.includes('formtable')) {
            return "Cp";
        }
        return "Cpa";
    } else if (key == "CurrentTermPolicyListed" && tableName != "Table 4") {
        if (text && text?.toLowerCase()?.includes('current policy endorsement listed')) {
            return "CpEl";
        }
        return "Cpl";
    } else if (key == "PriorTermPolicyAttached") {
        return "Pp";
    } else if (key == "CurrentTermPolicyAttached" && (tableName == "FormTable 2" || tableName == "FormTable 3")) {
        return "Cp";
    } else if (key == "CurrentTermPolicyListed" || key == "CurrentTermPolicy" || key == "CurrentTermPolicyListed1") {
        return "Cp";
    } else if (key == "PriorTermPolicyListed" && (tableName == "Table 4" || tableName == "Table 5")) {
        return "Pp";
    } else if (key == "PriorTermPolicyListed") {
        return "Ppl";
    } else if (key == "PriorTermPolicy") {
        return "Pp";
    } else if (key == "Quote" || key == "QuoteListed") {
        return "Quo";
    } else if (key == "Proposal" || key == "ProposalListed") {
        return "Prop";
    } else if (key == "Binder" || key == "BinderListed") {
        return "Bin";
    } else if (key == "Application" || key == "ApplicationListed") {
        return "App";
    } else if (key == "Schedule" || key == "ScheduleListed") {
        return "Sch";
    } else {
        //need to handle other appilcations
    }
    return '';
}

export const getObservationKey = (key, tableName) => {
    if (tableName?.toLowerCase()?.includes('formtable')) {
        if (key == 'PriorTermPolicyAttached') {
            return "Prior term: ";
        }
        if (key == 'CurrentTermPolicyAttached') {
            return "Current term: "
        }
    }
    if (key == "CurrentTermPolicy" || (key == "CurrentTermPolicyListed" && tableName == "Table 4") || key == "CurrentTermPolicyListed1") {
        return "Current term content(s): ";
    } else if (key == "CurrentTermPolicyListed") {
        return "Current term listed content: ";
    } else if (key == "CurrentTermPolicyAttached") {
        return "Current term attached content: ";
    } else if (key == "PriorTermPolicyAttached") {
        return "Prior term attached content: ";
    } else if (key == "PriorTermPolicy" || key == "PriorTermPolicyListed") {
        return "Prior term content(s): ";
    } else if (key == "Quote" || key == "QuoteListed") {
        return "Quote content(s): ";
    } else if (key == "Binder" || key == "BinderListed") {
        return "Binder content(s): ";
    }
    else if (key == "Proposal" || key == "ProposalListed") {
        return "Proposal content(s): ";
    }
    else if (key == "Application" || key == "ApplicationListed") {
        return "Application content(s): ";
    }
    else if (key == "Schedule" || key == "ScheduleListed") {
        return "Schedule content(s): ";
    }
    else {
        //need to handle other scenarios
        return "";
    }
}

export const onUpdateDataValidationProcess = (data, tableName) => {
    return data;
}

export const getOtherApplications = (columnData) => {
    const keys = Object.keys(columnData);
    const otherApplications = [];
    if (keys.length > 0) {
        const observationIndex = columnData?.Observation;
        keys.forEach((key) => {
            if (key != "CurrentTermPolicy" && key != "PriorTermPolicy" && columnData[key] != 0 && columnData[key] > 3 && columnData[key] < observationIndex) {
                otherApplications.push(key);
            }
        });
        return otherApplications;
    }
    return [];

}

export const getIndexForForms = (columnData) => {
    const keys = Object.keys(columnData);
    let data = {};
    if (keys.length > 0) {
        keys.forEach((f) => {
            if (columnData[f] == 3) {
                data['column1'] = f;
                data['columnIndex1'] = columnData[f];
            } else if (columnData[f] == 4) {
                data['column2'] = f;
                data['columnIndex2'] = columnData[f];
            }
        });
    }
    return data;
}

export const getTableApplicationColumns = (tableName) => {
    if (tableName === "Table 2") {
        return ["CurrentTermPolicy", "PriorTermPolicy", "Application", "Binder", "Proposal", "Quote", "Schedule"];
    }
    else if (tableName === "Table 3") {
        return ["CurrentTermPolicy", "PriorTermPolicy", "Application", "Binder", "Proposal", "Quote", "Schedule"];
    }
    else if (tableName === "Table 4") {
        return ["CurrentTermPolicyListed", "PriorTermPolicyListed", "ProposalListed", "BinderListed", "ScheduleListed", "QuoteListed", "ApplicationListed"];
    }
    else if (tableName === "Table 5") {
        return ["CurrentTermPolicyListed", "CurrentTermPolicyListed1", "PriorTermPolicyListed", "CurrentTermPolicyAttached"];
    }
    else if (tableName === "Table 6") {
        return ["CurrentTermPolicyAttached", "CurrentTermPolicyListed"];
    }
    else if (tableName === "Table 7") {
        return ["CurrentTermPolicyAttached", "CurrentTermPolicyListed"];
    } else if (tableName === "endorsement") {
        return ["Endorsement Page", "endorsement page", "Endorsement page", "endorsement Page", "Page"];
    } else if (tableName === "endorsementCPLA") {
        return ["Endorsement Page", "endorsement page", "Endorsement page", "endorsement Page",
            "Current Policy Endorsement Attached", "Current policy endorsement attached", "current policy endorsement attached",
            "Current Policy Endorsement Listed", "Current policy endorsement listed", "current policy endorsement listed",
            "Current policy listed", "current policy listed", "Current policy attached", "current policy attached",
            "Page"]
    } else {
        return [];
    }
}

/*
function to get the existing page key if available.
this will be used if the function content is in matched.
*/
export const getExistingPageKey = (pageData, key) => {
    if (!pageData) {
        return "NO RECORDS";
    }
    let text = getText(pageData, false);
    let trimmedText = text.replaceAll(' ', '');
    trimmedText = trimmedText?.replaceAll('\r\n', '');
    const splittedText = trimmedText.split(key);
    console.log("splittedText", splittedText);
    if (trimmedText.includes(key)) {
        let pageNumber = '';
        const loopinText = splittedText[1];
        console.log("loopinText", loopinText);
        const length = loopinText?.length;
        console.log("length", length);
        if (length > 0) {
            for (let i = 0; i < length; i++) {
                if (/^\d+$/.test(loopinText[i])) {
                    pageNumber += loopinText[i];
                } else {
                    break;
                }
            }
        }
        if (pageNumber && pageNumber?.length > 0) {
            return pageNumber;
        } else {
            return "NO RECORDS";
        }
    } else {
        return "NO RECORDS";
    }
}

export const isARType = (columnData, tableName) => {
    let allColumns = [];
    let fieldsOfAR = [];
    if (tableName === 'Table 2' || tableName === 'Table 3') {
        allColumns = ["Application", "Binder", "CurrentTermPolicy", "PriorTermPolicy", "Proposal", "Quote", "Schedule"];
        fieldsOfAR = ["CurrentTermPolicy", "PriorTermPolicy"];
    } else if (tableName === 'Table 4') {
        allColumns = getTableApplicationColumns(tableName);
        fieldsOfAR = ["CurrentTermPolicyListed", "PriorTermPolicyListed"];
    } else if (tableName === 'Table 5') {
        allColumns = getTableApplicationColumns(tableName);
        fieldsOfAR = ["CurrentTermPolicyListed", "PriorTermPolicyListed"];
    }
    const keys = Object.keys(columnData);
    const presentColumns = [];
    const presentARColumns = [];
    keys.forEach((key) => {
        if (columnData && columnData[key] > 0 && allColumns?.includes(key)) {
            presentColumns.push(key);
            if (fieldsOfAR?.includes(key)) {
                presentARColumns.push(key);
            }
        }
    });
    if (presentColumns?.length == presentARColumns?.length) {
        return { isAR: true, presentColumns }
    } else {
        return { isAR: false, presentColumns }
    }
}

export const getEndIdex = (tableRecords, currentIndex) => {
    console.log(tableRecords);
    const columnKeys = Object.keys(tableRecords);
    let hasFoundTableData = false;
    let endIndex = 0;
    columnKeys.forEach((f) => {
        const data = tableRecords[f];
        if (!hasFoundTableData && data?.range?.start <= currentIndex && data?.range?.end >= currentIndex) {
            hasFoundTableData = true;
            endIndex = data?.range?.end;
        }
    });
    return { hasFoundTableData, endIndex };

}

export const setDocumentDetails = async (jobId, token) => {
    try {
        sessionStorage.removeItem('jobDocumentData');
        const Token = await processAndUpdateToken(token);
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const response = await axios.get(baseUrl + '/api/Defaultdatum/GetDocumentDetails?jobId=' + jobId, {
            headers
        });
        if (response.status !== 200) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }
        // const response = [
        //     {
        //         "JobId": "1099csr122112024082206",
        //         "FileName": "2024-2025_RENL_BOP_10SBABD9403_HIG_20231212_0032_00861.pdf",
        //         "FileFor": "Current Term Policy",
        //         "FilePath": "D:\\PolicyCheck\\1099csr122112024082206\\Input\\"
        //     },
        //     {
        //         "JobId": "1099csr122112024082206",
        //         "FileName": "2024-2025_RENL_BOP_10SBABD9403_HIG_20231212_0032_00861.pdf",
        //         "FileFor": "Current Endorsement",
        //         "FilePath": "D:\\PolicyCheck\\1099csr122112024082206\\Input\\"
        //     },
        //     {
        //         "JobId": "1099csr122112024082206",
        //         "FileName": "2023-2024_RENL_BOP_10SBABD9403_HIG_20221028_0031_00830.pdf",
        //         "FileFor": "Prior Term Policy",
        //         "FilePath": "D:\\PolicyCheck\\1099csr122112024082206\\Input\\"
        //     }
        // ];
        sessionStorage.setItem('jobDocumentData', response?.data?.length > 0 ? JSON.stringify(response?.data) : response?.data);
        return response?.data;
    } catch (error) {
        //error happened block
    }
}

export const getDocumentDetails = async (jobId, token) => {
    const documentData = sessionStorage.getItem('jobDocumentData');
    if (documentData && documentData?.length > 0) {
        return documentData;
    } else {
        return await setDocumentDetails(jobId, token);
    }
}

export const validateEndorsementEntry = async (rowData, columnDetails, tableName, jobId, token) => {
    let hasInvalidEntry = false;
    if (rowData && columnDetails && tableName) {
        let documentData = await getDocumentDetails(jobId, token);
        documentData = documentData ? Array.isArray(documentData) ? documentData : JSON.parse(documentData) : [];
        const currentEndorsementDData = documentData.filter((f) => f?.FileFor?.includes('Endorsement') && f?.FileFor?.includes('Current'));
        const priorEndorsementDData = documentData.filter((f) => f?.FileFor?.includes('Endorsement') && f?.FileFor?.includes('Prior'));
        const tableColumnDetails = getTableApplicationColumns(tableName);
        let hasCTE = false;
        let hasPTE = false;
        tableColumnDetails.forEach((f) => {
            if (!hasInvalidEntry && columnDetails[f] && columnDetails[f] > 0) {
                const columnData = rowData[columnDetails[f]];
                if (columnData) {
                    const text = getTextWithoutAnyChnages(columnData);
                    const possibleEndorsementEntry = ["endorsement page", "current policy endorsement listed", "current policy endorsement attached"];
                    if (text) {
                        possibleEndorsementEntry?.forEach((item) => {
                            if (text?.toLowerCase()?.includes(item)) {
                                if (currentEndorsementDData?.length == 0 && priorEndorsementDData?.length == 0) {
                                    const hasseen = true;
                                }
                                if (f?.toLowerCase()?.includes('current') && currentEndorsementDData?.length == 0) {
                                    hasInvalidEntry = true;
                                }
                                else if (f?.toLowerCase()?.includes('prior') && priorEndorsementDData?.length == 0) {
                                    hasInvalidEntry = true;
                                }
                                else if (!f?.toLowerCase()?.includes('current') && !f?.toLowerCase()?.includes('prior')) {
                                    hasInvalidEntry = true;
                                } else {
                                    //no need to do anything
                                }
                            }
                        });
                    }
                }
            }
        });
    }
    return hasInvalidEntry;
}

export const getTextWithoutAnyChnages = (data) => {
    if (data?.ct?.s && data?.ct?.s?.length > 0) {
        let text = '';
        data?.ct?.s?.forEach((f) => {
            text += f?.v;
        });
        text = text?.trim();
        // return text?.replace( /\s+/g, ' ' );
        return text;
    } else if (!data?.ct?.s && (data?.m || data?.v)) {
        return data?.m ? data?.m?.replace(/\s+/g, ' ') : data?.v;
    }
    return '';
}

export const setLOBSplitMasterKeys = async (token) => {
    try {
        sessionStorage.removeItem('LobSplitMasterKeys');
        const Token = await processAndUpdateToken(token);
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const response = await axios.get(baseUrl + '/api/Defaultdatum/GetMasterLOBForSplit', {
            headers
        });
        if (response.status !== 200) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }
        sessionStorage.setItem('LobSplitMasterKeys', response?.data?.length > 0 ? JSON.stringify(response?.data) : response?.data);
        return response?.data;
    } catch (error) {
        //error happened block
    }
}

export const getLOBSplitMasterKeys = async (token) => {
    const LobSplitMasterKeys = sessionStorage.getItem('LobSplitMasterKeys');
    if (LobSplitMasterKeys && LobSplitMasterKeys?.length > 0) {
        return LobSplitMasterKeys;
    } else {
        return await setLOBSplitMasterKeys(token);
    }
}

export const mapLOBColumns = async (dataSet, token, jobId) => {
    let LobSplitMasterKeys = await getLOBSplitMasterKeys(token);
    LobSplitMasterKeys = LobSplitMasterKeys ? Array.isArray(LobSplitMasterKeys) ? LobSplitMasterKeys : JSON.parse(LobSplitMasterKeys) : [];
    let mappedData = dataSet;
    const JobCoveragePossibleColumns = getTableApplicationColumns("Table 3");
    if (mappedData && mappedData?.length > 0) {
        mappedData = mappedData.map((e) => {
            let findLOBName = LobSplitMasterKeys?.find((f) => f?.CHECKLIST_LOB_CODE?.toLowerCase() == e["Lob"]?.toLowerCase()?.trim());
            if (findLOBName) {
                e["PolicyLob"] = findLOBName?.Checklist_LOB_Master;
            }
            return e;
        });
    }
    //below code is mandatory --by gokul
    let checklistQuestionMaster = await GetCheckListQuestionMasterData(token, jobId);
    checklistQuestionMaster = checklistQuestionMaster ? Array.isArray(checklistQuestionMaster) ? checklistQuestionMaster : JSON.parse(checklistQuestionMaster) : [];
    const mappedDataForUniquPolicyLob = mappedData.filter((f) => f?.Lob != null && f?.Lob !== undefined && f?.Lob?.trim() !== "");
    const policyLob = mappedDataForUniquPolicyLob.map((e) => e?.PolicyLob);
    const uniquePolicyLob = Array.from(new Set(policyLob));
    if (uniquePolicyLob?.length > 0) {
        uniquePolicyLob.forEach((f, index) => {
            const policyQuestionMaster = checklistQuestionMaster.filter((item) => item?.LobName && item?.LobName == f);
            if (policyQuestionMaster?.length > 0) {
                let filteredJobCoverageData = mappedData?.filter((p) => p?.PolicyLob && p?.PolicyLob == f);
                if (filteredJobCoverageData && filteredJobCoverageData?.length > 0) {
                    filteredJobCoverageData = filteredJobCoverageData.filter((f) => f?.Lob);
                    if (filteredJobCoverageData && filteredJobCoverageData.length > 0) {
                        policyQuestionMaster.forEach((qm) => {
                            const findQuestionIsPresent = filteredJobCoverageData?.filter((md) => {
                                const questionCode = md?.ChecklistQuestions?.slice(0, 3);
                                if (md?.CoverageSpecificationsMaster?.toLowerCase()?.includes(qm?.Question?.toLowerCase()) ||
                                    questionCode?.toLowerCase() == qm?.ShortQuestion?.toLowerCase()) {
                                    return true;
                                }
                                return false;
                            });
                            if (findQuestionIsPresent?.length == 0) {
                                let data = {};
                                const keys = Object.keys(filteredJobCoverageData[0]);
                                keys.map((key) => {
                                    if (key == "CoverageSpecificationsMaster") {
                                        data["CoverageSpecificationsMaster"] = qm?.Question;
                                    } else if (key == "Lob") {
                                        data["Lob"] = qm?.LobCode;
                                    } else if (key == "PolicyLob") {
                                        data["PolicyLob"] = qm?.LobName;
                                    } else if (key == "ChecklistQuestions") {
                                        data["ChecklistQuestions"] = qm?.CheckListQuestion == null ? "-" : qm?.CheckListQuestion;
                                    } else if (key == "Columnid") {
                                        data["Columnid"] = mappedData?.length;
                                    } else if (JobCoveragePossibleColumns?.includes(key)) {
                                        data[key] = "Details not available in the document";
                                    }
                                    data["IsDataForSp"] = true;
                                });
                                mappedData.push(data);
                            }

                        });
                    }
                }
            }
        });
        const acitiveLobList = checklistQuestionMaster.filter((f) => f?.LobName && uniquePolicyLob?.includes(f.LobName));
        console.log(uniquePolicyLob);
        console.log(acitiveLobList);
    }
    return mappedData;
}

export const SetCheckListQuestionMasterData = async (token, jobId) => {
    try {
        sessionStorage.removeItem('LobSplitQuestionData');
        const Token = await processAndUpdateToken(token);
        const headers = {
            'Authorization': `Bearer ${Token}`,
            "Content-Type": "application/json",
        };
        const response = await axios.get(baseUrl + '/api/Defaultdatum/GetCheckListQuestionMasterData?jobId=' + jobId, {
            headers
        });
        if (response.status !== 200) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }
        sessionStorage.setItem('LobSplitQuestionData', response?.data?.length > 0 ? JSON.stringify(response?.data) : response?.data);
        return response?.data;
    } catch (error) {
        //error happened block
    }
}

export const GetCheckListQuestionMasterData = async (token, jobId) => {
    const LobSplitMasterKeys = sessionStorage.getItem('LobSplitQuestionData');
    if (LobSplitMasterKeys && LobSplitMasterKeys?.length > 0) {
        return LobSplitMasterKeys;
    } else {
        return await SetCheckListQuestionMasterData(token, jobId);
    }
}

//validating if the job has endorsement fn
export const endorsementCheck = (data) => {
    if (data && data?.JobFileInfo?.length > 0) {
        const keys = ["JobCoverage", "JobCommonDeclaration", "TblChecklistForm1", "TblChecklistForm1", "TblChecklistForm1", "TblChecklistForm1"];
        let hasEndorsement = false;
        keys.forEach((key, index) => {
            const templateData = data[key];
            if (!hasEndorsement && templateData && templateData?.length > 0) {
                const tableName = "Table " + (index + 2);
                const applicableColumns = getTableApplicationColumns(tableName);
                applicableColumns.forEach((column) => {
                    const columnData = templateData.map((e) => e[column]).filter((f) => f != null && f != undefined);
                    // const filterContent = ["matched","details not available"];
                    if (columnData?.length > 0) {
                        const filteredInvalidData = columnData.filter((f) => !(f.toLowerCase() == "matched" || f.toLowerCase().includes("details not available")));
                        const filterByEndorsement = filteredInvalidData?.length > 0 ? filteredInvalidData.filter((f) => f?.toLowerCase()?.includes("endorsement page") || f?.toLowerCase()?.includes("current policy endorsement attched") ||
                            f?.toLowerCase()?.includes("current policy endorsement listed")) : [];
                        if (filterByEndorsement?.length > 0) {
                            hasEndorsement = true;
                        }
                    }
                });
            }
        });
    }
}

export const updateFormsPHData = async (jobId, data, token) => {
    const phTableRecords = [];
    if (jobId && data && data.length > 0) {
        let documentDetails = await getDocumentDetails(jobId, token);
        documentDetails = Array.isArray(documentDetails) ? documentDetails : JSON.parse(documentDetails);
        const CpfileName = documentDetails.find((f) => f.FileFor == "Current Term Policy")?.FileName;
        const PpfileName = documentDetails.find((f) => f.FileFor == "Prior Term Policy")?.FileName;
        const Filepath = documentDetails[0]?.FilePath;
        const unMatchedData = data.filter((f) => f?.IsMatched == false || f?.IsMatched == "false");
        if (unMatchedData?.length > 0) {
            unMatchedData.forEach((item) => {
                let Cpvaluetext = getCpPpText("CurrentTermPolicyAttached", item);
                let Ppvaluetext = getCpPpText("PriorTermPolicyAttached", item);
                let CppageNumber = getCpPpPageNo("CurrentTermPolicyAttached", item);
                let PppageNumber = getCpPpPageNo("PriorTermPolicyAttached", item);
                phTableRecords.push({
                    "Jobid": jobId,
                    "Filepath": Filepath,
                    "CpfileName": CpfileName,
                    "PpfileName": PpfileName,
                    "Cpvaluetext": Cpvaluetext,
                    "Ppvaluetext": Ppvaluetext,
                    "CppageNumber": CppageNumber,
                    "PppageNumber": PppageNumber
                });
            });
        }
    }
    return phTableRecords;
}

const getCpPpText = (key, item) => {
    let text = "NO RECORDS";
    if (key && item) {
        const keyText = item[key];
        const trimmedText = keyText?.trim();
        const pageKey = trimmedText?.includes("~~page") ? "~~page" : trimmedText?.includes("~~Page") ? "~~Page" : "~~page";
        if (!(trimmedText?.includes("details not available in the document") || trimmedText?.includes("Details not available"))) {
            const splittedText = trimmedText?.split(pageKey)[0];
            if (splittedText) {
                text = splittedText;
            }
        }
    }
    return text;
}

const getCpPpPageNo = (key, item) => {
    let pageNumber = "NO RECORDS";
    if (key && item) {
        const keyText = item[key];
        const trimmedText = keyText?.trim()?.toLowerCase();
        if (!(trimmedText?.includes("details not available in the document") || trimmedText?.includes("Details not available")) && trimmedText?.includes("page #")) {
            const splittedText = trimmedText?.split("~~page")[1];
            const pageNumberText = splittedText ? splittedText?.match(/\d+/g) : "";
            if (splittedText && pageNumberText?.length == 1) {
                pageNumber = pageNumberText[0];
            }
        }
    }
    return pageNumber;
}

export const getQACData = async (jobId, token) => {
    const defaultData = { canRender: false, data: {} };
    if (jobId) {
        const renderQACFromMaster = sessionStorage.getItem('renderQACFromMaster');
        let qacCheckListSheetData = sessionStorage.getItem('qacCheckListSheetData');
        qacCheckListSheetData = JSON.parse(qacCheckListSheetData);
        if (renderQACFromMaster == 'true' && qacCheckListSheetData) {
            let qacApplicableBrokerIds = sessionStorage.getItem('qacApplicableBrokerIds');
            qacApplicableBrokerIds = JSON.parse(qacApplicableBrokerIds);
            const brokerId = jobId.slice(0, 4);
            if (qacApplicableBrokerIds && qacApplicableBrokerIds?.length > 0 && qacApplicableBrokerIds.includes(brokerId)) {
                return { canRender: true, data: qacCheckListSheetData };
            } else {
                return defaultData
            }
        } else {
            try {
                const Token = await processAndUpdateToken(token);
                const headers = {
                    'Authorization': `Bearer ${Token}`,
                    "Content-Type": "application/json",
                };
                const response = await axios.get(baseUrl + '/api/Defaultdatum/GetCheckListQACMasterData?JobID=' + jobId, {
                    headers
                });
                if (response.status !== 200) {
                    return defaultData;
                }
                if (response?.data && response?.data?.isQACApplicable) {
                    if (response?.data?.QACList && response?.data?.QACList?.length > 0) {
                        const structuredData = structureQACData(response?.data?.QACList);
                        return { canRender: true, data: structuredData };
                    } else {
                        return defaultData;
                    }
                } else {
                    return defaultData;
                }
            } catch (error) {
                return defaultData;
            }
        }
    }
    return defaultData;
}

const structureQACData = (data) => {
    const cellData = [];
    const defaultSheetData = qacMasterSet;
    const lobList = Array.from(new Set(data.map((e) => e.PolicyLob)));
    // console.log(lobList);
    let rowIndex = 0;

    lobList.forEach((lob, lobIndex) => {
        const lobData = data.filter((f) => f.PolicyLob == lob);
        if (lobData?.length > 0) {
            const initialRIndex = rowIndex;
            const lobDataLength = lobData?.length;
            //lob header for lob("TRUCKING AUTO LIABILITY")
            if (lob == "TRUCKING AUTO LIABILITY") {
                cellData.push({
                    "r": rowIndex,
                    "c": 0,
                    "v": {
                        "ct": {
                            "fa": "@",
                            "t": "inlineStr"
                        },
                        "m": "SUB-HAULER SECTION",
                        "v": "SUB-HAULER SECTION",
                        "merge": null,
                        "tb": "2",
                        "fs": 8,
                        "bl": 1,
                        "bg": "#000000",
                        "fc": "#ffffff",
                        "ht": "0"
                    }
                });
                rowIndex += 1;
            }
            //lob table headerSection
            cellData.push({
                "r": rowIndex,
                "c": 0,
                "v": {
                    "ct": {
                        "fa": "@",
                        "t": "inlineStr"
                    },
                    "m": lob,
                    "v": lob,
                    "merge": null,
                    "tb": "2",
                    "fs": 8,
                    "bl": 1,
                    "bg": "rgb(139,173,212)",
                }
            });
            rowIndex += 1;
            cellData.push({
                "r": rowIndex,
                "c": 0,
                "v": {
                    "ct": {
                        "fa": "@",
                        "t": "inlineStr"
                    },
                    "m": lobData[0]?.Heading,
                    "v": lobData[0]?.Heading,
                    "merge": null,
                    "tb": "2",
                    "fs": 8,
                    "bl": 1,
                    "bg": "rgb(139,173,212)",
                }
            });
            rowIndex += 1;
            //lob question data section
            lobData.map((item, index) => {
                let emptySet = {
                    "r": 0,
                    "c": 0,
                    "v": {
                        "ct": {
                            "fa": "@",
                            "t": "inlineStr"
                        },
                        "m": "",
                        "v": "",
                        "merge": null,
                        "tb": "2",
                        "fs": 7
                    }
                };
                emptySet["r"] = rowIndex;
                emptySet["v"]["m"] = item.Question;
                emptySet["v"]["v"] = item.Question;
                cellData.push(emptySet);
                rowIndex += 1;
                if (lobDataLength == (index + 1)) {
                    defaultSheetData["config"]["borderInfo"][lobIndex] = {
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
                                "row": [
                                    initialRIndex,
                                    rowIndex - 1
                                ],
                                "column": [
                                    0,
                                    0
                                ],
                                "row_focus": 0,
                                "column_focus": 0
                            }
                        ]
                    }
                    rowIndex += 2;
                }
            });
        }
    });
    defaultSheetData["celldata"] = cellData
    return defaultSheetData;
}

export const setMasterData = async () => {
    try {
        const hasMasterData = sessionStorage.getItem("HasMasterData");
        if (!(hasMasterData && (hasMasterData === "true" || hasMasterData === true))) {
            const masterDataResponse = await axios.get(baseUrl + '/api/Authentication/GetMasterData', {
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                }
            });
            if (masterDataResponse.status == 200) {
                const masterData = masterDataResponse.data;
                masterData.forEach(item => {
                    sessionStorage.setItem(item.FieldName, item.FieldValue);
                });
                sessionStorage.setItem("HasMasterData", true);
            } else {
                //handle error's
                // throw new Error( `HTTP error! Status: ${ response.status }` );
            }
        }
    } catch (error) {
        //handle error
    }
}

const getTextPageNumberForForms = (text) => {
    if (text && text != null && text != undefined && text?.length > 0 && text?.toLowerCase()?.includes('page')) {
        const splittedText = text?.toLowerCase().split('page')[1];
        if (splittedText && splittedText != null && splittedText != undefined) {
            const currentTextPageNumber = splittedText?.match(/\d+/g)[0];
            if (currentTextPageNumber && currentTextPageNumber != null && currentTextPageNumber != undefined) {
                return currentTextPageNumber;
            }
        }
        return '';
        // dataFormated.splice(index,1)
    }
    else {
        return '';
    }
}

export const formsCompareDuplicateDataRemover = (data) => {

    let formsDataFiltered = [];
    const originalData = data;
    try {
        if (originalData && originalData?.length > 0) {
            const idsToBePopped = [];
            const columns = ['CurrentTermPolicyAttached', 'PriorTermPolicyAttached'];
            columns.map((column) => {

                originalData.forEach((f, index) => {
                    const text = f[column];

                    const currentPageNumber = getTextPageNumberForForms(text);
                    if (currentPageNumber && currentPageNumber?.length > 0) {
                        // console.log(currentPageNumber);
                        originalData.forEach((f1, index1) => {
                            if (f.Id != f1.Id && !idsToBePopped.includes(f1.Id) && index1 > index) {

                                const text1 = f1[column];
                                const recordPageNumber = getTextPageNumberForForms(text1);
                                // console.log(recordPageNumber);
                                if (recordPageNumber && recordPageNumber?.length > 0 && recordPageNumber == currentPageNumber) {
                                    idsToBePopped.push(f1.Id);
                                }
                            }
                        });
                    }
                    // dataFormated.splice(index,1)

                });
            });
            if (idsToBePopped?.length > 0) {
                formsDataFiltered = originalData.filter((item) => !idsToBePopped.includes(item.Id));
                return formsDataFiltered;
            } else {
                return data;
            }
        }
    } catch (error) {
        return data;
    }
    return data;
}

export const findTableForIndex = (selectedIndex, tableDetails, excludedColumns) => {
    for (const tableName in tableDetails) {
        if (tableDetails.hasOwnProperty(tableName)) {
            const range = tableDetails[tableName].range;
            const columnNames = tableDetails[tableName].columnNames;
            if (typeof range.start === 'number' && typeof range.end === 'number') {
                if (selectedIndex >= range.start && selectedIndex <= range.end) {
                    if (columnNames && typeof columnNames === 'object') {
                        const validColumns = Object.keys(columnNames).filter(colName => !excludedColumns.includes(colName));
                        if (validColumns.length > 0) {
                            return tableName;
                        }
                    }
                }
            }
        }
    }
    return null;
}

export const findTblRowAllIndex = (index, tabledata) => {
    if (index != undefined && tabledata != undefined) {
        for (const tableName in tabledata) {
            const range = tabledata[tableName].range;
            if (index >= range.start && index <= range.end) {
                return tableName;
            }
        }
        return null;
    }
};

export const filterSelectedRowIndexForCopyPaste = (tblSelectedRow, selectedIndex) => {
    if (tblSelectedRow != undefined && selectedIndex != undefined) {
        const rangeStart = tblSelectedRow.range.start;
        const rangeEnd = tblSelectedRow.range.end;

        if (rangeStart <= selectedIndex && selectedIndex <= rangeEnd) {
            return selectedIndex;
        } else {
            return null;
        }
    }
}

export const getPreviewChecklistDataForUpdate = (jobId, sheetDataSet, columnPositions, columnsSet, tableName, policyLobToMap) => {
    try {
        const dataToUpdate = [];
        if (sheetDataSet && columnPositions && columnsSet) {
            if (tableName != "Table 1") {
                sheetDataSet.forEach((f, index) => {
                    let objectSet = {};
                    columnsSet.forEach((col) => {
                        const sheetObj = f[columnPositions[col]];
                        let text = getText(sheetObj, true);
                        text = text.replace(/\r\n/g, '~~');
                        text = text.replace(/\•/g, '.');
                        if (col === "Actions on Discrepancy") {
                            objectSet["ActionOnDiscrepancy"] = text;
                        }
                        else if (col === "Request Endorsement") {
                            objectSet["RequestEndorsement"] = text;
                        }
                        else if (col === "Notes(Free Fill)") {
                            objectSet["NotesFreeFill"] = text;
                        }
                        else {
                            objectSet[col] = text;
                        }
                    });
                    objectSet["PolicyLob"] = policyLobToMap;
                    objectSet["jobId"] = jobId;
                    if (tableName === "Table 2" || tableName === "Table 3") {
                        objectSet["Columnid"] = index;
                    }
                    dataToUpdate.push(objectSet);
                });
            } else {
                sheetDataSet.forEach((f, index) => {
                    let objectSet = {};
                    const sheetObjKey = f[1];
                    const textKey = getText(sheetObjKey, true);
                    const replaceKey = textKey.replace(/\r\n/g, '~~');
                    const sheetObjValue = f[2];
                    const textValue = getText(sheetObjValue, true);
                    const replaceValue = textValue.replace(/\r\n/g, '~~');
                    objectSet["Headers"] = replaceKey;
                    objectSet["NoColumnName"] = replaceValue;
                    objectSet["jobid"] = jobId;
                    objectSet["PolicyLob"] = policyLobToMap;
                    objectSet["HeaderId"] = index;
                    dataToUpdate.push(objectSet);
                });
            }
            return dataToUpdate;
        } else {
            return dataToUpdate;
        }
    } catch (error) {
        return error;
    }
}


export const ExportData = async (jobId, sheetType, token) => {

    try {
        const updatedToken = await processAndUpdateToken(token);
        token = updatedToken;

        const headers = {
            'Authorization': `Bearer ${updatedToken}`,
        };

        if (jobId && sheetType) {
            const response = await axios.get(`${baseUrl}/api/ProcedureData/Exportdata`, {
                params: { jobId, SheetType: sheetType },
                headers,
            });

            if (response.status === 200) {
                return response.data;
            } else {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }
        } else {
            throw new Error("Missing required parameters: jobId or sheetType");
        }
    } catch (error) {
        console.error('Error:', error);
        throw error;
    }
};


export const updateAllSheetsResultArray = (result, xRayData, flagCheck) => {
    if (flagCheck == "PolicyReviewChecklist" || flagCheck == "Forms Compare") {
        result.forEach(obj => {
            let xRayObj = obj.find(f => f?.value === "X-Ray");

            if (xRayObj) {
                xRayData.forEach(arr => {
                    arr?.Data.forEach(arrObj => {
                        let findIndex = arrObj?.dataToUpDate.find(e => e?.sheetPosition === xRayObj?.rowIndex);
                        if (findIndex) {
                            xRayObj.value = findIndex["Document Viewer"];
                        }
                    })
                })
            }
        })
        return result;
    } else if (flagCheck == "Red" || flagCheck == "Green") {
        result.forEach(obj => {
            let xRayObj = obj.find(f => f?.value === "X-Ray");

            if (xRayObj) {
                xRayData.forEach(arrObj => {
                    let findIndex = arrObj?.data.find(e => e?.sheetPosition === xRayObj?.rowIndex);
                    if (findIndex) {
                        xRayObj.value = findIndex["Document Viewer"];
                    }
                })
            }
        })
        return result;
    }
}


export const structureColumns = (key, data) => {
    let column = [];
    if (key) {
        const tableData = data.find((f) => f?.tableName == key)?.result || data.find((f) => f?.Tablename == key)?.TemplateData;
        if (tableData && tableData?.length > 0) {
            const columnSet = Object.keys(tableData[0]).filter(
                (f) =>
                    !["createdon", "updatedon", "columnid", "setid", "jobid", "isdataforsp", "id"]?.includes(
                        f?.trim()?.toLowerCase()
                    )
            );
            columnSet.forEach((col) => {
                column.push(getColumnObject(col));
            });
        }
    }
    return column;
};

export const getColumnObject = (col) => {
    const colSpec = {
        name: col,
        // selector: (row, index) => (
        //   <Tooltip title={col == "Id" ? index : row[col]}>
        //     {col == "Id" ? index : row[col]}
        //   </Tooltip>
        // ),
        selector: (row) => {
            const cellValue = row[col];
            const htmlContent = typeof cellValue === "string"
                ? cellValue?.replace(/~~/g, "<br />")
                : cellValue;
            return <span dangerouslySetInnerHTML={{ __html: htmlContent }} />;
        },
        sortable: true,
        wrap: true,
        style: { fontSize: "11px !important" },
    };
    if (col?.trim()?.toLowerCase() == "id") {
        colSpec["width"] = "100px";
    } else if (col?.trim()?.toLowerCase() == "jobid") {
        colSpec["width"] = "200px";
    } else if (
        ["policy lob", "page number", "checklist questions", "lob"].includes(
            col?.trim()?.toLowerCase()
        )
    ) {
        colSpec["width"] = "200px";
    } else if (
        [
            "coverage_specifications_master",
            "coverage_specification_master",
            "observation",
        ].includes(col?.trim()?.toLowerCase())
    ) {
        colSpec["width"] = "300px";
    } else {
        colSpec["width"] = "400px";
    }
    return colSpec;
};

export const CsrSaveHistoryApiCall = async (sheetType, jobId, userName, CsrSaveJobIdHistoryData, BrokerId, isAutoUpdate) => {
    let token = sessionStorage.getItem("token");
    document.body.classList.add('loading-indicator');
    const Token = await processAndUpdateToken(token);

    let message;
    let log;
    if(sheetType === "PreviewCheckList") {
        message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-PolicyCheck-initiated" : "CSRSaveHistory-Jobid-Update-PolicyCheck-initiated";
        log = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoSave-PolicyCheck" : "CSRSaveHistory-Jobid-Save-PolicyCheck";
    } else if(sheetType === "GradiationSheet") {
        message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-Gradiation-initiated" : "CSRSaveHistory-Jobid-Update-Gradiation-initiated";
        log = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoSave-Gradiation" : "CSRSaveHistory-Jobid-Save-Gradiation";
    }

    updateGridAuditLog(jobId, message, log, "");
    const headers = {
        'Authorization': `Bearer ${Token}`,
        "Content-Type": "application/json",
    };
    const apiUrl = `${baseUrl}/api/ProcedureData/CsrSaveHistory`;
    try {
        const response = await axios({
            method: "POST",
            url: apiUrl,
            headers: headers,
            data: {
                JobID: jobId,
                UserName: userName,
                ChecklistData: CsrSaveJobIdHistoryData,
                BrokerId: BrokerId
            }
        });
        if (response.status !== 200) {
            if(sheetType === "PreviewCheckList") {
                message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-PolicyCheck-Update-failed" : "CSRSaveHistory-Jobid-Update-PolicyCheck-Update-failed";
            } else if(sheetType === "GradiationSheet") {
                message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-Gradiation-Update-failed" : "CSRSaveHistory-Jobid-Update-Gradiation-Update-failed";
            }
            updateGridAuditLog(jobId, message, JSON.parse(response));
            return "error";
        }

        return response.data;
    } catch (error) {
        if(sheetType === "PreviewCheckList") {
            message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-PolicyCheck-Update-failed" : "CSRSaveHistory-Jobid-Update-PolicyCheck-Update-failed";
        } else if(sheetType === "GradiationSheet") {
            message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-Gradiation-Update-failed" : "CSRSaveHistory-Jobid-Update-Gradiation-Update-failed";
        }
        updateGridAuditLog(jobId, message, error);
        return "error";
    } finally {
        if(sheetType === "PreviewCheckList") {
            message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-PolicyCheck-completed" : "CSRSaveHistory-Jobid-Update-PolicyCheck-completed";
            log = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoSave-PolicyCheck-success" : "CSRSaveHistory-Jobid-Save-PolicyCheck-success";
        } else if(sheetType === "GradiationSheet") {
            message = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoUpdate-Gradiation-completed" : "CSRSaveHistory-Jobid-Update-Gradiation-completed";
            log = isAutoUpdate ? "CSRSaveHistory-Jobid-AutoSave-Gradiation-success" : "CSRSaveHistory-Jobid-Save-Gradiation-success";
        }
        updateGridAuditLog(jobId, message, log, "");
        document.body.classList.remove('loading-indicator');
        return "success";
    }
};

export const brokerIdsGetData = async () => {
    var token = sessionStorage.getItem("token");

    if (token && token.length >= 40) {
        document.body.classList.add("loading-indicator");
        const headers = {
            Authorization: `Bearer ${token}`, // Fix the string interpolation here
        };

        try {
            const response = await axios.get(
                `${baseUrl}/api/Defaultdatum/GetBrokerNames/GetBrokerNames`,
                { headers }
            );

            if (response.status !== 200) {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }

            return response.data;
        } catch (error) {
            console.error("Error fetching broker data:", error);
        } finally {
            document.body.classList.remove("loading-indicator");
        }
    }
};

export const CsrPendingReport = async (jobId, userName, CsrSaveJobIdHistoryData, Process) => {
    let token = sessionStorage.getItem("token");
    document.body.classList.add('loading-indicator');
    const Token = await processAndUpdateToken(token);
    updateGridAuditLog(jobId, "CsrPending-JobId-Update-initiated", "");
    const headers = {
        'Authorization': `Bearer ${Token}`,
        "Content-Type": "application/json",
    };
    const apiUrl = `${baseUrl}/api/ProcedureData/CsrPendingReport`;
    try {
        const response = await axios({
            method: "POST",
            url: apiUrl,
            headers: headers,
            data: {
                JobID: jobId,
                UserName: userName,
                ChecklistData: CsrSaveJobIdHistoryData,
                Process: Process
            }
        });
        if (response.status !== 200) {
            updateGridAuditLog(jobId, "CsrPending-JobId-Update-failed", JSON.parse(response));
            return "error";
        }

        return response.data;
    } catch (error) {
        // console.error( 'Error:', error );
        updateGridAuditLog(jobId, "CsrPending-JobId-Update-failed", error);
        return "error";
    } finally {
        updateGridAuditLog(jobId, "CsrPending-JobId-Update-completed", "");
        document.body.classList.remove('loading-indicator');
        return "success";
    }
};

export const formatTo12HourIST = (date) => {
    // Adjust the date to IST (UTC+5:30)
    date = new Date(date);
    // const istOffset = 5.5 * 60 * 60 * 1000; // 5 hours 30 minutes in milliseconds
    const istDate = new Date(date.getTime());

    // Get the date components
    const year = istDate.getFullYear();
    const month = (istDate.getMonth() + 1).toString().padStart(2, "0"); // Months are 0-based, so add 1
    const day = istDate.getDate().toString().padStart(2, "0");

    // Get the hours, minutes, and seconds
    let hours = istDate.getHours();
    const minutes = istDate.getMinutes();
    const seconds = istDate.getSeconds();

    // Determine AM/PM suffix
    const period = hours >= 12 ? "PM" : "AM";

    // Convert hours from 24-hour to 12-hour format
    hours = hours % 12;
    hours = hours ? hours : 12; // The hour '0' should be '12'

    // Format the date & time
    const formattedDate = `${year}-${month}-${day}`;
    const formattedTime = `${hours}:${minutes
        .toString()
        .padStart(2, "0")}:${seconds.toString().padStart(2, "0")} ${period}`;

    return `${formattedDate} ${formattedTime}`;
};

export const groupNumbers = (data) => {
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


export const qacTblRangeStructureFun = (tableData) => {
    const structuredTblObj = tableData.reduce((acc, curr) => {
        const tableKey = Object.keys(curr)[0];
        acc[tableKey] = curr[tableKey];
        return acc;
    }, {});
    return structuredTblObj;
};

export const getConfidenceScoreConfigStatus = (data, key, Question) => {
    if(key === "question check"){
        const questionCode = Question?.trim()?.slice(0,3) ;
        if(questionCode && data?.length > 0){
            const hasQuestionInStp = data?.filter((f) => f.ShortQuestion?.toLowerCase() === questionCode.toLowerCase())?.length > 0;
            if(hasQuestionInStp){
                return true;
            }
        }
        return false;
    }else{
        if(data && data?.length > 0){            
            return data?.find(f => f?.Key === key)?.Value;
        }
        return "";
    }
}

export const getCsRespectiveColumn = ( csColName ) => {
    if(csColName){
        const data = {
            "CurrentTermPolicyCs":"CurrentTermPolicy",
            "PriorTermPolicyCs":"PriorTermPolicy",
            "ProposalCs":"Proposal",
            "ApplicationCs":"Application",
            "QuoteCs":"Quote",
            "BinderCs":"Binder",
            "ScheduleCs":"Schedule",
            "CurrentTermPolicyListedCs":"CurrentTermPolicyListed",
            "PriorTermPolicyListedCs":"PriorTermPolicyListed",
            "ProposalListedCs":"ProposalListed",
            "ApplicationListedCs":"ApplicationListed",
            "QuoteListedCs":"QuoteListed",
            "BinderListedCs":"BinderListed",
            "ScheduleListedCs":"ScheduleListed",
            "CurrentTermPolicyAttachedCs":"CurrentTermPolicyAttached",
            "PriorTermPolicyAttachedCs":"PriorTermPolicyAttached",
            "CurrentTermPolicyListedCs1":"CurrentTermPolicyListed1",

            "CurrentTermPolicy":"CurrentTermPolicyCs",
            "PriorTermPolicy":"PriorTermPolicyCs",
            "Proposal":"ProposalCs",
            "Application":"ApplicationCs",
            "Quote":"QuoteCs",
            "Binder":"BinderCs",
            "Schedule":"ScheduleCs",
            "CurrentTermPolicyListed":"CurrentTermPolicyListedCs",
            "PriorTermPolicyListed":"PriorTermPolicyListedCs",
            "ProposalListed":"ProposalListedCs",
            "ApplicationListed":"ApplicationListedCs",
            "QuoteListed":"QuoteListedCs",
            "BinderListed":"BinderListedCs",
            "ScheduleListed":"ScheduleListedCs",
            "CurrentTermPolicyAttached":"CurrentTermPolicyAttachedCs",
            "PriorTermPolicyAttached":"PriorTermPolicyAttachedCs",
            "CurrentTermPolicyListed1":"CurrentTermPolicyListedCs1",
        };
        return data[csColName];
    }else{
        return "";
    }
}

export const getKeyByValue = (obj, value) => {
    for (const [key, val] of Object.entries(obj)) {
      if (val === value) {
        return key;
      }
    }
    return null; // Return null if no match is found
  };