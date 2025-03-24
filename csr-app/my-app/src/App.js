import * as React from "react";
import "./App.css";
import Luckysheet from "./SpreadSheet/Luckysheet";
import { useEffect, useState } from "react";
import { useLocation } from "react-router-dom";
import {
  baseUrl,
  staticExclusionData,
  staticExclusionCsrData,confidenceScoreConfigStaticData
} from "./Services/Constants";
import Util from "./Services/utils";
import axios from "axios";
import jwt_decode from "jwt-decode";
import {
  processAndUpdateToken,
  setMasterData,
  getObservationKey,
  getPageKey,
  isARType,
  splitPageKekFromTextForDataRendering,
  getTextByRequirement,
  SetCheckListQuestionMasterData
} from "./Services/CommonFunctions";
import CsrSheet from "./SpreadSheet/CsrSheet";
import DefaultDataSheet from "./SpreadSheet/DefaultDataSheet";
import {
  JobHeaderData,
  JobCDData,
  JobJCData,
  JobF1Data,
  JobF2Data,
  JobF3Data,
  JobF4Data,
  JobConfigData,
  JobPRTotalCount,
  ComponentRenderApi
} from "./Services/PreviewChecklistDataService";
import { auditProcessNames } from "./Services/enums";
import { updateGridAuditLog, GetConfidenceScoreConfig } from "./Services/PreviewChecklistDataService";
import { SimpleSnackbarWithOutTimeOut } from "./Components/SnackBar";

function App(props) {
  const reference = React.useRef();
  const [checklistData, setChecklistData] = useState();
  const [csrChecklistData, setCsrChecklistData] = useState();
  const [csrFormsData, setCsrFormsData] = useState();
  const [csrQacData, setCsrQacData] = useState();
  const [formData, setFormData] = useState();
  const [exclusionData, setExclusionData] = useState();
  const [defaultSheetNameData, setDefaultSheetNameData] = useState();
  const [isValue, setIsValue] = useState(false);
  const [isReadonly, setIsReadonly] = useState(false);
  const [msgVisible, setMsgVisible] = useState(false);
  const [sheetRenderConfig, setSheetRenderConfig] = useState({
    PolicyReviewChecklist: "false",
    FormsCompare: "false",
    Exclusion: "false",
    QAC_not_answered_questions: "false",
  });
  const [msgClass, setMsgClass] = useState("");
  const [msgText, setMsgText] = useState("");
  const [Qacdata, setQacdata] = useState("");
  const [jobId, setJobId] = useState();
  const [gradtionData, setGradtionData] = useState();
  const [sheetListOptionSet, setSheetListOptionSet] = useState();
  const [selectedSheet, setSelectedSheet] = useState("PolicyReviewChecklist");
  const [confidenceScoreConfig, setConfidenceScoreConfig] = useState([]);
  const [enableExclusionCellLock, setEnableExclusionCellLock] = useState(true);
  const location = useLocation();
  const renderBrokerIdsQac = ["1003", "1162"];

  let token = sessionStorage.getItem("token");

  useEffect(() => {
    sessionStorage.removeItem("jobDocumentData");

    let urlString;

    // Detect if running in Electron/Desktop mode
    if (typeof process !== "undefined" && process.argv) {
        const arg = process.argv.find((arg) => arg.startsWith("csragent://"));
        if (arg) {
            urlString = arg.replace("csragent://", "https://"); // Convert to a valid URL
        }
    } 
    
    // Fallback to browser URL
    if (!urlString) {
        urlString = window.location.href;
    }

    // Fix possible double slash issue
    urlString = urlString.replace(/([^:]\/)\/+/g, "$1");

    const url = new URL(urlString);
    const jobId = url.searchParams.get("jobId") || url.searchParams.get("jobid");
    const urlParamToken = url.searchParams.get("token");

    if (!jobId) return;

    setJobId(jobId);
    const brokerId = jobId.slice(0, 4);
    const currentPath = location.pathname;

    if (urlParamToken?.length > 15 && (currentPath === "/csrView" || currentPath === "/csrview")) {
        updateCsrEntry(jobId, 'CsrJobValidation-Initial-Entry', false);
        csrAuthenticationTokenApi();
    }

    SetCheckListQuestionMasterData(token, jobId);

    const apiCalls = [
        { condition: currentPath === "/xlpage", fn: () => dataRender(jobId, "PolicyReviewChecklist") },
        { condition: currentPath === "/xlpage", fn: () => formRender(jobId, "") },
        { condition: canRenderExclusion(brokerId), fn: () => exclusionRender(jobId, currentPath.includes("csrView") ? "csrExclusion" : "") },
        { condition: currentPath.includes("csrView"), fn: () => GetGradationData(jobId, currentPath) },
        { condition: currentPath.includes("csrView"), fn: () => csrPolicyData(jobId, "9", "0") },
        { condition: currentPath.includes("csrView"), fn: () => csrFormData(jobId) },
        { condition: currentPath.includes("csrView"), fn: () => ExportQac() },
        { condition: currentPath === "/xlpage", fn: () => ExportQac() },
        { condition: currentPath === "/xlpage", fn: () => getSheetListsInMasterData(jobId) },
        { condition: currentPath.includes("csrView") && renderBrokerIdsQac?.includes(brokerId), fn: () => QACDataRender(jobId, token) },
    ];

    Promise.all(apiCalls.filter(({ condition }) => condition).map(({ fn }) => fn()))
        .then(() => {
            updateGridAuditLog(
                jobId,
                auditProcessNames.JobFetchProcessCompleted,
                ""
            );
            setIsValue(true);
            setIsReadonly(true);
            document.body.classList.remove("loading-indicator");
        });

}, []);


  const updateCsrEntry = async (jobid, message, needRedirection) => {
    try {
      await updateGridAuditLog(jobid, message, window.location.href);
      if (needRedirection === true) {
        window.location.href = "/UnAuthorizedUser";
      }
    } catch (error) {
      if (needRedirection === true) {
        window.location.href = "/UnAuthorizedUser";
      }
    }
  }

  const canRenderExclusion = (brokerId) => {
    const ApplicablebrokerIds = sessionStorage.getItem("exclusionApplicableBrokerIds");
    try {
      const parsedData = ApplicablebrokerIds && ApplicablebrokerIds?.length > 0 ? JSON.parse(ApplicablebrokerIds) : [];
      return parsedData?.length > 0 && parsedData?.includes(brokerId) ? true : false;
    } catch (error) { return false }
  }

  const ComponentRenderApiFun = async (jobId) => {
    try {
      let userName = (await ComponentRenderApi(jobId, location.pathname));
      // userName['data'] = "Dispatched" 
      if ((location.pathname === "/xlpage" || location.pathname === "/csrView") && userName?.data && userName?.data === "Dispatched") {
        window.location.href = "/DispatchedJob";
      }
      setDefaultSheetNameData(userName?.data);
      return userName?.data;
    } catch (error) {
      return "";
    }
  }

  const csrAuthenticationTokenApi = async () => {
    document.body.classList.add("loading-indicator");
    let url = new URL(window.location.href);
    let jobId = url?.searchParams.get("jobId") || url?.searchParams.get("jobid");
    try {
      let urlParamToken = url?.searchParams.get("token");
      const decode_UserName = jwt_decode(urlParamToken || "");
      const csrAuthUserName = decode_UserName?.preferred_username;

      if (jobId != undefined && csrAuthUserName != undefined) {
        sessionStorage.setItem("csrAuthUserName", csrAuthUserName);
        const dataForcsrValidata = {
          JobId: jobId,
          csrName: csrAuthUserName,
          token: urlParamToken
        }
        let csrView_UserName = await new Promise((resolve, reject) => {
          axios
            .post(
              `${baseUrl}/api/Authentication/csrValidate`, dataForcsrValidata, {
              headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
              }
            }
            )
            .then((response) => {
              if (response.data == false) {
                if (location.pathname === "/csrView" || location.pathname === "/csrview") {
                  window.location.href = "/UnAuthorizedUser";
                }
              } else {
                resolve(response.data);
              }
            })
            .catch((error) => {
              sessionStorage.setItem("csrValidateErrors", typeof error === 'object' ? JSON.stringify(error?.response || error) : error);
              reject(error);
            });
        });
      } else {
        await updateGridAuditLog(jobId, 'CsrJobValidation-Error', urlParamToken);
        if (location.pathname === "/csrView" || location.pathname === "/csrview") {
          window.location.href = "/UnAuthorizedUser";
        }
      }
    } catch (error) {
      sessionStorage.setItem("csrValidateError", error?.response?.data);
      await updateGridAuditLog(jobId,  'CsrJobValidation-Error',  url?.searchParams.get("token"));
      if (location.pathname === "/csrView" || location.pathname === "/csrview") {
        window.location.href = "/UnAuthorizedUser";
      }
      return error;
    } finally {
      // document.body.classList.remove('loading-indicator');
      return "success";
    }
  };

  const getSheetListsInMasterData = async (jobId) => {
    try {
      const updatedToken = await processAndUpdateToken(token);
      token = updatedToken;
      const headers = {
        Authorization: `Bearer ${updatedToken}`,
      };
      if (jobId) {
        let GetSheetList = await new Promise((resolve, reject) => {
          axios
            .get(`${baseUrl}/api/Master/GetSheetList?jobId=${jobId}`, {
              headers,
            })
            .then((response) => {
              if (response.status !== 200) {
                reject(new Error(`HTTP error! Status: ${response.status}`));
              } else {
                resolve(response.data);
              }
            })
            .catch((error) => {
              reject(error);
            });
        });
        setSheetListOptionSet(GetSheetList);
      }
    } catch (error) {
      console.error("Error:", error);
      return error;
    } finally {
      return "success";
    }
  };

  const csrPolicyData = async (jobId, templateId, sheetType) => {
    try {
      setMasterData();
    } catch (error) { }
    try {
      const updatedToken = await processAndUpdateToken(token);
      token = updatedToken;

      const headers = {
        Authorization: `Bearer ${updatedToken}`,
      };

      if (jobId && templateId && sheetType) {
        const body = {
          JOBID: jobId,
          TEMPLATEID: templateId,
          SheetType: sheetType,
        };

        let csrViewPolicyData = await new Promise((resolve, reject) => {
          axios
            .get(
              `${baseUrl}/api/ProcedureData/GetChecklistDataByjobid?JOBID=${body.JOBID}&TEMPLATEID=${body.TEMPLATEID}&SheetType=${body.SheetType}`,
              { headers }
            )
            .then((response) => {
              if (response.status !== 200) {
                reject(new Error(`HTTP error! Status: ${response.status}`));
              } else {
                resolve(response.data);
              }
            })
            .catch((error) => {
              reject(error);
            });
        });
        if (csrViewPolicyData?.IsApplicable) {
          setCsrChecklistData(csrViewPolicyData?.Response);
        }
      } else {
        throw new Error(
          "Missing required parameters: jobId, templateId, or sheetType"
        );
      }
    } catch (error) {
      console.error("Error:", error);
      return error;
    }
  };

  const csrFormData = async (jobId) => {
    try {
      const updatedToken = await processAndUpdateToken(token);
      token = updatedToken;

      const headers = {
        Authorization: `Bearer ${updatedToken}`,
      };

      if (jobId) {
        const body = {
          JOBID: jobId,
        };

        let csrViewFormData = await new Promise((resolve, reject) => {
          axios
            .get(
              `${baseUrl}/api/ProcedureData/GetGridChecklistFormData?JOBID=${body.JOBID}`,
              { headers }
            )
            .then((response) => {
              if (response.status !== 200) {
                reject(new Error(`HTTP error! Status: ${response.status}`));
              } else {
                resolve(response.data);
              }
            })
            .catch((error) => {
              reject(error);
            });
        });
        if (csrViewFormData?.IsApplicable) {
          if (
            csrViewFormData?.Response &&
            csrViewFormData?.Response?.length > 0
          ) {
            csrViewFormData?.Response.map((e, index) => {
              if (index === 2) {
                let matechedSectionData =
                  typeof e?.TemplateData === "string"
                    ? JSON.parse(e?.TemplateData)
                    : e?.TemplateData;
                if (matechedSectionData && matechedSectionData?.length > 0) {
                  matechedSectionData = matechedSectionData.map(
                    ({ "Document Viewer": _, ...rest }) => rest
                  );
                }
                e["TemplateData"] = JSON.stringify(matechedSectionData);
              }
              return e;
            });
          }
          setCsrFormsData(csrViewFormData?.Response);
        }
      } else {
        throw new Error(
          "Missing required parameters: jobId, templateId, or sheetType"
        );
      }
    } catch (error) {
      console.error("Error:", error);
      return error;
    }
  };

  const dataRender = async (jobId, processName) => {
    try {
      setMasterData();
    } catch (error) { }
    try {
      if (processName == "PolicyReviewChecklist") {
        const lineItemCount = await JobPRTotalCount(jobId);
        if (lineItemCount && lineItemCount > 2000) {
          sessionStorage.setItem("JobPRTotalCount", lineItemCount);
          reference?.current?.showSnackbar(
            `This job has ${lineItemCount} line items. Please be Patience, the data will be rendered soon.`,
            "success",
            true
          );
        }
        const updatedToken = await processAndUpdateToken(token);
        token = updatedToken;
        const headers = { Authorization: `Bearer ${updatedToken}` };
        let jobResponse;
        let csConfigData = [];
        try{
          csConfigData = await GetConfidenceScoreConfig(jobId);
          setConfidenceScoreConfig(csConfigData);
        }catch(error){
          setConfidenceScoreConfig(confidenceScoreConfigStaticData);
        }
        if (jobId) {
          const listOfCalls = [2, 3, 4, 5, 6, 7, 8];
          const headerData = await JobHeaderData(jobId, token);
          let responseDataSet = await Promise.all(
            listOfCalls.map(async (e, index) => {
              // if(e === 1){
              //   document.body.classList.add( 'loading-indicator' );
              //   return { e, data: await JobHeaderData( jobId, token ) };
              // } else
              if (e === 2) {
                document.body.classList.add("loading-indicator");
                return { e, data: await JobCDData(jobId, token) };
              } else if (e === 3) {
                document.body.classList.add("loading-indicator");
                return { e, data: await JobJCData(jobId, token) };
              } else if (e === 4) {
                document.body.classList.add("loading-indicator");
                return { e, data: await JobF1Data(jobId, token) };
              } else if (e === 5) {
                document.body.classList.add("loading-indicator");
                return { e, data: await JobF2Data(jobId, token) };
              } else if (e === 6) {
                document.body.classList.add("loading-indicator");
                return { e, data: await JobF3Data(jobId, token) };
              } else if (e === 7) {
                document.body.classList.add("loading-indicator");
                return { e, data: await JobF4Data(jobId, token) };
              } else if (e === 8) {
                document.body.classList.add("loading-indicator");
                return { e, data: await JobConfigData(jobId, token) };
              }
            })
          );
          responseDataSet = [{ e: 1, data: headerData }, ...responseDataSet];
          let jobResponse = [];
          const listOfCallsSub = [1, 2, 3, 4, 5, 6, 7, 8];
          listOfCallsSub.forEach((f) => {
            if (f != 8) {
              const findData = responseDataSet.find((item) => item?.e === f);
              jobResponse.push(findData?.data);
            } else {
              const find8Data = responseDataSet.find((item) => item?.e === f);
              jobResponse = [...jobResponse, ...find8Data?.data];
            }
          });
          let orderedData = jobResponse.map((e) => {
            if (
              e?.TemplateData &&
              e?.TemplateData?.length > 0 &&
              (e?.Tablename == "Table 2" || e?.Tablename == "Table 3")
            ) {
              const hasColumnId = e.TemplateData.filter(
                (f) => f?.Columnid == null || f?.Columnid == ""
              );
              if (hasColumnId && hasColumnId?.length == 0) {
                e.TemplateData = e.TemplateData.sort(
                  (a, b) => parseInt(a.Columnid) - parseInt(b.Columnid)
                );
              }
              if (e?.Tablename == "Table 3") {
                e.TemplateData = e?.TemplateData?.filter(
                  (f) => !f?.IsDataForSp
                );
              }
            }
            return e;
          });

         
          //updating the OBSERVATION AND PAGENUMBER COLUMNS ON RENDERING ****BEGIN****
          const requiredDataSet = [
            "Table 1",
            "Table 2",
            "Table 3",
            "Table 4",
            "Table 5",
            "Table 6",
            "Table 7",
          ];
          const masterDataForColumns = orderedData.filter(
            (f) => f && f?.Tablename && !requiredDataSet?.includes(f?.Tablename)
          );
          orderedData.map((item, oIndex) => {
            if (
              item &&
              item?.Tablename &&
              requiredDataSet?.includes(item?.Tablename)
            ) {
              if (item?.TemplateData && item?.TemplateData?.length > 0) {
                let applicableColumns = [];
                let appilcableColumnMaster = [];
                if (item?.Tablename === "Table 1") {
                  appilcableColumnMaster = masterDataForColumns.find(
                    (f) => f?.Tablename === "JobHeader"
                  );
                } else if (item?.Tablename === "Table 2") {
                  appilcableColumnMaster = masterDataForColumns.find(
                    (f) => f?.Tablename === "JobCommonDeclaration"
                  );
                } else if (item?.Tablename === "Table 3") {
                  appilcableColumnMaster = masterDataForColumns.find(
                    (f) => f?.Tablename === "JobCoverages"
                  );
                } else if (item?.Tablename === "Table 4") {
                  appilcableColumnMaster = masterDataForColumns.find(
                    (f) => f?.Tablename === "Tbl_ChecklistForm1"
                  );
                } else if (item?.Tablename === "Table 5") {
                  appilcableColumnMaster = masterDataForColumns.find(
                    (f) => f?.Tablename === "Tbl_ChecklistForm2"
                  );
                } else if (item?.Tablename === "Table 6") {
                  appilcableColumnMaster = masterDataForColumns.find(
                    (f) => f?.Tablename === "Tbl_ChecklistForm3"
                  );
                } else if (item?.Tablename === "Table 7") {
                  appilcableColumnMaster = masterDataForColumns.find(
                    (f) => f?.Tablename === "Tbl_ChecklistForm4"
                  );
                }

                if (
                  item?.Tablename !== "Table 1" &&
                  appilcableColumnMaster?.TemplateData
                ) {
                  const columnDetails = isARType(
                    item.TemplateData[0],
                    item?.Tablename
                  );
                  let orderedColumns = getColumnsAvailableInMasterOrder(
                    appilcableColumnMaster?.TemplateData,
                    item?.Tablename
                  );
                  if (orderedColumns && orderedColumns?.length > 0) {
                    item.TemplateData = item.TemplateData.map((e) => {
                      const questionCode =
                        e?.ChecklistQuestions == null || undefined
                          ? ""
                          : e?.ChecklistQuestions.trim().slice(0, 2);
                      let observationText = "";
                      let pageText = "";
                      let originalColumnData = "";
                      orderedColumns.forEach((column) => {
                        const columnKey = getObservationKey(
                          column,
                          item?.Tablename
                        );
                        if (e[column]) {
                          originalColumnData = e[column];
                          let columnDataText = e[column];
                          if (
                            columnDataText &&
                            columnDataText?.toLowerCase()?.trim() !=
                            "matched" &&
                            !columnDataText
                              ?.toLowerCase()
                              ?.includes(
                                "details not available in the document"
                              )
                          ) {
                            columnDataText =
                              splitPageKekFromTextForDataRendering(
                                columnDataText
                              );
                            observationText +=
                              columnKey + columnDataText + " ~~ ";
                          } else if (observationText != "") {
                            observationText +=
                              columnKey + "NO RECORDS" + " ~~ ";
                          }
                        } else {
                          observationText += columnKey + "NO RECORDS" + " ~~ ";
                        }
                        //page number updation
                        const columnPageKeyCode = getPageKey(
                          column,
                          item?.Tablename,
                          originalColumnData
                        );
                        if (columnPageKeyCode) {
                          const pageNumber = getTextByRequirement(
                            originalColumnData,
                            "getPage",
                            column
                          );
                          originalColumnData = originalColumnData?.replace(
                            /\s+/g,
                            " "
                          );
                          if (
                            originalColumnData &&
                            originalColumnData
                              ?.toLowerCase()
                              ?.includes("endorsement page #")
                          ) {
                            pageText +=
                              columnPageKeyCode?.trim() +
                              "E" +
                              questionCode?.trim() +
                              ":" +
                              pageNumber +
                              " ~~ ";
                          } else {
                            pageText +=
                              columnPageKeyCode?.trim() +
                              questionCode?.trim() +
                              ":" +
                              pageNumber +
                              " ~~ ";
                          }
                        }
                      });
                      e["Observation"] = observationText;
                      if (observationText != "") {
                        e["PageNumber"] = pageText;
                      } else {
                        e["PageNumber"] = "";
                      }
                      return e;
                    });
                  }
                }
              }
              return item;
            }
          });
          //updating the OBSERVATION AND PAGENUMBER COLUMNS ON RENDERING ****END****
          let EnableExclusionCellLockVlaue = true;
          orderedData = orderedData?.map((e) => {
            if(e.Tablename == "JobHeader" && e?.EnableCS && e?.MetaData){
                const parsedData = JSON.parse(e?.MetaData);
                const usr_name = sessionStorage.getItem("userName");
                if(parsedData && parsedData?.QuestionCode?.length > 0 && parsedData?.UserName && 
                  parsedData?.UserName?.toLowerCase()?.includes(usr_name?.toLowerCase())
                ){
                    let sqCodes = parsedData.QuestionCode.split(',');
                    if(sqCodes?.length > 0){
                        sqCodes = sqCodes.map((code) => {return code?.trim()?.toLowerCase()});
                        e.StpMappings = e.StpMappings.filter((f) => !sqCodes?.includes(f?.ShortQuestion?.trim()?.toLowerCase())  )
                    }
                }
              }
              if(e.Tablename == "JobHeader"){
                EnableExclusionCellLockVlaue = e?.EnableExclusionCellLock;
              }
              return e;
            });
          console.log("orderedData",orderedData);
          setChecklistData(orderedData);
          setEnableExclusionCellLock(EnableExclusionCellLockVlaue);
          setSheetRenderConfig({
            PolicyReviewChecklist: "true",
            FormsCompare: "false",
            Exclusion: "false",
            QAC_not_answered_questions: "false"
          });
        }
        reference?.current?.showSnackbar("Data Rendered", "success", true);
        setTimeout(() => {
          reference?.current.hideSnackbar();
        }, 500);
      }
    } catch (error) {
      console.error("Error:", error);
      setIsValue(true);
      setMsgVisible(true);
      setMsgClass("alert error");
      setMsgText(`Error fetching data: ${error}`);
      setTimeout(() => {
        setMsgVisible(false);
        setMsgText("");
      }, 3500);
      return error;
    } finally {
      document.body.classList.remove("loading-indicator");
      return "success";
    }
  };

  const getColumnsAvailableInMasterOrder = (masterData, tableName) => {
    if (tableName && masterData) {
      const orderedColumnData = [];
      masterData.forEach((f) => {
        if (
          f === "Current Term Policy" &&
          (tableName === "Table 2" || tableName === "Table 3")
        )
          orderedColumnData.push("CurrentTermPolicy");
        if (
          f === "Prior Term Policy" &&
          (tableName === "Table 2" || tableName === "Table 3")
        )
          orderedColumnData.push("PriorTermPolicy");
        if (
          f === "Proposal" &&
          (tableName === "Table 2" || tableName === "Table 3")
        )
          orderedColumnData.push("Proposal");
        if (
          f === "Quote" &&
          (tableName === "Table 2" || tableName === "Table 3")
        )
          orderedColumnData.push("Quote");
        if (
          f === "Schedule" &&
          (tableName === "Table 2" || tableName === "Table 3")
        )
          orderedColumnData.push("Schedule");
        if (
          f === "Application" &&
          (tableName === "Table 2" || tableName === "Table 3")
        )
          orderedColumnData.push("Application");
        if (
          f === "Binder" &&
          (tableName === "Table 2" || tableName === "Table 3")
        )
          orderedColumnData.push("Binder");
        if (f === "Current Term Policy - Listed" && tableName === "Table 4")
          orderedColumnData.push("CurrentTermPolicyListed");
        if (
          f === "Prior Term Policy - Listed" &&
          (tableName === "Table 4" || tableName === "Table 5")
        )
          orderedColumnData.push("PriorTermPolicyListed");
        if (
          f === "Proposal - Listed" &&
          (tableName === "Table 4" || tableName === "Table 5")
        )
          orderedColumnData.push("ProposalListed");
        if (
          f === "Binder - Listed" &&
          (tableName === "Table 4" || tableName === "Table 5")
        )
          orderedColumnData.push("BinderListed");
        if (
          f === "Schedule - Listed" &&
          (tableName === "Table 4" || tableName === "Table 5")
        )
          orderedColumnData.push("ScheduleListed");
        if (
          f === "Quote - Listed" &&
          (tableName === "Table 4" || tableName === "Table 5")
        )
          orderedColumnData.push("QuoteListed");
        if (
          f === "Application - Listed" &&
          (tableName === "Table 4" || tableName === "Table 5")
        )
          orderedColumnData.push("ApplicationListed");
        if (f === "Current Term Policy - Listed" && tableName === "Table 5")
          orderedColumnData.push("CurrentTermPolicyListed1");
        if (
          f === "Current Term Policy Attached" &&
          (tableName === "Table 5" ||
            tableName === "Table 6" ||
            tableName === "Table 7")
        )
          orderedColumnData.push("CurrentTermPolicyAttached");
        if (
          f === "Current Term Policy Listed" &&
          (tableName === "Table 5" ||
            tableName === "Table 6" ||
            tableName === "Table 7")
        )
          orderedColumnData.push("CurrentTermPolicyListed");
      });
      return orderedColumnData;
    }
    return [];
  };

  const formRender = async (jobId, processName) => {
    try {
      const updatedToken = await processAndUpdateToken(token);
      token = updatedToken;
      const headers = {
        Authorization: `Bearer ${updatedToken}`,
      };
      if (jobId && processName == "Forms Compare") {
        let csConfigData = [];
        try{
          csConfigData = await GetConfidenceScoreConfig(jobId);
          setConfidenceScoreConfig(csConfigData);
        }catch(error){
          setConfidenceScoreConfig(confidenceScoreConfigStaticData);
        }
        let formData = await new Promise((resolve, reject) => {
          axios
            .get(
              `${baseUrl}/api/ProcedureData/GetFormDataByJobId?jobId=${jobId}`,
              { headers }
            )
            .then((response) => {
              if (response.status !== 200) {
                reject(new Error(`HTTP error! Status: ${response.status}`));
              } else {
                resolve(response.data);
              }
            })
            .catch((error) => {
              reject(error);
            });
        });
        if (
          formData[0]?.Tablename == "FormTable 1" &&
          formData[0]?.TemplateData &&
          Array.isArray(formData[0]?.TemplateData)
        ) {
          let templateData = formData[0]?.TemplateData;
          const modifiedTemplateData = [];
          const hasNameInsured = templateData?.filter((f) =>
            f?.Headers?.toLowerCase()?.includes("named insured")
          );
          const hasTerm = templateData?.filter((f) =>
            f?.Headers?.toLowerCase()?.includes("term")
          );
          const hasLob = templateData?.filter((f) =>
            f?.Headers?.toLowerCase()?.includes("lob")
          );
          const hasPol = templateData?.filter((f) =>
            f?.Headers?.toLowerCase()?.includes("pol#")
          );
          const hasCarrier = templateData?.filter((f) =>
            f?.Headers?.toLowerCase()?.includes("carrier name")
          );
          if (hasNameInsured?.length == 0) {
            modifiedTemplateData.push({
              Id: 0,
              HeaderId: 0,
              Jobid: templateData[0]?.Jobid,
              PolicyLob: templateData[0]?.PolicyLob,
              Headers: "Named Insured",
              NoColumnName: "",
            });
          }
          if (hasTerm?.length == 0) {
            modifiedTemplateData.push({
              Id: 0,
              HeaderId: 0,
              Jobid: templateData[0]?.Jobid,
              PolicyLob: templateData[0]?.PolicyLob,
              Headers: "Term",
              NoColumnName: "",
            });
          }
          if (hasLob?.length == 0) {
            modifiedTemplateData.push({
              Id: 0,
              HeaderId: 0,
              Jobid: templateData[0]?.Jobid,
              PolicyLob: templateData[0]?.PolicyLob,
              Headers: "LOB",
              NoColumnName: "",
            });
          }
          if (hasPol?.length == 0) {
            modifiedTemplateData.push({
              Id: 0,
              HeaderId: 0,
              Jobid: templateData[0]?.Jobid,
              PolicyLob: templateData[0]?.PolicyLob,
              Headers: "Pol#",
              NoColumnName: "",
            });
          }
          if (hasCarrier?.length == 0) {
            modifiedTemplateData.push({
              Id: 0,
              HeaderId: 0,
              Jobid: templateData[0]?.Jobid,
              PolicyLob: templateData[0]?.PolicyLob,
              Headers: "Carrier Name",
              NoColumnName: "",
            });
          }
          formData[0].TemplateData = [...modifiedTemplateData, ...templateData];
        }
        // if ( formData[ 1 ]?.Tablename == "FormTable 2" && formData[ 1 ]?.TemplateData?.length > 0 )
        // {
        //   formData[ 1 ].TemplateData = formsCompareDuplicateDataRemover( formData[ 1 ].TemplateData );
        //   formData[ 1 ].TemplateData = formData[ 1 ].TemplateData.map( ( e ) => {
        //     e[ "ChecklistQuestions" ] = "CA2";
        //     e[ "CoverageSpecificationsMaster" ] = "Attached Forms";
        //     return e;
        //   } );
        // }
        // if ( formData[ 2 ]?.Tablename == "FormTable 3" && formData[ 2 ]?.TemplateData?.length > 0 )
        // {
        //   formData[ 2 ].TemplateData = formsCompareDuplicateDataRemover( formData[ 2 ].TemplateData );
        //   formData[ 2 ].TemplateData = formData[ 2 ].TemplateData.map( ( e ) => {
        //     e[ "ChecklistQuestions" ] = "CA2";
        //     e[ "CoverageSpecificationsMaster" ] = "Attached Forms";
        //     return e;
        //   } );
        // }
        formData = formData.map((e) => {
          if(e?.Tablename == "FormTable 1" && e?.EnableCS && e?.MetaData){
            const parsedData = JSON.parse(e?.MetaData);
            const usr_name = sessionStorage.getItem("userName");
            if(parsedData && parsedData?.QuestionCode?.length > 0 && parsedData?.UserName && 
              parsedData?.UserName?.toLowerCase()?.includes(usr_name?.toLowerCase())
            ){
                let sqCodes = parsedData.QuestionCode.split(',');
                if(sqCodes?.length > 0){
                    sqCodes = sqCodes.map((code) => {return code?.trim()?.toLowerCase()});
                    e.StpMappings = e.StpMappings.filter((f) => !sqCodes?.includes(f?.ShortQuestion?.trim()?.toLowerCase())  )
                }
            }
          }
          return e;
        });
        setFormData(formData);
        setSheetRenderConfig({
          PolicyReviewChecklist: "false",
          FormsCompare: "true",
          Exclusion: "false",
          QAC_not_answered_questions: "false",
        });
      }
    } catch (error) {
      console.error("Error:", error);
      return error;
    } finally {
      return "success";
    }
  };

  const exclusionRender = async (jobId, processName) => {
    try {
      const updatedToken = await processAndUpdateToken(token);
      token = updatedToken;
      const headers = {
        Authorization: `Bearer ${updatedToken}`,
      };
      if (jobId && (processName == "Exclusion" || processName == "csrExclusion")) {
        let exclusionData = await new Promise((resolve, reject) => {
          axios
            .get(
              `${baseUrl}/api/ProcedureData/GetExclusionData?jobId=${jobId}`,
              { headers }
            )
            .then((response) => {
              if (response.status !== 200) {
                reject(new Error(`HTTP error! Status: ${response.status}`));
              } else {
                resolve(response.data);
              }
            })
            .catch((error) => {
              reject(error);
            });
        });
        if (exclusionData && exclusionData?.length > 0) {
          if(processName != "csrExclusion"){
            exclusionData = exclusionData?.map((Obj, index) => {
              if(!Obj?.ConfidenceScore){
                Obj["ConfidenceScore"] = "Details not available in the document";
              }
              return Obj;
            });
          }          
        } else {
          if (processName == "Exclusion") {
            exclusionData = staticExclusionData
          } else if (processName == "csrExclusion") {
            exclusionData = staticExclusionCsrData;
          }
        }
        console.log(exclusionData)
        setExclusionData(exclusionData);
        setSheetRenderConfig({
          PolicyReviewChecklist: "false",
          FormsCompare: "false",
          Exclusion: "true",
          QAC_not_answered_questions: "false",
        });
      }
    } catch (error) {
      console.error("Error:", error);
      return error;
    } finally {
      return "success";
    }
  };

  const QACDataRender = async (jobId, token) => {
    const defaultData = { data: {} };
    if (jobId) {
      try {
        const Token = await processAndUpdateToken(token);
        const headers = {
          'Authorization': `Bearer ${Token}`,
          "Content-Type": "application/json",
        };
        const response = await axios.get(baseUrl + '/api/Defaultdatum/GetQacChecklistData?jobId=' + jobId, {
          headers
        });
        if (response.status !== 200) {
          return defaultData;
        }
        if (response?.data && response?.data?.JsonData) {
          if (response?.data?.JsonData) {
            setCsrQacData(response?.data?.JsonData);
            return { data: response?.data?.JsonData };
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
    return defaultData;
  };

  const renderComponent = () => {
    const getUserName = defaultSheetNameData;
    console.log("getUserName", getUserName)
    let sessionUserName = sessionStorage.getItem('userName');

    if ((location.pathname === "/csrView" || location.pathname === "/csrview") && jobId) {
      return (
        <CsrSheet
          data={csrChecklistData}
          formCompareData={csrFormsData}
          exclusionRenderData={exclusionData}
          qacData={csrQacData}
          gradtionDataSet={gradtionData}
          qacdataapi={Qacdata}
          defaultSheetUserNameData={defaultSheetNameData}
          selectedJob={jobId}
          selectedJobName={"jobName"}
        />
      );
    }

    const commonProps = {
      data: checklistData,
      enableCs: checklistData?.find((f) => f.Tablename === "JobHeader")?.EnableCS,
      enableCellLock: checklistData?.find((f) => f.Tablename === "JobHeader")?.EnableCellLock,
      enableExclusionCellLock: enableExclusionCellLock,
      csMeatData: checklistData?.find((f) => f.Tablename === "JobHeader")?.MetaData,
      formCompareData: formData,
      formsCompareHeaderData: formData?.find((f) => f.Tablename === "FormTable 1"),
      exclusionRenderData: exclusionData,
      selectedJob: jobId,
      qacdataapi: Qacdata,
      gradtionDataSet: gradtionData,
      selectedJobName: "",
      sheetOptionSet: sheetListOptionSet,
      selectedSheet: selectedSheet,
      sheetRenderConfig: sheetRenderConfig,
      confidenceScoreConfig: confidenceScoreConfig,
      selectChange: async (value) => {
        setIsValue(false);
        setIsReadonly(false);
        await ComponentRenderApiFun();
        if (value === "Forms Compare") {
          await formRender(jobId, "Forms Compare");
        } else if (value === "PolicyReviewChecklist") {
          await dataRender(jobId, "PolicyReviewChecklist");
        } else if (value === "Exclusion") {
          await exclusionRender(jobId, "Exclusion");
        } else if (value === "QAC not answered questions") {
          setSheetRenderConfig({
            PolicyReviewChecklist: "false",
            FormsCompare: "false",
            Exclusion: "false",
            QAC_not_answered_questions: "true",
          });
        }
        setSelectedSheet(value);
        setIsValue(true);
        setIsReadonly(true);
      },
    };

    return (getUserName === sessionUserName || !getUserName) ? (
      <Luckysheet {...commonProps} />
    ) : (
      <DefaultDataSheet {...commonProps} />
    );
  };

  const GetGradationData = async (jobId, path) => {
    try {
      if (jobId) {
        const updatedToken = await processAndUpdateToken(token);
        token = updatedToken;
        const headers = { Authorization: `Bearer ${updatedToken}` };
        let url = "";
        if (path === "/csrView" || path === "/csrview") {
          url = "/api/ProcedureData/GetGradiationGridData?JobId=";
        } else {
          url = "/api/ProcedureData/GetGradationDataByJobId?JobId=";
        }
        const response = await axios.get(baseUrl + url + jobId, {
          headers,
        });
        if (response.status == 200) {
          if (path === "/csrView" || path === "/csrview") {
            if (response?.data?.IsApplicable) {
              setGradtionData(response?.data?.Response);
            }
          } else {
            setGradtionData(response.data);
          }
          return "success";
        } else {
          //handle error's
          // throw new Error( `HTTP error! Status: ${ response.status }` );
        }
      }
    } catch (error) {
      //handle error
    }
  };

  const ExportQac = async () => {
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
    };
    axios
      .get(baseUrl + "/api/Excel/SplitDataByPolicyLob", { headers })
      .then((response) => {
        if (response.status !== 200) {
          throw new Error(`HTTP error! Status: ${response.status}`);
        }
        const qacdata = response.data;
        setQacdata(qacdata);
      })
      .catch((error) => {
        // Handle errors...
      })
      .finally(() => {
        // document.body.classList.remove('loading-indicator');
      });
  };

  return (
    <div className="main-container">
      <div className="app-container">
        {/* <div>
          <AppSidebar />
        </div> */}
        <div>
          {msgVisible && (
            <div className="alert-container">
              <div className={msgClass}>{msgText}</div>
            </div>
          )}
          {isValue && isReadonly && renderComponent()}
          <br />
        </div>
        <SimpleSnackbarWithOutTimeOut ref={reference} />
      </div>
    </div>
  );
}
export default App;