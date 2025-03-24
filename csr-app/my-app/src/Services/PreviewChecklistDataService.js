import axios from "axios";
import { baseUrl } from "./Constants";
import { processAndUpdateToken } from "./CommonFunctions";
import { SimpleSnackbar } from "../Components/SnackBar";
import { auditProcessNames } from "./enums";

export const JobHeaderData = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistHeaderData?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

export const JobCDData = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistCdData?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

export const JobJCData = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistJcData?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

const getFormsQuestionBasedOnLob = (data) => {
  const isPT =
    data?.filter(
      (f) =>
        f["PolicyLob"]?.toLowerCase()?.trim() ===
          "do the forms, endorsements and edition dates match the source documents?" ||
        f["PolicyLob"]?.toLowerCase()?.trim() ===
          "do the forms, endorsements and edition dates match the expiring policy?"
    )?.length > 0;
  if (isPT) {
    return "PT1:Do the Forms, Endorsements and Edition Dates match the expiring policy and Source Document?";
  }

  const isCL =
    data?.filter(
      (f) =>
        f["PolicyLob"]?.toLowerCase()?.trim() ===
        "are the forms and endorsements listed, attached in current term policy?"
    )?.length > 0;
  if (isCL) {
    return "CL1:Are the Forms and Endorsements listed, attached in current term policy?";
  }

  const isCA =
    data?.filter(
      (f) =>
        f["PolicyLob"]?.toLowerCase()?.trim() ===
        "are the forms and endorsements attached, listed in current term policy?"
    )?.length > 0;
  if (isCA) {
    return "CA1:Are the Forms and Endorsements attached, listed in current term policy?";
  }
};

export const JobF1Data = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistForm1Data?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      if (response?.data?.TemplateData?.length > 0) {
        const columnKeys = [
          "CurrentTermPolicyListed",
          "ProposalListed",
          "QuoteListed",
          "PriorTermPolicyListed",
          "ScheduleListed",
          "BinderListed",
          "ApplicationListed",
        ];
        const question = getFormsQuestionBasedOnLob(
          response?.data?.TemplateData
        );
        response?.data?.TemplateData?.map((item) => {
          columnKeys.forEach((key) => {
            if (key) {
              const value = item[key];
              if (value) {
                const cs_value = item[key + "Cs"];
                if (!cs_value) {
                  item[key + "Cs"] = "Details not available in the document";
                }
                // if(value?.trim()?.toLowerCase() === "details not available in the document"){
                //   item[key + "Cs"] = "Details not available in the document";
                // }else if(value?.trim()?.toLowerCase() === "matched"){
                //   item[key + "Cs"] = "";
                // }else{
                //   if(!item[key + "Cs"]){
                //     item[key + "Cs"] = "Details not available in the document";
                //   }
                // }
              }
              // else{
              //   item[key + "Cs"] = "";
              // }
            }
          });
          if (question) {
            item["ChecklistQuestions"] = question;
          }
          return item;
        });
      }
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

export const JobF2Data = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistForm2Data?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      // CS - data formatting
      if (response?.data?.TemplateData?.length > 0) {
        const columnKeys = [
          "CurrentTermPolicyListed",
          "CurrentTermPolicyListed1",
          "CurrentTermPolicyAttached",
          "PriorTermPolicyListed",
        ];
        const question = getFormsQuestionBasedOnLob(
          response?.data?.TemplateData
        );
        response?.data?.TemplateData?.map((item) => {
          columnKeys.forEach((key) => {
            if (key) {
              const value = item[key];
              if (value) {
                if (key === "CurrentTermPolicyListed1") {
                  const cs_value = item["CurrentTermPolicyListedCs1"];
                  if (!cs_value) {
                    item["CurrentTermPolicyListedCs1"] =
                      "Details not available in the document";
                  }
                } else {
                  const cs_value = item[key + "Cs"];
                  if (!cs_value) {
                    item[key + "Cs"] = "Details not available in the document";
                  }
                }
                // if(value?.trim()?.toLowerCase() === "details not available in the document"){
                //   if(key === "CurrentTermPolicyListed1"){
                //     item["CurrentTermPolicyListedCs1"] = "Details not available in the document";
                //   }else{
                //     item[key + "Cs"] = "Details not available in the document";
                //   }
                // }else if(value?.trim()?.toLowerCase() === "matched"){
                //   if(key === "CurrentTermPolicyListed1"){
                //     item["CurrentTermPolicyListedCs1"] = "";
                //   }else{
                //     item[key + "Cs"] = "";
                //   }
                // }else{
                //   if(key === "CurrentTermPolicyListed1"){
                //     item["CurrentTermPolicyListedCs1"] = "Details not available in the document";
                //   }else{
                //     item[key + "Cs"] = "Details not available in the document";
                //   }
                // }
              }
              // else{
              //   if(key === "CurrentTermPolicyListed1"){
              //     item["CurrentTermPolicyListedCs1"] = "";
              //   }else{
              //     item[key + "Cs"] = "";
              //   }
              // }
            }
          });
          if (question) {
            item["ChecklistQuestions"] = question;
          }
          return item;
        });
      }
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

export const JobF3Data = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistForm3Data?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      if (response?.data?.TemplateData?.length > 0) {
        const columnKeys = [
          "CurrentTermPolicyListed",
          "CurrentTermPolicyAttached",
        ];
        const question = getFormsQuestionBasedOnLob(
          response?.data?.TemplateData
        );
        response?.data?.TemplateData?.map((item) => {
          columnKeys.forEach((key) => {
            if (key) {
              const value = item[key];
              if (value) {
                const cs_value = item[key + "Cs"];
                if (!cs_value) {
                  item[key + "Cs"] = "Details not available in the document";
                }
                // if(value?.trim()?.toLowerCase() === "details not available in the document"){
                //   item[key + "Cs"] = "Details not available in the document";
                // }else if(value?.trim()?.toLowerCase() === "matched"){
                //   item[key + "Cs"] = "";
                // }else{
                //   item[key + "Cs"] = "Details not available in the document";
                // }
              }
              // else{
              //   item[key + "Cs"] = "";
              // }
            }
          });
          if (question) {
            item["ChecklistQuestions"] = question;
          }
          return item;
        });
      }
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

export const JobF4Data = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistForm4Data?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      if (response?.data?.TemplateData?.length > 0) {
        const columnKeys = [
          "CurrentTermPolicyListed",
          "CurrentTermPolicyAttached",
        ];
        const question = getFormsQuestionBasedOnLob(
          response?.data?.TemplateData
        );
        response?.data?.TemplateData?.map((item) => {
          columnKeys.forEach((key) => {
            if (key) {
              const value = item[key];
              if (value) {
                const cs_value = item[key + "Cs"];
                if (!cs_value) {
                  item[key + "Cs"] = "Details not available in the document";
                }
                // if(value?.trim()?.toLowerCase() === "details not available in the document"){
                //   item[key + "Cs"] = "Details not available in the document";
                // }else if(value?.trim()?.toLowerCase() === "matched"){
                //   item[key + "Cs"] = "";
                // }else{
                //   item[key + "Cs"] = "Details not available in the document";
                // }
              }
              // else{
              //   item[key + "Cs"] = "";
              // }
            }
          });
          if (question) {
            item["ChecklistQuestions"] = question;
          }
          return item;
        });
      }
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};
export const JobConfigData = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetChecklistAppConfigData?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};


export const FormJobConfigData = async (jobId, token) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/ProcedureData/GetFormChecklistAppConfigData?JobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

let token = sessionStorage.getItem("token");
const formCompareUpdateTable1 = async (jobId, tableName, json) => {
  // document.body.classList.add("loading-indicator");
  const Token = await processAndUpdateToken(token);
  token = Token;
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };
  const apiUrl = `${baseUrl}/api/ProcedureData/UpdateFormHeaderData?jobId=${jobId}`;

  axios({
    method: "POST",
    url: apiUrl,
    headers: headers,
    data: {
      JobId: jobId,
      TableName: tableName,
      NewTemplateData: json,
    },
  })
    .then((response) => {
      if (response.status !== 200) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      return response.data;
    })
    .then((data) => {
      if (data?.status == 400) {
        //    console.log(data);
      }
    })
    .finally(() => {
      // document.body.classList.remove("loading-indicator");
    });
};

const updateHeaderData = async (data, token, needLoader) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/UpdateJobHeader",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(
        data?.JobId,
        "Header - " + needLoader
          ? auditProcessNames.JobUpdateProcessFailed
          : auditProcessNames.JobUpdateProcessAutoSaveFailed,
        message
      );
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};

const updateJDData = async (data, token, needLoader) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/UpdateJobCommonDeclaration",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(
        data?.JobId,
        "JD - " + needLoader
          ? auditProcessNames.JobUpdateProcessFailed
          : auditProcessNames.JobUpdateProcessAutoSaveFailed,
        message
      );
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};

const updateJCData = async (data, token, needLoader) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/UpdateJobCoverage",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(
        data?.JobId,
        "JC - " + needLoader
          ? auditProcessNames.JobUpdateProcessFailed
          : auditProcessNames.JobUpdateProcessAutoSaveFailed,
        message
      );
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};
const updateJobForm1Data = async (data, token, needLoader) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/UpdateJobForm1",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(
        data?.JobId,
        "F1 - " + needLoader
          ? auditProcessNames.JobUpdateProcessFailed
          : auditProcessNames.JobUpdateProcessAutoSaveFailed,
        message
      );
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};
const updateJobForm2Data = async (data, token, needLoader) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/UpdateJobForm2",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(
        data?.JobId,
        "F2 - " + needLoader
          ? auditProcessNames.JobUpdateProcessFailed
          : auditProcessNames.JobUpdateProcessAutoSaveFailed,
        message
      );
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};
const updateJobForm3Data = async (data, token, needLoader) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/UpdateJobForm3",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(
        data?.JobId,
        "F3 - " + needLoader
          ? auditProcessNames.JobUpdateProcessFailed
          : auditProcessNames.JobUpdateProcessAutoSaveFailed,
        message
      );
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};
const updateJobForm4Data = async (data, token, needLoader) => {
  try {
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/UpdateJobForm4",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(
        data?.JobId,
        "F4 - " + needLoader
          ? auditProcessNames.JobUpdateProcessFailed
          : auditProcessNames.JobUpdateProcessAutoSaveFailed,
        message
      );
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};

const commonFunTable1 = async (data, token, needLoader) => {
  await updateHeaderData(data, token, needLoader);
  // let columHeaderchange = data.NewTemplateData;
  // columHeaderchange = columHeaderchange.map((item) => {
  //   const { NoColumnName, HeaderId, jobid, ...rest } = item;
  //   return { ...rest, "": NoColumnName, HeaderID: HeaderId, JOBID: jobid };
  // });
  // await formCompareUpdateTable1(
  //   data?.JobId,
  //   "FormTable 1",
  //   JSON.stringify(columHeaderchange)
  // );
};

export const apiCallSwitch = async (data, token, needLoader) => {
  try {
    if (
      data?.TableName === "Table 1" &&
      data?.NewTemplateData &&
      data?.NewTemplateData?.length > 0
    ) {
      // return await updateHeaderData( data, token );
      await commonFunTable1(data, token, needLoader);
    } else if (
      data?.TableName === "Table 2" &&
      data?.NewTemplateData &&
      data?.NewTemplateData?.length > 0
    ) {
      return await updateJDData(data, token, needLoader);
    } else if (
      data?.TableName === "Table 3" &&
      data?.NewTemplateData &&
      data?.NewTemplateData?.length > 0
    ) {
      return await updateJCData(data, token, needLoader);
    } else if (
      data?.TableName === "Table 4" &&
      data?.NewTemplateData &&
      data?.NewTemplateData?.length > 0
    ) {
      return await updateJobForm1Data(data, token, needLoader);
    } else if (
      data?.TableName === "Table 5" &&
      data?.NewTemplateData &&
      data?.NewTemplateData?.length > 0
    ) {
      return await updateJobForm2Data(data, token, needLoader);
    } else if (
      data?.TableName === "Table 6" &&
      data?.NewTemplateData &&
      data?.NewTemplateData?.length > 0
    ) {
      return await updateJobForm3Data(data, token, needLoader);
    } else if (
      data?.TableName === "Table 7" &&
      data?.NewTemplateData &&
      data?.NewTemplateData?.length > 0
    ) {
      return await updateJobForm4Data(data, token, needLoader);
    }
  } catch (error) {
    return error;
  }
};

export const UpdateJobPreviewStatus = async (jobid, token) => {
  try {
    if (jobid) {
      const Token = await processAndUpdateToken(token);
      const user = sessionStorage.getItem("csrAuthUserName");

      const headers = {
        Authorization: `Bearer ${Token}`,
        "Content-Type": "application/json",
      };

      // Use the correct method and URL structure
      const url = new URL(`${baseUrl}/api/Master/UpdateJobPreviewStatus`);
      url.searchParams.append("user", user);
      url.searchParams.append("jobId", jobid);

      // Make the PUT request
      const response = await axios.put(url.toString(), null, { headers });
      if (response.status !== 200) {
        return "failed";
      }
      if (response?.data) {
        return response.data;
      } else {
        return "failed";
      }
    }
  } catch (error) {
    console.error("Error updating job preview status:", error);
    // Optionally, handle the error appropriately
  }
};

export const UpdateJobSendPolicyInsured = async (jobid, token) => {
  try {
    if (jobid) {
      const Token = await processAndUpdateToken(token);
      const user = sessionStorage.getItem("csrAuthUserName");

      const headers = {
        Authorization: `Bearer ${Token}`,
        "Content-Type": "application/json",
      };

      // Use the correct method and URL structure
      const url = new URL(`${baseUrl}/api/Master/UpdateJobSecoundaryStatus`);
      url.searchParams.append("user", user);
      url.searchParams.append("jobId", jobid);

      // Make the PUT request
      const response = await axios.put(url.toString(), null, { headers });
      if (response.status !== 200) {
        return "failed";
      }
      if (response?.data) {
        return response.data;
      } else {
        return "failed";
      }
    }
  } catch (error) {
    console.error("Error updating job preview status:", error);
    // Optionally, handle the error appropriately
  }
};

export const updateGridAuditLog = async (
  jobId,
  processName,
  Message,
  userName
) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const data = {
      JobId: jobId,
      CreatedBy:
        userName ||
        sessionStorage.getItem("userName") ||
        sessionStorage.getItem("csrAuthUserName"),
      ProcessName: processName,
      Message: Message,
    };
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/updateGridAuditLog",
      data,
      {
        headers,
      }
    );
  } catch (error) {
    console.log(error);
  }
};

export const TriggerBackUp = async (jobId, SheetName) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const userName = sessionStorage.getItem("userName");
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/ProcedureData/TransferDataToGridBackUpTables",
      {
        JobId: jobId,
        User: userName,
        SheetName: SheetName,
      },
      {
        headers,
      }
    );
    if (response.status !== 200) {
      let message = JSON?.parse(response);
      updateGridAuditLog(jobId, "Data transfer failed", message);
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};

export const JobPRTotalCount = async (jobId) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/Defaultdatum/GetTotalCount?jobId=" + jobId,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "failed";
  }
};

export const ComponentRenderApi = async (jobId, location) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const path =
      location === "/xlpage"
        ? "/api/Defaultdatum/GetGridJobActiveUserById?jobId="
        : "/api/Defaultdatum/GetCsrGridJobActiveUserById?jobId=";

    const response = await axios.get(baseUrl + path + jobId, {
      headers,
    });
    if (response.status !== 200) {
      return "failed";
    }
    return response;
  } catch (error) {
    return "failed";
  }
};

export const GetConfidenceScoreConfig = async (jobId) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const response = await axios.get(
      baseUrl + `/api/JobConfiguration/GetConfidenceScoreConfig?jobId=${jobId}`,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    return response?.data;
  } catch (error) {
    return "failed";
  }
};

export const UpdateAppConfigMeta = async (data) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/JobConfiguration/UpdateAppConfigMeta",
      data,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    if (response?.data) {
      return response.data;
    } else {
      return "failed";
    }
  } catch (error) {
    return "update failed";
  }
};

export const GetAppConfigMeta = async (jobId) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const response = await axios.get(
      baseUrl + `/api/JobConfiguration/GetAppConfigMeta?JobId=${jobId}`,
      {
        headers,
      }
    );
    if (response.status !== 200) {
      return "failed";
    }
    return response?.data;
  } catch (error) {
    return "failed";
  }
}; 