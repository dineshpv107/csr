import axios from "axios";
import { baseUrl } from "./Constants";
import { processAndUpdateToken } from "./CommonFunctions";

export const quoteCompareTblsData = [
  {
    "HeaderID": "11",
    "JOBID": " ",
    "Headers": "Named Insured",
    "": ""
  },
  {
    "HeaderID": "21",
    "JOBID": " ",
    "Headers": "Term",
    "": ""
  },
  {
    "HeaderID": "31",
    "JOBID": " ",
    "Headers": "LOB",
    "": ""
  },
  {
    "HeaderID": "41",
    "JOBID": " ",
    "Headers": "Pol#",
    "": ""
  },
  {
    "HeaderID": "51",
    "JOBID": " ",
    "Headers": "Carrier Name",
    "": ""
  },
  {
    "HeaderID": "61",
    "JOBID": " ",
    "Headers": "Account Manager",
    "": " "
  },
  {
    "HeaderID": "71",
    "JOBID": " ",
    "Headers": "Checked by",
    "": ""
  },
  {
    "HeaderID": "81",
    "JOBID": " ",
    "Headers": "Checked on",
    "": ""
  },
  {
    "HeaderID": "91",
    "JOBID": " ",
    "Headers": "Documents used for policy review",
    "": ""
  }
]

export const GetAllJobInfo = async () => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/QuoteCompare/GetAllBrokerQCJobs",
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

export const GetJobInfo = async (brokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/QuoteCompare/GetJobs?brokerId=" + brokerId,
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
export const GetDiscrepancyInfo = async (brokerId, JobId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/QuoteCompare/GetDiscrepancrJobs?brokerId=" + brokerId + "&jobId=" + JobId,
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
export const GetBrokerList = async () => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/Defaultdatum/GetBrokerNames/GetBrokerNames",
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

export const GetTbl1Data = async (BrokerId, JobId) => {

  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl +
      `/api/QuoteCompare/GetJobInfo?brokerid=${BrokerId}&jobid=${JobId}`,
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

export const GetJobFiles = async (JobId, BrokerId) => {
  
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl +
      `/api/QuoteCompare/GetJobFiles?JobId=${JobId}&brokerId=${BrokerId}`,
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

export const GetJobQCData = async (Id, BrokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + `/api/QuoteCompare/GetJobData?id=${Id}&brokerId=${BrokerId}`,
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

export const GetLobQuestionsByJobId = async (Id, BrokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl +
      `/api/QuoteCompare/GetLobQuestionsByJobId?id=${Id}&brokerId=${BrokerId}`,
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

export const GetLobQuestionsByLobId = async (Id, BrokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl +
      `/api/QuoteCompare/GetLobQuestionsByLobId?id=${Id}&brokerId=${BrokerId}`,
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

export const GetLobsByBrokerId = async (BrokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + `/api/QuoteCompare/GetLobsByBrokerId?brokerId=${BrokerId}`,
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

export const GetChecklistQuestionsByBrokerId = async (BrokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl +
      `/api/QuoteCompare/GetChecklistQuestionsByBrokerId?brokerId=${BrokerId}`,
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

export const GetCoverageSpecificationsByBrokerId = async (BrokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl +
      `/api/QuoteCompare/GetCoverageSpecificationsByBrokerId?brokerId=${BrokerId}`,
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