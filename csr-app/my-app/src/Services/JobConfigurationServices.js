import axios from "axios";
import { baseUrl } from "./Constants";
import { processAndUpdateToken } from "./CommonFunctions";

export const UpdateFormsComapreConfig = async (JobId, enable) => {
  if (JobId) {
    try {
      const token = sessionStorage.getItem("token");
      const Token = await processAndUpdateToken(token);
      const headers = {
        Authorization: `Bearer ${Token}`,
        "Content-Type": "application/json",
      };
      const response = await axios.put(
        baseUrl +
          `/api/JobConfiguration/UpdateFormsComapreConfig?jobId=${JobId}&enable=${enable}`,
        {},
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
  }
};

export const UpdateLobSplitConfig = async (JobId, enable) => {
  if (JobId) {
    try {
      const token = sessionStorage.getItem("token");
      const Token = await processAndUpdateToken(token);
      const headers = {
        Authorization: `Bearer ${Token}`,
        "Content-Type": "application/json",
      };
      const response = await axios.put(
        baseUrl +
          `/api/JobConfiguration/UpdateLobSplitConfig?jobId=${JobId}&enable=${enable}`,
        {},
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
  }
};

export const GetAvailableLob = async (JobId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + "/api/JobConfiguration/GetAvailableLob?jobId=" + JobId,
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

export const UpdateColumnConfig = async (data) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/JobConfiguration/UpdateColumnConfig",
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

export const AddTableInfoInAppConfig = async (data) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/JobConfiguration/AddTableInfoInAppConfig",
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

export const DeleteFormsTableConfig = async (jobId, tableName) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.delete(
      baseUrl +
        `/api/JobConfiguration/DeleteFormsTableConfig?jobId=${jobId}&tableName=${tableName}`,
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

export const UpdateCsConfigByJobId = async (data) => {
  try {
    let token = sessionStorage.getItem("token");
    token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.post(
      baseUrl + "/api/JobConfiguration/UpdateCsConfigByJobId",
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

export const GetAmsDataByJobId = async (JobId, FromDate, ToDate, BrokerId) => {
  try {
    const token = sessionStorage.getItem("token");
    const Token = await processAndUpdateToken(token);
    const headers = {
      Authorization: `Bearer ${Token}`,
      "Content-Type": "application/json",
    };
    const response = await axios.get(
      baseUrl + `/api/JobConfiguration/GetAmsReviewData?JobId=${JobId}&FromDate=${FromDate}&ToDate=${ToDate}&BrokerId=${BrokerId}`,
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

export const UpdateJobReviewAmsStatus = async (JobId) => {
  if (JobId) {
    try {
      const token = sessionStorage.getItem("token");
      const Token = await processAndUpdateToken(token);
      const headers = {
        Authorization: `Bearer ${Token}`,
        "Content-Type": "application/json",
      };
      const response = await axios.put(
        baseUrl +
          `/api/Master/UpdateJobReviewAmsStatus?user=${sessionStorage.getItem("userName")}&jobId=${JobId}`,
        {},
        {
          headers,
        }
      );
      if (response.status !== 200) {
        return "failed";
      }else{
        return "success";
      }
    } catch (error) {
      return "failed";
    }
  }
};