import React, { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import AppRoutes from "./AppRoutes";

const AppLayout = () => {
  const navigate = useNavigate();

  const [selectedOption, setSelectedOption] = useState(
    sessionStorage.getItem("selectedOption") || "PolicyCheck"
  );

  useEffect(() => {
    sessionStorage.setItem("selectedOption", selectedOption);
  }, [selectedOption, navigate]);

  useEffect(() => {
    let urlString = "";

        if (window.__TAURI__) {
            // For Tauri app, get the current window label or pathname
            urlString = window.location.pathname; 
        } else {
            // For Web app
            urlString = window.location.href;
            urlString = urlString.replace(/([^:]\/)\/+/g, "$1");
        }

        if (urlString) {
            const url = new URL(window.location.origin + urlString);
            const jobId = url.searchParams.get("jobId") || url.searchParams.get("jobid");
            const urlParamToken = url.searchParams.get("token");

            if (!jobId && !urlParamToken) {
                navigate("/UnAuthorizedUser");
            } else {
                navigate(`/csrView?jobId=${jobId}&token=${urlParamToken}`);
            }
        } else {
            navigate("/UnAuthorizedUser");
        }
  }, []);

  return (
    <div>
      <div className="app-container">
        <div style={{ width: "100%" }}>
          <AppRoutes selectedOption={selectedOption} />
        </div>
      </div>
    </div>
  );
};

export default AppLayout;