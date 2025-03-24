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
    if (window.location.pathname === '/') {
      navigate("/csrView?jobId=1003csr133142024043332&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJzYW5kZWVwX2t1bWFyQGV4ZGlvbi5jb20iLCJqdGkiOiJiZjJkOTE1OC1kMTgzLTQ2ZDYtODIwMS02MGFlNTViZjExNjAiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiJzYW5kZWVwX2t1bWFyQGV4ZGlvbi5jb20iLCJleHAiOjIwNDQwNjkxMjMsImlzcyI6ImNzclZhbGlkYXRlSXNzdWVyIiwiYXVkIjoiY3NyVmFsaWRhdGVBdWRpZW5jZSJ9.fBOIIZq8HMSoSScTRjuYTXREqYpl8R39LDFxgmSG1xg");
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