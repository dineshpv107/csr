import React from "react";
import { Route, Routes } from "react-router-dom";
import App from "../App";
import { baseData } from "../Services/Constants";
import UnAuthorizedUser from "../Components/UnAuthorizedUser";
import AccessRequired from "../Components/AccessRequired";
import DispatchedJobScreen from "../Components/DispatchedJobScreen";

const AppRoutes = ({ selectedOption }) => {
  return (
    <>
      <Routes>
        {/* {selectedOption === "PolicyCheck" ?
          <Route path="/" element={<Navigate to="/job" />} />
          : <Route path="/" element={<Navigate to="/quoteCompare" />} />
        } */}
        <Route path="/UnAuthorizedUser" element={<UnAuthorizedUser />} />
        <Route path="/xlpage" element={<App baseData={baseData} />} />
        <Route path="/csrView" element={<App baseData={baseData} />} />
        <Route path="/DispatchedJob" element={<DispatchedJobScreen />} />
        <Route path="/AccessDenied" element={<AccessRequired />} />

        {/* {selectedOption === "QuoteCompare" && (
          <>
            <Route path="/quoteCompare" element={<QuoteCompare />} />
            <Route path="/qcJob" element={<QCDataRenderer />} />
            <Route path="/qcAllJobs" element={<GetAllQCJobs />} />
            

          </>
        )} */}

      </Routes>
    </>
  );
};

export default AppRoutes;