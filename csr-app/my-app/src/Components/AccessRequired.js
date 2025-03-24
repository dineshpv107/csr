import React, { useEffect } from 'react';

const AccessRequired = () => {

    return (
        <div className="unbody">
            <div className="uacontainer">
                <h1 className="unh1">Access Denied</h1>
                <p className="unp">Sorry, you do not have enough role access to access this application. Please ask the admin team to provide the necessary role access.</p>
            </div>
        </div>
    );
}

export default AccessRequired;