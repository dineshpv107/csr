import React, { useEffect } from 'react';

const UnAuthorizedUser = () => {

    return (
        <div className="unbody">
            <div className="uacontainer">
                <h1 className="unh1">Unauthorized Access</h1>
                <p className="unp">Sorry, you are not authorized to access this page.</p>
            </div>
        </div>
    );
}

export default UnAuthorizedUser;