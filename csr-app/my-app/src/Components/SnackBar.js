// SimpleSnackbar.js
import * as React from "react";
import Button from "@mui/material/Button";
import Snackbar from "@mui/material/Snackbar";
import Alert from "@mui/material/Alert";
import { FontIcon } from "@fluentui/react/lib/Icon";

export const SimpleSnackbar = React.forwardRef((props, ref) => {
  const [open, setOpen] = React.useState(false);
  const [messageToBeShown, setMessageToBeShown] = React.useState("");
  const [messageType, setMessageType] = React.useState(""); //"success","error","info","warning"

  // React.useImperativeHandle( ref, () => ( {
  const showSnackbar = (message, msgType, msgStatus) => {
    setMessageToBeShown(message);
    setMessageType(msgType);
    setOpen(msgStatus);

    if (msgStatus) {
      setTimeout(() => {
        hideSnackbar();
      }, 1000); // Hide snackbar after 1 second
    }
  };
  // } ) );

  const handleClose = (event, reason) => {
    if (reason === "clickaway") {
      return;
    }
    setOpen(false);
  };

  const hideSnackbar = () => {
    setOpen(false);
  };

  React.useImperativeHandle(ref, () => ({
    showSnackbar,
    hideSnackbar,
  }));

  const action = (
    <React.Fragment>
      <FontIcon
        aria-label="Compass"
        style={{ cursor: "pointer" }}
        iconName="Cancel"
        onClick={(e) => {
          handleClose(e, "");
        }}
      />
    </React.Fragment>
  );

  return (
    <Snackbar
      anchorOrigin={{ vertical: "top", horizontal: "center" }}
      open={open}
      onClose={handleClose}
      action={action}
    >
      <Alert
        onClose={handleClose}
        severity={messageType}
        variant="filled"
        sx={{ width: "100%" }}
      >
        {messageToBeShown}
      </Alert>
    </Snackbar>
  );
});

export const SimpleSnackbarWithOutTimeOut = React.forwardRef((props, ref) => {
  const [open, setOpen] = React.useState(false);
  const [messageToBeShown, setMessageToBeShown] = React.useState("");
  const [messageType, setMessageType] = React.useState(""); //"success","error","info","warning"

  // React.useImperativeHandle( ref, () => ( {
  const showSnackbar = (message, msgType, msgStatus) => {
    setMessageToBeShown(message);
    setMessageType(msgType);
    setOpen(msgStatus);
  };
  // } ) );

  const handleClose = (event, reason) => {
    if (reason === "clickaway") {
      return;
    }
    setOpen(false);
  };

  const hideSnackbar = () => {
    setOpen(false);
  };

  React.useImperativeHandle(ref, () => ({
    showSnackbar,
    hideSnackbar,
  }));

  const action = (
    <React.Fragment>
      <FontIcon
        aria-label="Compass"
        style={{ cursor: "pointer" }}
        iconName="Cancel"
        onClick={(e) => {
          handleClose(e, "");
        }}
      />
    </React.Fragment>
  );

  return (
    <Snackbar
      anchorOrigin={{ vertical: "top", horizontal: "center" }}
      open={open}
      onClose={handleClose}
      action={action}
    >
      <Alert
        onClose={handleClose}
        severity={messageType}
        variant="filled"
        sx={{ width: "100%" }}
      >
        {messageToBeShown}
      </Alert>
    </Snackbar>
  );
});
