import React, { useEffect, useState, useRef } from "react";

const QuoteCompareSheet = (data) => {
  const container = useRef();
  const luckysheet = window.luckysheet;

  const luckyCss = {
    margin: "0px",
    padding: "0px",
    position: "absolute",
    width: "100% !important",
    height: "50%",
    left: "0px",
    top: "0px",
  };

  const handleVisibilityChange = () => {
    if (luckysheet) {
      luckysheet.refresh();
      luckysheet.exitEditMode();
    }
  };

  useEffect(() => {
    document.addEventListener(
      "visibilitychange",
      handleVisibilityChange,
      false
    );
  });

  return (
    <div>
      <p>Sheet will render here....</p>
      <div className="csrSheet" id="luckysheet2" ref={luckyCss} />
    </div>
  );
};
export default QuoteCompareSheet;

export const Policy_appDataConfig = {
  demo: {
    name: "PolicyReviewChecklist",
    color: "",
    config: {
      merge: {
        "0_1": {
          rs: 1,
          cs: 6,
          r: 0,
          c: 1,
        },
        "1_1": {
          rs: 1,
          cs: 2,
          r: 0,
          c: 1,
        },
      },
      borderInfo: [],
      rowlen: {
        0: 20,
        1: 20,
        2: 20,
        3: 35,
        4: 50,
        5: 35,
        6: 35,
        7: 35,
        8: 35,
        9: 35,
        10: 50,
        11: 50,
        12: 60,
        13: 20,
        14: 20,
        15: 20,
        16: 20,
        17: 31,
      },
      columnlen: {
        0: 63,
        1: 280,
        2: 250,
        3: 250,
        4: 250,
        5: 250,
        6: 250,
        7: 250,
        8: 250,
        9: 250,
        10: 250,
        11: 250,
        12: 250,
        13: 250,
        14: 250,
      },
      curentsheetView: "viewPage",
      sheetViewZoom: {
        viewNormalZoomScale: 0.6,
        viewPageZoomScale: 0.6,
      },
    },

    chart: [],
    status: "1",
    order: "0",
    hide: 0,
    column: 20,
    celldata: [],
    ch_width: 2322,
    rh_height: 949,
    scrollLeft: 0,
    scrollTop: 0,
    luckysheet_select_save: [],
    calcChain: [],
    luckysheet_alternateformat_save: [],
    luckysheet_alternateformat_save_modelCustom: [],
    sheets: [],
  },
};
