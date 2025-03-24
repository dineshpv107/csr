
export const column2 = [
        {
          name: "S.No",
          selector: (row) =>  row.index,
          sortable: true,
          width:'10%'
        },
        {
          name: "Job Id",
          cell: (row) => row?.JobId,
          width:'22.5%',
          sortable: true,
        },
        {
          name: "Message Logged On",
          selector: (row) => {
            if (row.CreatedOn != null && row.CreatedOn != undefined) {
              const createdDate = new Date(row.CreatedOn).toString();
              return createdDate;
            } else {
              return "";
            }
          },
          sortable: true,
        },
        {
          name: "Message",
          cell: (row) =>{ if(row.Message==null){
              return row?.ProcessName;
          }else{ return (
              <span title={row.Message}>{row.Message.length > 70 ? `${row.Message.substring(0, 60)}...` : row.Message}</span>
          )
          }
          },
          sortable: true,
          width:'40%',
        },
      ];
