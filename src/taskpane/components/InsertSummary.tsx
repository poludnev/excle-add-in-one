import React from "react";
import { makeStyles } from "@fluentui/react-components";

import { createSheetWithName } from "../../commands/data";
import { fillSummaryDefaultValues, insertSummaryHeaders } from "../../commands/summary";
import { isSheetByNameExists } from "../../utilities/utils";

const useStyles = makeStyles({
  data: {
    backgroundColor: "#fffbca",
    padding: "10px",
    paddingRight: "20px",
    display: "flex",
    flexDirection: "column",
    gap: "5px",
  },
  title: {
    // backgroundColor: "#4CAF50",
    margin: "0px",
    // color: "white",
    // padding: "10px",
  },
  buttons: {
    // paddingTop: "10px",
    display: "flex",
    gap: "10px",
    // marginBottom: "10px",
  },
  addButton: {
    flexGrow: "1",
    display: "block",
    padding: "5px",
  },
  updateButton: {
    flexGrow: "1",
    display: "block",
    padding: "5px",
  },
  errorSection: {
    color: "red",
    position: "fixed",
    bottom: "0",
    left: "0",
    // transform: "translateY(100%)",
    padding: "10px",
    backgroundColor: "#ffe5e5",
    width: "100%",
    boxSizing: "border-box",
  },
  closeButton: {
    position: "absolute",
    top: "5px",
    right: "2px",
    background: "none",
    border: "none",
    fontSize: "16px",
    cursor: "pointer",
  },

  dialog: {
    position: "absolute",
    bottom: "0",
    left: "0",
    backgroundColor: "#fff3cd",
    width: "100%",
    // height: "80px",
    boxSizing: "border-box",
    border: "none",
    padding: "10px",
  },
  p: {
    margin: "0px",
    padding: "0px",
  },
  confirm: {
    display: "flex",
    justifyContent: "space-between",
    gap: "10px",
    // marginTop: "10px",
    paddingTop: "10px",
  },
  confirmButton: {
    backgroundColor: "#d42000",
    color: "white",
    border: "none",
    padding: "10px",
    cursor: "pointer",
    display: "block",
    width: "100%",
  },
  cancelButton: {
    backgroundColor: "#6c757d",
    color: "white",
    border: "none",
    padding: "10px",
    cursor: "pointer",
    display: "block",
    width: "100%",
  },
});

export const InsertSummary = () => {
  const [error, setError] = React.useState<string | null>(null);
  const [isBlockingButtons, setIsBlockingButtons] = React.useState(false);

  const [showConfirmFillSummaryDialog, setShowConfirmFillSummaryDialog] = React.useState(false);
  const styles = useStyles();

  const insertSheetHandler = async () => {
    try {
      setError(null);
      const dataSheetExists = await isSheetByNameExists("data");
      if (!dataSheetExists) {
        setError('Please create a sheet named "data" before adding the summary sheet.');
        return;
      }
      const summarySheetExists = await isSheetByNameExists("summary");
      // if (summarySheetExists) {
      //   setShowConfirmFillSummaryDialog(true);
      //   setError(
      //     'A sheet named "summary" already exists. Please delete or rename it before adding a new one.'
      //   );
      //   // return;
      // }
      await createSheetWithName("summary");
      await insertSummaryHeaders();
      await fillSummaryDefaultValues();
    } catch (error) {
      setError("Failed to create summary sheet.");
      console.error("Error creating summary sheet:", error);
    }
  };
  const insertHeadingsHandler = async () => {
    console.log("handler");
    await insertSummaryHeaders();
    await fillSummaryDefaultValues();
  };
  return (
    <div className={styles.data}>
      <h3 className={styles.title}>Add Summary sheet</h3>
      <details>
        <summary>Instructions</summary>
        <ol>
          <li>A new sheet named "data" will be created.</li>
          <li>
            If a sheet named "data" already exists, it will be activated instead of creating a new
            one.
          </li>
          <li>
            Rows width and sum formulas are needed to be updated, use "Update formats" button.
          </li>
        </ol>
      </details>
      {error && (
        <div className={styles.errorSection}>
          <button className={styles.closeButton} onClick={() => setError(null)}>
            x
          </button>
          <div style={{ color: "red" }}>CMR Error: {error}</div>
        </div>
      )}

      <div className={styles.buttons}>
        <button
          onClick={insertSheetHandler}
          disabled={isBlockingButtons}
          className={styles.addButton}
        >
          Insert Summary Sheet
        </button>
      </div>
      {/* <div>
        <button onClick={insertHeadingsHandler}>Insert Summary HEADINGS</button>
      </div> */}
    </div>
  );
};
