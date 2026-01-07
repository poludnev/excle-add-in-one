import React from "react";
import { createSheetWithName } from "../../commands/data";
import {
  // fillCMR_data_constants,
  fillCMR_data_values,
  isCMRDataSheetExists,
  // fillCMRData,
  // fillCMRTemplate,
  makeCMRs,
} from "../../commands/cmr";
import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  cmr: {
    position: "relative",
    backgroundColor: "#cafff2",
    padding: "10px",
    paddingRight: "20px",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    marginBottom: "20px",
  },
  cmrTitle: {
    margin: "0px",
  },
  cmrButton: {
    padding: "5px",
    display: "block",
    width: "100%",
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
export const InsertCMR = () => {
  const [error, setError] = React.useState<string | null>(null);
  const [showConfirmWillCMRdataDialog, setShowConfirmWillCMRdataDialog] = React.useState(false);

  const [isBlockingButtons, setIsBlockingButtons] = React.useState(false);

  const styles = useStyles();

  const insertSheetHandler = async () => {
    await createSheetWithName("cmr");
    await createSheetWithName("cmr_data");
  };

  // const fillCMRTemplateHandler = async () => {
  //   await fillCMRTemplate();
  // };

  const fillCMRDataSourceHandler = async (ignoreConfirmation: boolean = false) => {
    try {
      setIsBlockingButtons(true);
      if (!ignoreConfirmation) {
        const cmrDataSheetExists = await isCMRDataSheetExists();
        if (cmrDataSheetExists) {
          setShowConfirmWillCMRdataDialog(true);
          return;
        }
      }

      setError(null);
      // await fillCMR_data_constants();
      const result = await fillCMR_data_values();
      if (result.success === false) {
        console.error("Error filling CMR data values:", result.error);
        setError(result.error?.message || "Unknown error");
        return;
      }
    } catch (error) {
      console.error("Error filling CMR data values:", error);
      setError((error as Error).message);
    } finally {
      setIsBlockingButtons(false);
    }
  };

  // const fillCMRDataHandler = async () => {
  //   await fillCMRData();
  // };
  const makeCMRShandler = async () => {
    try {
      setIsBlockingButtons(true);
      setError(null);
      // await makeCMRSheets();
      // try {
      const result = await makeCMRs();
      if (result.success === false) {
        console.error("Error making CMRs in front run:", result);
        setError(result.error.message);
        // throw Error("Error making CMRs in front run");
        return;
      }
    } catch (error) {
      console.error("Error making CMRs: front", error);
      // throw error;
    } finally {
      setIsBlockingButtons(false);
    }
  };

  return (
    <div className={styles.cmr}>
      <dialog open={showConfirmWillCMRdataDialog} className={styles.dialog}>
        <p className={styles.p}>cmr_data sheet already exists, data will be rewritten, confirm.</p>
        <div className={styles.confirm}>
          <button
            className={styles.confirmButton}
            onClick={() => {
              fillCMRDataSourceHandler(true);
              setShowConfirmWillCMRdataDialog(false);
            }}
          >
            Confirm
          </button>
          <button
            className={styles.cancelButton}
            onClick={() => setShowConfirmWillCMRdataDialog(false)}
          >
            Cancel
          </button>
        </div>
      </dialog>
      <h3 className={styles.cmrTitle}>Add CMR </h3>
      <details>
        <summary>Instructions</summary>
        <ol>
          <li>a sheet named "instructions" should exist.</li>
          <li>
            the "instructions" sheet should contain cmr number (column W) and invoice date (column
            X).
          </li>
          <li>"DEFAULT_" fields in the "cmr_data" sheet should be filled as needed.</li>
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
      {/* <div>
        <button onClick={insertSheetHandler}>Add CMR sheet</button>
      </div> */}
      {/* <div>
        <button onClick={fillCMRTemplateHandler}>Fill CMR template</button>
      </div> */}
      {/* <div>
        <button disabled onClick={fillCMRDataHandler}>
          Fill CMR data
        </button>
      </div> */}
      <div>
        <button
          disabled={isBlockingButtons}
          className={styles.cmrButton}
          onClick={() => fillCMRDataSourceHandler()}
        >
          Fill CMR data source
        </button>
      </div>
      <div>
        <button disabled={isBlockingButtons} className={styles.cmrButton} onClick={makeCMRShandler}>
          Make CMRs
        </button>
      </div>
    </div>
  );
};
