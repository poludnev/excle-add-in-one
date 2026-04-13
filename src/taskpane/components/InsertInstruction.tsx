import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import { createSheetWithName } from "../../utilities/utils";
import { fillInstructionData, fillInstuctionHeading } from "../../commands/instruction";
// import { createSheetWithName, fillSummaryHeading } from "../../commands/summary";

const useStyles = makeStyles({
  data: {
    backgroundColor: "#ffe8ca",
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
});
export const InsertInstruction = () => {
  const [error, setError] = React.useState<string | null>(null);
  const [isBlockingButtons, setIsBlockingButtons] = React.useState(false);

  const styles = useStyles();

  const addInstructionSheetHandler = async () => {
    try {
      setIsBlockingButtons(true);
      const { worksheet, context } = await createSheetWithName("instruction");
      console.log("addInstructionSheetHandler 2");
      if (worksheet) {
        worksheet.activate();
        await context.sync();
        await fillInstuctionHeading();
        await fillInstructionData();
      }
    } catch (error) {
      console.error("Error in addInstructionSheetHandler:", error);
      setError(error instanceof Error ? error.message : "Unknown error");
    } finally {
      setIsBlockingButtons(false);
    }
  };

  const addSummaryHeadings = async () => {
    // fillSummaryHeading();
    fillInstuctionHeading();
  };

  const fillInstrustionDataHandler = async () => {
    fillInstructionData();
  };
  return (
    <div className={styles.data}>
      <h3 className={styles.title}>Add Instruction sheet</h3>
      {/* <details>
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
      </details> */}
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
          onClick={addInstructionSheetHandler}
          disabled={isBlockingButtons}
          className={styles.addButton}
        >
          Add instruction sheet
        </button>
      </div>
      {/* <div>
        <button onClick={addSummaryHeadings}>Add Instruction header</button>
      </div>
      <div>
        <button onClick={fillInstrustionDataHandler}>Fill Instruction data</button>
      </div> */}
    </div>
  );
};
