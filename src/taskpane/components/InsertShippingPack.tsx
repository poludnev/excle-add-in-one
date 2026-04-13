import React from "react";
import { createSheetWithName } from "../../utilities/utils";
import { makeStyles } from "@fluentui/react-components";
import { fillRazbivkaData, fillRazbivkaHeading } from "../../commands/razbivka";
import { fillPackingData, fillPackingHeading } from "../../commands/packing";
import { fillTansitData } from "../../commands/transit";

const useStyles = makeStyles({
  data: {
    backgroundColor: "#d0ffca",
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

export const InsertShippingPack = () => {
  const [error, setError] = React.useState<string | null>(null);
  const [isBlockingButtons, setIsBlockingButtons] = React.useState(false);
  const styles = useStyles();

  const addRazbvkaSheetHandler = async () => {
    const {} = await createSheetWithName("razbivka");
    await fillRazbivkaHeading();
    await fillRazbivkaData();
    const {} = await createSheetWithName("packing");
    await fillPackingHeading();
    await fillPackingData();
    const { worksheet, context } = await createSheetWithName("transit");
    await fillTansitData();
    worksheet.activate();
    await context.sync();
  };

  // const addRazbivkaHeader = () => {
  //   fillRazbivkaHeading();
  // };

  // const fillRazbivkaSheet = () => {
  //   fillRazbivkaData();
  // };
  return (
    <div className={styles.data}>
      <h3 className={styles.title}>Add Shipping Docs</h3>
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
          onClick={addRazbvkaSheetHandler}
          disabled={isBlockingButtons}
          className={styles.addButton}
        >
          Add Razbivka, Packing and Transit sheets
        </button>
      </div>
      {/* <div>
        <button onClick={addRazbivkaHeader}>Add Razbivka header</button>
      </div>
      <div>
        <button onClick={fillRazbivkaSheet}>Fill Razbivka data</button>
      </div> */}
    </div>
  );
};
