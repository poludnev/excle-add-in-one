import React from "react";
import { createSheetWithName } from "../../commands/data";
import {
  // fillCMR_data_constants,
  fillCMR_data_values,
  // fillCMRData,
  // fillCMRTemplate,
  makeCMRs,
} from "../../commands/cmr";
export const InsertCMR = () => {
  const [error, setError] = React.useState<string | null>(null);
  const insertSheetHandler = async () => {
    await createSheetWithName("cmr");
    await createSheetWithName("cmr_data");
  };

  // const fillCMRTemplateHandler = async () => {
  //   await fillCMRTemplate();
  // };

  const fillCMRDataSourceHandler = async () => {
    // await fillCMR_data_constants();
    await fillCMR_data_values();
  };

  // const fillCMRDataHandler = async () => {
  //   await fillCMRData();
  // };
  const makeCMRShandler = async () => {
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

    // } catch (error) {
    //   console.error("Error making CMRs: front", error);
    //   // throw error;
    // }
  };

  return (
    <div>
      Add CMR
      {error && (
        <div>
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
        <button onClick={fillCMRDataSourceHandler}>Fill CMR data source</button>
      </div>
      <div>
        <button onClick={makeCMRShandler}>Make CMRs</button>
      </div>
    </div>
  );
};
