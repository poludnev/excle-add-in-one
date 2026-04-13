import * as React from "react";
import { createSheetWithName, fillDataHeading, updateDataFormats } from "../../commands/data";
export const InsertDataDev = () => {
  const addSummarySheetHandler = async () => {
    const {} = await createSheetWithName("data");
  };

  const addDataHeadings = async () => {
    fillDataHeading();
  };
  const updateDataFormatsHandler = async () => {
    updateDataFormats();
  };
  return (
    <div>
      InsrtData
      <div>
        <button onClick={addSummarySheetHandler}>Add data sheet and header</button>
      </div>
      <div>
        <button onClick={addDataHeadings}>Add data header</button>
      </div>
      <div>
        <button onClick={updateDataFormatsHandler}>Update data formats</button>
      </div>
    </div>
  );
};
