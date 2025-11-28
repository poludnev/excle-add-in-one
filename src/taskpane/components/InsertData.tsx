import * as React from "react";
import { createSheetWithName, fillSummaryHeading, updateDataFormats } from "../../commands/data";
export const InsertData = () => {
  const addSummarySheetHandler = async () => {
    const {} = await createSheetWithName("data");
  };

  const addSummaryHeadings = async () => {
    fillSummaryHeading();
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
        <button onClick={addSummaryHeadings}>Add data header</button>
      </div>
      <div>
        <button onClick={updateDataFormatsHandler}>Update data formats</button>
      </div>
    </div>
  );
};
