import React from "react";
import { createSheetWithName } from "../../commands/data";
import { fillSummaryDefaultValues, insertSummaryHeaders } from "../../commands/summary";
export const InsertSummaryDev = () => {
  const insertSheetHandler = async () => {
    await createSheetWithName("summary");
  };
  const insertHeadingsHandler = async () => {
    console.log("handler");
    await insertSummaryHeaders();
    await fillSummaryDefaultValues();
  };
  return (
    <div>
      Insert Summary sheet Dev
      <div>
        <button onClick={insertSheetHandler}>Insert Summary Sheet</button>
      </div>
      <div>
        <button onClick={insertHeadingsHandler}>Insert Summary HEADINGS</button>
      </div>
    </div>
  );
};
