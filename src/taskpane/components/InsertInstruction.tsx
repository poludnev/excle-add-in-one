import * as React from "react";
import { createSheetWithName } from "../../utilities/utils";
import { fillInstructionData, fillInstuctionHeading } from "../../commands/instruction";
// import { createSheetWithName, fillSummaryHeading } from "../../commands/summary";
export const InsertInstruction = () => {
  const addSummarySheetHandler = async () => {
    const {} = await createSheetWithName("instruction");
  };

  const addSummaryHeadings = async () => {
    // fillSummaryHeading();
    fillInstuctionHeading();
  };

  const fillInstrustionDataHandler = async () => {
    fillInstructionData();
  };
  return (
    <div>
      InsrtSummary
      <div>
        <button onClick={addSummarySheetHandler}>Add instuction sheet and header</button>
      </div>
      <div>
        <button onClick={addSummaryHeadings}>Add Instruction header</button>
      </div>
      <div>
        <button onClick={fillInstrustionDataHandler}>Fill Instruction data</button>
      </div>
    </div>
  );
};
