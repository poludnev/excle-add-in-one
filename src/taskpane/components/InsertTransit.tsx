import React from "react";
import { createSheetWithName } from "../../utilities/utils";
import { fillTansitData } from "../../commands/transit";
export const InsertTransit = () => {
  const addTranzitSheetHandler = () => {
    createSheetWithName("transit");
  };

  // const addTransitkaHeader = () => {};

  const fillTransitkaSheet = () => {
    fillTansitData();
  };

  return (
    <div>
      InsrtSummary
      <div>
        <button onClick={addTranzitSheetHandler}>Add Razbivka sheet and header</button>
      </div>
      {/* <div>
        <button onClick={addTransitkaHeader}>Add Transitka header</button>
      </div> */}
      <div>
        <button onClick={fillTransitkaSheet}>Fill Transitka data</button>
      </div>
    </div>
  );
};
