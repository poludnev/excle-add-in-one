import React from "react";
import { createSheetWithName } from "../../utilities/utils";
import { fillRazbivkaData, fillRazbivkaHeading } from "../../commands/razbivka";

export const InsertRazbivka = () => {
  const addRazbvkaSheetHandler = async () => {
    const {} = await createSheetWithName("razbivka");
  };

  const addRazbivkaHeader = () => {
    fillRazbivkaHeading();
  };

  const fillRazbivkaSheet = () => {
    fillRazbivkaData();
  };
  return (
    <div>
      InsrtSummary
      <div>
        <button onClick={addRazbvkaSheetHandler}>Add Razbivka sheet and header</button>
      </div>
      <div>
        <button onClick={addRazbivkaHeader}>Add Razbivka header</button>
      </div>
      <div>
        <button onClick={fillRazbivkaSheet}>Fill Razbivka data</button>
      </div>
    </div>
  );
};
