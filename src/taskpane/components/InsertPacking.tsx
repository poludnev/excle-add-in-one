import React from "react";
import { createSheetWithName } from "../../utilities/utils";
import { fillPackingData, fillPackingHeading } from "../../commands/packing";

export const InsertPacking = () => {
  const addPackingSheetHandler = async () => {
    const {} = await createSheetWithName("packing");
  };

  const addPackingHeader = () => {
    fillPackingHeading();
  };

  const fillPackingDataHandler = () => {
    try {
      fillPackingData();
    } catch (e) {
      console.error(e);
    }
  };

  return (
    <div>
      Inset Packing
      <div>
        <button onClick={addPackingSheetHandler}>Add PAcking sheet</button>
      </div>
      <div>
        <button onClick={addPackingHeader}>Add PAcking sheet header</button>
      </div>
      <div>
        <button onClick={fillPackingDataHandler}>Add PAcking sheet data</button>
      </div>
    </div>
  );
};
