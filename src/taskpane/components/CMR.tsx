import React from "react";
import { createSheetWithName } from "../../commands/data";
import { fillCMRTemplate } from "../../commands/cmr";
export const InsertCMR = () => {
  const insertSheetHandler = async () => {
    await createSheetWithName("cmr");
  };

  const fillCMRTemplateHandler = async () => {
    await fillCMRTemplate();
  };

  return (
    <div>
      Add CMR
      <div>
        <button onClick={insertSheetHandler}>Add CMR sheet</button>
      </div>
      <div>
        <button onClick={fillCMRTemplateHandler}>Fill CMR template</button>
      </div>
      <div>
        <button>Fill CMR data</button>
      </div>
    </div>
  );
};
