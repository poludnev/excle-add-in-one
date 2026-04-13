import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";
import { InsertData } from "./InsertData";
import { InsertDataDev } from "./InsertDataDev";
import { InsertInstruction } from "./InsertInstruction";
import { InsertInstructionDev } from "./InsertInstructionDev";
import { InsertRazbivka } from "./InsertRazbivka";
import { InsertPacking } from "./InsertPacking";
import { InsertTransit } from "./InsertTransit";
import { InsertSummary } from "./InsertSummary";
import { InsertSummaryDev } from "./InsertSummaryDev";
import { InsertCMR } from "./CMR";
import { InsertShippingPack } from "./InsertShippingPack";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    minWidth: "270px",
    backgroundColor: "#f0f0f0",
  },
});

const App: React.FC<AppProps> = () =>
  // props: AppProps
  {
    const styles = useStyles();
    // The list items are static and won't change at runtime,
    // so this should be an ordinary const, not a part of state.
    const listItems: HeroListItem[] = [
      {
        icon: <Ribbon24Regular />,
        primaryText: "Achieve more with Office integration",
      },
      {
        icon: <LockOpen24Regular />,
        primaryText: "Unlock features and functionality",
      },
      {
        icon: <DesignIdeas24Regular />,
        primaryText: "Create and visualize like a pro",
      },
    ];

    // const save = async () => {
    //   Excel.run(async (context: Excel.RequestContext) => {
    //     context.workbook.save();
    //   });
    // };

    return (
      <div className={styles.root}>
        {/* <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" /> */}
        {/* <HeroList message="Discover what this add-in can do for you today!" items={listItems} /> */}
        {/* <TextInsertion insertText={insertText} /> */}
        <InsertData />
        <InsertInstruction />
        <InsertSummary />
        <InsertShippingPack />
        <InsertCMR />

        <details>
          <summary>Dev Section</summary>
          <div>
            <InsertDataDev />
            <InsertSummaryDev />
            <InsertInstructionDev />
            <InsertRazbivka />
            <InsertPacking />
            <InsertTransit />
          </div>
        </details>
      </div>
    );
  };

export default App;
