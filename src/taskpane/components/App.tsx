import React, { ReactElement, useEffect, useState } from "react";
import { DefaultButton } from "@fluentui/react";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Office, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App = ({ title, isOfficeInitialized }: AppProps): ReactElement => {
  const [listItems, setListItems] = useState<HeroListItem[]>([]);

  useEffect(() => {
    setListItems([
      {
        url: "https://cdn.glitch.global/04ac2eab-7093-47ad-976f-739938dcbb74/pulp-fiction-john-travolta.gif?v=1667742225377",
      },
      {
        url: "https://uploads-ssl.webflow.com/63d9004a7de2b71ce6d5f83b/63eca414da25f1e270f5deb8_6213b19fae7ebd8f1962e96c_2%255B1%255D.png",
      },
      {
        url: "https://media.tenor.com/HGPmx7TvfYAAAAAd/drawify-man.gif",
      },
      {
        url: "https://cdn.glitch.global/51637606-60d9-484c-a941-c3ad0567928a/7a299d06-06b7-4831-97ab-2fc34607fa81.png?v=1686167167966",
      },
    ]);
  }, []);

  const click = async () => {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      "Hello World!",
      {
        coercionType: Office.CoercionType.Text,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      {/* <Header /> */}
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
          Run
        </DefaultButton>
        {/* <DefaultButton
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={addImageToSlide}
        >
          Runss
        </DefaultButton> */}
      </HeroList>
    </div>
  );
};

export default App;
