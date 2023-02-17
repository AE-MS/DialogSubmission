import { useContext, useState } from "react";
import { Image, Menu } from "@fluentui/react-northstar";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { app, dialog, FrameContexts } from "@microsoft/teams-js";
import { AzureFunctions } from "./AzureFunctions";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useData } from "@microsoft/teamsfx-react";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { TeamsFxContext } from "../Context";

function openUrlDialog() {
  dialog.url.open({
    url: window.location.href,
    title: "Dialog Bug Test",
    size: { height: 600, width: 600 },
    },
    (result: dialog.ISdkResponse) => {
      const dialogResultElement = document.getElementById("submissionAcknowledgement")!;
      // document.getElementById("submissionAcknowledgement")!.style.display = "block";
      dialogResultElement.innerText = `Url Dialog submission occurred, result = ${result.result} err = ${result.err}`;
    }
  );
}

function submitUrlDialog() {
  dialog.url.submit("Everything is awesome");
}

export function Welcome(props: { showFunction?: boolean; environment?: string }) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const steps = ["local", "azure", "publish"];
  const friendlyStepsName: { [key: string]: string } = {
    local: "1. Build your app locally",
    azure: "2. Provision and Deploy to the Cloud",
    publish: "3. Publish to Teams",
  };
  const [selectedMenuItem, setSelectedMenuItem] = useState("local");
  const items = steps.map((step) => {
    return {
      key: step,
      content: friendlyStepsName[step] || "",
      onClick: () => setSelectedMenuItem(step),
    };
  });

  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  const userName = (loading || error) ? "": data!.displayName;
  const context = useData(async () => {
    await app.initialize();
    const context = await app.getContext();    
    return context;
  })?.data;

  const hubName: string | undefined = context?.app.host.name;
  const frameContext: FrameContexts | undefined = context?.page.frameContext;

    return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ", " + userName : ""}!</h1>
        {hubName && (
          <p className="center">Your app is running in {hubName}</p>
        )}
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>
        <p className="center">The frame context is "{frameContext}"</p>
        <p className="center">The current URL is {window.location.href}</p>
        <p id="submissionAcknowledgement" className="center"></p>
        { frameContext === FrameContexts.content && (
          <button onClick={openUrlDialog}>Open URL Dialog</button>
        )}
        { frameContext === FrameContexts.task && (
          <button onClick={submitUrlDialog}>Submit Dialog</button>
        )}

        <Menu defaultActiveIndex={0} items={items} underlined secondary />
        <div className="sections">
          {selectedMenuItem === "local" && (
            <div>
              <EditCode showFunction={showFunction} />
              <CurrentUser userName={userName} />
              <Graph />
              {showFunction && <AzureFunctions />}
            </div>
          )}
          {selectedMenuItem === "azure" && (
            <div>
              <Deploy />
            </div>
          )}
          {selectedMenuItem === "publish" && (
            <div>
              <Publish />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
