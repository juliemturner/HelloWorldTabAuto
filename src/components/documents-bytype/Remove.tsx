import { useContext, useEffect } from "react";
import { TeamsFxContext } from "../Context";
import * as microsoftTeams from "@microsoft/teams-js";

export default function Remove() {
  const LOG_SOURCE = "documents-bytype-remove";
  const { themeString } = useContext(TeamsFxContext);

  useEffect(() => {
    (async () => {
      try {
        microsoftTeams.pages.config.setValidityState(true);
        microsoftTeams.pages.config.registerOnRemoveHandler((removeEvent: microsoftTeams.pages.config.RemoveEvent) => {
          // cleanup logic
          removeEvent.notifySuccess();
        })
      } catch (err) {
        console.error(`${LOG_SOURCE} (useEffect) - ${err}`);
      }
    })();
  }, []);

  return <div className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}>
    <h1>Remove the Document By Type Tab</h1>
    <p>Click the button below to remove the tab.</p>
  </div>;
}