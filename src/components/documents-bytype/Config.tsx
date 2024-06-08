import { useContext, useEffect, useRef, useState } from "react";
import { TeamsFxContext } from "../Context";
import { useTeams } from "@microsoft/teamsfx-react";
import * as microsoftTeams from "@microsoft/teams-js";
import useDocuments from "../../hooks/useDocuments";
import { Dropdown, Label, Option, OptionOnSelectData, SelectionEvents } from "@fluentui/react-components";

export default function Config() {
  const LOG_SOURCE = "documents-bytype-config";
  const { themeString } = useContext(TeamsFxContext);

  const [{ context }] = useTeams();
  const { documentsByType, isError } = useDocuments();
  const [docExt, setDocExt] = useState<string>("");
  const [docExtOptions, setDocExtOptions] = useState<string[]>([]);
  const entityId = useRef("");
  const saveDocExt = useRef("");

  const _savePage = async (saveEvent: microsoftTeams.pages.config.SaveEvent): Promise<void> => {
    try {
      //docExt always is "", docExtOptions is always []
      const baseUrl = `https://${window.location.hostname}:${window.location.port}/index.html#`;
      await microsoftTeams.pages.config.setConfig({
        suggestedDisplayName: `Documents (${saveDocExt.current})`,
        entityId: entityId.current,
        contentUrl: `${baseUrl}/documents-bytype`,
        websiteUrl: `${baseUrl}/documents-bytype`,
        removeUrl: `${baseUrl}/documents-bytype-remove`
      });
      saveEvent.notifySuccess();
    } catch (err) {
      console.error(`${LOG_SOURCE} (_savePage) - ${err}`);
    }
  }

  const _optionSelected = (event: SelectionEvents, data: OptionOnSelectData) => {
    try {
      const docExt = (data.optionValue) ? data.optionValue.toString() : ".docx";
      setDocExt(docExt);
      saveDocExt.current = docExt;
      entityId.current = `${docExt}DocumentByTypePage`;
      
      microsoftTeams.pages.config.setValidityState(docExt.length > 0);
    } catch (err) {
      console.error(`${LOG_SOURCE} (  const _optionSelected = (event: SelectionEvents, data: OptionOnSelectData) => {
      ) - ${err}`);
    }
  }

  useEffect(() => {
    if (context) {
      (async () => {
        try {
          const currentConfig = await microsoftTeams.pages.getConfig();
          setDocExt(currentConfig.entityId?.replace("DocumentByTypePage", "") ?? "");
          
          entityId.current = currentConfig.entityId as string;

          microsoftTeams.pages.config.registerOnSaveHandler(_savePage);
          microsoftTeams.pages.config.setValidityState(docExt.length > 0);

          microsoftTeams.app.notifySuccess();
        } catch (err) {
          console.error(`${LOG_SOURCE} (useEffect) - ${err}`);
        }
      })();
    }
  }, [context]);

  useEffect(() => {
    if (documentsByType) {
      const docExtOptions = Object.keys(documentsByType);
      setDocExtOptions(docExtOptions);
    }
  }, [documentsByType]);

  if (!isError) {
    return <div className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}>
      <Label htmlFor="">Select the Document Type:</Label>
      <div>
        <Dropdown
          value={docExt}
          selectedOptions={docExtOptions}
          onOptionSelect={_optionSelected}>
          {docExtOptions != null && docExtOptions?.map((o) => {
            return (<Option key={o} value={o}>{o}</Option>);
          })}
        </Dropdown>
      </div>
      <p>You selected {docExt}</p>
    </div>;
  } else {
    return <div>Error loading Config screen.</div>;
  }
}