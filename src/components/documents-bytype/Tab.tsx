import { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import useDocuments from "../../hooks/useDocuments";
import { useTeams } from "@microsoft/teamsfx-react";
import * as microsoftTeams from "@microsoft/teams-js";

export default function DocumentsByType() {
  const LOG_SOURCE = "documents-bytype";
  const { themeString } = useContext(TeamsFxContext);
  const { documentsByType, isError } = useDocuments();

  const [{ context }] = useTeams();
  const [docExt, setDocExt] = useState<string>(".docx");

  useEffect(() => {
    try {
      if (context) {
        setDocExt(context.page.id?.replace("DocumentByTypePage", "") ?? ".docx");
      }
      microsoftTeams.app.notifySuccess();
    } catch (err) {
      console.error(`${LOG_SOURCE} (useEffect) - ${err}`);
    }
  }, [context]);

  if (!isError) {
    if (documentsByType == null || documentsByType[docExt] == null || documentsByType[docExt]?.length < 1) {
      return <p>No documents found.</p>
    } else {
      return (
        <div className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}>
          <h2>Documents with extension '{docExt}'</h2>
          <table width="100%">
            <thead>
              <tr>
                <td><strong>Title</strong></td>
                <td><strong>Name</strong></td>
                <td><strong>Size (KB)</strong></td>
              </tr>
            </thead>
            <tbody>
              {documentsByType[docExt]?.map((doc, idx) => {
                return (
                  <tr key={idx}>
                    <td>{doc.Title}</td>
                    <td>{doc.Name}</td>
                    <td>{(doc.Size / 1024).toFixed(2)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      );
    }
  } else {
    return (
      <div>Error loading!</div>
    )
  }
}
