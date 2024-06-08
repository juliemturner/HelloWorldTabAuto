import { useContext } from "react";
import { TeamsFxContext } from "./Context";
import useDocuments from "../hooks/useDocuments";

export default function ListDocumentsTab() {
  const { themeString } = useContext(TeamsFxContext);
  const { documents, totalSize, isError } = useDocuments();

  if (!isError) {
    return (
      <div className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}>
        <h2>PnPjs Demo Documents with {documents?.length.toString() ?? "0"} documents.</h2>
        <table width="100%">
          <thead>
            <tr>
              <td><strong>Title</strong></td>
              <td><strong>Name</strong></td>
              <td><strong>Size (KB)</strong></td>
            </tr>
          </thead>
          <tbody>
            {documents?.map((item, idx) => {
              return (
                <tr key={idx}>
                  <td>{item.Title}</td>
                  <td>{item.Name}</td>
                  <td>{(item.Size / 1024).toFixed(2)}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr>
              <td></td>
              <td>
                <strong>Total:</strong>
              </td>
              <td>
                <strong>{(totalSize / 1024).toFixed(2)}</strong>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    );
  } else {
    return (
      <div>Error loading!</div>
    )
  }
}
