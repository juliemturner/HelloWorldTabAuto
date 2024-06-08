import { useContext } from "react";
import { TeamsFxContext } from "./Context";

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  return (
    <div
      className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}
    >
      <p>Loaded</p>
    </div>
  );
}
