import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import App from "./App";

window.onload = async () => {
  if (typeof Office !== "undefined") {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        Office.actions.associate(
          "showTPWorkbook",
          async (event: Office.AddinCommands.Event) => {
            await Office.addin.showAsTaskpane();
            event.completed();
          }
        );
      }
    });
  }
  const root = ReactDOM.createRoot(
    document.getElementById("root") as HTMLElement
  );
  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
};
