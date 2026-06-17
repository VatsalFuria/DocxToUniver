import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

// Global reset — Univer manages its own styling inside its container
import "./index.css";

ReactDOM.createRoot(document.getElementById("root")!).render(
  // NOTE: StrictMode renders components twice in dev, which would create
  // two Univer instances. Disabled here to avoid that.
  <App />
);
