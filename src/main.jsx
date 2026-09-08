import "./storageShim.js";
import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import FinanceTracker from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <FinanceTracker />
  </React.StrictMode>
);
