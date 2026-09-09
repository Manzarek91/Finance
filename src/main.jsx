import "./storageShim.js";
import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import AuthGate from "./AuthGate.jsx";
import FinanceTracker from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <AuthGate>
      <FinanceTracker />
    </AuthGate>
  </React.StrictMode>
);
