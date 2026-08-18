import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

// App.jsx was built for Claude's artifact sandbox, which provides a `window.storage`
// key-value API for persistence. That API doesn't exist in a normal deployed site, so this
// shim backs it with localStorage instead. The interface (get/set/delete/list, all async,
// same shape of return values) matches exactly, so App.jsx needs no changes to work here.
if (!window.storage) {
  window.storage = {
    async get(key) {
      const value = window.localStorage.getItem(key);
      return value !== null ? { key, value, shared: false } : null;
    },
    async set(key, value) {
      window.localStorage.setItem(key, value);
      return { key, value, shared: false };
    },
    async delete(key) {
      window.localStorage.removeItem(key);
      return { key, deleted: true, shared: false };
    },
    async list(prefix) {
      const keys = Object.keys(window.localStorage).filter(k => !prefix || k.startsWith(prefix));
      return { keys, prefix, shared: false };
    }
  };
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
