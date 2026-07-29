import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import IOSInstallBanner from "./IOSInstallBanner.jsx";
import { supabaseStorageAdapter } from "./lib/supabaseStorageAdapter";

// O App.jsx foi escrito originalmente usando "window.storage" (o
// armazenamento chave-valor do ambiente de desenvolvimento). Ao atribuir
// o adaptador do Supabase nesse mesmo lugar, todo o app passa a gravar
// de verdade em um banco de dados na nuvem, sem precisar alterar
// nenhuma linha da lógica de negócio do App.jsx.
window.storage = supabaseStorageAdapter;

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
    <IOSInstallBanner />
  </React.StrictMode>
);
