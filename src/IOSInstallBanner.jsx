import { useEffect, useState } from "react";

const DISMISS_KEY = "geoacervo_ios_banner_dismissed";

function isIOS() {
  const ua = window.navigator.userAgent;
  const iOSDevice = /iPad|iPhone|iPod/.test(ua);
  // iPadOS 13+ se identifica como "Macintosh", mas tem suporte a toque
  const iPadOS13 = navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1;
  return iOSDevice || iPadOS13;
}

function isStandalone() {
  return window.navigator.standalone === true || window.matchMedia("(display-mode: standalone)").matches;
}

export default function IOSInstallBanner() {
  const [show, setShow] = useState(false);

  useEffect(() => {
    const dismissed = localStorage.getItem(DISMISS_KEY) === "1";
    if (isIOS() && !isStandalone() && !dismissed) setShow(true);
  }, []);

  const dismiss = () => {
    localStorage.setItem(DISMISS_KEY, "1");
    setShow(false);
  };

  if (!show) return null;

  return (
    <div style={styles.wrap}>
      <div style={styles.icon}>⬆️</div>
      <div style={styles.text}>
        <strong style={styles.title}>Instale o GeoAcervo no seu iPhone</strong>
        <span style={styles.desc}>
          Toque em <strong>Compartilhar</strong> (o ícone de seta) na barra do Safari e depois em{" "}
          <strong>"Adicionar à Tela de Início"</strong>.
        </span>
      </div>
      <button onClick={dismiss} style={styles.close} aria-label="Fechar">✕</button>
    </div>
  );
}

const styles = {
  wrap: {
    position: "fixed",
    left: 12,
    right: 12,
    bottom: 12,
    zIndex: 9999,
    background: "#0B3D2E",
    color: "#F7F9F6",
    borderRadius: 14,
    padding: "12px 14px",
    display: "flex",
    alignItems: "flex-start",
    gap: 10,
    boxShadow: "0 12px 30px rgba(0,0,0,0.3)",
    fontFamily: "Inter, sans-serif",
  },
  icon: { fontSize: 20, lineHeight: 1 },
  text: { display: "flex", flexDirection: "column", gap: 3, flex: 1 },
  title: { fontSize: 13.5 },
  desc: { fontSize: 12, color: "#cfe0d6", lineHeight: 1.5 },
  close: { background: "rgba(255,255,255,0.12)", border: "none", color: "#fff", borderRadius: 8, width: 26, height: 26, flexShrink: 0, cursor: "pointer" },
};
