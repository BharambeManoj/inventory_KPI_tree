import React, { useEffect, useRef, useState, useCallback } from "react";

/**
 * InventoryKPIDashboard
 *
 * Standalone, embeddable KPI Tree Dashboard Component.
 * Integrates directly into centralized React Dashboard Portals.
 *
 * Features:
 * - Full-width and full-height responsive layout (no navbar, no left sidebar).
 * - Compact top filter toolbar (Value Streams, Search, Tree/Tables View Switch, Year/Month filters, Actions).
 * - Complete 3-Tier KPI Tree hierarchy with dynamic SVG connector lines.
 * - Interactive Chart.js trend modals & responsible planner breakdown modals.
 * - Isolated CSS execution (prevents Tailwind/global style collisions with the parent portal).
 * - Clean mounting/unmounting with bi-directional postMessage state synchronization.
 *
 * Example Integration:
 * ```jsx
 * import InventoryKPIDashboard from "./dashboards/InventoryKPIDashboard";
 *
 * function DashboardPage() {
 *   return (
 *     <div style={{ width: "100%", height: "calc(100vh - 64px)" }}>
 *       <InventoryKPIDashboard endpoint="http://localhost:5000" />
 *     </div>
 *   );
 * }
 * ```
 *
 * @param {Object} props
 * @param {string} [props.endpoint=""] - URL where the dashboard backend is served (e.g., "http://localhost:5000" or "" for same origin).
 * @param {string} [props.valueStream] - Optional initial Value Stream (e.g., "ESP-ABS", "VDS", "ATC", "NBS", "OSS", "VM-IN").
 * @param {string} [props.view] - Initial view mode ("tree" or "tables").
 * @param {string} [props.className=""] - Optional class names for the outer container.
 * @param {React.CSSProperties} [props.style={}] - Optional inline styles for the outer container.
 * @param {Function} [props.onStateChange] - Callback invoked when dashboard filter or view state changes: `(state) => void`.
 */
export default function InventoryKPIDashboard({
  endpoint = "",
  valueStream = "ESP-ABS",
  view = "tree",
  className = "",
  style = {},
  onStateChange,
}) {
  const iframeRef = useRef(null);
  const [isReady, setIsReady] = useState(false);

  // Normalize base URL
  const baseUrl = endpoint ? endpoint.replace(/\/$/, "") : "";
  const iframeSrc = baseUrl ? `${baseUrl}/` : "/";

  // Handle messages sent from the embedded dashboard
  const handleIncomingMessage = useCallback(
    (event) => {
      if (!event.data || typeof event.data !== "object") return;

      if (event.data.type === "KPI_DASHBOARD_STATE") {
        if (typeof onStateChange === "function") {
          onStateChange(event.data.state);
        }
      }
    },
    [onStateChange]
  );

  useEffect(() => {
    window.addEventListener("message", handleIncomingMessage);
    return () => {
      window.removeEventListener("message", handleIncomingMessage);
    };
  }, [handleIncomingMessage]);

  // When iframe loads, sync initial props if customized
  const handleIframeLoad = () => {
    setIsReady(true);
    if (!iframeRef.current || !iframeRef.current.contentWindow) return;

    if (valueStream) {
      iframeRef.current.contentWindow.postMessage(
        { type: "SET_VALUE_STREAM", payload: valueStream },
        "*"
      );
    }

    if (view && view !== "tree") {
      iframeRef.current.contentWindow.postMessage(
        { type: "SWITCH_VIEW", payload: view },
        "*"
      );
    }
  };

  return (
    <div
      className={`kpi-dashboard-embed-container ${className}`}
      style={{
        width: "100%",
        height: "100%",
        minHeight: "550px",
        position: "relative",
        display: "flex",
        flexDirection: "column",
        overflow: "hidden",
        backgroundColor: "#f4f7f6",
        ...style,
      }}
    >
      <iframe
        ref={iframeRef}
        src={iframeSrc}
        title="Inventory KPI Tree Dashboard"
        onLoad={handleIframeLoad}
        style={{
          width: "100%",
          height: "100%",
          border: "none",
          flex: 1,
          display: "block",
          backgroundColor: "#f4f7f6",
        }}
      />
    </div>
  );
}

export { InventoryKPIDashboard };
