import React, { useState, useEffect } from "react";
import {
  Upload,
  FileSpreadsheet,
  BarChart3,
  Calculator,
  AlertCircle,
  CheckCircle2,
  Loader2,
  Server,
} from "lucide-react";

const ExcelProcessor = () => {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [data, setData] = useState(null);
  const [selectedSheet, setSelectedSheet] = useState("balance_sheet");
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [serverStatus, setServerStatus] = useState("checking");

  const sheetInfo = {
    balance_sheet: {
      name: "Balance Sheet",
      display: "Cân đối kế toán",
      icon: <FileSpreadsheet className="w-5 h-5" />,
      color: "bg-blue-500",
    },
    income_statement: {
      name: "Income Statement",
      display: "Kết quả kinh doanh",
      icon: <BarChart3 className="w-5 h-5" />,
      color: "bg-green-500",
    },
    financial_ratios: {
      name: "Financial Ratios",
      display: "Chỉ số tài chính",
      icon: <Calculator className="w-5 h-5" />,
      color: "bg-purple-500",
    },
  };

  // Check server status on component mount
  useEffect(() => {
    const checkServerStatus = async () => {
      try {
        const response = await fetch("http://127.0.0.1:5000/api/health");
        if (response.ok) {
          setServerStatus("online");
        } else {
          setServerStatus("offline");
        }
      } catch (err) {
        setServerStatus("offline");
      }
    };

    checkServerStatus();
    // Check every 10 seconds
    const interval = setInterval(checkServerStatus, 10000);
    return () => clearInterval(interval);
  }, []);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      const fileExtension = selectedFile.name.split(".").pop().toLowerCase();
      if (fileExtension === "xlsx" || fileExtension === "xls") {
        setFile(selectedFile);
        setError("");
        setData(null);
        setSuccess("");
      } else {
        setError("Please select a valid Excel file (.xlsx or .xls)");
        setFile(null);
      }
    }
  };

  const handleUpload = async () => {
    if (!file) {
      setError("Please select a file first");
      return;
    }

    setLoading(true);
    setError("");
    setSuccess("");

    const formData = new FormData();
    formData.append("file", file);

    try {
      console.log("Attempting to connect to Flask server...");
      const response = await fetch("http://127.0.0.1:5000/api/upload", {
        method: "POST",
        body: formData,
      });

      console.log("Response received:", response.status);
      const result = await response.json();

      if (response.ok) {
        setData(result.data);
        setSuccess(result.message);
        // Auto-select first available sheet
        const availableSheets = Object.keys(result.data);
        if (availableSheets.length > 0) {
          setSelectedSheet(availableSheets[0]);
        }
      } else {
        setError(result.error || "Failed to process file");
      }
    } catch (err) {
      console.error("Connection error:", err);
      setError(
        `Connection failed: ${err.message}. Check if Flask server is running on port 5000.`
      );
    } finally {
      setLoading(false);
    }
  };

  const renderDataFrame = () => {
    if (!data || !data[selectedSheet]) {
      return (
        <div
          style={{
            textAlign: "center",
            padding: "48px 12px",
            color: "#6B7280",
          }}
        >
          <FileSpreadsheet
            style={{
              width: "64px",
              height: "64px",
              margin: "0 auto 16px",
              opacity: "0.5",
            }}
          />
          <p>No data available for the selected sheet</p>
        </div>
      );
    }

    const sheetData = data[selectedSheet];
    const { columns, data: rows, shape } = sheetData;

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            flexWrap: "wrap",
            gap: "10px",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            {sheetInfo[selectedSheet]?.icon}
            <h3
              style={{
                fontSize: "1.125rem",
                fontWeight: "600",
                color: "#1F2937",
                margin: 0,
              }}
            >
              {sheetInfo[selectedSheet]?.display || "Data"}
            </h3>
          </div>
          <div
            style={{
              fontSize: "0.875rem",
              color: "#6B7280",
              background: "#F3F4F6",
              padding: "4px 8px",
              borderRadius: "4px",
            }}
          >
            Shape: {shape[0]} rows × {shape[1]} columns
          </div>
        </div>

        <div
          style={{
            overflowX: "auto",
            overflowY: "auto",
            maxHeight: "500px",
            border: "1px solid #E5E7EB",
            borderRadius: "8px",
            boxShadow: "0 1px 3px 0 rgba(0, 0, 0, 0.1)",
          }}
        >
          <table
            style={{
              minWidth: "100%",
              borderCollapse: "collapse",
              fontSize: "0.875rem",
            }}
          >
            <thead
              style={{
                background: "#F9FAFB",
                position: "sticky",
                top: 0,
                zIndex: 1,
              }}
            >
              <tr>
                {columns.map((col, index) => (
                  <th
                    key={index}
                    style={{
                      padding: "12px 16px",
                      textAlign: "left",
                      fontSize: "0.75rem",
                      fontWeight: "600",
                      color: "#374151",
                      textTransform: "uppercase",
                      letterSpacing: "0.05em",
                      borderRight: "1px solid #E5E7EB",
                      background: "#F9FAFB",
                      minWidth: index === 0 ? "200px" : "120px",
                      position: "sticky",
                      top: 0,
                    }}
                  >
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody style={{ background: "white" }}>
              {rows.slice(0, 100).map((row, rowIndex) => (
                <tr
                  key={rowIndex}
                  style={{
                    background: rowIndex % 2 === 0 ? "white" : "#F9FAFB",
                    transition: "background-color 0.2s",
                  }}
                  onMouseOver={(e) => {
                    e.currentTarget.style.backgroundColor = "#EBF8FF";
                  }}
                  onMouseOut={(e) => {
                    e.currentTarget.style.backgroundColor =
                      rowIndex % 2 === 0 ? "white" : "#F9FAFB";
                  }}
                >
                  {row.map((cell, cellIndex) => (
                    <td
                      key={cellIndex}
                      style={{
                        padding: "12px 16px",
                        color: "#1F2937",
                        borderRight: "1px solid #E5E7EB",
                        borderBottom: "1px solid #F3F4F6",
                        maxWidth: cellIndex === 0 ? "250px" : "150px",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                        fontWeight: cellIndex === 0 ? "500" : "normal",
                      }}
                      title={cell?.toString() || ""}
                    >
                      {cell !== null && cell !== undefined
                        ? typeof cell === "number"
                          ? cell.toLocaleString()
                          : cell.toString()
                        : "-"}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {rows.length > 100 && (
          <p
            style={{
              fontSize: "0.875rem",
              color: "#6B7280",
              textAlign: "center",
              margin: 0,
              padding: "8px",
              background: "#F9FAFB",
              borderRadius: "4px",
            }}
          >
            Showing first 100 rows of {rows.length} total rows
          </p>
        )}
      </div>
    );
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
        padding: "20px",
        fontFamily:
          '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
      }}
    >
      <div style={{ maxWidth: "1200px", margin: "0 auto" }}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: "40px" }}>
          <h1
            style={{
              fontSize: "2.5rem",
              fontWeight: "bold",
              color: "white",
              marginBottom: "10px",
              textShadow: "2px 2px 4px rgba(0,0,0,0.3)",
            }}
          >
            Financial Data Processor
          </h1>
          <p
            style={{
              color: "rgba(255,255,255,0.9)",
              fontSize: "1.1rem",
              marginBottom: "20px",
            }}
          >
            Upload Excel files and view financial dataframes
          </p>

          {/* Server Status Indicator */}
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: "8px",
              marginTop: "15px",
            }}
          >
            <Server className="w-4 h-4" style={{ color: "white" }} />
            <span style={{ fontSize: "0.9rem", color: "white" }}>
              Flask Server:
              <span
                style={{
                  marginLeft: "5px",
                  fontWeight: "600",
                  color:
                    serverStatus === "online"
                      ? "#10B981"
                      : serverStatus === "offline"
                      ? "#EF4444"
                      : "#F59E0B",
                }}
              >
                {serverStatus === "online"
                  ? "🟢 Online"
                  : serverStatus === "offline"
                  ? "🔴 Offline"
                  : "🟡 Checking..."}
              </span>
            </span>
          </div>
        </div>

        {/* Upload Section */}
        <div
          style={{
            background: "white",
            borderRadius: "16px",
            boxShadow:
              "0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)",
            padding: "30px",
            marginBottom: "30px",
          }}
        >
          <div
            style={{
              border: "2px dashed #D1D5DB",
              borderRadius: "12px",
              padding: "40px 20px",
              textAlign: "center",
              transition: "border-color 0.2s",
              cursor: "pointer",
            }}
            onDragOver={(e) => {
              e.preventDefault();
              e.currentTarget.style.borderColor = "#3B82F6";
            }}
            onDragLeave={(e) => {
              e.currentTarget.style.borderColor = "#D1D5DB";
            }}
          >
            <Upload
              style={{
                width: "48px",
                height: "48px",
                margin: "0 auto 20px",
                color: "#9CA3AF",
              }}
            />

            <div style={{ marginBottom: "20px" }}>
              <label htmlFor="file-upload" style={{ cursor: "pointer" }}>
                <span
                  style={{
                    background: "#3B82F6",
                    color: "white",
                    padding: "12px 24px",
                    borderRadius: "8px",
                    border: "none",
                    fontSize: "1rem",
                    fontWeight: "500",
                    cursor: "pointer",
                    transition: "background-color 0.2s",
                    display: "inline-block",
                  }}
                  onMouseOver={(e) =>
                    (e.target.style.backgroundColor = "#2563EB")
                  }
                  onMouseOut={(e) =>
                    (e.target.style.backgroundColor = "#3B82F6")
                  }
                >
                  Choose Excel File
                </span>
                <input
                  id="file-upload"
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileChange}
                  style={{ display: "none" }}
                />
              </label>
            </div>

            {file && (
              <div
                style={{
                  fontSize: "0.9rem",
                  color: "#6B7280",
                  marginBottom: "15px",
                }}
              >
                Selected: {file.name} ({(file.size / 1024 / 1024).toFixed(2)}{" "}
                MB)
              </div>
            )}

            <button
              onClick={handleUpload}
              disabled={!file || loading}
              style={{
                background: !file || loading ? "#9CA3AF" : "#10B981",
                color: "white",
                padding: "12px 32px",
                borderRadius: "8px",
                border: "none",
                fontSize: "1rem",
                fontWeight: "500",
                cursor: !file || loading ? "not-allowed" : "pointer",
                display: "flex",
                alignItems: "center",
                gap: "8px",
                margin: "0 auto",
                transition: "background-color 0.2s",
              }}
              onMouseOver={(e) => {
                if (!(!file || loading)) {
                  e.target.style.backgroundColor = "#059669";
                }
              }}
              onMouseOut={(e) => {
                if (!(!file || loading)) {
                  e.target.style.backgroundColor = "#10B981";
                }
              }}
            >
              {loading && <Loader2 className="w-4 h-4 animate-spin" />}
              <span>{loading ? "Processing..." : "Process File"}</span>
            </button>
          </div>

          {/* Status Messages */}
          {error && (
            <div
              style={{
                marginTop: "20px",
                padding: "16px",
                background: "#FEF2F2",
                border: "1px solid #FECACA",
                borderRadius: "8px",
                display: "flex",
                alignItems: "center",
                gap: "8px",
              }}
            >
              <AlertCircle
                style={{ width: "20px", height: "20px", color: "#EF4444" }}
              />
              <span style={{ color: "#B91C1C" }}>{error}</span>
            </div>
          )}

          {success && (
            <div
              style={{
                marginTop: "20px",
                padding: "16px",
                background: "#F0FDF4",
                borderRadius: "8px",
                display: "flex",
                alignItems: "center",
                gap: "8px",
              }}
            >
              <CheckCircle2
                style={{ width: "20px", height: "20px", color: "#10B981" }}
              />
              <span style={{ color: "#15803D" }}>{success}</span>
            </div>
          )}
        </div>

        {/* Sheet Selection */}
        {data && (
          <div
            style={{
              background: "white",
              borderRadius: "16px",
              boxShadow: "0 20px 25px -5px rgba(0, 0, 0, 0.1)",
              padding: "30px",
              marginBottom: "30px",
            }}
          >
            <h2
              style={{
                fontSize: "1.25rem",
                fontWeight: "600",
                marginBottom: "20px",
                color: "#1F2937",
              }}
            >
              Select Financial Statement
            </h2>
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: "12px",
              }}
            >
              {Object.keys(data).map((sheetKey) => {
                const sheet = sheetInfo[sheetKey];
                const isSelected = selectedSheet === sheetKey;
                const backgroundColor = isSelected
                  ? sheetKey === "balance_sheet"
                    ? "#3B82F6"
                    : sheetKey === "income_statement"
                    ? "#10B981"
                    : "#8B5CF6"
                  : "white";

                return (
                  <button
                    key={sheetKey}
                    onClick={() => setSelectedSheet(sheetKey)}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "8px",
                      padding: "12px 16px",
                      borderRadius: "8px",
                      border: isSelected ? "none" : "1px solid #D1D5DB",
                      background: backgroundColor,
                      color: isSelected ? "white" : "#374151",
                      cursor: "pointer",
                      transition: "all 0.2s",
                      fontSize: "0.95rem",
                      fontWeight: "500",
                    }}
                    onMouseOver={(e) => {
                      if (!isSelected) {
                        e.target.style.backgroundColor = "#F9FAFB";
                      }
                    }}
                    onMouseOut={(e) => {
                      if (!isSelected) {
                        e.target.style.backgroundColor = "white";
                      }
                    }}
                  >
                    {sheet?.icon}
                    <span>{sheet?.name}</span>
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* Data Display */}
        {data && (
          <div
            style={{
              background: "white",
              borderRadius: "16px",
              boxShadow: "0 20px 25px -5px rgba(0, 0, 0, 0.1)",
              padding: "30px",
            }}
          >
            <h2
              style={{
                fontSize: "1.25rem",
                fontWeight: "600",
                marginBottom: "20px",
                color: "#1F2937",
              }}
            >
              DataFrame View
            </h2>
            {renderDataFrame()}
          </div>
        )}
      </div>
    </div>
  );
};

export default ExcelProcessor;
