"use client";

import { ChangeEvent, useState } from "react";
import * as XLSX from "xlsx";

type SheetRow = Record<string, string>;
type DiscoveryResponses = Record<number, Record<string, string>>;

const DEFAULT_DISCOVERY_QUESTIONS = [
  "What is the primary business need or goal for this procurement?",
  "What is the desired timeline for completion?",
  "Are there any specific vendors or solutions already in consideration?",
  "What is the estimated budget range?",
  "Who are the key stakeholders involved in this decision?",
];

export default function Home() {
  const [rows, setRows] = useState<SheetRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [selectedRow, setSelectedRow] = useState(0);
  const [status, setStatus] = useState("Upload an Excel file to import and fill out your procurement information.");
  const [fileName, setFileName] = useState("");
  const [discoveryQuestions, setDiscoveryQuestions] = useState<string[]>(DEFAULT_DISCOVERY_QUESTIONS);
  const [discoveryResponses, setDiscoveryResponses] = useState<DiscoveryResponses>({});
  const [submitStatus, setSubmitStatus] = useState<string>("");

  const handleFileUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setFileName(file.name);
    setStatus("Reading spreadsheet...");

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
        defval: "",
      });

      const normalizedRows = jsonRows.map((row) => {
        return Object.fromEntries(
          Object.entries(row).map(([key, value]) => [String(key).trim(), String(value ?? "")])
        ) as SheetRow;
      });

      if (normalizedRows.length === 0) {
        setRows([]);
        setHeaders([]);
        setSelectedRow(0);
        setStatus("The file did not contain any rows to edit.");
        return;
      }

      const nextHeaders = Object.keys(normalizedRows[0]);
      setHeaders(nextHeaders);
      setRows(normalizedRows);
      setSelectedRow(0);
      setStatus(`Loaded ${normalizedRows.length} item${normalizedRows.length === 1 ? "" : "s"} from ${file.name}.`);
    } catch (error) {
      console.error(error);
      setStatus("Unable to read that file. Please use a .xlsx, .xls, or .csv file.");
    }
  };

  const updateRowField = (rowIndex: number, column: string, value: string) => {
    setRows((currentRows) =>
      currentRows.map((row, index) => (index === rowIndex ? { ...row, [column]: value } : row))
    );
  };

  const addNewRow = () => {
    const emptyRow = Object.fromEntries(headers.map((header) => [header, ""])) as SheetRow;
    setRows((currentRows) => [...currentRows, emptyRow]);
    setSelectedRow(rows.length);
  };

  const updateDiscoveryResponse = (rowIndex: number, questionIndex: number, value: string) => {
    setDiscoveryResponses((prev) => ({
      ...prev,
      [rowIndex]: {
        ...(prev[rowIndex] || {}),
        [questionIndex]: value,
      },
    }));
  };

  const handleSubmitForm = () => {
    if (rows.length === 0) {
      setSubmitStatus("Please upload and fill out at least one record before submitting.");
      return;
    }

    const submissionData = {
      timestamp: new Date().toISOString(),
      records: rows.map((row, index) => ({
        recordData: row,
        discoveryAnswers: discoveryResponses[index] || {},
      })),
    };

    console.log("Form submission:", submissionData);
    setSubmitStatus("✓ Form submitted successfully! Check the console for details.");
    
    // Reset status after 5 seconds
    setTimeout(() => setSubmitStatus(""), 5000);
  };

  const currentRow = rows[selectedRow];
  const currentDiscoveryResponses = discoveryResponses[selectedRow] || {};

  return (
    <main className="min-h-screen bg-slate-950 px-4 py-10 text-slate-100 sm:px-6 lg:px-8">
      <div className="mx-auto flex max-w-6xl flex-col gap-6">
        <section className="rounded-3xl border border-slate-800 bg-slate-900/80 p-8 shadow-2xl shadow-black/20">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <div className="max-w-2xl">
              <p className="mb-2 text-sm font-semibold uppercase tracking-[0.3em] text-cyan-400">Procurement intake</p>
              <h1 className="text-3xl font-semibold sm:text-4xl">Upload an Excel sheet and fill out the details in a guided form.</h1>
              <p className="mt-3 text-base text-slate-400">
                Import rows from a spreadsheet, review each record, and update the information without leaving the page.
              </p>
            </div>
            <label className="flex cursor-pointer flex-col items-start rounded-2xl border border-dashed border-cyan-500/60 bg-slate-800/70 px-4 py-3 text-sm font-medium text-cyan-300 transition hover:border-cyan-400 hover:bg-slate-800">
              <span>Choose Excel file</span>
              <input className="hidden" type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} />
            </label>
          </div>

          <div className="mt-6 rounded-2xl border border-slate-800 bg-slate-950/70 p-4 text-sm text-slate-300">
            <p className="font-medium text-slate-200">{status}</p>
            {fileName ? <p className="mt-1 text-slate-400">Selected file: {fileName}</p> : null}
          </div>
        </section>

        <section className="grid gap-6 lg:grid-cols-[280px_minmax(0,1fr)]">
          <aside className="rounded-3xl border border-slate-800 bg-slate-900/80 p-5 shadow-xl shadow-black/10">
            <div className="flex items-center justify-between">
              <h2 className="text-lg font-semibold">Records</h2>
              <span className="rounded-full bg-cyan-500/15 px-2.5 py-1 text-xs font-medium text-cyan-300">{rows.length}</span>
            </div>
            <div className="mt-4 space-y-2">
              {rows.length === 0 ? (
                <p className="rounded-2xl border border-dashed border-slate-700 p-3 text-sm text-slate-400">
                  No records loaded yet.
                </p>
              ) : (
                rows.map((row, index) => {
                  const firstValue = Object.values(row)[0] ?? "Untitled record";
                  return (
                    <button
                      key={`${firstValue}-${index}`}
                      type="button"
                      onClick={() => setSelectedRow(index)}
                      className={`w-full rounded-2xl border px-3 py-2 text-left text-sm transition ${
                        selectedRow === index
                          ? "border-cyan-500 bg-cyan-500/10 text-cyan-200"
                          : "border-slate-800 bg-slate-950/70 text-slate-300 hover:border-slate-700"
                      }`}
                    >
                      <div className="font-medium">{firstValue || `Record ${index + 1}`}</div>
                      <div className="mt-1 text-xs text-slate-500">{headers[0] || "Imported row"}</div>
                    </button>
                  );
                })
              )}
            </div>
            <button
              type="button"
              onClick={addNewRow}
              className="mt-4 w-full rounded-2xl border border-slate-700 bg-slate-800 px-3 py-2 text-sm font-medium text-slate-200 transition hover:border-cyan-500 hover:text-cyan-200"
            >
              Add empty record
            </button>
          </aside>

          <section className="rounded-3xl border border-slate-800 bg-slate-900/80 p-6 shadow-xl shadow-black/10">
            {rows.length === 0 ? (
              <div className="rounded-2xl border border-dashed border-slate-700 bg-slate-950/50 p-8 text-center text-slate-400">
                Upload a spreadsheet to start filling the form.
              </div>
            ) : (
              <>
                <div className="mb-6 flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <p className="text-sm font-semibold uppercase tracking-[0.3em] text-cyan-400">Editable form</p>
                    <h2 className="text-2xl font-semibold text-white">Record {selectedRow + 1}</h2>
                  </div>
                  <div className="rounded-full border border-slate-700 bg-slate-950/70 px-3 py-1 text-sm text-slate-400">
                    {headers.length} fields available
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-2">
                  {headers.map((header) => (
                    <label key={header} className="flex flex-col gap-2 text-sm text-slate-300">
                      <span className="font-medium text-slate-200">{header}</span>
                      <input
                        className="rounded-2xl border border-slate-700 bg-slate-950/70 px-3 py-2 text-slate-100 outline-none ring-0 transition focus:border-cyan-500"
                        value={currentRow?.[header] ?? ""}
                        onChange={(event) => updateRowField(selectedRow, header, event.target.value)}
                        placeholder={`Enter ${header}`}
                      />
                    </label>
                  ))}
                </div>

                <div className="mt-8 border-t border-slate-800 pt-8">
                  <div className="mb-6">
                    <p className="text-sm font-semibold uppercase tracking-[0.3em] text-cyan-400">Discovery questions</p>
                    <h3 className="text-xl font-semibold text-white">Fill out the discovery questionnaire</h3>
                  </div>

                  <div className="space-y-4">
                    {discoveryQuestions.map((question, qIndex) => (
                      <label key={qIndex} className="flex flex-col gap-2 text-sm text-slate-300">
                        <span className="font-medium text-slate-200">
                          {qIndex + 1}. {question}
                        </span>
                        <textarea
                          className="rounded-2xl border border-slate-700 bg-slate-950/70 px-3 py-2 text-slate-100 outline-none ring-0 transition focus:border-cyan-500"
                          value={currentDiscoveryResponses[qIndex] ?? ""}
                          onChange={(event) => updateDiscoveryResponse(selectedRow, qIndex, event.target.value)}
                          placeholder="Enter your response..."
                          rows={3}
                        />
                      </label>
                    ))}
                  </div>
                </div>

                <div className="mt-8 flex flex-col gap-3 sm:flex-row sm:justify-between">
                  <div>
                    {submitStatus && (
                      <p className={`text-sm font-medium ${submitStatus.includes("✓") ? "text-green-400" : "text-amber-400"}`}>
                        {submitStatus}
                      </p>
                    )}
                  </div>
                  <button
                    type="button"
                    onClick={handleSubmitForm}
                    className="rounded-2xl bg-cyan-600 px-6 py-3 font-medium text-white transition hover:bg-cyan-500 active:bg-cyan-700"
                  >
                    Submit Form
                  </button>
                </div>
              </>
            )}
          </section>
        </section>
      </div>
    </main>
  );
}
