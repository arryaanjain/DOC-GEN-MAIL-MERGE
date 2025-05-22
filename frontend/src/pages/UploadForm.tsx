import { useState } from "react";

export default function UploadForm() {
  const [file, setFile] = useState<File | null>(null);
  const [message, setMessage] = useState("");

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFile(e.target.files?.[0] || null);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) {
      setMessage("Please select a DOCX file.");
      return;
    }
    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch(`${import.meta.env.VITE_API_URL}/upload`, {
        method: "POST",
        body: formData,
      });
      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "converted.xlsx";
        a.click();
        setMessage("File converted and downloaded!");
      } else {
        const err = await response.json();
        setMessage(err.error || "Conversion failed.");
      }
    } catch (err) {
      setMessage("Error uploading file.");
    }
  };

  return (
    <div className="flex items-center justify-center min-h-screen bg-gray-50">
      <form onSubmit={handleSubmit} className="bg-white p-8 rounded shadow-md w-96">
        <h2 className="text-2xl font-bold mb-6 text-center">Convert DOCX to Excel</h2>
        <input
          className="w-full mb-4"
          type="file"
          accept=".docx"
          onChange={handleFileChange}
        />
        <button className="w-full bg-purple-600 text-white py-2 rounded hover:bg-purple-700" type="submit">
          Convert & Download
        </button>
        {message && <p className="mt-4 text-center text-red-500">{message}</p>}
      </form>
    </div>
  );
}