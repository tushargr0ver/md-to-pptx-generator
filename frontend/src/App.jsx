import React, { useState, useCallback } from 'react';

function App() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [statusText, setStatusText] = useState("");
  const [provider, setProvider] = useState("gemini");
  const [model, setModel] = useState("gemini-3.1-flash");
  const [apiKey, setApiKey] = useState("");

  const onDragOver = useCallback((e) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const onDragLeave = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.name.endsWith('.md')) {
        setFile(droppedFile);
      } else {
        alert('Please upload a Markdown (.md) file');
      }
    }
  }, []);

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
  };

  const handleGenerate = async () => {
    if (!file) return;

    setLoading(true);
    setStatusText("Initializing Agents...");

    const formData = new FormData();
    formData.append("markdown_file", file);
    formData.append("provider", provider);
    formData.append("model", model);
    formData.append("api_key", apiKey);

    try {
      // Step 1: Preprocessing
      setTimeout(() => setStatusText(`Structuring Storyline with ${provider === 'gemini' ? 'Gemini' : 'OpenAI'}...`), 1000);
      
      const response = await fetch("http://localhost:8000/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Failed to generate presentation");
      }

      setStatusText("Rendering PPTX Layouts...");
      
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.style.display = "none";
      a.href = url;
      a.download = file.name.replace(".md", ".pptx");
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      
    } catch (err) {
      alert("Error: " + err.message);
    } finally {
      setLoading(false);
      setStatusText("");
    }
  };

  return (
    <div className="min-h-screen relative overflow-hidden flex flex-col items-center justify-center p-6">
      {/* Background Blobs */}
      <div className="absolute top-[-10%] left-[-10%] w-[40vw] h-[40vw] rounded-full bg-brand-600/20 blur-[100px] animate-blob"></div>
      <div className="absolute bottom-[-10%] right-[-10%] w-[50vw] h-[50vw] rounded-full bg-purple-600/20 blur-[120px] animate-blob" style={{ animationDelay: '2s' }}></div>
      <div className="absolute top-[20%] right-[10%] w-[30vw] h-[30vw] rounded-full bg-pink-600/10 blur-[80px] animate-blob" style={{ animationDelay: '4s' }}></div>

      <main className="z-10 w-full max-w-4xl glass-panel rounded-3xl p-8 md:p-14 relative">
        <header className="mb-12 text-center">
          <h1 className="text-5xl md:text-6xl font-black mb-4 tracking-tight" style={{ fontFamily: 'var(--font-display)' }}>
            Code <span className="gradient-text">EZ</span> Hackathon
          </h1>
          <p className="text-xl text-gray-400 font-light max-w-2xl mx-auto">
            Drag & drop your comprehensive Markdown to automatically generate a visually stunning, perfectly structured PowerPoint deck.
          </p>
        </header>

        <section 
          className={`
            border-2 border-dashed rounded-2xl p-12 flex flex-col items-center justify-center text-center transition-all duration-300 relative overflow-hidden
            ${isDragging ? 'border-brand-500 bg-brand-500/10 scale-[1.02]' : 'border-gray-700 hover:border-gray-500'}
            ${file ? 'bg-gray-800/40 border-green-500/50' : ''}
          `}
          onDragOver={onDragOver}
          onDragLeave={onDragLeave}
          onDrop={onDrop}
        >
          {loading && (
            <div className="absolute inset-0 z-20 bg-dark-900/80 backdrop-blur-sm flex flex-col items-center justify-center rounded-2xl">
              <div className="w-16 h-16 border-4 border-gray-700 border-t-brand-500 rounded-full animate-spin mb-4"></div>
              <p className="text-lg font-medium gradient-text animate-pulse">{statusText}</p>
            </div>
          )}

          {!file ? (
            <>
              <div className="w-20 h-20 bg-gray-800 rounded-full flex items-center justify-center mb-6 shadow-xl">
                <svg className="w-8 h-8 text-brand-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                </svg>
              </div>
              <h3 className="text-2xl font-bold mb-2">Upload your .md file</h3>
              <p className="text-gray-400 mb-6">Or click to browse from your computer</p>
              <label className="cursor-pointer bg-white text-black px-6 py-3 rounded-full font-semibold hover:bg-gray-200 transition-colors shadow-lg shadow-white/10">
                Browse Files
                <input type="file" className="hidden" accept=".md" onChange={handleFileChange} />
              </label>
            </>
          ) : (
            <div className="flex flex-col items-center">
              <div className="w-20 h-20 bg-green-500/20 text-green-400 rounded-full flex items-center justify-center mb-6 shadow-xl border border-green-500/30">
                 <svg className="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              </div>
              <h3 className="text-2xl font-bold mb-2">{file.name}</h3>
              <p className="text-green-400/80 mb-8 pb-4 border-b border-gray-700 w-full">Ready for generation</p>
              
              <div className="w-full max-w-sm mb-8 space-y-4 text-left">
                <div>
                  <label className="block text-sm font-medium text-gray-400 mb-1">AI Provider</label>
                  <select 
                    value={provider} 
                    onChange={(e) => {
                      setProvider(e.target.value);
                      setModel(e.target.value === 'gemini' ? 'gemini-3.1-flash' : 'gpt-5.4-mini');
                    }}
                    className="w-full bg-dark-900 border border-gray-700 rounded-lg px-4 py-2 text-white outline-none focus:border-brand-500"
                  >
                    <option value="gemini">Google Gemini</option>
                    <option value="openai">OpenAI</option>
                  </select>
                </div>
                
                <div>
                  <label className="block text-sm font-medium text-gray-400 mb-1">Model</label>
                  <select 
                    value={model} 
                    onChange={(e) => setModel(e.target.value)}
                    className="w-full bg-dark-900 border border-gray-700 rounded-lg px-4 py-2 text-white outline-none focus:border-brand-500"
                  >
                    {provider === 'gemini' ? (
                      <>
                        <option value="gemini-3.1-pro">Gemini 3.1 Pro</option>
                        <option value="gemini-3.1-flash">Gemini 3.1 Flash</option>
                        <option value="gemini-3.1-nano">Gemini 3.1 Nano</option>
                      </>
                    ) : (
                      <>
                        <option value="gpt-5.4">GPT-5.4</option>
                        <option value="gpt-5.4-mini">GPT-5.4 Mini</option>
                        <option value="gpt-5.4-nano">GPT-5.4 Nano</option>
                        <option value="o3">o3</option>
                        <option value="o4-mini">o4-mini</option>
                      </>
                    )}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-400 mb-1">API Key</label>
                  <input 
                    type="text"
                    placeholder={`Mandatory: Enter your ${provider === 'gemini' ? 'Google AI Studio' : 'OpenAI'} API Key`}
                    value={apiKey}
                    onChange={(e) => setApiKey(e.target.value)}
                    className="w-full bg-dark-900 border border-gray-700 rounded-lg px-4 py-2 text-white outline-none focus:border-brand-500 placeholder-gray-600"
                  />
                  <p className="text-xs text-brand-500/80 mt-2">
                    Required: We don't store your keys. Your key is directly used per execution.
                  </p>
                </div>
              </div>
              
              <div className="flex gap-4">
                <button 
                  onClick={() => setFile(null)}
                  className="px-6 py-3 rounded-full font-medium border border-gray-600 hover:bg-gray-800 transition-colors"
                >
                  Clear
                </button>
                <button 
                  onClick={handleGenerate}
                  className="px-8 py-3 rounded-full font-bold bg-gradient-to-r from-brand-600 to-purple-600 hover:from-brand-500 hover:to-purple-500 transition-all shadow-lg shadow-purple-500/20 hover:shadow-purple-500/40 relative overflow-hidden group"
                >
                  <span className="relative z-10">Generate PPTX</span>
                  <div className="absolute inset-0 h-full w-full bg-white/20 group-hover:translate-x-full transition-transform duration-500 -translate-x-full rotate-12"></div>
                </button>
              </div>
            </div>
          )}
        </section>
      </main>
    </div>
  );
}

export default App;
