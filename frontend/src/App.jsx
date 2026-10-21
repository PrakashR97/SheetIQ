import { useState } from 'react'
import './App.css'

function App() {
  const [file, setFile] = useState(null)
  const [dateFormat, setDateFormat] = useState('MM/DD/YYYY')
  const [loading, setLoading] = useState(false)
  const [result, setResult] = useState(null)
  const [error, setError] = useState('')
  const [page, setPage] = useState('upload')

  async function handleClean() {
    if (!file) {
      setError('Please select an Excel file.')
      return
    }

    setLoading(true)
    setError('')
    setResult(null)

    const formData = new FormData()
    formData.append('file', file)

    try {
      const response = await fetch('/api/clean', {
        method: 'POST',
        body: formData,
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.detail || 'Cleaning failed.')
      }

      setResult(data)
      setPage('results')
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  function handleBack() {
    setPage('upload')
    setFile(null)
    setResult(null)
    setError('')
  }

  return (
    <main className='page'>
      <header className='header'>
        <div className='logo'>EC / AI</div>
        <span className='header-label'>DATA QUALITY WORKSPACE</span>
      </header>

      <section className='workspace'>
        {page === 'upload' && (
          <>
            <div className='eyebrow'>EXCEL DATA CLEANING</div>

            <h1>Clean data. Clear results.</h1>

            <p className='intro'>
              Upload your Excel workbook, clean inconsistent data, and download
              a detailed CSV report.
            </p>

            <div className='panel'>
              <div className='section-heading'>
                <span>01 / UPLOAD WORKBOOK</span>
                <span>.XLSX</span>
              </div>

              <label className='upload-zone'>
                <input
                  type='file'
                  accept='.xlsx'
                  onChange={event => {
                    setFile(event.target.files?.[0] || null)
                    setResult(null)
                    setError('')
                  }}
                />

                <span className='upload-symbol'>+</span>

                <strong>{file ? file.name : 'Choose an Excel file'}</strong>

                <small>Click to browse your files</small>
              </label>

              <div className='section-heading settings-heading'>
                <span>02 / CLEANING SETTINGS</span>
              </div>

              <label className='field-label' htmlFor='date-format'>
                Date output format
              </label>

              <select
                id='date-format'
                value={dateFormat}
                onChange={event => setDateFormat(event.target.value)}
              >
                <option value='MM/DD/YYYY'>MM/DD/YYYY</option>
                <option value='DD/MM/YYYY'>DD/MM/YYYY</option>
                <option value='MM-DD-YYYY'>MM-DD-YYYY</option>
              </select>

              <p className='note'>
                Date format selection is a preview setting. Backend integration
                is coming next.
              </p>

              <div className='rules'>
                <span>✓ Remove exact duplicates</span>
                <span>✓ Replace missing values</span>
                <span>✓ Trim whitespace</span>
                <span>✓ Process all sheets</span>
              </div>

              <button
                className='primary-button'
                onClick={handleClean}
                disabled={loading || !file}
              >
                {loading ? 'CLEANING...' : 'START CLEANING →'}
              </button>

              {error && <p className='error'>{error}</p>}
            </div>
          </>
        )}

        {page === 'results' && result && (
          <div className='panel result-panel'>
            <div className='section-heading'>
              <span>03 / RESULTS</span>
              <span>COMPLETE</span>
            </div>

            <h2>Cleaning completed.</h2>

            <p>Your files are ready to download.</p>

            <div className='download-actions'>
              <a href={`/api${result.cleaned_excel_url}`}>
                ↓ Download cleaned Excel
              </a>

              {result.cleaned_csv_url && (
                <a href={`/api${result.cleaned_csv_url}`}>
                  ↓ Download cleaned CSV
                </a>
              )}

              <a href={`/api${result.csv_report_url}`}>↓ Download CSV report</a>
            </div>

            <button
              className='primary-button'
              style={{ marginTop: '24px' }}
              onClick={handleBack}
            >
              ← BACK TO UPLOAD
            </button>
          </div>
        )}
      </section>

      <footer className='footer'>
        <span>DEVELOPED BY PRAKASH R</span>
        <span>
          <a href='mailto:prakashramasamy97@gmail.com'>
            prakashramasamy97@gmail.com
          </a>
        </span>
      </footer>
    </main>
  )
}

export default App
