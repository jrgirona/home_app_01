import { useState } from 'react'
import './App.css'

function App() {
  const [year, setYear] = useState(new Date().getFullYear())
  const [reportType, setReportType] = useState('IRPF')
  const [loading, setLoading] = useState(false)
  const [message, setMessage] = useState('')
  const [error, setError] = useState(false)

  // Generate an array of years from 2025 to the current year
  const currentYear = new Date().getFullYear()
  const years = []
  for (let y = 2025; y <= Math.max(2025, currentYear); y++) {
    years.push(y)
  }

  const handleGenerate = async () => {
    setLoading(true)
    setMessage('')
    setError(false)

    try {
      const response = await fetch('http://127.0.0.1:8000/api/generate', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          year: parseInt(year),
          report_type: reportType
        })
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.detail || 'Failed to generate report')
      }

      setMessage(data.message || 'Report generated successfully!')
    } catch (err) {
      setError(true)
      setMessage(err.message || 'An error occurred during generation.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="app-container">
      <div className="glass-panel">
        <h1 className="title">MatildeMartin Finances</h1>
        <p className="subtitle">Automated Report Generator</p>

        <div className="form-group">
          <label htmlFor="year-select">Select Year</label>
          <select 
            id="year-select"
            value={year} 
            onChange={(e) => setYear(e.target.value)}
            className="modern-select"
          >
            {years.map(y => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>
        </div>

        <div className="form-group">
          <label>Report Type</label>
          <div className="radio-group">
            <button 
              className={`type-btn ${reportType === 'IRPF' ? 'active' : ''}`}
              onClick={() => setReportType('IRPF')}
            >
              IRPF
            </button>
            <button 
              className={`type-btn ${reportType === 'Inmobiliaria' ? 'active' : ''}`}
              onClick={() => setReportType('Inmobiliaria')}
            >
              Inmobiliaria
            </button>
          </div>
        </div>

        <button 
          className={`generate-btn ${loading ? 'loading' : ''}`} 
          onClick={handleGenerate}
          disabled={loading}
        >
          {loading ? <span className="spinner"></span> : 'Generate Report'}
        </button>

        {message && (
          <div className={`status-message ${error ? 'error' : 'success'}`}>
            {message}
          </div>
        )}
      </div>
    </div>
  )
}

export default App
