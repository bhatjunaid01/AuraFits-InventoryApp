import { StrictMode } from 'react'
import { Component } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.jsx'
import { installDbBridge } from './tauriDb.js'

installDbBridge()

class ErrorBoundary extends Component {
  constructor(props) {
    super(props)
    this.state = { error: null }
  }

  static getDerivedStateFromError(error) {
    return { error }
  }

  componentDidCatch(error, info) {
    console.error(error, info)
  }

  render() {
    if (this.state.error) {
      return (
        <div style={{ minHeight: '100vh', background: '#0A0A0A', color: '#E8E4D9', padding: 32, fontFamily: 'system-ui, sans-serif' }}>
          <h1 style={{ color: '#C9A84C', fontSize: 24, marginBottom: 12 }}>Aura Fits hit an error</h1>
          <p style={{ color: '#aaa', marginBottom: 16 }}>Please restart the app. Error details:</p>
          <pre style={{ whiteSpace: 'pre-wrap', background: '#141414', border: '1px solid #2a2a2a', borderRadius: 8, padding: 16 }}>
            {this.state.error.message}
          </pre>
        </div>
      )
    }

    return this.props.children
  }
}

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
)
