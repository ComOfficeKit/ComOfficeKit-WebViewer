import { useState } from 'react';
import DocumentViewer from './DocumentViewer.jsx';

export default function App() {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState('Choose a document.');
  const [page, setPage] = useState(0);

  return (
    <main style={{ maxWidth: 1100, margin: '0 auto', padding: 16, font: '15px system-ui, sans-serif' }}>
      <h1 style={{ fontSize: 20 }}>ComOffice viewer in React</h1>
      <p>
        <input type="file" accept=".docx,.xlsx,.xlsm,.pptx,.csv,.tsv" onChange={e => setFile(e.target.files[0] ?? null)} />{' '}
        <span role="status">{status}{page > 0 ? `, page ${page}` : ''}</span>
      </p>
      <DocumentViewer
        file={file}
        style={{ height: '70vh' }}
        onLoad={info => { setStatus(`${info.pages} pages`); setPage(1); }}
        onError={problem => { setStatus(problem.message); setPage(0); }}
        onPageChange={setPage}
      />
    </main>
  );
}
