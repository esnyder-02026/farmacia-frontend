import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { Reportes } from './pages/Reportes';

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/reportes" element={<Reportes />} />
        <Route path="*" element={<Navigate to="/reportes" replace />} />
      </Routes>
    </BrowserRouter>
  );
}

export default App;
