import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles.css';

// 딥링크 데모: ?step=dashboard 등으로 특정 화면을 샘플 데이터로 바로 열 수 있음
const VALID_STEPS = ['upload', 'recognition', 'progress', 'dashboard', 'detail', 'notice'] as const;
type Step = (typeof VALID_STEPS)[number];
const requested = new URLSearchParams(window.location.search).get('step');
const initialStep = (VALID_STEPS as readonly string[]).includes(requested ?? '') ? (requested as Step) : 'upload';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App initialStep={initialStep} />
  </StrictMode>,
);
