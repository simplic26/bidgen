import type { ButtonHTMLAttributes, InputHTMLAttributes, ReactNode } from 'react';
import type { ValidationStatus } from '../types';
import { statusLabels } from '../utils/results';

interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'primary' | 'secondary' | 'ghost' | 'danger';
  icon?: ReactNode;
}

export function Button({ variant = 'secondary', icon, children, className = '', ...props }: ButtonProps) {
  return (
    <button className={`button button-${variant} ${className}`} {...props}>
      {icon}
      <span>{children}</span>
    </button>
  );
}

export function Panel({ children, className = '' }: { children: ReactNode; className?: string }) {
  return <section className={`panel ${className}`}>{children}</section>;
}

export function StatusBadge({ status }: { status: ValidationStatus }) {
  return <span className={`status-badge status-${status.toLowerCase()}`}>{statusLabels[status]}</span>;
}

export function SummaryCard({
  label,
  value,
  tone = 'neutral',
}: {
  label: string;
  value: string | number;
  tone?: 'neutral' | 'error' | 'review' | 'ok' | 'warning';
}) {
  return (
    <div className={`summary-card summary-${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

interface UploadPanelProps extends Omit<InputHTMLAttributes<HTMLInputElement>, 'title'> {
  label: string;
  title: ReactNode;
  description: string;
  action: string;
  meta?: ReactNode;
}

export function UploadPanel({ label, title, description, action, meta, ...props }: UploadPanelProps) {
  return (
    <div className="upload-panel">
      <div>
        <h3>{title}</h3>
        <p>{description}</p>
      </div>
      <label className="upload-action">
        {action}
        <input aria-label={label} type="file" {...props} />
      </label>
      {meta ? <div className="upload-meta">{meta}</div> : null}
    </div>
  );
}
