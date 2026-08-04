interface IconProps {
  size?: number;
  className?: string;
}

function SvgIcon({ size = 18, className = '', children }: IconProps & { children: React.ReactNode }) {
  return (
    <svg
      aria-hidden="true"
      className={className}
      fill="none"
      height={size}
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      viewBox="0 0 24 24"
      width={size}
    >
      {children}
    </svg>
  );
}

export function ShieldCheck(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M12 3 20 6v6c0 5-3.4 8.4-8 9-4.6-.6-8-4-8-9V6l8-3Z" />
      <path d="m9 12 2 2 4-5" />
    </SvgIcon>
  );
}

export function FileSpreadsheet(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z" />
      <path d="M14 2v6h6" />
      <path d="M8 13h8M8 17h8M10 11v8M14 11v8" />
    </SvgIcon>
  );
}

export function FileText(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z" />
      <path d="M14 2v6h6" />
      <path d="M8 13h8M8 17h6" />
    </SvgIcon>
  );
}

export function Play(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="m8 5 11 7-11 7Z" />
    </SvgIcon>
  );
}

export function ArrowLeft(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M19 12H5M12 19l-7-7 7-7" />
    </SvgIcon>
  );
}

export function ArrowRight(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M5 12h14M12 5l7 7-7 7" />
    </SvgIcon>
  );
}

export function CheckCircle2(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <circle cx="12" cy="12" r="9" />
      <path d="m8.5 12.5 2.2 2.2 4.8-5.2" />
    </SvgIcon>
  );
}

export function TriangleAlert(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="m12 3 10 18H2Z" />
      <path d="M12 9v5M12 18h.01" />
    </SvgIcon>
  );
}

export function Circle(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <circle cx="12" cy="12" r="9" />
    </SvgIcon>
  );
}

export function LoaderCircle(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M21 12a9 9 0 0 1-9 9" />
      <path d="M3 12a9 9 0 0 1 9-9" />
    </SvgIcon>
  );
}

export function FileSearch(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h7" />
      <path d="M14 2v6h6" />
      <circle cx="16" cy="16" r="3" />
      <path d="m18.5 18.5 2 2" />
    </SvgIcon>
  );
}

export function RotateCcw(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M3 12a9 9 0 1 0 3-6.7" />
      <path d="M3 4v6h6" />
    </SvgIcon>
  );
}

export function CheckSquare(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <rect height="16" rx="2" width="16" x="4" y="4" />
      <path d="m8.5 12.5 2 2 5-5" />
    </SvgIcon>
  );
}
