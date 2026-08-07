/* ============================================================
   SEGUROCONTROL – App Principal
   React 18 + Babel Standalone (sin bundler)
   ============================================================ */
const { useState, useEffect, useRef, useCallback, useMemo } = React;

// ─── Helpers ─────────────────────────────────────────────────
const formatMoney = (n) =>
  new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(n || 0);

const formatDate = (d) => {
  if (!d) return '—';
  const dt = typeof d === 'string' ? new Date(d + 'T12:00:00') : d;
  return dt.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' });
};

const calcNextDate = (current, forma) => {
  if (!current) return todayISO();
  const d = new Date(current + 'T12:00:00');
  if (forma === 'MENSUAL') d.setMonth(d.getMonth() + 1);
  else if (forma === 'TRIMESTRAL') d.setMonth(d.getMonth() + 3);
  else if (forma === 'SEMESTRAL') d.setMonth(d.getMonth() + 6);
  else if (forma === 'CONTADO') d.setFullYear(d.getFullYear() + 1);
  return d.toISOString().split('T')[0];
};

// ── Utils Excel Parsing ───────────────────────────────────────
const normalize = (s) => String(s).toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/[^a-z0-9 ]/g, ' ').trim();

const findCol = (row, fragments) => {
  const keys = Object.keys(row);
  for (const frag of fragments) {
    const nf = normalize(frag);
    const key = keys.find(k => normalize(k).includes(nf));
    if (key !== undefined && row[key] !== undefined && row[key] !== '') return row[key];
  }
  for (const frag of fragments) {
    const nf = normalize(frag);
    const key = keys.find(k => normalize(k).split(' ').some(w => w.startsWith(nf)));
    if (key !== undefined && row[key] !== undefined && row[key] !== '') return row[key];
  }
  return undefined;
};

const parseMonto = (val) => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  let s = String(val).replace(/[^0-9.,]/g, '');
  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } else if (s.includes(',')) {
    const parts = s.split(',');
    if (parts[1] && parts[1].length <= 2) s = s.replace(',', '.');
    else s = s.replace(/,/g, '');
  }
  return Number(s) || 0;
};

const getEffectiveMonto = (p) => {
  if (!p) return 0;
  if (p.formaPago !== 'CONTADO' && p.montoSubsecuente && p.fechaInicioVigencia && p.fechaPago > p.fechaInicioVigencia) {
    return Number(p.montoSubsecuente);
  }
  return Number(p.monto || 0);
};

const daysUntil = (dateStr) => {
  if (!dateStr) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(dateStr + 'T00:00:00');
  return Math.round((target - today) / 86400000);
};

// Para saber si una póliza está VERDADERAMENTE vencida respetando periodo de gracia (solo primer recibo)
const isExpiredEffective = (p) => {
  if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
  // Periodo de gracia solo aplica si es igual o posterior a la fechaPago actual (primer recibo)
  const validGracia = (p.periodoGracia && p.periodoGracia >= p.fechaPago) ? p.periodoGracia : null;
  const expiryDate = validGracia || p.fechaPago;
  const d = daysUntil(expiryDate);
  return d !== null && d < 0;
};

// Si está dentro de los 4 días antes de su fechaPago original o periodo de gracia (para recordatorio)
const isUpcomingReminder = (p) => {
  if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
  if (isExpiredEffective(p)) return false; // ya venció
  const validGracia = (p.periodoGracia && p.periodoGracia >= p.fechaPago) ? p.periodoGracia : null;
  const targetDate = validGracia || p.fechaPago;
  const d = daysUntil(targetDate);
  return d !== null && d >= 0 && d <= 4;
};

// Si está liquidada y se acerca su fecha de renovación (<= 15 días)
const getRenewalDate = (p) => {
  if (p.fechaInicioVigencia) {
    const d = new Date(p.fechaInicioVigencia + 'T00:00:00');
    d.setFullYear(d.getFullYear() + 1);
    return d.toISOString().split('T')[0];
  }
  return p.fechaPago;
};

const isUpcomingRenewal = (p) => {
  if (p.estatus !== 'LIQUIDADO') return false;
  const d = daysUntil(getRenewalDate(p));
  return d !== null && d <= 15;
};

const todayISO = () => {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const generateId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

// Hook para cerrar modales al presionar la tecla Esc
function useEscapeKey(onClose) {
  useEffect(() => {
    if (!onClose) return;
    const handleKeyDown = (e) => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [onClose]);
}

// ─── Datos de muestra ─────────────────────────────────────────
// Sin datos de ejemplo — la app inicia vacía para importar el Excel real
const SAMPLE_DATA = [];

// ─── Plantillas de Mensajes Predeterminadas ───────────────────
const DEFAULT_TEMPLATES = {
  whatsapp: `Estimado(a) asegurado(a) *{nombre}* 👋

Le contactamos de parte de *PRE & PRO CONSULTORES* para recordarle que su póliza *{poliza}* {estado_vencimiento}.

📋 *Detalles de pago:*
• Unidad: {bien}
• Monto a pagar: *{monto}*

Le pedimos de favor realizar su pago antes de la fecha límite para mantener su cobertura vigente 🛡️

Si ya realizó su pago, le pedimos de favor nos envíe su comprobante de pago para su respectiva aplicación 📄

¡Gracias por su confianza! 😊
*PRE & PRO CONSULTORES*`,

  email_asunto: 'Recordatorio de Pago – Póliza {poliza} | PRE & PRO CONSULTORES',
  email_cuerpo: `Estimado(a) asegurado(a) {nombre},

Por medio del presente correo le recordamos amablemente que su póliza de seguro {poliza} {estado_vencimiento}.

DETALLES DE SU PÓLIZA:
━━━━━━━━━━━━━━━━━━━━━━━
• Póliza N°: {poliza}
• Unidad: {bien}
• Monto a pagar: {monto}

Para mantener la vigencia de su cobertura, le solicitamos realizar el pago antes de la fecha indicada.

Si ya realizó su pago, le pedimos nos envíe su comprobante para su respectiva aplicación.

Si tiene alguna duda, no dude en contactarnos.

Atentamente,
PRE & PRO CONSULTORES`,
};

// ─── Contexto Global (Toast) ──────────────────────────────────
function useToast() {
  const [toasts, setToasts] = useState([]);
  const toast = useCallback((msg, type = 'info') => {
    const id = generateId();
    setToasts(t => [...t, { id, msg, type }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 3500);
  }, []);
  return { toasts, toast };
}

// ─── SVG Icons ────────────────────────────────────────────────
const Icons = {
  Dashboard: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/>
      <rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/>
    </svg>
  ),
  Policies: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/>
      <line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/>
    </svg>
  ),
  Alert: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/>
      <path d="M13.73 21a2 2 0 0 1-3.46 0"/>
    </svg>
  ),
  Templates: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
    </svg>
  ),
  Import: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
    </svg>
  ),
  Plus: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" width="16" height="16">
      <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
    </svg>
  ),
  Edit: () => <span style={{fontSize:14}}>✏️</span>,
  Delete: () => <span style={{fontSize:14}}>🗑️</span>,
  WhatsApp: () => <span style={{fontSize:15}}>💬</span>,
  Email: () => <span style={{fontSize:14}}>📧</span>,
  Check: () => <span style={{fontSize:14}}>✅</span>,
  Eye: () => <span style={{fontSize:14}}>👁️</span>,
  Search: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="search-icon">
      <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
    </svg>
  ),
  Sort: ({ dir }) => (
    <span className={`sort-icon ${dir ? 'active' : ''}`}>
      {dir === 'asc' ? '↑' : dir === 'desc' ? '↓' : '↕'}
    </span>
  ),
  Close: () => <span>×</span>,
  Export: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16" strokeLinecap="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
    </svg>
  ),
  Shield: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20" strokeLinecap="round">
      <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
    </svg>
  ),
  Receipt: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1Z"/>
      <path d="M16 8h-6"/>
      <path d="M16 12h-8"/>
      <path d="M16 16h-8"/>
    </svg>
  ),
  Heart: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"/>
    </svg>
  ),
  Briefcase: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <rect x="2" y="7" width="20" height="14" rx="2" ry="2"/>
      <path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/>
    </svg>
  ),
  Home: () => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20" strokeLinecap="round" strokeLinejoin="round" className="nav-icon">
      <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>
      <polyline points="9 22 9 12 15 12 15 22"/>
    </svg>
  ),
};

// ─── Toast Component ──────────────────────────────────────────
function ToastContainer({ toasts }) {
  const icons = { success: '✅', error: '❌', info: 'ℹ️', warning: '⚠️' };
  return (
    <div className="toast-container">
      {toasts.map(t => (
        <div key={t.id} className={`toast toast-${t.type}`}>
          <span className="toast-icon">{icons[t.type]}</span>
          <span className="toast-msg">{t.msg}</span>
        </div>
      ))}
    </div>
  );
}

// ─── Status Badge ─────────────────────────────────────────────
function StatusBadge({ policy }) {
  const rawEstatus = policy.estatus;
  // Si la póliza está pendiente pero ya pasó su fecha (o periodo de gracia), mostrar como VENCIDO
  const estatus = (rawEstatus === 'PENDIENTE' && isExpiredEffective(policy)) ? 'VENCIDO' : rawEstatus;
  const map = {
    PAGADO: 'pagado', PENDIENTE: 'pendiente',
    VENCIDO: 'vencido', CANCELADO: 'cancelado', LIQUIDADO: 'liquidado'
  };
  const cls = map[estatus] || 'pendiente';
  
  let tooltip = '';
  if (estatus === 'LIQUIDADO') {
    const renewal = getRenewalDate(policy);
    if (renewal) {
      tooltip = `Vigencia hasta: ${formatDate(renewal)}`;
    }
  }

  return (
    <span className={`badge badge-${cls}`} title={tooltip}>
      <span className="badge-dot" />
      {estatus}
    </span>
  );
}

// ─── Agent Badge ──────────────────────────────────────────────
function AgentBadge({ policy, agente }) {
  const ag = agente || policy?.agente || policy?.aseguradora || 'SIN CLAVE';
  return (
    <span className={`agent-badge agent-${ag?.toLowerCase()}`}>
      {ag === 'DANIEL' ? '👤' : '👥'} {ag}
    </span>
  );
}

// ─── Ramo Badge ──────────────────────────────────────────────
function RamoBadge({ policy }) {
  if (policy?._isCaro) return <span style={{fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'rgba(139, 92, 246, 0.2)', color: '#a78bfa', fontWeight: 'bold'}}>AUTOS QUALITAS CARO</span>;
  if (policy?._isGmm) return <span style={{fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'rgba(16, 185, 129, 0.2)', color: '#34d399', fontWeight: 'bold'}}>GMM</span>;
  if (policy?._isAutos) return <span style={{fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'rgba(245, 158, 11, 0.2)', color: '#fbbf24', fontWeight: 'bold'}}>AUTOS</span>;
  if (policy?._isVida) return <span style={{fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'rgba(236, 72, 153, 0.2)', color: '#f472b6', fontWeight: 'bold'}}>VIDA</span>;
  if (policy?._isDanos) return <span style={{fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'rgba(14, 165, 233, 0.2)', color: '#38bdf8', fontWeight: 'bold'}}>DAÑOS</span>;
  if (policy?._isHogar) return <span style={{fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'rgba(168, 85, 247, 0.2)', color: '#c084fc', fontWeight: 'bold'}}>HOGAR</span>;
  return <span style={{fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'rgba(23, 113, 197, 0.2)', color: '#60a5fa', fontWeight: 'bold'}}>AUTOS QUALITAS DANI</span>;
}

// ─── Modal de Resumen de Póliza (Doble Clic) ─────────────────
function PolicySummaryModal({ policy: p, onClose }) {
  useEscapeKey(onClose);
  if (!p) return null;
  const bienLabel = p._isVida ? 'Producto' : p._isGmm ? 'Plan' : p._isHogar ? 'Inmueble' : p._isDanos ? 'Bien Asegurado' : 'Unidad / Vehículo';
  const ramoLabel = p._isCaro ? 'Autos Qualitas Caro' : p._isGmm ? 'GMM' : p._isAutos ? 'Autos (Otras)' : p._isVida ? 'Vida' : p._isDanos ? 'Daños' : p._isHogar ? 'Hogar' : 'Autos Qualitas Dani';

  return (
    <div className="modal-overlay" onClick={onClose} style={{zIndex: 999999}}>
      <div className="modal" style={{maxWidth: 440, padding: 0, overflow: 'hidden'}} onClick={e => e.stopPropagation()}>
        <div className="modal-header" style={{background: 'var(--bg-secondary)', padding: '16px 20px', borderBottom: '1px solid var(--border)'}}>
          <div style={{display: 'flex', alignItems: 'center', gap: 10}}>
            <span style={{fontSize: 20}}>📄</span>
            <div>
              <h3 style={{margin: 0, fontSize: 16, color: 'var(--text-primary)'}}>Resumen Póliza #{p.poliza}</h3>
              <span style={{fontSize: 11, color: '#38bdf8', fontWeight: 600}}>{ramoLabel}</span>
            </div>
          </div>
          <button className="modal-close" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body" style={{padding: 20, display: 'flex', flexDirection: 'column', gap: 10, fontSize: 13}}>
          <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
            <span style={{color: 'var(--text-muted)'}}>Asegurado:</span>
            <strong style={{color: 'var(--text-primary)', textAlign: 'right'}}>{p.nombre}</strong>
          </div>
          {p.perteneceA && (
            <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
              <span style={{color: 'var(--text-muted)'}}>Esta póliza pertenece a:</span>
              <strong style={{color: '#38bdf8', textAlign: 'right'}}>👤 {p.perteneceA}</strong>
            </div>
          )}
          {p.bien && (
            <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
              <span style={{color: 'var(--text-muted)'}}>{bienLabel}:</span>
              <strong style={{color: 'var(--text-primary)', textAlign: 'right', maxWidth: 220, wordBreak: 'break-word'}}>{p.bien}</strong>
            </div>
          )}
          <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
            <span style={{color: 'var(--text-muted)'}}>Agente:</span>
            <strong style={{color: 'var(--text-primary)'}}>{p.agente || 'N/A'}</strong>
          </div>
          {(p._isGmm || p._isAutos || p._isVida || p._isDanos || p._isHogar) && p.aseguradora && (
            <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
              <span style={{color: 'var(--text-muted)'}}>Aseguradora:</span>
              <strong style={{color: 'var(--text-primary)'}}>{p.aseguradora}</strong>
            </div>
          )}
          <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
            <span style={{color: 'var(--text-muted)'}}>Forma de Pago:</span>
            <span className="forma-badge">{p.formaPago}</span>
          </div>
          <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
            <span style={{color: 'var(--text-muted)'}}>Próxima Fecha Pago:</span>
            <strong style={{color: '#fbbf24'}}>{formatDate(p.fechaPago)}</strong>
          </div>
          <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
            <span style={{color: 'var(--text-muted)'}}>Monto:</span>
            <strong style={{color: '#34d399', fontSize: 15}}>{formatMoney(p.monto)}</strong>
          </div>
          <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
            <span style={{color: 'var(--text-muted)'}}>Estatus:</span>
            <StatusBadge policy={p} />
          </div>
          {p.fechaInicioVigencia && (
            <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
              <span style={{color: 'var(--text-muted)'}}>Inicio Vigencia:</span>
              <span>{formatDate(p.fechaInicioVigencia)}</span>
            </div>
          )}
          {p.periodoGracia && (
            <div style={{display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: 6}}>
              <span style={{color: 'var(--text-muted)'}}>Gracia activa hasta:</span>
              <span style={{color: '#818cf8', fontWeight: 600}}>📌 {formatDate(p.periodoGracia)}</span>
            </div>
          )}
          {p.notas && (
            <div style={{marginTop: 6, padding: 10, borderRadius: 8, background: 'var(--bg-secondary)', border: '1px dashed var(--border)', fontSize: 12, color: 'var(--text-secondary)'}}>
              <strong style={{color: 'var(--text-primary)'}}>📝 Notas Internas:</strong>
              <div style={{marginTop: 4}}>{p.notas}</div>
            </div>
          )}
        </div>
        <div className="modal-footer" style={{background: 'var(--bg-secondary)', padding: '12px 20px', display: 'flex', justifyContent: 'flex-end'}}>
          <button className="btn btn-primary btn-sm" onClick={onClose}>Entendido</button>
        </div>
      </div>
    </div>
  );
}

function DateCell({ dateStr, estatus, periodoGracia }) {
  // Solo considerar periodoGracia si es posterior o igual a la fechaPago (primer recibo)
  const activeGracia = (periodoGracia && periodoGracia >= dateStr) ? periodoGracia : null;

  const days = daysUntil(dateStr);
  const daysGracia = activeGracia ? daysUntil(activeGracia) : null;

  if (!dateStr) return <span className="text-muted">—</span>;
  if (estatus === 'LIQUIDADO') {
    return (
      <div style={{display:'flex', flexDirection:'column'}}>
        <span className="date-normal" style={{color: 'var(--text-primary)'}}>{formatDate(dateStr)}</span>
        <span style={{fontSize:10, color:'var(--accent-blue-light)'}}>Último pago</span>
      </div>
    );
  }
  if (estatus === 'PAGADO' || estatus === 'CANCELADO') {
    return <span className="date-normal">{formatDate(dateStr)}</span>;
  }
  if (days === null) return <span>{formatDate(dateStr)}</span>;

  const graciaChip = activeGracia ? (
    <span className="urgency-chip" style={{background:'rgba(99,102,241,0.15)', color:'#818cf8', marginLeft: 4, whiteSpace:'nowrap'}}>
      📌 Gracia: {formatDate(activeGracia)}
    </span>
  ) : null;

  // Dentro de periodo de gracia (ya pasó fechaPago pero aún no vence gracia)
  if (activeGracia && days < 0 && daysGracia !== null && daysGracia >= 0) {
    return (
      <span className="date-soon" title={`Período de gracia hasta ${formatDate(activeGracia)}`}>
        {formatDate(dateStr)}
        {graciaChip}
      </span>
    );
  }
  // Vencido efectivo (pasó periodo de gracia o fechaPago si no hay gracia)
  if (days < 0 && (!activeGracia || daysGracia < 0)) return (
    <span className="date-urgent" title={`Vencido hace ${Math.abs(days)} día(s)`}>
      {formatDate(dateStr)} <span className="urgency-chip">⚠ {Math.abs(days)}d venc.</span>
    </span>
  );
  // Recordatorio 4 días antes
  if (days <= 4) return (
    <span className="date-urgent" title={`Vence en ${days} día(s)`}>
      {formatDate(dateStr)} <span className="urgency-chip">🔴 {days}d</span>
      {graciaChip}
    </span>
  );
  if (days <= 10) return (
    <span className="date-soon" title={`Vence en ${days} días`}>
      {formatDate(dateStr)} <span className="urgency-chip" style={{background:'rgba(245,158,11,0.15)',color:'#fcd34d'}}>🟡 {days}d</span>
      {graciaChip}
    </span>
  );
  return (
    <span className="date-normal">
      {formatDate(dateStr)}
      {graciaChip}
    </span>
  );
}

// ─── Fill Template ────────────────────────────────────────────
function fillTemplate(tpl, policy, isWA = false) {
  if (!tpl) return '';
  const isVencido = policy.estatus === 'VENCIDO' || isExpiredEffective(policy);
  const fDate = formatDate(policy.fechaPago);

  const dateFormatted = isWA ? `*${fDate}*` : fDate;
  const estadoVencimientoText = isVencido 
    ? `venció el ${dateFormatted}` 
    : `está próxima a vencer el ${dateFormatted}`;

  let processedTpl = tpl
    .replace('vence el próximo *{fechaPago}*', `está próxima a vencer el *${fDate}*`)
    .replace('vence el próximo {fechaPago}', `está próxima a vencer el ${fDate}`)
    .replace('está próxima a vencer el *{fechaPago}*', `está próxima a vencer el *${fDate}*`)
    .replace('está próxima a vencer el {fechaPago}', `está próxima a vencer el ${fDate}`)
    .replace('tiene programado su próximo vencimiento el día {fechaPago}', `está próxima a vencer el ${fDate}`);

  return processedTpl
    .replace(/{estado_vencimiento}/g, estadoVencimientoText)
    .replace(/{nombre}/g, policy.nombre || '')
    .replace(/{poliza}/g, policy.poliza || '')
    .replace(/{bien}/g, policy.bien || '')
    .replace(/{monto}/g, formatMoney(policy.monto))
    .replace(/{formaPago}/g, policy.formaPago || '')
    .replace(/{agente}/g, policy.agente || '')
    .replace(/{fechaPago}/g, formatDate(policy.fechaPago))
    .replace(/{correo}/g, policy.correo || '')
    .replace(/{telefono}/g, policy.telefono || '');
}

// ─── Field wrapper (fuera del modal para evitar re-montar inputs) ───
function FieldGroup({ label, id, required, error, children }) {
  return (
    <div className="form-group">
      <label className="form-label" htmlFor={id}>
        {label}{required && <span className="required">*</span>}
      </label>
      {children}
      {error && <span style={{fontSize:11, color:'var(--accent-red)'}}>{error}</span>}
    </div>
  );
}

// ─── Modal: Nueva / Editar Póliza ────────────────────────────
function PolicyModal({ policy, onSave, onClose, toast, agentOptions, isGmm = false, isAutos = false, isVida = false, isDanos = false, isHogar = false }) {
  useEscapeKey(onClose);
  const gmmAseguradoras = ['AXA', 'MAPFRE', 'GNP', 'CHUBB', 'SURA', 'PLAN SEGUROS', 'ZURICH', 'QUALITAS', 'AIG', 'BANORTE', 'OTRO'];
  const autosAseguradoras = ['ZURICH', 'AXA', 'HDI', 'GNP', 'QUALITAS', 'AIG', 'MAPFRE', 'BANORTE', 'ANA', 'SEGUROS ARGO', 'OTRO'];
  const defaultOpts = (isGmm || isAutos || isVida || isDanos || isHogar) 
    ? ['DANIEL', 'OTRO'] 
    : ['DANIEL', 'MARTIN'];
  const optsList = agentOptions || defaultOpts;

  const isEdit = !!policy?.id;
  const [form, setForm] = useState(() => {
    if (policy) {
      const isKnown = optsList.includes(policy.agente);
      return {
        ...policy,
        agente: isKnown ? policy.agente : 'OTRO',
        agenteCustom: isKnown ? '' : (policy.agente || ''),
        aseguradora: policy.aseguradora || (isGmm ? gmmAseguradoras[0] : autosAseguradoras[0]),
        perteneceA: policy.perteneceA || ''
      };
    }
    return {
      nombre: '', poliza: '', bien: '', formaPago: 'MENSUAL',
      agente: optsList[0], fechaPago: todayISO(), monto: '',
      estatus: 'PENDIENTE', correo: '', telefono: '', notas: '',
      periodoGracia: '', fechaInicioVigencia: '',
      aseguradora: isGmm ? gmmAseguradoras[0] : autosAseguradoras[0],
      agenteCustom: '',
      perteneceA: ''
    };
  });
  const [errors, setErrors] = useState({});

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const validate = () => {
    const e = {};
    if (!form.nombre.trim()) e.nombre = 'Requerido';
    if (!form.poliza.trim()) e.poliza = 'Requerido';
    if (!form.fechaPago) e.fechaPago = 'Requerido';
    if (!form.monto || isNaN(Number(form.monto))) e.monto = 'Monto inválido';
    if (form.formaPago !== 'CONTADO') {
      if (!form.montoSubsecuente || isNaN(Number(form.montoSubsecuente))) e.montoSubsecuente = 'Requerido';
    }
    if (form.correo && !/\S+@\S+\.\S+/.test(form.correo)) e.correo = 'Correo inválido';
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const handleSave = () => {
    if (!validate()) { toast('Por favor corrige los errores', 'error'); return; }
    const isSubsequent = form.formaPago !== 'CONTADO' && form.montoSubsecuente && form.fechaInicioVigencia && form.fechaPago > form.fechaInicioVigencia;
    const activeMonto = isSubsequent ? Number(form.montoSubsecuente) : Number(form.monto);
    const saved = {
      ...form,
      id: form.id || generateId(),
      monto: activeMonto,
      ...(form.formaPago !== 'CONTADO' && form.montoSubsecuente ? { montoSubsecuente: Number(form.montoSubsecuente) } : {}),
      ...((isAutos || isGmm || isVida || isDanos || isHogar) && form.agente === 'OTRO' && form.agenteCustom ? { agente: form.agenteCustom } : {})
    };
    onSave(saved);
    toast(isEdit ? 'Póliza actualizada ✅' : 'Póliza registrada ✅', 'success');
    onClose();
  };

  return (
    <div className="modal-overlay">
      <div className="modal modal-wide" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>{isEdit ? '✏️ Editar Póliza' : '➕ Nueva Póliza'}</h2>
          <button className="modal-close" onClick={onClose}><Icons.Close /></button>
        </div>
        <div className="modal-body">
          <div className="form-grid">
            <FieldGroup label="Nombre del Asegurado" id="nombre" required error={errors.nombre}>
              <input id="nombre" className={`input ${errors.nombre ? 'input-error' : ''}`}
                value={form.nombre} onChange={e => set('nombre', e.target.value)}
                placeholder="Nombre completo" />
            </FieldGroup>
            <FieldGroup label="Número de Póliza" id="poliza" required error={errors.poliza}>
              <input id="poliza" className="input" value={form.poliza}
                onChange={e => set('poliza', e.target.value)} placeholder="POL-2024-000" />
            </FieldGroup>
            <div className="form-group full-width">
              <label className="form-label">{isVida ? 'Producto' : isGmm ? 'Plan' : isHogar ? 'Ubicación / Inmueble' : isDanos ? 'Bien Asegurado' : 'Vehículo / Bien Asegurado'}</label>
              <input className="input" value={form.bien}
                onChange={e => set('bien', e.target.value)}
                placeholder={isVida ? 'Ej: Orvi 99, Segubeca, Vida Individual...' : isGmm ? 'Ej: Salud Global Esencial...' : isHogar ? 'Ej: Casa Habitación Residencial...' : isDanos ? 'Ej: Edificio Industrial, Maquinaria...' : 'Ej: Toyota Corolla 2022 – ABC-123-X'} />
            </div>
            <div className="form-group full-width">
              <label className="form-label" htmlFor="perteneceA">Esta póliza le pertenece a...</label>
              <input id="perteneceA" className="input" value={form.perteneceA || ''}
                onChange={e => set('perteneceA', e.target.value)}
                placeholder="Nombre de la persona o titular a quien pertenece..." />
            </div>
            <FieldGroup label="Forma de Pago" id="formaPago">
              <select id="formaPago" className="select" value={form.formaPago}
                onChange={e => set('formaPago', e.target.value)}>
                <option value="CONTADO">CONTADO</option>
                <option value="MENSUAL">MENSUAL</option>
                <option value="TRIMESTRAL">TRIMESTRAL</option>
                <option value="SEMESTRAL">SEMESTRAL</option>
              </select>
            </FieldGroup>
            <FieldGroup label="Clave de Agente" id="agente">
              <select id="agente" className="select" value={form.agente}
                onChange={e => set('agente', e.target.value)}>
                {Array.from(new Set([...optsList, ...(form.agente && !optsList.includes(form.agente) && form.agente !== 'OTRO' ? [form.agente] : [])])).map(opt => (
                  <option key={opt} value={opt}>{opt}</option>
                ))}
              </select>
              {(isAutos || isGmm || isVida || isDanos || isHogar) && form.agente === 'OTRO' && (
                <input className="input" style={{marginTop: 8}} value={form.agenteCustom || ''}
                  onChange={e => set('agenteCustom', e.target.value)}
                  placeholder="Escribe el nombre del agente..." />
              )}
            </FieldGroup>
            {(isAutos || isGmm || isVida || isDanos || isHogar) && (
              <FieldGroup label="Aseguradora" id="aseguradora">
                <select id="aseguradora" className="select" 
                  value={form.aseguradora || (isGmm ? gmmAseguradoras[0] : autosAseguradoras[0])}
                  onChange={e => set('aseguradora', e.target.value)}>
                  {(isGmm ? gmmAseguradoras : autosAseguradoras).map(opt => (
                    <option key={opt} value={opt}>{opt}</option>
                  ))}
                </select>
              </FieldGroup>
            )}
            <FieldGroup label="Inicio de Vigencia" id="fechaInicioVigencia">
              <input id="fechaInicioVigencia" type="date" className="input" value={form.fechaInicioVigencia || ''}
                onChange={e => set('fechaInicioVigencia', e.target.value)} />
            </FieldGroup>
            <FieldGroup label="Próxima Fecha de Pago" id="fechaPago" required error={errors.fechaPago}>
              <input id="fechaPago" type="date" className="input" value={form.fechaPago}
                onChange={e => set('fechaPago', e.target.value)} />
            </FieldGroup>
            <FieldGroup label="Periodo de Gracia (solo primer recibo)" id="periodoGracia">
              <input id="periodoGracia" type="date" className="input" value={form.periodoGracia || ''}
                onChange={e => set('periodoGracia', e.target.value)}
                title="Fecha hasta la cual la póliza sigue activa aunque ya pasó la fecha de pago" />
              {form.periodoGracia && (
                <span style={{fontSize:11, color:'#818cf8', marginTop:2}}>
                  📌 No vencerá hasta el {formatDate(form.periodoGracia)}
                </span>
              )}
            </FieldGroup>
            <FieldGroup label="Monto 1er Recibo ($)" id="monto" required error={errors.monto}>
              <input id="monto" type="number" className="input" value={form.monto}
                onChange={e => set('monto', e.target.value)} placeholder="0.00" min="0" step="0.01" />
            </FieldGroup>
            {form.formaPago !== 'CONTADO' && (
              <FieldGroup label="Monto Subsecuente ($) (2º recibo en adelante)" id="montoSubsecuente" required error={errors.montoSubsecuente}>
                <input id="montoSubsecuente" type="number" className={`input ${errors.montoSubsecuente ? 'input-error' : ''}`}
                  value={form.montoSubsecuente || ''}
                  onChange={e => set('montoSubsecuente', e.target.value)} placeholder="Ej: 1880.82" min="0" step="0.01" />
                <span style={{fontSize:11, color:'var(--text-muted)', marginTop:2}}>
                  📌 El sistema cobrará esta cantidad automáticamente a partir del 2º pago.
                </span>
              </FieldGroup>
            )}
            <FieldGroup label="Estatus" id="estatus">
              <select id="estatus" className="select" value={form.estatus}
                onChange={e => set('estatus', e.target.value)}>
                <option value="PENDIENTE">PENDIENTE</option>
                <option value="PAGADO">PAGADO</option>
                <option value="VENCIDO">VENCIDO</option>
                <option value="CANCELADO">CANCELADO</option>
                <option value="LIQUIDADO">LIQUIDADO</option>
              </select>
            </FieldGroup>
            <FieldGroup label="Correo Electrónico" id="correo" error={errors.correo}>
              <input id="correo" type="email" className="input" value={form.correo}
                onChange={e => set('correo', e.target.value)} placeholder="ejemplo@correo.com" />
            </FieldGroup>
            <FieldGroup label="Teléfono / WhatsApp 1 (con lada)" id="telefono">
              <div style={{display:'flex', gap:8}}>
                <select className="select" style={{width:90}}
                  value={form.lada || '+52'} onChange={e => set('lada', e.target.value)}>
                  <option value="+52">🇲🇽 +52</option>
                  <option value="+1">🇺🇸 +1</option>
                  <option value="+34">🇪🇸 +34</option>
                </select>
                <input id="telefono" type="tel" className="input" value={form.telefono}
                  onChange={e => set('telefono', e.target.value.replace(/\D/g, ''))}
                  placeholder="10 dígitos" maxLength={10} />
              </div>
            </FieldGroup>
            <FieldGroup label="Teléfono / WhatsApp 2 (Opcional)" id="telefono2">
              <div style={{display:'flex', gap:8}}>
                <select className="select" style={{width:90}}
                  value={form.lada2 || '+52'} onChange={e => set('lada2', e.target.value)}>
                  <option value="+52">🇲🇽 +52</option>
                  <option value="+1">🇺🇸 +1</option>
                  <option value="+34">🇪🇸 +34</option>
                </select>
                <input id="telefono2" type="tel" className="input" value={form.telefono2 || ''}
                  onChange={e => set('telefono2', e.target.value.replace(/\D/g, ''))}
                  placeholder="2do número (opcional)" maxLength={10} />
              </div>
            </FieldGroup>
            <div className="form-group full-width">
              <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom: 6}}>
                <label className="form-label" style={{margin:0}}>Notas Internas</label>
                <button
                  type="button"
                  className="btn btn-ghost btn-sm"
                  style={{fontSize: 11, padding: '2px 8px', color: 'var(--accent-blue)', background: 'rgba(23,113,197,0.1)', border: '1px solid rgba(23,113,197,0.2)', borderRadius: 6}}
                  onClick={() => {
                    const texto = 'Ya mande recordatorio';
                    set('notas', form.notas ? (form.notas.includes(texto) ? form.notas : form.notas + '\n' + texto) : texto);
                  }}
                >
                  📌 + Ya mande recordatorio
                </button>
              </div>
              <textarea className="input" rows={3} value={form.notas}
                onChange={e => set('notas', e.target.value)}
                placeholder="Observaciones, acuerdos, historial..." />
            </div>
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-outline" onClick={onClose}>Cancelar</button>
          <button className="btn btn-primary" onClick={handleSave}>
            {isEdit ? '💾 Guardar Cambios' : '➕ Registrar Póliza'}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Modal: Marcar como Pagado ────────────────────────────────
function MarkPaidModal({ policy, onConfirm, onClose, toast }) {
  useEscapeKey(onClose);
  const nextDate = policy.formaPago !== 'CONTADO'
    ? calcNextDate(policy.fechaPago, policy.formaPago) : null;
  const [comprobante, setComprobante] = useState(null);
  const [nextMonto, setNextMonto] = useState(policy.montoSubsecuente || policy.monto || '');

  let isLastPayment = false;
  if (policy.formaPago !== 'CONTADO' && policy.fechaInicioVigencia && nextDate) {
    const startD = new Date(policy.fechaInicioVigencia + 'T00:00:00');
    const endOfCoverage = new Date(startD);
    endOfCoverage.setFullYear(endOfCoverage.getFullYear() + 1);
    
    const nextD = new Date(nextDate + 'T00:00:00');
    if (nextD >= endOfCoverage) {
      isLastPayment = true;
    }
  }

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setComprobante(reader.result);
      };
      reader.readAsDataURL(file);
    }
  };

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="modal-header">
          <h2>✅ Registrar Pago</h2>
          <button className="modal-close" onClick={onClose}><Icons.Close /></button>
        </div>
        <div className="modal-body">
          <div className="info-grid">
            <div className="info-card">
              <div className="info-card-label">Asegurado</div>
              <div className="info-card-value">{policy.nombre}</div>
            </div>
            <div className="info-card">
              <div className="info-card-label">Póliza</div>
              <div className="info-card-value">{policy.poliza}</div>
            </div>
            <div className="info-card">
              <div className="info-card-label">Monto Pagado</div>
              <div className="info-card-value" style={{color:'var(--accent-green)'}}>{formatMoney(policy.monto)}</div>
            </div>
            <div className="info-card">
              <div className="info-card-label">Forma de Pago</div>
              <div className="info-card-value">{policy.formaPago}</div>
            </div>
          </div>

          {policy.formaPago === 'CONTADO' || isLastPayment ? (
            <div style={{
              background:'rgba(139,92,246,0.1)', border:'1px solid rgba(139,92,246,0.3)',
              borderRadius:'var(--radius-md)', padding:16, marginTop:12
            }}>
              <p style={{fontSize:14, color:'#c4b5fd'}}>
                🎉 <strong>{policy.formaPago === 'CONTADO' ? 'Póliza de CONTADO' : 'Último pago del ciclo'}</strong> — Al confirmar, la póliza quedará marcada como <strong>LIQUIDADA</strong> hasta su fecha de renovación anual.
              </p>
            </div>
          ) : (
            <div style={{
              background:'rgba(16,185,129,0.08)', border:'1px solid rgba(16,185,129,0.25)',
              borderRadius:'var(--radius-md)', padding:16, marginTop:12
            }}>
              <p style={{fontSize:13, color:'var(--text-secondary)', marginBottom:8}}>
                🔄 <strong>Re-agendamiento automático</strong>
              </p>
              <p style={{fontSize:14, color:'#6ee7b7'}}>
                La próxima fecha de pago se calculará automáticamente:
              </p>
              <p style={{fontSize:18, fontWeight:800, color:'var(--accent-green)', marginTop:8}}>
                📅 {formatDate(nextDate)}
              </p>
              <p style={{fontSize:12, color:'var(--text-muted)', marginTop:4}}>
                El estatus regresará a <strong>PENDIENTE</strong> para el siguiente ciclo.
              </p>
              <div style={{marginTop:12, paddingTop:12, borderTop:'1px solid rgba(255,255,255,0.1)'}}>
                <label className="form-label" style={{fontSize:12, marginBottom:4, color:'var(--text-primary)'}}>
                  Monto del Siguiente Recibo ($)
                </label>
                <input type="number" className="input" value={nextMonto} onChange={e => setNextMonto(e.target.value)} placeholder="0.00" step="0.01" />
                <span style={{fontSize:11, color:'var(--text-muted)', marginTop:4, display:'block'}}>
                  💡 Ajusta este monto si los recibos subsecuentes cambian respecto al 1er pago.
                </span>
              </div>
            </div>
          )}

          <div style={{marginTop: 16}}>
            <label className="form-label" style={{display:'block', marginBottom:8}}>Comprobante de pago (Opcional)</label>
            <input type="file" accept="image/*,.pdf" className="input" onChange={handleFileChange} />
            {comprobante && (
              <div style={{marginTop: 8, fontSize: 12, color: 'var(--accent-green)'}}>
                ✅ Archivo adjunto listo para guardar
              </div>
            )}
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-outline" onClick={onClose}>Cancelar</button>
          <button className="btn btn-success" onClick={() => {
            onConfirm(policy, nextDate, comprobante, isLastPayment, nextMonto);
            toast('Pago registrado y fecha actualizada ✅', 'success');
            onClose();
          }}>
            ✅ Confirmar Pago
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Modal: Calendario Interactivo con Indicadores ────────────
function CustomCalendarPickerModal({ policies, caroPolicies, onClose, onSelectDate }) {
  useEscapeKey(onClose);
  const [currentDate, setCurrentDate] = useState(new Date());

  const year = currentDate.getFullYear();
  const month = currentDate.getMonth();

  const monthNames = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];

  const formatYYYYMMDD = (y, m, d) => {
    const mm = String(m + 1).padStart(2, '0');
    const dd = String(d).padStart(2, '0');
    return `${y}-${mm}-${dd}`;
  };

  const allPolicies = useMemo(() => [...policies, ...caroPolicies], [policies, caroPolicies]);

  const dateIndicators = useMemo(() => {
    const map = {};
    allPolicies.forEach(p => {
      if (p.estatus === 'CANCELADO') return;
      const isPaid = p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO';
      const isExpired = isExpiredEffective(p);

      const targetDate = isPaid ? (p.fechaUltimoPago || p.fechaPago) : p.fechaPago;

      if (targetDate) {
        if (!map[targetDate]) map[targetDate] = { pending: 0, expired: 0, paid: 0 };
        if (isPaid) {
          map[targetDate].paid += 1;
        } else if (isExpired) {
          map[targetDate].expired += 1;
        } else {
          map[targetDate].pending += 1;
        }
      }
    });
    return map;
  }, [allPolicies]);

  const firstDayIndex = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  const daysGrid = [];
  const prevMonthDays = new Date(year, month, 0).getDate();
  for (let i = firstDayIndex - 1; i >= 0; i--) {
    daysGrid.push({ day: prevMonthDays - i, currentMonth: false });
  }
  for (let d = 1; d <= daysInMonth; d++) {
    daysGrid.push({ day: d, currentMonth: true });
  }
  const remaining = 42 - daysGrid.length;
  const nextPadding = remaining < 7 ? remaining : remaining - 7;
  for (let i = 1; i <= nextPadding; i++) {
    daysGrid.push({ day: i, currentMonth: false });
  }

  const prevMonth = () => setCurrentDate(new Date(year, month - 1, 1));
  const nextMonth = () => setCurrentDate(new Date(year, month + 1, 1));
  const goToday = () => setCurrentDate(new Date());

  const todayStr = todayISO();

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()} style={{ zIndex: 1200 }}>
      <div className="modal" style={{ maxWidth: 440, width: '100%', padding: '20px 24px', borderRadius: 'var(--radius-lg)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontSize: 17, fontWeight: 700 }}>
            📅 {monthNames[month]} {year}
          </h3>
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <button className="btn btn-ghost btn-sm" onClick={goToday} style={{ fontSize: 12, padding: '3px 8px' }}>
              Hoy
            </button>
            <button className="btn btn-outline btn-sm" onClick={prevMonth} style={{ padding: '3px 8px', fontSize: 13 }}>
              ◀
            </button>
            <button className="btn btn-outline btn-sm" onClick={nextMonth} style={{ padding: '3px 8px', fontSize: 13 }}>
              ▶
            </button>
            <button className="modal-close" onClick={onClose} style={{ marginLeft: 4 }}>
              <Icons.Close />
            </button>
          </div>
        </div>

        {/* Leyenda */}
        <div style={{ display: 'flex', gap: 14, justifyContent: 'center', marginBottom: 16, fontSize: 12, background: 'rgba(255,255,255,0.03)', padding: '8px 12px', borderRadius: 8, border: '1px solid var(--border)' }}>
          <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
            <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#f59e0b', display: 'inline-block' }} />
            Por pagar
          </span>
          <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
            <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#ef4444', display: 'inline-block' }} />
            Vencido
          </span>
          <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
            <span style={{ width: 8, height: 8, borderRadius: '50%', background: '#10b981', display: 'inline-block' }} />
            Pagado
          </span>
        </div>

        {/* Encabezado días de la semana */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', textAlign: 'center', fontWeight: 600, fontSize: 12, color: 'var(--text-muted)', marginBottom: 8 }}>
          <span>Dom</span><span>Lun</span><span>Mar</span><span>Mié</span><span>Jue</span><span>Vie</span><span>Sáb</span>
        </div>

        {/* Cuadrícula del calendario */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
          {daysGrid.map((item, idx) => {
            if (!item.currentMonth) {
              return (
                <div key={idx} style={{ padding: '10px 0', textAlign: 'center', fontSize: 12, color: 'var(--text-muted)', opacity: 0.3, userSelect: 'none' }}>
                  {item.day}
                </div>
              );
            }

            const dateStr = formatYYYYMMDD(year, month, item.day);
            const isToday = dateStr === todayStr;
            const indicators = dateIndicators[dateStr] || { pending: 0, expired: 0, paid: 0 };
            const hasDots = indicators.pending > 0 || indicators.expired > 0 || indicators.paid > 0;

            return (
              <button
                key={idx}
                onClick={() => {
                  onSelectDate(dateStr);
                  onClose();
                }}
                style={{
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  padding: '8px 0',
                  borderRadius: 8,
                  border: isToday ? '2px solid var(--accent-blue)' : '1px solid var(--border)',
                  background: isToday ? 'rgba(23, 113, 197, 0.12)' : hasDots ? 'rgba(255,255,255,0.03)' : 'transparent',
                  cursor: 'pointer',
                  transition: 'all 0.15s ease',
                  outline: 'none',
                  minHeight: 46
                }}
              >
                <span style={{ fontSize: 13, fontWeight: isToday ? 700 : 500, color: isToday ? 'var(--accent-blue)' : 'var(--text-main)' }}>
                  {item.day}
                </span>
                <div style={{ display: 'flex', gap: 3, marginTop: 4, height: 6, alignItems: 'center' }}>
                  {indicators.expired > 0 && (
                    <span
                      title={`${indicators.expired} póliza(s) vencida(s)`}
                      style={{ width: 6, height: 6, borderRadius: '50%', background: '#ef4444', display: 'inline-block' }}
                    />
                  )}
                  {indicators.pending > 0 && (
                    <span
                      title={`${indicators.pending} póliza(s) por pagar`}
                      style={{ width: 6, height: 6, borderRadius: '50%', background: '#f59e0b', display: 'inline-block' }}
                    />
                  )}
                  {indicators.paid > 0 && (
                    <span
                      title={`${indicators.paid} póliza(s) pagada(s)`}
                      style={{ width: 6, height: 6, borderRadius: '50%', background: '#10b981', display: 'inline-block' }}
                    />
                  )}
                </div>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ─── Modal: Pagos por Día ─────────────────────────────────────
function DailyPaymentsModal({ dateStr, policies, caroPolicies, onClose, onEdit, onDelete, onMarkPaid, onWhatsApp, onEmail }) {
  useEscapeKey(onClose);
  const isPolicyForDate = useCallback((p) => {
    if (p.estatus === 'CANCELADO') return false;
    const isPaid = p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO';
    if (isPaid) {
      return (p.fechaUltimoPago === dateStr) || (p.fechaPago === dateStr);
    }
    return p.fechaPago === dateStr;
  }, [dateStr]);

  const duePolicies = useMemo(() => policies.filter(isPolicyForDate), [policies, isPolicyForDate]);
  const annotatedCaro = useMemo(() => caroPolicies.map(p => ({ ...p, _isCaro: true })), [caroPolicies]);
  const dueCaroPolicies = useMemo(() => annotatedCaro.filter(isPolicyForDate), [annotatedCaro, isPolicyForDate]);

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-wide" style={{ maxWidth: 900 }}>
        <div className="modal-header">
          <h2>📅 Pagos programados para el {formatDate(dateStr)}</h2>
          <button className="modal-close" onClick={onClose}><Icons.Close /></button>
        </div>
        <div className="modal-body" style={{ maxHeight: '70vh', overflowY: 'auto', paddingBottom: 24 }}>
          {duePolicies.length === 0 && dueCaroPolicies.length === 0 ? (
            <div style={{ padding: 40, textAlign: 'center', color: 'var(--text-muted)' }}>
              No hay pagos programados para esta fecha.
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
              {duePolicies.length > 0 && (
                <div>
                  <h3 style={{ fontSize: 14, color: 'var(--text-secondary)', marginBottom: 12 }}>Pólizas Generales</h3>
                  <PoliciesTable policies={duePolicies} compact={true} onEdit={onEdit} onDelete={onDelete} onMarkPaid={onMarkPaid} onWhatsApp={onWhatsApp} onEmail={onEmail} />
                </div>
              )}
              {dueCaroPolicies.length > 0 && (
                <div>
                  <h3 style={{ fontSize: 14, color: '#8b5cf6', marginBottom: 12 }}>Pólizas Clave Caro</h3>
                  <PoliciesTable policies={dueCaroPolicies} compact={true} onEdit={onEdit} onDelete={onDelete} onMarkPaid={onMarkPaid} onWhatsApp={onWhatsApp} onEmail={onEmail} />
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Modal: WhatsApp / Correo ─────────────────────────────────
function ContactModal({ policy, type, templates, onClose }) {
  useEscapeKey(onClose);
  const lada = policy.lada || '+52';
  const initialPhone = (policy.telefono || '').replace(/\D/g, '');
  const [editablePhone, setEditablePhone] = useState(initialPhone);
  const waNumber = lada.replace('+', '') + editablePhone;

  const msgText = fillTemplate(templates.whatsapp, policy, true);
  const emailAsunto = fillTemplate(templates.email_asunto, policy, false);
  const emailCuerpo = fillTemplate(templates.email_cuerpo, policy, false);

  const openWA = () => {
    const url = `https://wa.me/${waNumber}?text=${encodeURIComponent(msgText)}`;
    window.open(url, '_blank');
  };

  const openEmail = () => {
    const url = `https://mail.google.com/mail/u/0/?view=cm&fs=1&to=${encodeURIComponent(policy.correo || '')}&cc=${encodeURIComponent('dagarso79@hotmail.com')}&su=${encodeURIComponent(emailAsunto)}&body=${encodeURIComponent(emailCuerpo)}`;
    window.open(url, '_blank');
  };

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-wide">
        <div className="modal-header">
          <h2>{type === 'whatsapp' ? '💬 Enviar WhatsApp' : '📧 Enviar Correo'}</h2>
          <button className="modal-close" onClick={onClose}><Icons.Close /></button>
        </div>
        <div className="modal-body">
          <div className="info-grid" style={{marginBottom:16}}>
            <div className="info-card">
              <div className="info-card-label">Destinatario</div>
              <div className="info-card-value" style={{fontSize:13}}>{policy.nombre}</div>
            </div>
            {type === 'whatsapp' ? (
              <div className="info-card" style={{padding: '8px 12px'}}>
                <div className="info-card-label">WhatsApp (Selecciona el número a enviar)</div>
                <div style={{display: 'flex', flexDirection: 'column', gap: 6, marginTop: 4}}>
                  <div style={{display: 'flex', alignItems: 'center', gap: 6}}>
                    <input 
                      type="radio" 
                      id="opt_tel1" 
                      name="wa_select" 
                      checked={editablePhone === (policy.telefono || '').replace(/\D/g, '')}
                      onChange={() => setEditablePhone((policy.telefono || '').replace(/\D/g, ''))}
                    />
                    <label htmlFor="opt_tel1" style={{fontSize: 12, cursor: 'pointer', fontWeight: 500}}>
                      Tel 1: {policy.lada || '+52'} {policy.telefono || '—'}
                    </label>
                  </div>
                  {policy.telefono2 && (
                    <div style={{display: 'flex', alignItems: 'center', gap: 6}}>
                      <input 
                        type="radio" 
                        id="opt_tel2" 
                        name="wa_select" 
                        checked={editablePhone === (policy.telefono2 || '').replace(/\D/g, '')}
                        onChange={() => setEditablePhone((policy.telefono2 || '').replace(/\D/g, ''))}
                      />
                      <label htmlFor="opt_tel2" style={{fontSize: 12, cursor: 'pointer', fontWeight: 500}}>
                        Tel 2: {policy.lada2 || '+52'} {policy.telefono2}
                      </label>
                    </div>
                  )}
                  <input 
                    type="text" 
                    className="input" 
                    style={{padding: '4px 8px', width: '100%', marginTop: 4}}
                    value={editablePhone}
                    onChange={(e) => setEditablePhone(e.target.value.replace(/[^0-9]/g, ''))}
                    placeholder="Número personalizado"
                  />
                </div>
              </div>
            ) : (
              <div className="info-card">
                <div className="info-card-label">Correo</div>
                <div className="info-card-value" style={{fontSize:13, wordBreak:'break-all'}}>
                  {policy.correo || '—'}
                </div>
              </div>
            )}
          </div>

          {type === 'whatsapp' ? (
            <>
              <p className="form-label" style={{marginBottom:8}}>Vista previa del mensaje:</p>
              <div className="template-preview">{msgText}</div>
              {!editablePhone && (
                <p style={{fontSize:12, color:'var(--accent-red)', marginTop:8}}>
                  ⚠️ Introduce un número de teléfono válido.
                </p>
              )}
            </>
          ) : (
            <>
              <div style={{marginBottom:12}}>
                <p className="form-label" style={{marginBottom:6}}>Asunto:</p>
                <div className="template-preview" style={{padding:'10px 14px', fontSize:14, fontWeight:600}}>
                  {emailAsunto}
                </div>
              </div>
              <p className="form-label" style={{marginBottom:8}}>Cuerpo del correo:</p>
              <div className="template-preview">{emailCuerpo}</div>
              {!policy.correo && (
                <p style={{fontSize:12, color:'var(--accent-red)', marginTop:8}}>
                  ⚠️ Esta póliza no tiene correo electrónico registrado.
                </p>
              )}
            </>
          )}
        </div>
        <div className="modal-footer">
          <button className="btn btn-outline" onClick={onClose}>Cerrar</button>
          {type === 'whatsapp' ? (
            <button className="btn btn-success" onClick={openWA} disabled={!editablePhone}>
              💬 Abrir en WhatsApp
            </button>
          ) : (
            <button className="btn btn-primary" onClick={openEmail} disabled={!policy.correo}>
              📧 Abrir en Correo
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Tabla Principal de Pólizas ───────────────────────────────
function PoliciesTable({ policies, onEdit, onDelete, onMarkPaid, onWhatsApp, onEmail, compact = true, showSectionTag = false }) {
  const [sort, setSort] = useState({ key: 'fechaPago', dir: 'asc' });
  const [summaryPolicy, setSummaryPolicy] = useState(null);

  const toggleSort = (key) => {
    setSort(s => ({ key, dir: s.key === key && s.dir === 'asc' ? 'desc' : 'asc' }));
  };

  const sorted = useMemo(() => {
    return [...policies].sort((a, b) => {
      let av = a[sort.key], bv = b[sort.key];
      if (sort.key === 'ramo') {
        const getRamoStr = p => p?._isCaro ? 'AUTOS QUALITAS CARO' : p?._isGmm ? 'GMM' : p?._isAutos ? 'AUTOS' : p?._isVida ? 'VIDA' : p?._isDanos ? 'DAÑOS' : p?._isHogar ? 'HOGAR' : 'AUTOS QUALITAS DANI';
        av = getRamoStr(a);
        bv = getRamoStr(b);
      }
      if (sort.key === 'monto') { av = Number(av); bv = Number(bv); }
      if (typeof av === 'string') av = av.toLowerCase();
      if (typeof bv === 'string') bv = bv.toLowerCase();
      if (av < bv) return sort.dir === 'asc' ? -1 : 1;
      if (av > bv) return sort.dir === 'asc' ? 1 : -1;
      return 0;
    });
  }, [policies, sort]);

  const Th = ({ k, label }) => (
    <th className="sortable" onClick={() => toggleSort(k)}>
      {label}
      <Icons.Sort dir={sort.key === k ? sort.dir : null} />
    </th>
  );

  const isUrgent = (p) => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    return isUpcomingReminder(p) || isExpiredEffective(p);
  };

  if (sorted.length === 0) {
    return (
      <div className="empty-state">
        <div className="empty-state-icon">📋</div>
        <h3>Sin resultados</h3>
        <p>No hay pólizas que coincidan con los filtros aplicados.</p>
      </div>
    );
  }

  return (
    <div className="table-wrapper">
      <table>
        <thead>
          <tr>
            <Th k="nombre" label="Asegurado" />
            <Th k="poliza" label="Póliza" />
            {!compact && <Th k="bien" label="Unidad" />}
            <Th k="agente" label="Agente" />
            {showSectionTag && <Th k="ramo" label="Ramo" />}
            <Th k="formaPago" label="Forma Pago" />
            <Th k="fechaPago" label="Fecha Límite" />
            <Th k="monto" label="Monto" />
            <Th k="estatus" label="Estatus" />
            <th>Acciones</th>
          </tr>
        </thead>
        <tbody>
          {sorted.map(p => (
            <tr key={p.id} className={isUrgent(p) ? 'urgent-row' : ''}>
              <td>
                <div style={{fontWeight:600, fontSize:13}}>{p.nombre}</div>
                {p.perteneceA && <div style={{fontSize:11, color:'#38bdf8', marginTop:2}}>👤 Pertenece a: {p.perteneceA}</div>}
                {p.notas && <div style={{fontSize:11, color:'var(--text-muted)', marginTop:2}}>📝 {p.notas.slice(0,40)}{p.notas.length > 40 ? '…' : ''}</div>}
              </td>
              <td>
                <code 
                  onDoubleClick={() => setSummaryPolicy(p)}
                  title="Haz doble clic para ver el resumen completo"
                  style={{fontSize:12, color:'var(--text-secondary)', background:'rgba(255,255,255,0.08)', padding:'4px 8px', borderRadius:4, cursor:'pointer', userSelect:'text'}}
                >
                  📋 {p.poliza}
                </code>
              </td>
              {!compact && <td style={{maxWidth:200}}><div className="truncate" style={{fontSize:12, color:'var(--text-secondary)'}} title={p.bien}>{p.bien || '—'}</div></td>}
              <td><AgentBadge policy={p} agente={p.agente || p.aseguradora} /></td>
              {showSectionTag && <td><RamoBadge policy={p} /></td>}
              <td><span className="forma-badge">{p.formaPago}</span></td>
              <td><DateCell dateStr={p.fechaPago} estatus={p.estatus} periodoGracia={p.periodoGracia} /></td>
              <td><span style={{fontWeight:600}}>{formatMoney(getEffectiveMonto(p))}</span></td>
              <td><StatusBadge policy={p} /></td>
              <td>
                <div className="action-btns">
                  <button className="action-btn action-btn-status" title="Confirmar pago / subir comprobante"
                    onClick={() => onMarkPaid(p)}>✅</button>
                  <button className="action-btn action-btn-whatsapp" title="Enviar WhatsApp"
                    onClick={() => onWhatsApp(p)}>
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                      <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.888-.788-1.487-1.761-1.66-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413Z"/>
                    </svg>
                  </button>
                  <button className="action-btn action-btn-email" title="Enviar correo"
                    onClick={() => onEmail(p)}>✉️</button>
                  <button className="action-btn action-btn-edit" title="Editar"
                    onClick={() => onEdit(p)}>✏️</button>
                  <button className="action-btn action-btn-delete" title="Eliminar"
                    onClick={() => onDelete(p)}>🗑️</button>
                </div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {summaryPolicy && <PolicySummaryModal policy={summaryPolicy} onClose={() => setSummaryPolicy(null)} />}
    </div>
  );
}

// ─── Página: Dashboard ────────────────────────────────────────
function DashboardPage({ policies, onMarkPaid, onWhatsApp, onEmail, onEdit, onDelete }) {
  const [filterEstatus, setFilterEstatus] = useState('TODOS');

  const stats = useMemo(() => {
    const total = policies.length;
    const pagados = policies.filter(p => p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO').length;
    const pendientes = policies.filter(p => p.estatus === 'PENDIENTE' && !isExpiredEffective(p)).length;
    const vencidos = policies.filter(p => p.estatus === 'VENCIDO' || isExpiredEffective(p)).length;
    const cancelados = policies.filter(p => p.estatus === 'CANCELADO').length;
    const montoTotal = policies.filter(p => p.estatus !== 'CANCELADO').reduce((s, p) => s + Number(p.monto || 0), 0);
    const urgentes = policies.filter(p => isUpcomingReminder(p)).length;
    const renovaciones = policies.filter(p => isUpcomingRenewal(p)).length;
    return { total, pagados, pendientes, vencidos, cancelados, montoTotal, urgentes, renovaciones };
  }, [policies]);

  const filteredByStat = useMemo(() => {
    if (filterEstatus === 'TODOS') return policies;
    if (filterEstatus === 'RENOVACIONES') return policies.filter(p => isUpcomingRenewal(p));
    if (filterEstatus === 'VENCIDO') return policies.filter(p => p.estatus === 'VENCIDO' || isExpiredEffective(p));
    if (filterEstatus === 'PENDIENTE') return policies.filter(p => p.estatus === 'PENDIENTE' && !isExpiredEffective(p));
    if (filterEstatus === 'PAGADO') return policies.filter(p => p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO');
    return policies.filter(p => p.estatus === filterEstatus);
  }, [policies, filterEstatus]);

  const vencidas = useMemo(() => policies.filter(p => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    return isExpiredEffective(p);
  }), [policies]);

  const proximas = useMemo(() => policies.filter(p => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    return isUpcomingReminder(p);
  }), [policies]);

  const renovaciones = useMemo(() => policies.filter(p => isUpcomingRenewal(p)), [policies]);

  return (
    <div className="page-fade-enter">
      {/* KPIs */}
      <div className="stats-grid">
        {[
          { label: 'Total Pólizas', value: stats.total, icon: '🛡️', cls: 'stat-blue', filter: 'TODOS' },
          { label: 'Pendientes', value: stats.pendientes, icon: '⏳', cls: 'stat-yellow', filter: 'PENDIENTE' },
          { label: 'Vencidos', value: stats.vencidos, icon: '🔴', cls: 'stat-red', filter: 'VENCIDO' },
          { label: 'Renovaciones', value: stats.renovaciones, icon: '🔄', cls: 'stat-purple', filter: 'RENOVACIONES' },
          { label: 'Pagados (ciclo)', value: stats.pagados, icon: '✅', cls: 'stat-green', filter: 'PAGADO' },
          { label: 'Cancelados', value: stats.cancelados, icon: '❌', cls: 'stat-gray', filter: 'CANCELADO' },
          { label: 'Cobranza Total', value: formatMoney(stats.montoTotal), icon: '💰', cls: 'stat-orange', filter: 'TODOS' },
        ].map(s => (
          <div key={s.label} className={`stat-card ${s.cls}`} 
            style={{
              cursor: 'pointer',
              opacity: filterEstatus === s.filter || filterEstatus === 'TODOS' ? 1 : 0.5,
              border: filterEstatus === s.filter && s.filter !== 'TODOS' ? '2px solid currentColor' : '1px solid transparent',
              transition: 'all 0.2s ease'
            }} 
            onClick={() => setFilterEstatus(filterEstatus === s.filter ? 'TODOS' : s.filter)}>
            <div className="stat-card-icon">{s.icon}</div>
            <div className="stat-card-value">{s.value}</div>
            <div className="stat-card-label">{s.label}</div>
          </div>
        ))}
      </div>

      {/* Si hay un filtro seleccionado en el Dashboard, mostrar tabla concentrada filtrada */}
      {filterEstatus !== 'TODOS' ? (
        <div className="card" style={{marginTop: 20}}>
          <div className="card-header">
            <span className="card-title">📋 Pólizas Concentradas — Estatus: {filterEstatus} ({filteredByStat.length})</span>
            <button className="btn btn-ghost btn-sm" onClick={() => setFilterEstatus('TODOS')}>↩ Ver Vista General</button>
          </div>
          <PoliciesTable
            policies={filteredByStat}
            compact={true}
            showSectionTag={true}
            onEdit={onEdit}
            onDelete={onDelete}
            onMarkPaid={onMarkPaid}
            onWhatsApp={onWhatsApp}
            onEmail={onEmail}
          />
        </div>
      ) : (
        <>
          {/* Vencidas */}
          {vencidas.length > 0 && (
            <div className="card" style={{border: '1px solid rgba(239, 68, 68, 0.3)'}}>
              <div className="card-header" style={{borderBottom: '1px solid rgba(239, 68, 68, 0.2)'}}>
                <span className="card-title" style={{color: 'var(--accent-red)'}}>🛑 Pólizas Vencidas — {vencidas.length} póliza(s) con pago atrasado</span>
                <span style={{fontSize:12, color:'var(--text-muted)'}}>Contacta de inmediato a estos asegurados</span>
              </div>
              <PoliciesTable
                policies={vencidas}
                compact={true}
                showSectionTag={true}
                onEdit={onEdit}
                onDelete={onDelete}
                onMarkPaid={onMarkPaid}
                onWhatsApp={onWhatsApp}
                onEmail={onEmail}
              />
            </div>
          )}

          {/* Próximas a vencer */}
          {proximas.length > 0 && (
            <div className="card">
              <div className="card-header" style={{background: 'rgba(245, 158, 11, 0.05)', borderBottom: '1px solid rgba(245, 158, 11, 0.2)'}}>
                <span className="card-title" style={{color: 'var(--accent-yellow)'}}>⚠️ Próximas a vencer (en 4 días o menos)</span>
                <span style={{fontSize:12, color:'var(--text-muted)'}}>
                  {proximas.length} póliza(s)
                </span>
              </div>
              <PoliciesTable
                policies={proximas}
                compact={true}
                showSectionTag={true}
                onEdit={onEdit}
                onDelete={onDelete}
                onMarkPaid={onMarkPaid}
                onWhatsApp={onWhatsApp}
                onEmail={onEmail}
              />
            </div>
          )}

          {/* Renovaciones */}
          {renovaciones.length > 0 && (
            <div className="card">
              <div className="card-header" style={{background: 'rgba(139, 92, 246, 0.05)', borderBottom: '1px solid rgba(139, 92, 246, 0.2)'}}>
                <span className="card-title" style={{color: '#8b5cf6'}}>🔄 Próximas a Renovar</span>
                <span style={{fontSize:12, color:'var(--text-muted)'}}>
                  {renovaciones.length} póliza(s) (ya liquidadas, vence su ciclo anual en &lt;= 15 días)
                </span>
              </div>
              <PoliciesTable
                policies={renovaciones}
                compact={true}
                showSectionTag={true}
                onEdit={onEdit}
                onDelete={onDelete}
                onMarkPaid={onMarkPaid}
                onWhatsApp={onWhatsApp}
                onEmail={onEmail}
              />
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ─── Página: Qualitas D&M (Pólizas Separadas) ─────────────────
function PoliciesPage({ policies, onEdit, onDelete, onMarkPaid, onWhatsApp, onEmail, onNew, onUpdatePolicy, defaultEstatus = 'TODOS' }) {
  const [selectedImg, setSelectedImg] = useState(null);

  const [search, setSearch] = useState('');
  const [filterAgente, setFilterAgente] = useState('TODOS');
  const [filterEstatus, setFilterEstatus] = useState(defaultEstatus);
  const [filterForma, setFilterForma] = useState('TODOS');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  useEffect(() => {
    setFilterEstatus(defaultEstatus);
  }, [defaultEstatus]);

  const stats = useMemo(() => {
    const list = policies || [];
    const total = list.length;
    const pagados = list.filter(p => p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO').length;
    const pendientes = list.filter(p => p.estatus === 'PENDIENTE').length;
    const vencidos = list.filter(p => isExpiredEffective(p)).length;
    const urgentes = list.filter(p => isUpcomingReminder(p)).length;
    const renovaciones = list.filter(p => isUpcomingRenewal(p)).length;
    const comprobantes = list.filter(p => p.comprobante).length;
    return { total, pagados, pendientes, vencidos, urgentes, renovaciones, comprobantes };
  }, [policies]);

  const agentOpts = useMemo(() => {
    const list = policies || [];
    return Array.from(new Set(list.map(p => p.agente).filter(Boolean)));
  }, [policies]);

  const filtered = useMemo(() => {
    const list = policies || [];
    return list.filter(p => {
      if (!p) return false;
      const q = search.toLowerCase().trim();
      const matchName = (p.nombre || '').toLowerCase().includes(q);
      const matchPoliza = (p.poliza || '').toLowerCase().includes(q);
      const matchBien = (p.bien || '').toLowerCase().includes(q);
      if (q && !matchName && !matchPoliza && !matchBien) return false;
      if (filterAgente !== 'TODOS' && p.agente !== filterAgente && p.aseguradora !== filterAgente) return false;
      if (filterEstatus !== 'TODOS') {
        if (filterEstatus === 'RENOVACIONES') {
          if (!isUpcomingRenewal(p)) return false;
        } else if (filterEstatus === 'URGENTES') {
          if (!isUpcomingReminder(p)) return false;
        } else if (filterEstatus === 'VENCIDO') {
          if (!isExpiredEffective(p)) return false;
        } else if (filterEstatus === 'COMPROBANTES') {
          if (!p.comprobante) return false;
        } else if (filterEstatus === 'PAGADO') {
          if (p.estatus !== 'PAGADO' && p.estatus !== 'LIQUIDADO') return false;
        } else {
          if (p.estatus !== filterEstatus) return false;
        }
      }
      if (filterForma !== 'TODOS' && p.formaPago !== filterForma) return false;
      if (dateFrom && p.fechaPago < dateFrom) return false;
      if (dateTo && p.fechaPago > dateTo) return false;
      return true;
    });
  }, [policies, search, filterAgente, filterEstatus, filterForma, dateFrom, dateTo]);

  const clearFilters = () => {
    setSearch(''); setFilterAgente('TODOS'); setFilterEstatus('TODOS');
    setFilterForma('TODOS'); setDateFrom(''); setDateTo('');
  };

  const activeFilters = filterAgente !== 'TODOS' || filterEstatus !== 'TODOS' ||
    filterForma !== 'TODOS' || dateFrom || dateTo || search;

  return (
    <div className="page-fade-enter">
      {/* Tarjetas KPI como filtros */}
      <div className="stats-grid" style={{marginBottom: 20}}>
        {[
          { label: 'Total Autos Qualitas', value: stats.total, icon: '📋', cls: 'stat-blue', filter: 'TODOS' },
          { label: 'Pendientes', value: stats.pendientes, icon: '⏳', cls: 'stat-yellow', filter: 'PENDIENTE' },
          { label: 'Próx. a Vencer (4d)', value: stats.urgentes, icon: '🔴', cls: 'stat-orange', filter: 'URGENTES' },
          { label: 'Vencidos', value: stats.vencidos, icon: '🛑', cls: 'stat-red', filter: 'VENCIDO' },
          { label: 'Renovaciones', value: stats.renovaciones, icon: '🔄', cls: 'stat-purple', filter: 'RENOVACIONES' },
          { label: 'Pagados', value: stats.pagados, icon: '✅', cls: 'stat-green', filter: 'PAGADO' },
          { label: 'Comprobantes', value: stats.comprobantes, icon: '🧾', cls: 'stat-orange', filter: 'COMPROBANTES' },
        ].map(s => (
          <div key={s.label} className={`stat-card ${s.cls}`} 
            style={{
              cursor: 'pointer', 
              opacity: filterEstatus === s.filter || filterEstatus === 'TODOS' ? 1 : 0.5,
              border: filterEstatus === s.filter ? '2px solid currentColor' : '1px solid transparent',
              transition: 'all 0.2s ease'
            }} 
            onClick={() => setFilterEstatus(s.filter)}>
            <div className="stat-card-icon">{s.icon}</div>
            <div className="stat-card-value">{s.value}</div>
            <div className="stat-card-label">{s.label}</div>
          </div>
        ))}
      </div>

      {/* Barra de Filtros */}
      <div className="card" style={{marginBottom: 20}}>
        <div className="card-header" style={{flexDirection:'column', alignItems:'flex-start', gap:14}}>
          <div className="flex justify-between w-full items-center">
            <span className="card-title">📋 Pólizas Qualitas D&M ({filtered.length})</span>
            <div className="flex gap-2">
              {activeFilters && (
                <button className="btn btn-ghost btn-sm" onClick={clearFilters}>✕ Limpiar Filtros</button>
              )}
              <button className="btn btn-primary btn-sm" onClick={onNew}>
                <Icons.Plus /> Nueva Póliza
              </button>
            </div>
          </div>
          <div className="filters-bar">
            <div className="search-wrapper">
              <Icons.Search />
              <input className="input input-search" value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Buscar nombre, póliza, bien..." />
            </div>
            <div className="filter-group">
              <span className="filter-label">Agente</span>
              <select className="select" style={{minWidth:130}} value={filterAgente}
                onChange={e => setFilterAgente(e.target.value)}>
                <option value="TODOS">Todos</option>
                {agentOpts.map(opt => <option key={opt} value={opt}>{opt}</option>)}
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Estatus</span>
              <select className="select" style={{minWidth:140}} value={filterEstatus}
                onChange={e => setFilterEstatus(e.target.value)}>
                <option value="TODOS">Todos</option>
                <option value="PENDIENTE">PENDIENTE</option>
                <option value="URGENTES">PRÓX. A VENCER (4D)</option>
                <option value="VENCIDO">VENCIDO</option>
                <option value="PAGADO">PAGADO</option>
                <option value="CANCELADO">CANCELADO</option>
                <option value="LIQUIDADO">LIQUIDADO</option>
                <option value="RENOVACIONES">RENOVACIONES (Próximas)</option>
                <option value="COMPROBANTES">COMPROBANTES</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Forma de Pago</span>
              <select className="select" style={{minWidth:140}} value={filterForma}
                onChange={e => setFilterForma(e.target.value)}>
                <option value="TODOS">Todas</option>
                <option value="CONTADO">CONTADO</option>
                <option value="MENSUAL">MENSUAL</option>
                <option value="TRIMESTRAL">TRIMESTRAL</option>
                <option value="SEMESTRAL">SEMESTRAL</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha desde</span>
              <input type="date" className="input" style={{width:140}} value={dateFrom}
                onChange={e => setDateFrom(e.target.value)} />
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha hasta</span>
              <input type="date" className="input" style={{width:140}} value={dateTo}
                onChange={e => setDateTo(e.target.value)} />
            </div>
          </div>
        </div>
      </div>

      {filterEstatus !== 'COMPROBANTES' && (
        <div className="card">
          <div className="card-header">
            <span className="card-title">
              {filterEstatus === 'TODOS' ? 'Todas las Pólizas Autos Qualitas' : `Pólizas (${filterEstatus})`} 
              {' '}({filtered.length})
            </span>
            {activeFilters && (
              <button className="btn btn-ghost btn-sm" onClick={clearFilters}>↩ Mostrar Todas</button>
            )}
          </div>
          <PoliciesTable
            policies={filtered}
            compact={true}
            onEdit={onEdit}
            onDelete={onDelete}
            onMarkPaid={onMarkPaid}
            onWhatsApp={onWhatsApp}
            onEmail={onEmail}
          />
          {filtered.length > 0 && (
            <div style={{padding:'12px 24px', borderTop:'1px solid var(--border)', fontSize:12, color:'var(--text-muted)', display:'flex', justifyContent:'space-between'}}>
              <span>{filtered.length} registro(s) encontrado(s)</span>
              <span>Total filtrado: <strong style={{color:'var(--accent-green)'}}>{formatMoney(filtered.reduce((s,p) => s+Number(p.monto||0), 0))}</strong></span>
            </div>
          )}
        </div>
      )}

      {/* Vista de Comprobantes Qualitas D&M (cuando se filtra por COMPROBANTES) */}
      {filterEstatus === 'COMPROBANTES' && (() => {
        const withComprobantes = (policies || []).filter(p => p.comprobante);
        if (withComprobantes.length === 0) return (
          <div className="card" style={{marginTop: 20}}>
            <div style={{padding: 40, textAlign: 'center', color: 'var(--text-muted)'}}>
              🧾 Aún no hay comprobantes guardados en Qualitas D&M.
            </div>
          </div>
        );
        const grouped = {};
        withComprobantes.forEach(p => {
          const dStr = p.fechaUltimoPago || new Date().toISOString().split('T')[0];
          const date = new Date(dStr + 'T12:00:00');
          const monthYear = date.toLocaleDateString('es-MX', { month: 'long', year: 'numeric' });
          const capitalized = monthYear.charAt(0).toUpperCase() + monthYear.slice(1);
          if (!grouped[capitalized]) grouped[capitalized] = [];
          grouped[capitalized].push(p);
        });
        return Object.entries(grouped).map(([monthName, groupPolicies]) => (
          <div key={monthName} className="card" style={{marginTop: 20}}>
            <div className="card-header">
              <span className="card-title">📁 {monthName}</span>
              <span style={{fontSize:12, color:'var(--text-muted)'}}>{groupPolicies.length} comprobante(s)</span>
            </div>
            <div style={{padding: '20px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '20px'}}>
              {groupPolicies.map(p => (
                <div key={p.id} style={{border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: 12, background: 'var(--bg-card)'}}>
                  <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start'}}>
                    <div style={{fontWeight: 600, fontSize: 14, marginBottom: 4}}>{p.nombre}</div>
                    <button 
                      title="Eliminar comprobante"
                      onClick={() => { if (confirm('¿Eliminar este comprobante?')) onEdit({ ...p, comprobante: null }); }}
                      style={{background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 6, padding: '4px 8px', cursor: 'pointer', fontSize: 13, color: '#ef4444', flexShrink: 0}}
                    >🗑️</button>
                  </div>
                  <div style={{fontSize: 12, color: 'var(--text-secondary)', marginBottom: 8}}>
                    <strong>Póliza:</strong> {p.poliza}
                  </div>
                  <div style={{fontSize: 11, color: 'var(--text-muted)', marginBottom: 4}}>
                    <strong>Fecha límite:</strong> {formatDate(p.fechaPagoAnterior || p.fechaPago)}
                  </div>
                  <div style={{fontSize: 11, color: 'var(--accent-green)', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6}}>
                    <strong>Fecha pagado:</strong> 
                    <input type="date" 
                      defaultValue={p.fechaUltimoPago || todayISO()} 
                      onBlur={e => {
                        const val = e.target.value;
                        if (val && val !== p.fechaUltimoPago) {
                          onUpdatePolicy({ ...p, fechaUltimoPago: val });
                          if (toast) toast('Fecha de pago actualizada ✅', 'success');
                        }
                      }}
                      style={{fontSize: 11, padding: '2px 6px', border: '1px solid var(--accent-green)', borderRadius: 4, background: 'var(--bg-input)', color: 'var(--accent-green)', fontWeight: 'bold'}}
                    />
                  </div>
                  <div style={{width: '100%', height: 200, borderRadius: 'var(--radius-sm)', overflow: 'hidden', border: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg-secondary)', cursor: 'pointer'}}
                    onClick={() => setSelectedImg(p.comprobante)}>
                    {p.comprobante.startsWith('data:application/pdf') ? (
                      <embed src={p.comprobante} width="100%" height="100%" type="application/pdf" style={{pointerEvents: 'none'}} />
                    ) : (
                      <img src={p.comprobante} alt="Comprobante" style={{width: '100%', height: '100%', objectFit: 'contain'}} />
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ));
      })()}

      {selectedImg && <ImageModal src={selectedImg} onClose={() => setSelectedImg(null)} />}
    </div>
  );
}

// ─── Página: Urgentes ─────────────────────────────────────────
function UrgentPage({ policies, onEdit, onDelete, onMarkPaid, onWhatsApp, onEmail }) {
  const urgent = useMemo(() => policies.filter(p => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    const d = daysUntil(p.fechaPago);
    return d !== null && d <= 4;
  }).sort((a, b) => (a.fechaPago || '') < (b.fechaPago || '') ? -1 : 1), [policies]);

  return (
    <div className="page-fade-enter">
      {urgent.length === 0 ? (
        <div className="empty-state" style={{paddingTop:100}}>
          <div className="empty-state-icon">🎉</div>
          <h3>¡Sin urgencias!</h3>
          <p>No hay pólizas con vencimiento en los próximos 4 días. ¡Todo al día!</p>
        </div>
      ) : (
        <>
          <div className="alert-banner" style={{marginBottom:20}}>
            <span className="alert-icon">🚨</span>
            <div className="alert-content">
              <h3>{urgent.length} póliza(s) requieren atención inmediata</h3>
              <p>Estas pólizas vencen dentro de 4 días o ya están vencidas. Envía recordatorios ahora.</p>
            </div>
            <div style={{marginLeft:'auto', display:'flex', gap:8}}>
              <button className="btn btn-warning btn-sm" onClick={() => urgent.forEach(p => p.telefono && window.open(`https://wa.me/${(p.lada||'+52').replace('+','')}${p.telefono}?text=${encodeURIComponent(fillTemplate('Hola {nombre}, le recordamos que su póliza {poliza} vence el {fechaPago} por {monto}. Favor de realizar su pago. Gracias.', p))}`, '_blank'))}>
                💬 WA Masivo
              </button>
            </div>
          </div>
          <div className="card">
            <div className="card-header">
              <span className="card-title">⚡ Recordatorios Urgentes</span>
              <span style={{fontSize:12, color:'var(--text-muted)'}}>{urgent.length} registros</span>
            </div>
            <PoliciesTable
              policies={urgent}
              compact={false}
              onEdit={onEdit}
              onDelete={onDelete}
              onMarkPaid={onMarkPaid}
              onWhatsApp={onWhatsApp}
              onEmail={onEmail}
            />
          </div>
        </>
      )}
    </div>
  );
}


// ─── Página: Clave Caro (Pólizas Separadas) ───────────────────
function CaroPoliciesPage({ policies, onSave, onDelete, onMarkPaid, onWhatsApp, onEmail, toast }) {
  const [modalNew, setModalNew] = useState(false);
  const [modalEdit, setModalEdit] = useState(null);
  const [modalPaid, setModalPaid] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [selectedImg, setSelectedImg] = useState(null);

  useEscapeKey(deleteConfirm ? () => setDeleteConfirm(null) : null);

  const [search, setSearch] = useState('');
  const [filterAgente, setFilterAgente] = useState('TODOS');
  const [filterEstatus, setFilterEstatus] = useState('TODOS');
  const [filterForma, setFilterForma] = useState('TODOS');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  const stats = useMemo(() => {
    const list = policies || [];
    const total = list.length;
    const pagados = list.filter(p => p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO').length;
    const pendientes = list.filter(p => p.estatus === 'PENDIENTE').length;
    const vencidos = list.filter(p => isExpiredEffective(p)).length;
    const urgentes = list.filter(p => isUpcomingReminder(p)).length;
    const renovaciones = list.filter(p => isUpcomingRenewal(p)).length;
    const comprobantes = list.filter(p => p.comprobante).length;
    return { total, pagados, pendientes, vencidos, urgentes, renovaciones, comprobantes };
  }, [policies]);

  const agentOpts = useMemo(() => {
    const list = policies || [];
    return Array.from(new Set(list.map(p => p.agente).filter(Boolean)));
  }, [policies]);

  const filtered = useMemo(() => {
    const list = policies || [];
    return list.filter(p => {
      if (!p) return false;
      const q = search.toLowerCase().trim();
      const matchName = (p.nombre || '').toLowerCase().includes(q);
      const matchPoliza = (p.poliza || '').toLowerCase().includes(q);
      const matchBien = (p.bien || '').toLowerCase().includes(q);
      if (q && !matchName && !matchPoliza && !matchBien) return false;
      if (filterAgente !== 'TODOS' && p.agente !== filterAgente && p.aseguradora !== filterAgente) return false;
      if (filterEstatus !== 'TODOS') {
        if (filterEstatus === 'RENOVACIONES') {
          if (!isUpcomingRenewal(p)) return false;
        } else if (filterEstatus === 'URGENTES') {
          if (!isUpcomingReminder(p)) return false;
        } else if (filterEstatus === 'VENCIDO') {
          if (!isExpiredEffective(p)) return false;
        } else if (filterEstatus === 'COMPROBANTES') {
          if (!p.comprobante) return false;
        } else if (filterEstatus === 'PAGADO') {
          if (p.estatus !== 'PAGADO' && p.estatus !== 'LIQUIDADO') return false;
        } else {
          if (p.estatus !== filterEstatus) return false;
        }
      }
      if (filterForma !== 'TODOS' && p.formaPago !== filterForma) return false;
      if (dateFrom && p.fechaPago < dateFrom) return false;
      if (dateTo && p.fechaPago > dateTo) return false;
      return true;
    });
  }, [policies, search, filterAgente, filterEstatus, filterForma, dateFrom, dateTo]);

  const clearFilters = () => {
    setSearch(''); setFilterAgente('TODOS'); setFilterEstatus('TODOS');
    setFilterForma('TODOS'); setDateFrom(''); setDateTo('');
  };

  const activeFilters = filterAgente !== 'TODOS' || filterEstatus !== 'TODOS' ||
    filterForma !== 'TODOS' || dateFrom || dateTo || search;

  return (
    <div className="page-fade-enter">
      {/* Tarjetas KPI como filtros */}
      <div className="stats-grid" style={{marginBottom: 20}}>
        {[
          { label: 'Total Pólizas', value: stats.total, icon: '🛡️', cls: 'stat-blue', filter: 'TODOS' },
          { label: 'Pendientes', value: stats.pendientes, icon: '⏳', cls: 'stat-yellow', filter: 'PENDIENTE' },
          { label: 'Próx. a Vencer (4d)', value: stats.urgentes, icon: '🔴', cls: 'stat-orange', filter: 'URGENTES' },
          { label: 'Vencidos', value: stats.vencidos, icon: '🛑', cls: 'stat-red', filter: 'VENCIDO' },
          { label: 'Renovaciones', value: stats.renovaciones, icon: '🔄', cls: 'stat-purple', filter: 'RENOVACIONES' },
          { label: 'Pagados', value: stats.pagados, icon: '✅', cls: 'stat-green', filter: 'PAGADO' },
          { label: 'Comprobantes', value: stats.comprobantes, icon: '🧾', cls: 'stat-orange', filter: 'COMPROBANTES' },
        ].map(s => (
          <div key={s.label} className={`stat-card ${s.cls}`} 
            style={{
              cursor: 'pointer', 
              opacity: filterEstatus === s.filter || filterEstatus === 'TODOS' ? 1 : 0.5,
              border: filterEstatus === s.filter ? '2px solid currentColor' : '1px solid transparent',
              transition: 'all 0.2s ease'
            }} 
            onClick={() => setFilterEstatus(s.filter)}>
            <div className="stat-card-icon">{s.icon}</div>
            <div className="stat-card-value">{s.value}</div>
            <div className="stat-card-label">{s.label}</div>
          </div>
        ))}
      </div>

      {/* Barra de Filtros */}
      <div className="card" style={{marginBottom: 20}}>
        <div className="card-header" style={{flexDirection:'column', alignItems:'flex-start', gap:14}}>
          <div className="flex justify-between w-full items-center">
            <span className="card-title">🛡️ Pólizas Clave Caro ({filtered.length})</span>
            <div className="flex gap-2">
              {activeFilters && (
                <button className="btn btn-ghost btn-sm" onClick={clearFilters}>✕ Limpiar Filtros</button>
              )}
              <button className="btn btn-primary btn-sm" onClick={() => setModalNew(true)}>
                <Icons.Plus /> Nueva Póliza (Caro)
              </button>
            </div>
          </div>
          <div className="filters-bar">
            <div className="search-wrapper">
              <Icons.Search />
              <input className="input input-search" value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Buscar nombre, póliza, bien..." />
            </div>
            <div className="filter-group">
              <span className="filter-label">Agente</span>
              <select className="select" style={{minWidth:130}} value={filterAgente}
                onChange={e => setFilterAgente(e.target.value)}>
                <option value="TODOS">Todos</option>
                {agentOpts.map(opt => <option key={opt} value={opt}>{opt}</option>)}
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Estatus</span>
              <select className="select" style={{minWidth:140}} value={filterEstatus}
                onChange={e => setFilterEstatus(e.target.value)}>
                <option value="TODOS">Todos</option>
                <option value="PENDIENTE">PENDIENTE</option>
                <option value="URGENTES">PRÓX. A VENCER (4D)</option>
                <option value="VENCIDO">VENCIDO</option>
                <option value="PAGADO">PAGADO</option>
                <option value="CANCELADO">CANCELADO</option>
                <option value="LIQUIDADO">LIQUIDADO</option>
                <option value="RENOVACIONES">RENOVACIONES (Próximas)</option>
                <option value="COMPROBANTES">COMPROBANTES</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Forma de Pago</span>
              <select className="select" style={{minWidth:140}} value={filterForma}
                onChange={e => setFilterForma(e.target.value)}>
                <option value="TODOS">Todas</option>
                <option value="CONTADO">CONTADO</option>
                <option value="MENSUAL">MENSUAL</option>
                <option value="TRIMESTRAL">TRIMESTRAL</option>
                <option value="SEMESTRAL">SEMESTRAL</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha desde</span>
              <input type="date" className="input" style={{width:140}} value={dateFrom}
                onChange={e => setDateFrom(e.target.value)} />
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha hasta</span>
              <input type="date" className="input" style={{width:140}} value={dateTo}
                onChange={e => setDateTo(e.target.value)} />
            </div>
          </div>
        </div>
      </div>

      {filterEstatus !== 'COMPROBANTES' && (
        <div className="card">
          <div className="card-header">
            <span className="card-title">
              {filterEstatus === 'TODOS' ? 'Todas las Pólizas de Caro' : `Pólizas (${filterEstatus})`} 
              {' '}({filtered.length})
            </span>
            {activeFilters && (
              <button className="btn btn-ghost btn-sm" onClick={clearFilters}>↩ Mostrar Todas</button>
            )}
          </div>
          <PoliciesTable 
            policies={filtered}
            onEdit={setModalEdit}
            onDelete={setDeleteConfirm}
            onMarkPaid={setModalPaid}
            onWhatsApp={onWhatsApp}
            onEmail={onEmail}
          />
          {filtered.length > 0 && (
            <div style={{padding:'12px 24px', borderTop:'1px solid var(--border)', fontSize:12, color:'var(--text-muted)', display:'flex', justifyContent:'space-between'}}>
              <span>{filtered.length} registro(s) encontrado(s)</span>
              <span>Total filtrado: <strong style={{color:'var(--accent-green)'}}>{formatMoney(filtered.reduce((s,p) => s+Number(p.monto||0), 0))}</strong></span>
            </div>
          )}
        </div>
      )}

      {/* Vista de Comprobantes (cuando se filtra por COMPROBANTES) */}
      {filterEstatus === 'COMPROBANTES' && (() => {
        const withComprobantes = policies.filter(p => p.comprobante);
        if (withComprobantes.length === 0) return (
          <div className="card" style={{marginTop: 20}}>
            <div style={{padding: 40, textAlign: 'center', color: 'var(--text-muted)'}}>
              🧾 Aún no hay comprobantes guardados en Clave Caro.
            </div>
          </div>
        );
        const grouped = {};
        withComprobantes.forEach(p => {
          const dStr = p.fechaUltimoPago || new Date().toISOString().split('T')[0];
          const date = new Date(dStr + 'T12:00:00');
          const monthYear = date.toLocaleDateString('es-MX', { month: 'long', year: 'numeric' });
          const capitalized = monthYear.charAt(0).toUpperCase() + monthYear.slice(1);
          if (!grouped[capitalized]) grouped[capitalized] = [];
          grouped[capitalized].push(p);
        });
        return Object.entries(grouped).map(([monthName, groupPolicies]) => (
          <div key={monthName} className="card" style={{marginTop: 20}}>
            <div className="card-header">
              <span className="card-title">📁 {monthName}</span>
              <span style={{fontSize:12, color:'var(--text-muted)'}}>{groupPolicies.length} comprobante(s)</span>
            </div>
            <div style={{padding: '20px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '20px'}}>
              {groupPolicies.map(p => (
                <div key={p.id} style={{border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: 12, background: 'var(--bg-card)'}}>
                  <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start'}}>
                    <div style={{fontWeight: 600, fontSize: 14, marginBottom: 4}}>{p.nombre}</div>
                    <button 
                      title="Eliminar comprobante"
                      onClick={() => { if (confirm('¿Eliminar este comprobante?')) onSave({ ...p, comprobante: null }); }}
                      style={{background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 6, padding: '4px 8px', cursor: 'pointer', fontSize: 13, color: '#ef4444', flexShrink: 0}}
                    >🗑️</button>
                  </div>
                  <div style={{fontSize: 12, color: 'var(--text-secondary)', marginBottom: 8}}>
                    <strong>Póliza:</strong> {p.poliza}
                  </div>
                  <div style={{fontSize: 11, color: 'var(--text-muted)', marginBottom: 4}}>
                    <strong>Fecha límite:</strong> {formatDate(p.fechaPagoAnterior || p.fechaPago)}
                  </div>
                  <div style={{fontSize: 11, color: 'var(--accent-green)', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6}}>
                    <strong>Fecha pagado:</strong> 
                    <input type="date" 
                      defaultValue={p.fechaUltimoPago || todayISO()} 
                      onBlur={e => {
                        const val = e.target.value;
                        if (val && val !== p.fechaUltimoPago) {
                          onSave({ ...p, fechaUltimoPago: val });
                          if (toast) toast('Fecha de pago actualizada ✅', 'success');
                        }
                      }}
                      style={{fontSize: 11, padding: '2px 6px', border: '1px solid var(--accent-green)', borderRadius: 4, background: 'var(--bg-input)', color: 'var(--accent-green)', fontWeight: 'bold'}}
                    />
                  </div>
                  <div style={{width: '100%', height: 200, borderRadius: 'var(--radius-sm)', overflow: 'hidden', border: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg-secondary)', cursor: 'pointer'}}
                    onClick={() => setSelectedImg(p.comprobante)}>
                    {p.comprobante.startsWith('data:application/pdf') ? (
                      <embed src={p.comprobante} width="100%" height="100%" type="application/pdf" style={{pointerEvents: 'none'}} />
                    ) : (
                      <img src={p.comprobante} alt="Comprobante" style={{width: '100%', height: '100%', objectFit: 'contain'}} />
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ));
      })()}

      {modalNew && <PolicyModal agentOptions={['DANIEL', 'JULIO']} onSave={onSave} onClose={() => setModalNew(false)} toast={toast} />}
      {modalEdit && <PolicyModal agentOptions={['DANIEL', 'JULIO']} policy={modalEdit} onSave={onSave} onClose={() => setModalEdit(null)} toast={toast} />}
      {modalPaid && <MarkPaidModal policy={modalPaid} onConfirm={(p, n, c, isLast) => { onMarkPaid(p, n, c, isLast); setModalPaid(null); }} onClose={() => setModalPaid(null)} toast={toast} />}
      
      {deleteConfirm && (
        <div className="modal-overlay">
          <div className="modal" style={{maxWidth: 400}}>
            <div className="modal-body" style={{textAlign: 'center', padding: '30px 20px'}}>
              <div style={{fontSize:40, marginBottom:16}}>⚠️</div>
              <h3 style={{marginBottom:10}}>¿Eliminar Póliza?</h3>
              <p style={{color:'var(--text-secondary)', marginBottom:24}}>
                Se borrará permanentemente la póliza de <strong>{deleteConfirm.nombre}</strong>.
              </p>
              <div className="flex gap-2" style={{justifyContent:'center'}}>
                <button className="btn btn-outline" onClick={() => setDeleteConfirm(null)}>Cancelar</button>
                <button className="btn btn-danger" onClick={() => { onDelete(deleteConfirm.id); setDeleteConfirm(null); }}>Sí, eliminar</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {selectedImg && <ImageModal src={selectedImg} onClose={() => setSelectedImg(null)} />}
    </div>
  );
}

// ─── Página: Gastos Médicos Mayores (GMM - Pólizas Separadas) ──
function GmmPoliciesPage({ policies, onSave, onDelete, onMarkPaid, onWhatsApp, onEmail, toast }) {
  const [modalNew, setModalNew] = useState(false);
  const [modalEdit, setModalEdit] = useState(null);
  const [modalPaid, setModalPaid] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [selectedImg, setSelectedImg] = useState(null);

  useEscapeKey(deleteConfirm ? () => setDeleteConfirm(null) : null);

  const [search, setSearch] = useState('');
  const [filterAgente, setFilterAgente] = useState('TODOS');
  const [filterEstatus, setFilterEstatus] = useState('TODOS');
  const [filterForma, setFilterForma] = useState('TODOS');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  const stats = useMemo(() => {
    const list = policies || [];
    const total = list.length;
    const pagados = list.filter(p => p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO').length;
    const pendientes = list.filter(p => p.estatus === 'PENDIENTE').length;
    const vencidos = list.filter(p => isExpiredEffective(p)).length;
    const urgentes = list.filter(p => isUpcomingReminder(p)).length;
    const renovaciones = list.filter(p => isUpcomingRenewal(p)).length;
    const comprobantes = list.filter(p => p.comprobante).length;
    return { total, pagados, pendientes, vencidos, urgentes, renovaciones, comprobantes };
  }, [policies]);

  const agentOpts = useMemo(() => {
    const list = policies || [];
    return Array.from(new Set(list.flatMap(p => [p.agente, p.aseguradora]).filter(Boolean)));
  }, [policies]);

  const filtered = useMemo(() => {
    const list = policies || [];
    return list.filter(p => {
      if (!p) return false;
      const q = search.toLowerCase().trim();
      const matchName = (p.nombre || '').toLowerCase().includes(q);
      const matchPoliza = (p.poliza || '').toLowerCase().includes(q);
      const matchBien = (p.bien || '').toLowerCase().includes(q);
      if (q && !matchName && !matchPoliza && !matchBien) return false;
      if (filterAgente !== 'TODOS' && p.agente !== filterAgente && p.aseguradora !== filterAgente) return false;
      if (filterEstatus !== 'TODOS') {
        if (filterEstatus === 'RENOVACIONES') {
          if (!isUpcomingRenewal(p)) return false;
        } else if (filterEstatus === 'URGENTES') {
          if (!isUpcomingReminder(p)) return false;
        } else if (filterEstatus === 'VENCIDO') {
          if (!isExpiredEffective(p)) return false;
        } else if (filterEstatus === 'COMPROBANTES') {
          if (!p.comprobante) return false;
        } else if (filterEstatus === 'PAGADO') {
          if (p.estatus !== 'PAGADO' && p.estatus !== 'LIQUIDADO') return false;
        } else {
          if (p.estatus !== filterEstatus) return false;
        }
      }
      if (filterForma !== 'TODOS' && p.formaPago !== filterForma) return false;
      if (dateFrom && p.fechaPago < dateFrom) return false;
      if (dateTo && p.fechaPago > dateTo) return false;
      return true;
    });
  }, [policies, search, filterAgente, filterEstatus, filterForma, dateFrom, dateTo]);

  const clearFilters = () => {
    setSearch(''); setFilterAgente('TODOS'); setFilterEstatus('TODOS');
    setFilterForma('TODOS'); setDateFrom(''); setDateTo('');
  };

  const activeFilters = filterAgente !== 'TODOS' || filterEstatus !== 'TODOS' ||
    filterForma !== 'TODOS' || dateFrom || dateTo || search;

  return (
    <div className="page-fade-enter">
      {/* Tarjetas KPI como filtros */}
      <div className="stats-grid" style={{marginBottom: 20}}>
        {[
          { label: 'Total Pólizas GMM', value: stats.total, icon: '🏥', cls: 'stat-blue', filter: 'TODOS' },
          { label: 'Pendientes', value: stats.pendientes, icon: '⏳', cls: 'stat-yellow', filter: 'PENDIENTE' },
          { label: 'Próx. a Vencer (4d)', value: stats.urgentes, icon: '🔴', cls: 'stat-orange', filter: 'URGENTES' },
          { label: 'Vencidos', value: stats.vencidos, icon: '🛑', cls: 'stat-red', filter: 'VENCIDO' },
          { label: 'Renovaciones', value: stats.renovaciones, icon: '🔄', cls: 'stat-purple', filter: 'RENOVACIONES' },
          { label: 'Pagados', value: stats.pagados, icon: '✅', cls: 'stat-green', filter: 'PAGADO' },
          { label: 'Comprobantes', value: stats.comprobantes, icon: '🧾', cls: 'stat-orange', filter: 'COMPROBANTES' },
        ].map(s => (
          <div key={s.label} className={`stat-card ${s.cls}`} 
            style={{
              cursor: 'pointer', 
              opacity: filterEstatus === s.filter || filterEstatus === 'TODOS' ? 1 : 0.5,
              border: filterEstatus === s.filter ? '2px solid currentColor' : '1px solid transparent',
              transition: 'all 0.2s ease'
            }} 
            onClick={() => setFilterEstatus(s.filter)}>
            <div className="stat-card-icon">{s.icon}</div>
            <div className="stat-card-value">{s.value}</div>
            <div className="stat-card-label">{s.label}</div>
          </div>
        ))}
      </div>

      {/* Barra de Filtros */}
      <div className="card" style={{marginBottom: 20}}>
        <div className="card-header" style={{flexDirection:'column', alignItems:'flex-start', gap:14}}>
          <div className="flex justify-between w-full items-center">
            <span className="card-title">🏥 Pólizas Gastos Médicos Mayores (GMM) ({filtered.length})</span>
            <div className="flex gap-2">
              {activeFilters && (
                <button className="btn btn-ghost btn-sm" onClick={clearFilters}>✕ Limpiar Filtros</button>
              )}
              <button className="btn btn-primary btn-sm" onClick={() => setModalNew(true)}>
                <Icons.Plus /> Nueva Póliza GMM
              </button>
            </div>
          </div>
          <div className="filters-bar">
            <div className="search-wrapper">
              <Icons.Search />
              <input className="input input-search" value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Buscar asegurado, póliza, plan..." />
            </div>
            <div className="filter-group">
              <span className="filter-label">Agente / Aseg.</span>
              <select className="select" style={{minWidth:130}} value={filterAgente}
                onChange={e => setFilterAgente(e.target.value)}>
                <option value="TODOS">Todos</option>
                {agentOpts.map(opt => <option key={opt} value={opt}>{opt}</option>)}
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Estatus</span>
              <select className="select" style={{minWidth:140}} value={filterEstatus}
                onChange={e => setFilterEstatus(e.target.value)}>
                <option value="TODOS">Todos</option>
                <option value="PENDIENTE">PENDIENTE</option>
                <option value="URGENTES">PRÓX. A VENCER (4D)</option>
                <option value="VENCIDO">VENCIDO</option>
                <option value="PAGADO">PAGADO</option>
                <option value="CANCELADO">CANCELADO</option>
                <option value="LIQUIDADO">LIQUIDADO</option>
                <option value="RENOVACIONES">RENOVACIONES (Próximas)</option>
                <option value="COMPROBANTES">COMPROBANTES</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Forma de Pago</span>
              <select className="select" style={{minWidth:140}} value={filterForma}
                onChange={e => setFilterForma(e.target.value)}>
                <option value="TODOS">Todas</option>
                <option value="CONTADO">CONTADO</option>
                <option value="MENSUAL">MENSUAL</option>
                <option value="TRIMESTRAL">TRIMESTRAL</option>
                <option value="SEMESTRAL">SEMESTRAL</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha desde</span>
              <input type="date" className="input" style={{width:140}} value={dateFrom}
                onChange={e => setDateFrom(e.target.value)} />
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha hasta</span>
              <input type="date" className="input" style={{width:140}} value={dateTo}
                onChange={e => setDateTo(e.target.value)} />
            </div>
          </div>
        </div>
      </div>

      {filterEstatus !== 'COMPROBANTES' && (
        <div className="card">
          <div className="card-header">
            <span className="card-title">
              {filterEstatus === 'TODOS' ? 'Todas las Pólizas de GMM' : `Pólizas GMM (${filterEstatus})`} 
              {' '}({filtered.length})
            </span>
            {activeFilters && (
              <button className="btn btn-ghost btn-sm" onClick={clearFilters}>↩ Mostrar Todas</button>
            )}
          </div>
          <PoliciesTable 
            policies={filtered}
            onEdit={setModalEdit}
            onDelete={setDeleteConfirm}
            onMarkPaid={setModalPaid}
            onWhatsApp={onWhatsApp}
            onEmail={onEmail}
          />
          {filtered.length > 0 && (
            <div style={{padding:'12px 24px', borderTop:'1px solid var(--border)', fontSize:12, color:'var(--text-muted)', display:'flex', justifyContent:'space-between'}}>
              <span>{filtered.length} registro(s) encontrado(s)</span>
              <span>Total filtrado: <strong style={{color:'var(--accent-green)'}}>{formatMoney(filtered.reduce((s,p) => s+Number(p.monto||0), 0))}</strong></span>
            </div>
          )}
        </div>
      )}

      {/* Vista de Comprobantes GMM */}
      {filterEstatus === 'COMPROBANTES' && (() => {
        const withComprobantes = policies.filter(p => p.comprobante);
        if (withComprobantes.length === 0) return (
          <div className="card" style={{marginTop: 20}}>
            <div style={{padding: 40, textAlign: 'center', color: 'var(--text-muted)'}}>
              🧾 Aún no hay comprobantes guardados en GMM.
            </div>
          </div>
        );
        const grouped = {};
        withComprobantes.forEach(p => {
          const dStr = p.fechaUltimoPago || new Date().toISOString().split('T')[0];
          const date = new Date(dStr + 'T12:00:00');
          const monthYear = date.toLocaleDateString('es-MX', { month: 'long', year: 'numeric' });
          const capitalized = monthYear.charAt(0).toUpperCase() + monthYear.slice(1);
          if (!grouped[capitalized]) grouped[capitalized] = [];
          grouped[capitalized].push(p);
        });

        return (
          <div style={{marginTop: 24}}>
            <h3 style={{fontSize: 16, marginBottom: 16}}>🧾 Comprobantes GMM por Mes</h3>
            {Object.entries(grouped).map(([monthName, groupPolicies]) => (
              <div key={monthName} className="card" style={{marginBottom: 20}}>
                <div className="card-header" style={{background: 'var(--bg-secondary)', padding: '12px 20px'}}>
                  <span className="card-title">📁 {monthName}</span>
                  <span style={{fontSize:12, color:'var(--text-muted)'}}>{groupPolicies.length} comprobante(s)</span>
                </div>
                <div style={{padding: '20px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '20px'}}>
                  {groupPolicies.map(p => (
                    <div key={p.id} style={{border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: 12, background: 'var(--bg-card)'}}>
                      <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start'}}>
                        <div style={{fontWeight: 600, fontSize: 14, marginBottom: 4}}>{p.nombre}</div>
                        <button 
                          title="Eliminar comprobante"
                          onClick={() => { if (confirm('¿Eliminar este comprobante?')) onSave({ ...p, comprobante: null }); }}
                          style={{background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 6, padding: '4px 8px', cursor: 'pointer', fontSize: 13, color: '#ef4444', flexShrink: 0}}
                        >🗑️</button>
                      </div>
                      <div style={{fontSize: 12, color: 'var(--text-secondary)', marginBottom: 8}}>
                        <strong>Póliza:</strong> {p.poliza}
                      </div>
                      <div style={{fontSize: 11, color: 'var(--text-muted)', marginBottom: 4}}>
                        <strong>Fecha límite:</strong> {formatDate(p.fechaPagoAnterior || p.fechaPago)}
                      </div>
                      <div style={{fontSize: 11, color: 'var(--accent-green)', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6}}>
                        <strong>Fecha pagado:</strong> 
                        <input type="date" 
                          defaultValue={p.fechaUltimoPago || todayISO()} 
                          onBlur={e => {
                            const val = e.target.value;
                            if (val && val !== p.fechaUltimoPago) {
                              onSave({ ...p, fechaUltimoPago: val });
                              if (toast) toast('Fecha de pago actualizada ✅', 'success');
                            }
                          }}
                          style={{fontSize: 11, padding: '2px 6px', border: '1px solid var(--accent-green)', borderRadius: 4, background: 'var(--bg-input)', color: 'var(--accent-green)', fontWeight: 'bold'}}
                        />
                      </div>
                      <div style={{width: '100%', height: 200, borderRadius: 'var(--radius-sm)', overflow: 'hidden', border: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg-secondary)', cursor: 'pointer'}}
                           onClick={() => setSelectedImg(p.comprobante)}>
                        {p.comprobante.startsWith('data:application/pdf') ? (
                          <embed src={p.comprobante} width="100%" height="100%" type="application/pdf" style={{pointerEvents: 'none'}} />
                        ) : (
                          <img src={p.comprobante} alt={`Comprobante GMM`} style={{width: '100%', height: '100%', objectFit: 'contain'}} />
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        );
      })()}

      {/* Modales */}
      {modalNew && (
        <PolicyModal 
          isGmm={true}
          onSave={(p) => { onSave(p); setModalNew(false); }} 
          onClose={() => setModalNew(false)} 
          toast={toast}
          agentOptions={['DANIEL', 'OTRO']}
        />
      )}
      {modalEdit && (
        <PolicyModal 
          isGmm={true}
          policy={modalEdit} 
          onSave={(p) => { onSave(p); setModalEdit(null); }} 
          onClose={() => setModalEdit(null)} 
          toast={toast}
          agentOptions={['DANIEL', 'OTRO']}
        />
      )}
      {modalPaid && (
        <MarkPaidModal 
          policy={modalPaid} 
          onConfirm={(p, nextDate, comp, isLast) => { onMarkPaid(p, nextDate, comp, isLast); setModalPaid(null); }} 
          onClose={() => setModalPaid(null)} 
        />
      )}
      {deleteConfirm && (
        <div className="modal-overlay">
          <div className="modal" style={{maxWidth:420}} onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>🗑️ Confirmar Eliminación</h2>
              <button className="modal-close" onClick={() => setDeleteConfirm(null)}><Icons.Close /></button>
            </div>
            <div className="modal-body">
              <p>¿Eliminar la póliza de <strong>{deleteConfirm.nombre}</strong> (Póliza GMM: {deleteConfirm.poliza})?</p>
            </div>
            <div className="modal-footer">
              <button className="btn btn-outline" onClick={() => setDeleteConfirm(null)}>Cancelar</button>
              <button className="btn btn-danger" onClick={() => { onDelete(deleteConfirm.id); setDeleteConfirm(null); }}>Eliminar</button>
            </div>
          </div>
        </div>
      )}
      {selectedImg && <ImageModal src={selectedImg} onClose={() => setSelectedImg(null)} />}
    </div>
  );
}

// ─── Autos Otras Aseguradoras Page ────────────────────────────
function AutosOtrasPoliciesPage({ policies, onSave, onDelete, onMarkPaid, onWhatsApp, onEmail, toast }) {
  const [modalNew, setModalNew] = useState(false);
  const [modalEdit, setModalEdit] = useState(null);
  const [modalPaid, setModalPaid] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [selectedImg, setSelectedImg] = useState(null);

  useEscapeKey(deleteConfirm ? () => setDeleteConfirm(null) : null);

  const [search, setSearch] = useState('');
  const [filterAgente, setFilterAgente] = useState('TODOS');
  const [filterEstatus, setFilterEstatus] = useState('TODOS');
  const [filterForma, setFilterForma] = useState('TODOS');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  const stats = useMemo(() => {
    const list = policies || [];
    const total = list.length;
    const pagados = list.filter(p => p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO').length;
    const pendientes = list.filter(p => p.estatus === 'PENDIENTE').length;
    const vencidos = list.filter(p => isExpiredEffective(p)).length;
    const urgentes = list.filter(p => isUpcomingReminder(p)).length;
    const renovaciones = list.filter(p => isUpcomingRenewal(p)).length;
    const comprobantes = list.filter(p => p.comprobante).length;
    return { total, pagados, pendientes, vencidos, urgentes, renovaciones, comprobantes };
  }, [policies]);

  const agentOpts = useMemo(() => {
    const list = policies || [];
    return Array.from(new Set(list.flatMap(p => [p.agente, p.aseguradora]).filter(Boolean)));
  }, [policies]);

  const filtered = useMemo(() => {
    const list = policies || [];
    return list.filter(p => {
      if (!p) return false;
      const q = search.toLowerCase().trim();
      const matchName = (p.nombre || '').toLowerCase().includes(q);
      const matchPoliza = (p.poliza || '').toLowerCase().includes(q);
      const matchBien = (p.bien || '').toLowerCase().includes(q);
      if (q && !matchName && !matchPoliza && !matchBien) return false;
      if (filterAgente !== 'TODOS' && p.agente !== filterAgente && p.aseguradora !== filterAgente) return false;
      if (filterEstatus !== 'TODOS') {
        if (filterEstatus === 'RENOVACIONES') {
          if (!isUpcomingRenewal(p)) return false;
        } else if (filterEstatus === 'URGENTES') {
          if (!isUpcomingReminder(p)) return false;
        } else if (filterEstatus === 'VENCIDO') {
          if (!isExpiredEffective(p)) return false;
        } else if (filterEstatus === 'COMPROBANTES') {
          if (!p.comprobante) return false;
        } else if (filterEstatus === 'PAGADO') {
          if (p.estatus !== 'PAGADO' && p.estatus !== 'LIQUIDADO') return false;
        } else {
          if (p.estatus !== filterEstatus) return false;
        }
      }
      if (filterForma !== 'TODOS' && p.formaPago !== filterForma) return false;
      if (dateFrom && p.fechaPago < dateFrom) return false;
      if (dateTo && p.fechaPago > dateTo) return false;
      return true;
    });
  }, [policies, search, filterAgente, filterEstatus, filterForma, dateFrom, dateTo]);

  const clearFilters = () => {
    setSearch(''); setFilterAgente('TODOS'); setFilterEstatus('TODOS');
    setFilterForma('TODOS'); setDateFrom(''); setDateTo('');
  };

  const activeFilters = filterAgente !== 'TODOS' || filterEstatus !== 'TODOS' ||
    filterForma !== 'TODOS' || dateFrom || dateTo || search;

  return (
    <div className="page-fade-enter">
      {/* Tarjetas KPI como filtros */}
      <div className="stats-grid" style={{marginBottom: 20}}>
        {[
          { label: 'Total Pólizas Autos', value: stats.total, icon: '🚗', cls: 'stat-blue', filter: 'TODOS' },
          { label: 'Pendientes', value: stats.pendientes, icon: '⏳', cls: 'stat-yellow', filter: 'PENDIENTE' },
          { label: 'Próx. a Vencer (4d)', value: stats.urgentes, icon: '🔴', cls: 'stat-orange', filter: 'URGENTES' },
          { label: 'Vencidos', value: stats.vencidos, icon: '🛑', cls: 'stat-red', filter: 'VENCIDO' },
          { label: 'Renovaciones', value: stats.renovaciones, icon: '🔄', cls: 'stat-purple', filter: 'RENOVACIONES' },
          { label: 'Pagados', value: stats.pagados, icon: '✅', cls: 'stat-green', filter: 'PAGADO' },
          { label: 'Comprobantes', value: stats.comprobantes, icon: '🧾', cls: 'stat-orange', filter: 'COMPROBANTES' },
        ].map(s => (
          <div key={s.label} className={`stat-card ${s.cls}`} 
            style={{
              cursor: 'pointer', 
              opacity: filterEstatus === s.filter || filterEstatus === 'TODOS' ? 1 : 0.5,
              border: filterEstatus === s.filter ? '2px solid currentColor' : '1px solid transparent',
              transition: 'all 0.2s ease'
            }} 
            onClick={() => setFilterEstatus(s.filter)}>
            <div className="stat-card-icon">{s.icon}</div>
            <div className="stat-card-value">{s.value}</div>
            <div className="stat-card-label">{s.label}</div>
          </div>
        ))}
      </div>

      {/* Barra de Filtros */}
      <div className="card" style={{marginBottom: 20}}>
        <div className="card-header" style={{flexDirection:'column', alignItems:'flex-start', gap:14}}>
          <div className="flex justify-between w-full items-center">
            <span className="card-title">🚗 Autos (Otras Aseguradoras) ({filtered.length})</span>
            <div className="flex gap-2">
              {activeFilters && (
                <button className="btn btn-ghost btn-sm" onClick={clearFilters}>✕ Limpiar Filtros</button>
              )}
              <button className="btn btn-primary btn-sm" onClick={() => setModalNew(true)}>
                <Icons.Plus /> Nueva Póliza Auto
              </button>
            </div>
          </div>
          <div className="filters-bar">
            <div className="search-wrapper">
              <Icons.Search />
              <input className="input input-search" value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Buscar asegurado, póliza, vehículo..." />
            </div>
            <div className="filter-group">
              <span className="filter-label">Agente / Aseg.</span>
              <select className="select" style={{minWidth:130}} value={filterAgente}
                onChange={e => setFilterAgente(e.target.value)}>
                <option value="TODOS">Todos</option>
                {agentOpts.map(opt => <option key={opt} value={opt}>{opt}</option>)}
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Estatus</span>
              <select className="select" style={{minWidth:140}} value={filterEstatus}
                onChange={e => setFilterEstatus(e.target.value)}>
                <option value="TODOS">Todos</option>
                <option value="PENDIENTE">PENDIENTE</option>
                <option value="URGENTES">PRÓX. A VENCER (4D)</option>
                <option value="VENCIDO">VENCIDO</option>
                <option value="PAGADO">PAGADO</option>
                <option value="CANCELADO">CANCELADO</option>
                <option value="LIQUIDADO">LIQUIDADO</option>
                <option value="RENOVACIONES">RENOVACIONES (Próximas)</option>
                <option value="COMPROBANTES">COMPROBANTES</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Forma de Pago</span>
              <select className="select" style={{minWidth:140}} value={filterForma}
                onChange={e => setFilterForma(e.target.value)}>
                <option value="TODOS">Todas</option>
                <option value="CONTADO">CONTADO</option>
                <option value="MENSUAL">MENSUAL</option>
                <option value="TRIMESTRAL">TRIMESTRAL</option>
                <option value="SEMESTRAL">SEMESTRAL</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha desde</span>
              <input type="date" className="input" style={{width:140}} value={dateFrom}
                onChange={e => setDateFrom(e.target.value)} />
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha hasta</span>
              <input type="date" className="input" style={{width:140}} value={dateTo}
                onChange={e => setDateTo(e.target.value)} />
            </div>
          </div>
        </div>
      </div>

      {filterEstatus !== 'COMPROBANTES' && (
        <div className="card">
          <div className="card-header">
            <span className="card-title">
              {filterEstatus === 'TODOS' ? 'Todas las Pólizas de Autos' : `Pólizas Autos (${filterEstatus})`} 
              {' '}({filtered.length})
            </span>
            {activeFilters && (
              <button className="btn btn-ghost btn-sm" onClick={clearFilters}>↩ Mostrar Todas</button>
            )}
          </div>
          <PoliciesTable 
            policies={filtered}
            onEdit={setModalEdit}
            onDelete={setDeleteConfirm}
            onMarkPaid={setModalPaid}
            onWhatsApp={onWhatsApp}
            onEmail={onEmail}
          />
          {filtered.length > 0 && (
            <div style={{padding:'12px 24px', borderTop:'1px solid var(--border)', fontSize:12, color:'var(--text-muted)', display:'flex', justifyContent:'space-between'}}>
              <span>{filtered.length} registro(s) encontrado(s)</span>
              <span>Total filtrado: <strong style={{color:'var(--accent-green)'}}>{formatMoney(filtered.reduce((s,p) => s+Number(p.monto||0), 0))}</strong></span>
            </div>
          )}
        </div>
      )}

      {/* Vista de Comprobantes Autos */}
      {filterEstatus === 'COMPROBANTES' && (() => {
        const withComprobantes = policies.filter(p => p.comprobante);
        if (withComprobantes.length === 0) return (
          <div className="card" style={{marginTop: 20}}>
            <div style={{padding: 40, textAlign: 'center', color: 'var(--text-muted)'}}>
              🧾 Aún no hay comprobantes guardados en Autos.
            </div>
          </div>
        );
        const grouped = {};
        withComprobantes.forEach(p => {
          const dStr = p.fechaUltimoPago || new Date().toISOString().split('T')[0];
          const date = new Date(dStr + 'T12:00:00');
          const monthYear = date.toLocaleDateString('es-MX', { month: 'long', year: 'numeric' });
          const capitalized = monthYear.charAt(0).toUpperCase() + monthYear.slice(1);
          if (!grouped[capitalized]) grouped[capitalized] = [];
          grouped[capitalized].push(p);
        });

        return (
          <div style={{marginTop: 24}}>
            <h3 style={{fontSize: 16, marginBottom: 16}}>🧾 Comprobantes Autos por Mes</h3>
            {Object.entries(grouped).map(([monthName, groupPolicies]) => (
              <div key={monthName} className="card" style={{marginBottom: 20}}>
                <div className="card-header" style={{background: 'var(--bg-secondary)', padding: '12px 20px'}}>
                  <span className="card-title">📁 {monthName}</span>
                  <span style={{fontSize:12, color:'var(--text-muted)'}}>{groupPolicies.length} comprobante(s)</span>
                </div>
                <div style={{padding: '20px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '20px'}}>
                  {groupPolicies.map(p => (
                    <div key={p.id} style={{border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: 12, background: 'var(--bg-card)'}}>
                      <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start'}}>
                        <div style={{fontWeight: 600, fontSize: 14, marginBottom: 4}}>{p.nombre}</div>
                        <button 
                          title="Eliminar comprobante"
                          onClick={() => { if (confirm('¿Eliminar este comprobante?')) onSave({ ...p, comprobante: null }); }}
                          style={{background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 6, padding: '4px 8px', cursor: 'pointer', fontSize: 13, color: '#ef4444', flexShrink: 0}}
                        >🗑️</button>
                      </div>
                      <div style={{fontSize: 12, color: 'var(--text-secondary)', marginBottom: 8}}>
                        <strong>Póliza:</strong> {p.poliza}
                      </div>
                      <div style={{fontSize: 11, color: 'var(--text-muted)', marginBottom: 4}}>
                        <strong>Fecha límite:</strong> {formatDate(p.fechaPagoAnterior || p.fechaPago)}
                      </div>
                      <div style={{fontSize: 11, color: 'var(--accent-green)', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6}}>
                        <strong>Fecha pagado:</strong> 
                        <input type="date" 
                          defaultValue={p.fechaUltimoPago || todayISO()} 
                          onBlur={e => {
                            const val = e.target.value;
                            if (val && val !== p.fechaUltimoPago) {
                              onSave({ ...p, fechaUltimoPago: val });
                              if (toast) toast('Fecha de pago actualizada ✅', 'success');
                            }
                          }}
                          style={{fontSize: 11, padding: '2px 6px', border: '1px solid var(--accent-green)', borderRadius: 4, background: 'var(--bg-input)', color: 'var(--accent-green)', fontWeight: 'bold'}}
                        />
                      </div>
                      <div style={{width: '100%', height: 200, borderRadius: 'var(--radius-sm)', overflow: 'hidden', border: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg-secondary)', cursor: 'pointer'}}
                           onClick={() => setSelectedImg(p.comprobante)}>
                        {p.comprobante.startsWith('data:application/pdf') ? (
                          <embed src={p.comprobante} width="100%" height="100%" type="application/pdf" style={{pointerEvents: 'none'}} />
                        ) : (
                          <img src={p.comprobante} alt={`Comprobante Auto`} style={{width: '100%', height: '100%', objectFit: 'contain'}} />
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        );
      })()}

      {/* Modales */}
      {modalNew && (
        <PolicyModal 
          isAutos={true}
          onSave={(p) => { onSave(p); setModalNew(false); }} 
          onClose={() => setModalNew(false)} 
          toast={toast}
          agentOptions={['DANIEL', 'OTRO']}
        />
      )}
      {modalEdit && (
        <PolicyModal 
          isAutos={true}
          policy={modalEdit} 
          onSave={(p) => { onSave(p); setModalEdit(null); }} 
          onClose={() => setModalEdit(null)} 
          toast={toast}
          agentOptions={['DANIEL', 'OTRO']}
        />
      )}
      {modalPaid && (
        <MarkPaidModal 
          policy={modalPaid} 
          onConfirm={(p, nextDate, comp, isLast) => { onMarkPaid(p, nextDate, comp, isLast); setModalPaid(null); }} 
          onClose={() => setModalPaid(null)} 
        />
      )}
      {deleteConfirm && (
        <div className="modal-overlay">
          <div className="modal" style={{maxWidth:420}} onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>🗑️ Confirmar Eliminación</h2>
              <button className="modal-close" onClick={() => setDeleteConfirm(null)}><Icons.Close /></button>
            </div>
            <div className="modal-body">
              <p>¿Eliminar la póliza de <strong>{deleteConfirm.nombre}</strong> (Póliza Auto: {deleteConfirm.poliza})?</p>
            </div>
            <div className="modal-footer">
              <button className="btn btn-outline" onClick={() => setDeleteConfirm(null)}>Cancelar</button>
              <button className="btn btn-danger" onClick={() => { onDelete(deleteConfirm.id); setDeleteConfirm(null); }}>Eliminar</button>
            </div>
          </div>
        </div>
      )}
      {selectedImg && <ImageModal src={selectedImg} onClose={() => setSelectedImg(null)} />}
    </div>
  );
}

// ─── Página Genérica para Nuevas Secciones (Vida, Daños, Hogar, etc.) ────
function SectionPoliciesPage({ 
  title, 
  icon = '📋', 
  policies, 
  onSave, 
  onDelete, 
  onMarkPaid, 
  onWhatsApp, 
  onEmail, 
  toast,
  isVida = false,
  isDanos = false,
  isHogar = false,
  isGmm = false,
  isAutos = false
}) {
  const [modalNew, setModalNew] = useState(false);
  const [modalEdit, setModalEdit] = useState(null);
  const [modalPaid, setModalPaid] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [selectedImg, setSelectedImg] = useState(null);

  useEscapeKey(deleteConfirm ? () => setDeleteConfirm(null) : null);

  const [search, setSearch] = useState('');
  const [filterAgente, setFilterAgente] = useState('TODOS');
  const [filterEstatus, setFilterEstatus] = useState('TODOS');
  const [filterForma, setFilterForma] = useState('TODOS');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  const stats = useMemo(() => {
    const list = policies || [];
    const total = list.length;
    const pagados = list.filter(p => p.estatus === 'PAGADO' || p.estatus === 'LIQUIDADO').length;
    const pendientes = list.filter(p => p.estatus === 'PENDIENTE').length;
    const vencidos = list.filter(p => isExpiredEffective(p)).length;
    const urgentes = list.filter(p => isUpcomingReminder(p)).length;
    const renovaciones = list.filter(p => isUpcomingRenewal(p)).length;
    const comprobantes = list.filter(p => p.comprobante).length;
    return { total, pagados, pendientes, vencidos, urgentes, renovaciones, comprobantes };
  }, [policies]);

  const agentOpts = useMemo(() => {
    const list = policies || [];
    return Array.from(new Set(list.flatMap(p => [p.agente, p.aseguradora]).filter(Boolean)));
  }, [policies]);

  const filtered = useMemo(() => {
    const list = policies || [];
    return list.filter(p => {
      if (!p) return false;
      const q = search.toLowerCase().trim();
      const matchName = (p.nombre || '').toLowerCase().includes(q);
      const matchPoliza = (p.poliza || '').toLowerCase().includes(q);
      const matchBien = (p.bien || '').toLowerCase().includes(q);
      if (q && !matchName && !matchPoliza && !matchBien) return false;
      if (filterAgente !== 'TODOS' && p.agente !== filterAgente && p.aseguradora !== filterAgente) return false;
      if (filterEstatus !== 'TODOS') {
        if (filterEstatus === 'RENOVACIONES') {
          if (!isUpcomingRenewal(p)) return false;
        } else if (filterEstatus === 'URGENTES') {
          if (!isUpcomingReminder(p)) return false;
        } else if (filterEstatus === 'VENCIDO') {
          if (!isExpiredEffective(p)) return false;
        } else if (filterEstatus === 'COMPROBANTES') {
          if (!p.comprobante) return false;
        } else if (filterEstatus === 'PAGADO') {
          if (p.estatus !== 'PAGADO' && p.estatus !== 'LIQUIDADO') return false;
        } else {
          if (p.estatus !== filterEstatus) return false;
        }
      }
      if (filterForma !== 'TODOS' && p.formaPago !== filterForma) return false;
      if (dateFrom && p.fechaPago < dateFrom) return false;
      if (dateTo && p.fechaPago > dateTo) return false;
      return true;
    });
  }, [policies, search, filterAgente, filterEstatus, filterForma, dateFrom, dateTo]);

  const clearFilters = () => {
    setSearch(''); setFilterAgente('TODOS'); setFilterEstatus('TODOS');
    setFilterForma('TODOS'); setDateFrom(''); setDateTo('');
  };

  const activeFilters = filterAgente !== 'TODOS' || filterEstatus !== 'TODOS' ||
    filterForma !== 'TODOS' || dateFrom || dateTo || search;

  return (
    <div className="page-fade-enter">
      {/* Tarjetas KPI */}
      <div className="stats-grid" style={{marginBottom: 20}}>
        {[
          { label: `Total ${title}`, value: stats.total, icon: icon, cls: 'stat-blue', filter: 'TODOS' },
          { label: 'Pendientes', value: stats.pendientes, icon: '⏳', cls: 'stat-yellow', filter: 'PENDIENTE' },
          { label: 'Próx. a Vencer (4d)', value: stats.urgentes, icon: '🔴', cls: 'stat-orange', filter: 'URGENTES' },
          { label: 'Vencidos', value: stats.vencidos, icon: '🛑', cls: 'stat-red', filter: 'VENCIDO' },
          { label: 'Renovaciones', value: stats.renovaciones, icon: '🔄', cls: 'stat-purple', filter: 'RENOVACIONES' },
          { label: 'Pagados', value: stats.pagados, icon: '✅', cls: 'stat-green', filter: 'PAGADO' },
          { label: 'Comprobantes', value: stats.comprobantes, icon: '🧾', cls: 'stat-orange', filter: 'COMPROBANTES' },
        ].map(s => (
          <div key={s.label} className={`stat-card ${s.cls}`} 
            style={{
              cursor: 'pointer', 
              opacity: filterEstatus === s.filter || filterEstatus === 'TODOS' ? 1 : 0.5,
              border: filterEstatus === s.filter ? '2px solid currentColor' : '1px solid transparent',
              transition: 'all 0.2s ease'
            }} 
            onClick={() => setFilterEstatus(s.filter)}>
            <div className="stat-card-icon">{s.icon}</div>
            <div className="stat-card-value">{s.value}</div>
            <div className="stat-card-label">{s.label}</div>
          </div>
        ))}
      </div>

      {/* Barra de Filtros */}
      <div className="card" style={{marginBottom: 20}}>
        <div className="card-header" style={{flexDirection:'column', alignItems:'flex-start', gap:14}}>
          <div className="flex justify-between w-full items-center">
            <span className="card-title">{icon} Pólizas de {title} ({filtered.length})</span>
            <div className="flex gap-2">
              {activeFilters && (
                <button className="btn btn-ghost btn-sm" onClick={clearFilters}>✕ Limpiar Filtros</button>
              )}
              <button className="btn btn-primary btn-sm" onClick={() => setModalNew(true)}>
                <Icons.Plus /> Nueva Póliza ({title})
              </button>
            </div>
          </div>
          <div className="filters-bar">
            <div className="search-wrapper">
              <Icons.Search />
              <input className="input input-search" value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Buscar nombre, póliza..." />
            </div>
            <div className="filter-group">
              <span className="filter-label">Agente / Aseguradora</span>
              <select className="select" style={{minWidth:130}} value={filterAgente}
                onChange={e => setFilterAgente(e.target.value)}>
                <option value="TODOS">Todos</option>
                {agentOpts.map(opt => <option key={opt} value={opt}>{opt}</option>)}
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Estatus</span>
              <select className="select" style={{minWidth:140}} value={filterEstatus}
                onChange={e => setFilterEstatus(e.target.value)}>
                <option value="TODOS">Todos</option>
                <option value="PENDIENTE">PENDIENTE</option>
                <option value="URGENTES">PRÓX. A VENCER (4D)</option>
                <option value="VENCIDO">VENCIDO</option>
                <option value="PAGADO">PAGADO</option>
                <option value="CANCELADO">CANCELADO</option>
                <option value="LIQUIDADO">LIQUIDADO</option>
                <option value="RENOVACIONES">RENOVACIONES (Próximas)</option>
                <option value="COMPROBANTES">COMPROBANTES</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Forma de Pago</span>
              <select className="select" style={{minWidth:140}} value={filterForma}
                onChange={e => setFilterForma(e.target.value)}>
                <option value="TODOS">Todas</option>
                <option value="CONTADO">CONTADO</option>
                <option value="MENSUAL">MENSUAL</option>
                <option value="TRIMESTRAL">TRIMESTRAL</option>
                <option value="SEMESTRAL">SEMESTRAL</option>
              </select>
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha desde</span>
              <input type="date" className="input" style={{width:140}} value={dateFrom}
                onChange={e => setDateFrom(e.target.value)} />
            </div>
            <div className="filter-group">
              <span className="filter-label">Fecha hasta</span>
              <input type="date" className="input" style={{width:140}} value={dateTo}
                onChange={e => setDateTo(e.target.value)} />
            </div>
          </div>
        </div>
      </div>

      {filterEstatus !== 'COMPROBANTES' && (
        <div className="card">
          <div className="card-header">
            <span className="card-title">
              {filterEstatus === 'TODOS' ? `Todas las Pólizas de ${title}` : `Pólizas (${filterEstatus})`} 
              {' '}({filtered.length})
            </span>
            {activeFilters && (
              <button className="btn btn-ghost btn-sm" onClick={clearFilters}>↩ Mostrar Todas</button>
            )}
          </div>
          <PoliciesTable 
            policies={filtered}
            onEdit={setModalEdit}
            onDelete={setDeleteConfirm}
            onMarkPaid={setModalPaid}
            onWhatsApp={onWhatsApp}
            onEmail={onEmail}
          />
          {filtered.length > 0 && (
            <div style={{padding:'12px 24px', borderTop:'1px solid var(--border)', fontSize:12, color:'var(--text-muted)', display:'flex', justifyContent:'space-between'}}>
              <span>{filtered.length} registro(s) encontrado(s)</span>
              <span>Total filtrado: <strong style={{color:'var(--accent-green)'}}>{formatMoney(filtered.reduce((s,p) => s+Number(p.monto||0), 0))}</strong></span>
            </div>
          )}
        </div>
      )}

      {/* Vista de Comprobantes por Mes */}
      {filterEstatus === 'COMPROBANTES' && (() => {
        const withComprobantes = (policies || []).filter(p => p.comprobante);
        if (withComprobantes.length === 0) return (
          <div className="card" style={{marginTop: 20}}>
            <div style={{padding: 40, textAlign: 'center', color: 'var(--text-muted)'}}>
              🧾 Aún no hay comprobantes guardados en {title}.
            </div>
          </div>
        );
        const grouped = {};
        withComprobantes.forEach(p => {
          const dStr = p.fechaUltimoPago || todayISO();
          const date = new Date(dStr + 'T12:00:00');
          const monthYear = date.toLocaleDateString('es-MX', { month: 'long', year: 'numeric' });
          const capitalized = monthYear.charAt(0).toUpperCase() + monthYear.slice(1);
          if (!grouped[capitalized]) grouped[capitalized] = [];
          grouped[capitalized].push(p);
        });

        return (
          <div style={{marginTop: 24}}>
            <h3 style={{fontSize: 16, marginBottom: 16}}>🧾 Comprobantes {title} por Mes</h3>
            {Object.entries(grouped).map(([monthName, groupPolicies]) => (
              <div key={monthName} className="card" style={{marginBottom: 20}}>
                <div className="card-header" style={{background: 'var(--bg-secondary)', padding: '12px 20px'}}>
                  <span className="card-title">📁 {monthName}</span>
                  <span style={{fontSize:12, color:'var(--text-muted)'}}>{groupPolicies.length} comprobante(s)</span>
                </div>
                <div style={{padding: '20px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '20px'}}>
                  {groupPolicies.map(p => (
                    <div key={p.id} style={{border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: 12, background: 'var(--bg-card)'}}>
                      <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start'}}>
                        <div style={{fontWeight: 600, fontSize: 14, marginBottom: 4}}>{p.nombre}</div>
                        <button 
                          title="Eliminar comprobante"
                          onClick={() => { if (confirm('¿Eliminar este comprobante?')) onSave({ ...p, comprobante: null }); }}
                          style={{background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 6, padding: '4px 8px', cursor: 'pointer', fontSize: 13, color: '#ef4444', flexShrink: 0}}
                        >🗑️</button>
                      </div>
                      <div style={{fontSize: 12, color: 'var(--text-secondary)', marginBottom: 8}}>
                        <strong>Póliza:</strong> {p.poliza}
                      </div>
                      <div style={{fontSize: 11, color: 'var(--text-muted)', marginBottom: 4}}>
                        <strong>Fecha límite:</strong> {formatDate(p.fechaPagoAnterior || p.fechaPago)}
                      </div>
                      <div style={{fontSize: 11, color: 'var(--accent-green)', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6}}>
                        <strong>Fecha pagado:</strong> 
                        <input type="date" 
                          defaultValue={p.fechaUltimoPago || todayISO()} 
                          onBlur={e => {
                            const val = e.target.value;
                            if (val && val !== p.fechaUltimoPago) {
                              onSave({ ...p, fechaUltimoPago: val });
                              if (toast) toast('Fecha de pago actualizada ✅', 'success');
                            }
                          }}
                          style={{fontSize: 11, padding: '2px 6px', border: '1px solid var(--accent-green)', borderRadius: 4, background: 'var(--bg-input)', color: 'var(--accent-green)', fontWeight: 'bold'}}
                        />
                      </div>
                      <div style={{width: '100%', height: 200, borderRadius: 'var(--radius-sm)', overflow: 'hidden', border: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg-secondary)', cursor: 'pointer'}}
                           onClick={() => setSelectedImg(p.comprobante)}>
                        {p.comprobante.startsWith('data:application/pdf') ? (
                          <embed src={p.comprobante} width="100%" height="100%" type="application/pdf" style={{pointerEvents: 'none'}} />
                        ) : (
                          <img src={p.comprobante} alt={`Comprobante`} style={{width: '100%', height: '100%', objectFit: 'contain'}} />
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        );
      })()}

      {/* Modales */}
      {modalNew && (
        <PolicyModal 
          policy={null} 
          isVida={isVida}
          isDanos={isDanos}
          isHogar={isHogar}
          isGmm={isGmm}
          isAutos={isAutos}
          onSave={(p) => { onSave(p); setModalNew(false); }} 
          onClose={() => setModalNew(false)} 
          toast={toast} 
        />
      )}
      {modalEdit && (
        <PolicyModal 
          policy={modalEdit} 
          isVida={isVida}
          isDanos={isDanos}
          isHogar={isHogar}
          isGmm={isGmm}
          isAutos={isAutos}
          onSave={(p) => { onSave(p); setModalEdit(null); }} 
          onClose={() => setModalEdit(null)} 
          toast={toast} 
        />
      )}
      {modalPaid && (
        <MarkPaidModal policy={modalPaid} onConfirm={(p, nextDate, comp, isLast) => { onMarkPaid(p, nextDate, comp, isLast); setModalPaid(null); }} onClose={() => setModalPaid(null)} toast={toast} />
      )}
      {deleteConfirm && (
        <div className="modal-overlay">
          <div className="modal" style={{maxWidth: 400}}>
            <div className="modal-body" style={{textAlign: 'center', padding: '30px 20px'}}>
              <div style={{fontSize:40, marginBottom:16}}>⚠️</div>
              <h3 style={{marginBottom:10}}>¿Eliminar Póliza?</h3>
              <p style={{color:'var(--text-secondary)', marginBottom:24}}>
                Se borrará permanentemente la póliza de <strong>{deleteConfirm.nombre}</strong>.
              </p>
              <div className="flex gap-2" style={{justifyContent:'center'}}>
                <button className="btn btn-outline" onClick={() => setDeleteConfirm(null)}>Cancelar</button>
                <button className="btn btn-danger" onClick={() => {
                  onDelete(deleteConfirm.id);
                  setDeleteConfirm(null);
                }}>Sí, eliminar</button>
              </div>
            </div>
          </div>
        </div>
      )}
      {selectedImg && (
        <ImageModal src={selectedImg} onClose={() => setSelectedImg(null)} />
      )}
    </div>
  );
}

// ─── Página: Plantillas ───────────────────────────────────────
function TemplatesPage({ templates, onSave, toast }) {
  const [waText, setWaText] = useState(templates.whatsapp);
  const [emailAsunto, setEmailAsunto] = useState(templates.email_asunto);
  const [emailCuerpo, setEmailCuerpo] = useState(templates.email_cuerpo);
  const [activeTab, setActiveTab] = useState('whatsapp');
  const [saved, setSaved] = useState(false);

  const VARS = ['{nombre}', '{poliza}', '{bien}', '{monto}', '{formaPago}', '{agente}', '{fechaPago}', '{correo}', '{telefono}', '{estado_vencimiento}'];

  const insertVar = (v, setter) => {
    setter(t => t + v);
  };

  const handleSave = () => {
    onSave({ whatsapp: waText, email_asunto: emailAsunto, email_cuerpo: emailCuerpo });
    setSaved(true);
    toast('Plantillas guardadas ✅', 'success');
    setTimeout(() => setSaved(false), 2000);
  };

  const handleReset = () => {
    setWaText(DEFAULT_TEMPLATES.whatsapp);
    setEmailAsunto(DEFAULT_TEMPLATES.email_asunto);
    setEmailCuerpo(DEFAULT_TEMPLATES.email_cuerpo);
    toast('Plantillas restauradas al valor predeterminado', 'info');
  };

  const previewPolicy = {
    nombre: 'María Fernández', poliza: 'POL-2024-001',
    bien: 'Toyota Corolla 2022', monto: 1850,
    formaPago: 'MENSUAL', agente: 'DANIEL',
    fechaPago: todayISO(), correo: 'maria@gmail.com', telefono: '5512345678'
  };

  return (
    <div className="page-fade-enter">
      <div className="flex gap-4" style={{flexWrap:'wrap'}}>
        {/* Editor */}
        <div style={{flex:'1 1 400px'}}>
          <div className="card" style={{marginBottom:16}}>
            <div className="card-header">
              <span className="card-title">✏️ Editor de Plantillas</span>
              <div className="flex gap-2">
                <button className="btn btn-ghost btn-sm" onClick={handleReset}>↩ Restaurar</button>
                <button className="btn btn-primary btn-sm" onClick={handleSave}>
                  {saved ? '✅ Guardado' : '💾 Guardar'}
                </button>
              </div>
            </div>
            <div style={{padding:'16px 20px'}}>
              <div className="tabs" style={{marginBottom:16}}>
                <button className={`tab-btn ${activeTab === 'whatsapp' ? 'active' : ''}`}
                  onClick={() => setActiveTab('whatsapp')}>💬 WhatsApp</button>
                <button className={`tab-btn ${activeTab === 'email' ? 'active' : ''}`}
                  onClick={() => setActiveTab('email')}>📧 Correo</button>
              </div>

              <div style={{marginBottom:10}}>
                <p className="form-label" style={{marginBottom:6}}>Variables disponibles (click para insertar):</p>
                <div className="var-list">
                  {VARS.map(v => (
                    <span key={v} className="var-chip"
                      onClick={() => activeTab === 'whatsapp' ? insertVar(v, setWaText) : insertVar(v, setEmailCuerpo)}>
                      {v}
                    </span>
                  ))}
                </div>
              </div>

              {activeTab === 'whatsapp' ? (
                <div className="form-group">
                  <label className="form-label">Mensaje WhatsApp</label>
                  <textarea className="input" rows={12} value={waText}
                    onChange={e => setWaText(e.target.value)}
                    style={{fontFamily:'monospace', fontSize:12}} />
                </div>
              ) : (
                <>
                  <div className="form-group" style={{marginBottom:12}}>
                    <label className="form-label">Asunto del correo</label>
                    <input className="input" value={emailAsunto}
                      onChange={e => setEmailAsunto(e.target.value)} />
                  </div>
                  <div className="form-group">
                    <label className="form-label">Cuerpo del correo</label>
                    <textarea className="input" rows={14} value={emailCuerpo}
                      onChange={e => setEmailCuerpo(e.target.value)}
                      style={{fontFamily:'monospace', fontSize:12}} />
                  </div>
                </>
              )}
            </div>
          </div>
        </div>

        {/* Preview */}
        <div style={{flex:'1 1 300px'}}>
          <div className="card">
            <div className="card-header">
              <span className="card-title">👁️ Vista Previa</span>
              <span style={{fontSize:11, color:'var(--text-muted)'}}>Con datos de ejemplo</span>
            </div>
            <div style={{padding:'16px 20px'}}>
              {activeTab === 'whatsapp' ? (
                <div className="template-preview" style={{
                  background:'#0b2027', border:'1px solid #25d36640',
                  borderRadius:'var(--radius-md)', color:'#e8f5e9', lineHeight:1.8
                }}>
                  {fillTemplate(waText, previewPolicy, true)}
                </div>
              ) : (
                <>
                  <div style={{marginBottom:12}}>
                    <p className="form-label" style={{marginBottom:6}}>Asunto:</p>
                    <div className="template-preview" style={{padding:'8px 12px', fontSize:13, fontWeight:600}}>
                      {fillTemplate(emailAsunto, previewPolicy, false)}
                    </div>
                  </div>
                  <div className="template-preview" style={{lineHeight:1.8}}>
                    {fillTemplate(emailCuerpo, previewPolicy, false)}
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Página: Importar / Exportar ──────────────────────────────
function ImportExportPage({ 
  policies, caroPolicies, gmmPolicies, autosPolicies, vidaPolicies, danosPolicies, hogarPolicies,
  onImport, toast 
}) {
  const fileRef = useRef();
  const [importing, setImporting] = useState(false);
  const [preview, setPreview] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [fileInfo, setFileInfo] = useState(null);
  const [targetCategory, setTargetCategory] = useState('policies');

  const getTargetPolicies = () => {
    switch (targetCategory) {
      case 'caroPolicies': return caroPolicies;
      case 'gmmPolicies': return gmmPolicies;
      case 'autosPolicies': return autosPolicies;
      case 'vidaPolicies': return vidaPolicies;
      case 'danosPolicies': return danosPolicies;
      case 'hogarPolicies': return hogarPolicies;
      default: return policies;
    }
  };

  const handleExport = () => {
    if (!window.XLSX) { toast('Librería XLSX no cargada', 'error'); return; }
    const XLSX = window.XLSX;
    const target = getTargetPolicies() || [];
    if (target.length === 0) { toast('No hay pólizas para exportar en esta categoría', 'error'); return; }
    
    const rows = target.map(p => ({
      'Nombre': p.nombre,
      'Póliza': p.poliza,
      'Vehículo': p.bien,
      'Forma de pago': p.formaPago,
      'Clave': p.agente,
      'Fecha de pago': p.fechaPago,
      'Monto': p.monto,
      'Estatus': p.estatus,
      'Correo': p.correo,
      'Teléfono': p.telefono,
      'Notas': p.notas || '',
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    ws['!cols'] = [{wch:30},{wch:15},{wch:35},{wch:12},{wch:10},{wch:14},{wch:12},{wch:12},{wch:30},{wch:14},{wch:40}];
    XLSX.utils.book_append_sheet(wb, ws, 'Pólizas');
    const fecha = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `SeguroControl_${fecha}.xlsx`);
    toast('Archivo Excel exportado ✅', 'success');
  };



  // ── Parsear fecha (Date obj, serial Excel, string variado) ────
  const parseDate = (val) => {
    if (!val && val !== 0) return todayISO();
    // Ya es un Date de JS (cellDates: true)
    if (val instanceof Date) {
      return val.toISOString().split('T')[0];
    }
    // Número serial de Excel (días desde 1899-12-30)
    if (typeof val === 'number') {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];
      return todayISO();
    }
    const s = String(val).trim();
    // ISO: 2026-07-20
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    // DD/MM/YYYY o DD-MM-YYYY
    let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    // MM/DD/YYYY (fallback)
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})$/);
    if (m) return `20${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
    // Intentar Date.parse como último recurso
    const dp = new Date(s);
    if (!isNaN(dp.getTime())) return dp.toISOString().split('T')[0];
    return todayISO();
  };



  // ── Normalizar forma de pago ──────────────────────────────────
  const parseFormaPago = (val) => {
    if (!val) return 'MENSUAL';
    const v = normalize(val);
    if (v.includes('cont')) return 'CONTADO';
    if (v.includes('trim')) return 'TRIMESTRAL';
    if (v.includes('sem')) return 'SEMESTRAL';
    if (v.includes('men') || v.includes('month')) return 'MENSUAL';
    // Si es exactamente una de las opciones válidas
    const up = String(val).toUpperCase().trim();
    if (['CONTADO','MENSUAL','TRIMESTRAL','SEMESTRAL'].includes(up)) return up;
    return 'MENSUAL';
  };

  // ── Normalizar estatus ────────────────────────────────────────
  const parseEstatus = (val) => {
    if (!val) return 'PENDIENTE';
    const v = normalize(val);
    if (v.includes('pagad') || v === 'pago' || v === 'pagado') return 'PAGADO';
    if (v.includes('venc')) return 'VENCIDO';
    if (v.includes('canc')) return 'CANCELADO';
    if (v.includes('liquid')) return 'LIQUIDADO';
    if (v.includes('pend')) return 'PENDIENTE';
    const up = String(val).toUpperCase().trim();
    if (['PAGADO','VENCIDO','PENDIENTE','CANCELADO','LIQUIDADO'].includes(up)) return up;
    return 'PENDIENTE';
  };

  // ── Normalizar agente ─────────────────────────────────────────
  const parseAgente = (val) => {
    if (!val) return 'DANIEL';
    const v = normalize(val);
    if (v.includes('mart') || v.includes('mtn')) return 'MARTIN';
    if (v.includes('dani') || v.includes('dan')) return 'DANIEL';
    return String(val).toUpperCase().trim();
  };

  // ── Procesar archivo ──────────────────────────────────────────
  const processFile = (file) => {
    if (!file) return;
    setFileInfo({ name: file.name, size: (file.size / 1024).toFixed(1) + ' KB' });
    setImporting(true);
    setPreview(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '' });

        if (rawRows.length === 0) {
          toast('El archivo está vacío', 'error');
          setImporting(false);
          return;
        }

        // Log de columnas para debug
        console.log('📊 Columnas detectadas:', Object.keys(rawRows[0]));
        console.log('📊 Primer registro raw:', rawRows[0]);

        const mapped = rawRows.map(r => {
          // Buscar cada campo con múltiples variantes posibles
          const nombre   = findCol(r, ['nombre', 'asegurado', 'cliente', 'titular', 'contratante', 'name']);
          const poliza   = findCol(r, ['poliza', 'policy', 'numero', 'no poliza', 'num']);
          const bien     = findCol(r, ['vehiculo', 'bien', 'auto', 'carro', 'objeto', 'descripcion', 'inmueble', 'unidad']);
          const rawForma = findCol(r, ['forma', 'forma de pago', 'periodicidad', 'periodo', 'frecuencia', 'tipo pago']);
          const rawAgent = findCol(r, ['clave', 'agente', 'asesor', 'vendedor', 'ejecutivo', 'clave agente', 'agent']);
          const rawFecha = findCol(r, ['fecha', 'fecha de pago', 'vencimiento', 'vigencia', 'limite', 'proximo', 'pago']);
          const rawMonto = findCol(r, ['monto', 'prima', 'importe', 'total', 'precio', 'costo', 'valor', 'amount']);
          const rawEstat = findCol(r, ['estatus', 'status', 'estado', 'situacion']);
          const correo   = findCol(r, ['correo', 'email', 'mail', 'e-mail', 'electronico']);
          const telefono = findCol(r, ['telefono', 'celular', 'movil', 'whatsapp', 'tel', 'contacto', 'phone', 'cel']);
          const notas    = findCol(r, ['nota', 'notas', 'observacion', 'comentario', 'remarks', 'obs']);

          return {
            id: generateId(),
            nombre:   String(nombre || '').trim(),
            poliza:   String(poliza || '').trim(),
            bien:     String(bien || '').trim(),
            formaPago: parseFormaPago(rawForma),
            agente:   parseAgente(rawAgent),
            fechaPago: parseDate(rawFecha),
            monto:    parseMonto(rawMonto),
            estatus:  parseEstatus(rawEstat),
            correo:   String(correo || '').trim(),
            telefono: String(telefono || '').replace(/\D/g, '').slice(-10),
            notas:    String(notas || '').trim(),
          };
        }).filter(r => r.nombre && r.nombre.length > 1);

        if (mapped.length === 0) {
          toast('No se encontraron registros. Verifica que tu Excel tenga una columna con "Nombre" o "Asegurado".', 'error');
          console.log('⚠️ Claves disponibles:', Object.keys(rawRows[0]));
        } else {
          setPreview(mapped);
          toast(`✅ ${mapped.length} pólizas detectadas desde "${file.name}"`, 'success');
          console.log('✅ Primer registro mapeado:', mapped[0]);
        }
      } catch (err) {
        toast('Error al leer el archivo: ' + err.message, 'error');
        console.error('Error importación:', err);
      }
      setImporting(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const confirmImport = (mode) => {
    if (!preview) return;
    onImport(preview, mode, targetCategory);
    setPreview(null);
    toast(`${preview.length} pólizas importadas en ${targetCategory} ✅`, 'success');
  };

  return (
    <div className="page-fade-enter">
      <div className="flex gap-4" style={{flexWrap:'wrap'}}>
        <div style={{flex:'1 1 300px'}}>
          <div className="card">
            <div className="card-header"><span className="card-title">📤 Exportar a Excel</span></div>
            <div style={{padding:24}}>
              <p style={{fontSize:13, color:'var(--text-secondary)', marginBottom:20}}>Descarga tus pólizas en formato .xlsx para respaldo o edición masiva.</p>
              
              <div style={{marginBottom:16}}>
                <label style={{display:'block', fontSize:13, fontWeight:600, color:'var(--text-secondary)', marginBottom:6}}>Categoría a Exportar:</label>
                <select className="form-input" value={targetCategory} onChange={(e) => setTargetCategory(e.target.value)}>
                  <option value="policies">Autos Qualitas (Daniel/Martín)</option>
                  <option value="caroPolicies">Autos Qualitas (Caro)</option>
                  <option value="gmmPolicies">Gastos Médicos Mayores (GMM)</option>
                  <option value="autosPolicies">Autos (Otras Aseguradoras)</option>
                  <option value="vidaPolicies">Vida</option>
                  <option value="danosPolicies">Daños</option>
                  <option value="hogarPolicies">Hogar</option>
                </select>
              </div>

              <button className="btn btn-success w-full" onClick={handleExport}>Descargar Excel</button>
            </div>
          </div>
        </div>

        <div style={{flex:'2 1 380px'}}>
          <div className="card">
            <div className="card-header">
              <span className="card-title">📥 Importar Pólizas</span>
              <span style={{fontSize:11, color:'var(--text-muted)'}}>Excel .xlsx / .xls / .csv</span>
            </div>
            <div style={{padding:24}}>
              <div style={{marginBottom:20}}>
                <label style={{display:'block', fontSize:13, fontWeight:600, color:'var(--text-secondary)', marginBottom:6}}>Categoría de Destino:</label>
                <select className="form-input" value={targetCategory} onChange={(e) => setTargetCategory(e.target.value)}>
                  <option value="policies">Autos Qualitas (Daniel/Martín)</option>
                  <option value="caroPolicies">Autos Qualitas (Caro)</option>
                  <option value="gmmPolicies">Gastos Médicos Mayores (GMM)</option>
                  <option value="autosPolicies">Autos (Otras Aseguradoras)</option>
                  <option value="vidaPolicies">Vida</option>
                  <option value="danosPolicies">Daños</option>
                  <option value="hogarPolicies">Hogar</option>
                </select>
              </div>
              {/* Zona Drag & Drop */}
              <div
                style={{
                  border: `2px dashed ${dragOver ? 'var(--accent-blue)' : 'var(--border)'}`,
                  borderRadius: 'var(--radius-lg)',
                  padding: '40px 24px',
                  textAlign: 'center',
                  cursor: importing ? 'wait' : 'pointer',
                  transition: 'all 0.2s ease',
                  background: dragOver ? 'rgba(59,130,246,0.08)' : 'rgba(255,255,255,0.02)',
                }}
                onClick={() => !importing && fileRef.current?.click()}
                onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDragOver(true); }}
                onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); setDragOver(true); }}
                onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setDragOver(false); }}
                onDrop={(e) => {
                  e.preventDefault(); e.stopPropagation();
                  setDragOver(false);
                  if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
                }}
              >
                {importing ? (
                  <div style={{display:'flex', flexDirection:'column', alignItems:'center', gap:14}}>
                    <div className="loading-spinner" style={{width:40, height:40, borderWidth:3}} />
                    <p style={{fontSize:14, fontWeight:600}}>Leyendo archivo…</p>
                    <p style={{fontSize:12, color:'var(--text-muted)'}}>Detectando columnas y datos</p>
                  </div>
                ) : dragOver ? (
                  <div style={{display:'flex', flexDirection:'column', alignItems:'center', gap:10}}>
                    <span style={{fontSize:56}}>📂</span>
                    <p style={{fontSize:16, fontWeight:700, color:'var(--accent-blue-light)'}}>¡Suelta el archivo aquí!</p>
                  </div>
                ) : (
                  <div style={{display:'flex', flexDirection:'column', alignItems:'center', gap:10}}>
                    <span style={{fontSize:50, opacity:0.45}}>📊</span>
                    <p style={{fontSize:15, fontWeight:700}}>Arrastra tu archivo Excel aquí</p>
                    <p style={{fontSize:13, color:'var(--text-muted)'}}>o haz clic para buscarlo</p>
                    <div style={{display:'flex', gap:8, marginTop:8, flexWrap:'wrap', justifyContent:'center'}}>
                      {['.xlsx','.xls','.csv'].map(f => (
                        <span key={f} style={{
                          padding:'3px 10px', background:'rgba(255,255,255,0.06)',
                          border:'1px solid var(--border)', borderRadius:6,
                          fontSize:12, color:'var(--text-secondary)'
                        }}>{f}</span>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.ods"
                style={{display:'none'}}
                onChange={(e) => { processFile(e.target.files[0]); e.target.value = ''; }} />

              <button className="btn btn-primary w-full" style={{marginTop:14}}
                onClick={() => fileRef.current?.click()} disabled={importing}>
                📂 Seleccionar Archivo desde mi PC
              </button>

              {fileInfo && !importing && (
                <div style={{
                  marginTop:12, padding:'10px 14px',
                  background:'rgba(59,130,246,0.08)', border:'1px solid rgba(59,130,246,0.2)',
                  borderRadius:'var(--radius-md)', fontSize:12, color:'var(--text-secondary)'
                }}>
                  📄 <strong style={{color:'var(--text-primary)'}}>{fileInfo.name}</strong> — {fileInfo.size}
                </div>
              )}

              <div style={{marginTop:14, padding:'10px 14px', background:'rgba(255,255,255,0.03)', borderRadius:'var(--radius-md)'}}>
                <p style={{fontSize:12, color:'var(--text-muted)', lineHeight:1.8}}>
                  💡 <strong style={{color:'var(--text-secondary)'}}>Detección automática</strong> — el sistema reconoce
                  cualquier nombre: "Nombre", "Asegurado", "Cliente", "Póliza", "No. Póliza",
                  "Prima", "Monto", "Fecha", "Vencimiento", "Celular", "Tel"…
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>

      {preview && (
        <div className="card" style={{marginTop:20}}>
          <div className="card-header">
            <span className="card-title">👁️ Vista Previa — {preview.length} registros encontrados</span>
            <div className="flex gap-2">
              <button className="btn btn-outline btn-sm" onClick={() => setPreview(null)}>✕ Cancelar</button>
              <button className="btn btn-warning btn-sm" onClick={() => confirmImport('reemplazar')}>
                🔄 Reemplazar Todo
              </button>
              <button className="btn btn-primary btn-sm" onClick={() => confirmImport('agregar')}>
                ➕ Agregar a Existentes
              </button>
            </div>
          </div>
          <div style={{padding:'0 0 16px'}}>
            <PoliciesTable
              policies={preview.slice(0, 10)}
              compact={false}
              onEdit={() => {}} onDelete={() => {}} onMarkPaid={() => {}}
              onWhatsApp={() => {}} onEmail={() => {}}
            />
            {preview.length > 10 && (
              <p style={{fontSize:12, color:'var(--text-muted)', padding:'8px 24px'}}>
                … y {preview.length - 10} registro(s) más
              </p>
            )}
          </div>
        </div>
      )}

      {/* Guía de columnas */}
      <div className="card" style={{marginTop:20}}>
        <div className="card-header"><span className="card-title">📖 Guía de Columnas para Importación</span></div>
        <div style={{padding:20}}>
          <div className="table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Columna en Excel</th>
                  <th>Valores aceptados</th>
                  <th>Obligatorio</th>
                </tr>
              </thead>
              <tbody>
                {[
                  ['Nombre del Asegurado','Texto libre','✅ Sí'],
                  ['Póliza','Texto libre (ej: POL-2024-001)','✅ Sí'],
                  ['Vehículo / Bien Asegurado','Texto libre','No'],
                  ['Forma de Pago','CONTADO, MENSUAL, TRIMESTRAL, SEMESTRAL','No'],
                  ['Clave de Agente','DANIEL, MARTIN','No'],
                  ['Fecha de Pago','YYYY-MM-DD (ej: 2026-07-31)','No'],
                  ['Monto ($)','Número (ej: 1850)','No'],
                  ['Estatus','PENDIENTE, PAGADO, VENCIDO, CANCELADO','No'],
                  ['Correo Electrónico','email@dominio.com','No'],
                  ['Teléfono / WhatsApp','10 dígitos sin lada','No'],
                ].map(([col, vals, req]) => (
                  <tr key={col}>
                    <td><code style={{fontSize:12, color:'var(--accent-blue-light)'}}>{col}</code></td>
                    <td style={{fontSize:12, color:'var(--text-secondary)'}}>{vals}</td>
                    <td style={{fontSize:12}}>{req}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Modal para Ver Imagen ──────────────────────────────────────
function ImageModal({ src, onClose }) {
  useEffect(() => {
    const handleEsc = (e) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleEsc);
    return () => window.removeEventListener('keydown', handleEsc);
  }, [onClose]);

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()} style={{zIndex: 2000, padding: 40}}>
      <div style={{position: 'relative', width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center'}}>
        <button className="modal-close" onClick={onClose} style={{position: 'absolute', top: -10, right: -10, background: 'var(--bg-card)', borderRadius: '50%', padding: 4, zIndex: 2010}}>
          <Icons.Close />
        </button>
        {src.startsWith('data:application/pdf') ? (
          <embed src={src} width="100%" height="100%" type="application/pdf" style={{borderRadius: 8}} />
        ) : (
          <img src={src} style={{maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', borderRadius: 8}} />
        )}
      </div>
    </div>
  );
}

// ─── Página: Comprobantes ───────────────────────────────────────
function ComprobantesPage({ policies, onUpdatePolicy }) {
  const [selectedImg, setSelectedImg] = useState(null);
  const withComprobantes = policies.filter(p => p.comprobante);

  if (withComprobantes.length === 0) {
    return (
      <div className="empty-state" style={{paddingTop:100}}>
        <div className="empty-state-icon" style={{fontSize: 48, marginBottom: 16}}>🧾</div>
        <h3>Sin comprobantes</h3>
        <p>Aún no se han adjuntado comprobantes de pago a ninguna póliza.</p>
      </div>
    );
  }

  // Group by month
  const grouped = {};
  withComprobantes.forEach(p => {
    // Si no tiene fechaUltimoPago (porque es viejo), se asume el mes actual
    const dStr = p.fechaUltimoPago || new Date().toISOString().split('T')[0];
    const date = new Date(dStr + 'T12:00:00');
    const monthYear = date.toLocaleDateString('es-MX', { month: 'long', year: 'numeric' });
    const capitalized = monthYear.charAt(0).toUpperCase() + monthYear.slice(1);
    
    if (!grouped[capitalized]) grouped[capitalized] = [];
    grouped[capitalized].push(p);
  });

  return (
    <div className="page-fade-enter">
      {Object.entries(grouped).map(([monthName, groupPolicies]) => (
        <div key={monthName} className="card" style={{marginBottom: 24}}>
          <div className="card-header">
            <span className="card-title">📁 {monthName}</span>
            <span style={{fontSize:12, color:'var(--text-muted)'}}>{groupPolicies.length} comprobante(s)</span>
          </div>
          <div style={{padding: '20px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '20px'}}>
            {groupPolicies.map(p => (
              <div key={p.id} style={{border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: 12, background: 'var(--bg-card)'}}>
                <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start'}}>
                  <div style={{fontWeight: 600, fontSize: 14, marginBottom: 4}}>{p.nombre}</div>
                  <button 
                    title="Eliminar comprobante"
                    onClick={() => { if (confirm('¿Eliminar este comprobante?')) onUpdatePolicy({ ...p, comprobante: null }); }}
                    style={{background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: 6, padding: '4px 8px', cursor: 'pointer', fontSize: 13, color: '#ef4444', flexShrink: 0}}
                  >🗑️</button>
                </div>
                <div style={{fontSize: 12, color: 'var(--text-secondary)', marginBottom: 8}}>
                  <strong>Póliza:</strong> {p.poliza}
                </div>
                <div style={{fontSize: 11, color: 'var(--text-muted)', marginBottom: 4}}>
                  <strong>Fecha límite:</strong> {formatDate(p.fechaPagoAnterior || p.fechaPago)}
                </div>
                <div style={{fontSize: 11, color: 'var(--accent-green)', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6}}>
                  <strong>Fecha pagado:</strong> 
                  <input type="date" 
                    defaultValue={p.fechaUltimoPago || todayISO()} 
                    onBlur={e => {
                      const val = e.target.value;
                      if (val && val !== p.fechaUltimoPago) {
                        onUpdatePolicy({ ...p, fechaUltimoPago: val });
                      }
                    }}
                    style={{fontSize: 11, padding: '2px 6px', border: '1px solid var(--accent-green)', borderRadius: 4, background: 'var(--bg-input)', color: 'var(--accent-green)', fontWeight: 'bold'}}
                  />
                </div>
                
                <div style={{width: '100%', height: 200, borderRadius: 'var(--radius-sm)', overflow: 'hidden', border: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--bg-secondary)', cursor: 'pointer'}}
                     onClick={() => setSelectedImg(p.comprobante)}>
                  {p.comprobante.startsWith('data:application/pdf') ? (
                    <embed src={p.comprobante} width="100%" height="100%" type="application/pdf" style={{pointerEvents: 'none'}} />
                  ) : (
                    <img src={p.comprobante} alt={`Comprobante`} style={{width: '100%', height: '100%', objectFit: 'contain'}} />
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}

      {selectedImg && (
        <ImageModal src={selectedImg} onClose={() => setSelectedImg(null)} />
      )}
    </div>
  );
}

// ─── Página: Cotizaciones ───────────────────────────────────────
function CotizacionesPage({ cotizaciones, onSave, onUpdateEstatus }) {
  const [showForm, setShowForm] = useState(false);
  const [filterEstatus, setFilterEstatus] = useState('TODOS');
  const [filterAgente, setFilterAgente] = useState('TODOS');
  const [search, setSearch] = useState('');

  const [form, setForm] = useState({
    fecha: todayISO(),
    unidad: '',
    agente: 'MARTÍN',
    cp: '',
    estatus: 'PENDIENTE',
    obs: ''
  });

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!form.unidad || !form.cp) { alert('Unidad y CP son obligatorios'); return; }
    onSave({ ...form, id: generateId() });
    setForm({ ...form, unidad: '', cp: '', obs: '', fecha: todayISO() });
    setShowForm(false);
  };

  const filtered = cotizaciones.filter(c => {
    if (filterEstatus !== 'TODOS' && c.estatus !== filterEstatus) return false;
    if (filterAgente !== 'TODOS' && c.agente !== filterAgente) return false;
    if (search && !c.unidad.toLowerCase().includes(search.toLowerCase()) && !c.obs.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  return (
    <div className="page-fade-enter">
      <div className="flex" style={{ justifyContent: 'space-between', marginBottom: 20, gap: 10, flexWrap: 'wrap' }}>
        <div className="flex gap-2">
          <div className="search-wrapper">
            <Icons.Search />
            <input className="input input-search" placeholder="Buscar unidad u obs..." value={search} onChange={e => setSearch(e.target.value)} />
          </div>
          <select className="select" value={filterEstatus} onChange={e => setFilterEstatus(e.target.value)}>
            <option value="TODOS">Todos los estatus</option>
            <option value="PENDIENTE">PENDIENTE</option>
            <option value="EMITIDA">EMITIDA</option>
            <option value="NO CONCRETADA">NO CONCRETADA</option>
          </select>
          <select className="select" value={filterAgente} onChange={e => setFilterAgente(e.target.value)}>
            <option value="TODOS">Todos los agentes</option>
            <option value="MARTÍN">MARTÍN</option>
            <option value="DANIEL">DANIEL</option>
            <option value="CARO">CARO</option>
          </select>
        </div>
        <button className="btn btn-primary" onClick={() => setShowForm(true)}>
          <Icons.Plus /> Nueva Cotización
        </button>
      </div>

      {showForm && (
        <div className="modal-overlay" onClick={() => setShowForm(false)}>
          <div className="modal" style={{maxWidth: 500}} onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <span className="modal-title">Registrar Cotización</span>
              <button className="btn btn-ghost" onClick={() => setShowForm(false)}><Icons.Close /></button>
            </div>
            <div className="modal-body">
              <form onSubmit={handleSubmit} className="form-grid">
                <div className="form-group">
                  <label className="form-label">Fecha <span className="required">*</span></label>
                  <input type="date" className="input" value={form.fecha} onChange={e => setForm({...form, fecha: e.target.value})} required />
                </div>
                <div className="form-group">
                  <label className="form-label">Agente <span className="required">*</span></label>
                  <select className="select" value={form.agente} onChange={e => setForm({...form, agente: e.target.value})} required>
                    <option value="MARTÍN">MARTÍN</option>
                    <option value="DANIEL">DANIEL</option>
                    <option value="CARO">CARO</option>
                  </select>
                </div>
                <div className="form-group full-width">
                  <label className="form-label">Datos de la Unidad <span className="required">*</span></label>
                  <input type="text" className="input" placeholder="Ej. VW Jetta 2023" value={form.unidad} onChange={e => setForm({...form, unidad: e.target.value})} required />
                </div>
                <div className="form-group">
                  <label className="form-label">Código Postal <span className="required">*</span></label>
                  <input type="text" className="input" placeholder="Ej. 11000" value={form.cp} onChange={e => setForm({...form, cp: e.target.value})} required />
                </div>
                <div className="form-group">
                  <label className="form-label">Estatus inicial</label>
                  <select className="select" value={form.estatus} onChange={e => setForm({...form, estatus: e.target.value})}>
                    <option value="PENDIENTE">PENDIENTE</option>
                    <option value="EMITIDA">EMITIDA</option>
                    <option value="NO CONCRETADA">NO CONCRETADA</option>
                  </select>
                </div>
                <div className="form-group full-width">
                  <label className="form-label">Observaciones</label>
                  <textarea className="input" rows="3" placeholder="Comentarios adicionales..." value={form.obs} onChange={e => setForm({...form, obs: e.target.value})}></textarea>
                </div>
                <div className="form-group full-width" style={{marginTop: 10}}>
                  <button type="submit" className="btn btn-primary" style={{width: '100%', justifyContent: 'center'}}>Guardar Cotización</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      <div className="card">
        <div className="card-header">
          <span className="card-title">Listado de Cotizaciones ({filtered.length})</span>
        </div>
        <div className="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Agente</th>
                <th>Unidad</th>
                <th>C.P.</th>
                <th>Observaciones</th>
                <th>Estatus</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan="6" style={{textAlign: 'center', padding: 30, color: 'var(--text-muted)'}}>No se encontraron cotizaciones</td></tr>
              ) : (
                filtered.map(c => (
                  <tr key={c.id}>
                    <td style={{fontSize: 12, color: 'var(--text-muted)'}}>{formatDate(c.fecha)}</td>
                    <td><AgentBadge agente={c.agente} /></td>
                    <td style={{fontWeight: 600, fontSize: 13}}>{c.unidad}</td>
                    <td style={{fontSize: 13}}>{c.cp}</td>
                    <td style={{fontSize: 12, maxWidth: 200, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis'}} title={c.obs}>{c.obs || '—'}</td>
                    <td>
                      <select className="select" style={{fontSize: 11, padding: '4px 8px'}} value={c.estatus} onChange={e => onUpdateEstatus(c.id, e.target.value)}>
                        <option value="PENDIENTE">⏳ PENDIENTE</option>
                        <option value="EMITIDA">✅ EMITIDA</option>
                        <option value="NO CONCRETADA">❌ NO CONCRETADA</option>
                      </select>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function SiniestrosPage({ siniestros, onImport, onUpdateEstatus }) {
  const [dragOverM, setDragOverM] = useState(false);
  const [dragOverD, setDragOverD] = useState(false);
  const [importing, setImporting] = useState(null); // 'MARTIN' or 'DANIEL'
  const [msgModal, setMsgModal] = useState(null);
  const fileRefM = useRef();
  const fileRefD = useRef();

  const handleFileDrop = (e, agente) => {
    e.preventDefault();
    if (agente === 'MARTIN') setDragOverM(false);
    else setDragOverD(false);
    
    const file = e.dataTransfer ? e.dataTransfer.files[0] : e.target.files[0];
    if (!file) return;

    setImporting(agente);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const XLSX = window.XLSX;
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        
        // Buscar pestaña SINIESTROS o similar
        const sheetName = wb.SheetNames.find(s => normalize(s).includes('siniestro')) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '' });

        if (rawRows.length === 0) {
          alert('El archivo está vacío o no se encontraron datos.');
          setImporting(null);
          return;
        }

        const mapped = rawRows.map(r => {
          const poliza   = findCol(r, ['poliza', 'policy', 'numero', 'no poliza']);
          const asegurado= findCol(r, ['nomaseg', 'asegurado', 'cliente', 'nombre']);
          const vehiculo = findCol(r, ['vehiculo', 'auto', 'unidad', 'bien', 'descripcion']);
          const tipo     = findCol(r, ['tiposin', 'tipo', 'siniestro']);
          const cvestro  = findCol(r, ['cvestro', 'no siniestro', 'numero de siniestro', 'siniestro', 'reporte', 'folio']);
          const causa    = findCol(r, ['causa', 'motivo', 'descripcion']);
          const costo    = parseMonto(findCol(r, ['sintotal', 'costo', 'total', 'monto']));
          
          const rvadm = parseMonto(findCol(r, ['rvadm', 'reserva adm']));
          const rvart = parseMonto(findCol(r, ['rvart', 'reserva art']));
          const rvarc = parseMonto(findCol(r, ['rvarc', 'reserva arc']));
          const rvagm = parseMonto(findCol(r, ['rvagm', 'reserva agm']));
          const rvaot = parseMonto(findCol(r, ['rvaot', 'reserva aot']));
          const reservas = rvadm + rvart + rvarc + rvagm + rvaot;

          return { 
            poliza: String(poliza||'').trim(), 
            asegurado, vehiculo, tipo, causa, costo, reservas, agente,
            cvestro: String(cvestro||'').trim() 
          };
        }).filter(r => r.poliza && r.poliza.length > 2); // Solo filas con póliza válida

        // Identificar el de mayor costo por póliza
        const grouped = {};
        mapped.forEach(s => {
          if (!grouped[s.poliza]) {
            grouped[s.poliza] = s;
          } else {
            if (s.costo > grouped[s.poliza].costo) {
              grouped[s.poliza] = s;
            }
          }
        });

        const finalSiniestros = Object.values(grouped);
        if (finalSiniestros.length > 0) {
          onImport(finalSiniestros);
        } else {
          alert('No se detectaron pólizas/siniestros en el archivo.');
        }
      } catch (err) {
        console.error(err);
        alert('Error al leer el archivo Excel.');
      }
      setImporting(null);
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="page-fade-enter">
      <div className="flex gap-4" style={{marginBottom: 20}}>
        {/* Importador MARTIN */}
        <div className="card" style={{flex: 1}}>
          <div className="card-header" style={{background: 'rgba(20,184,166,0.05)', borderBottom: '1px solid rgba(20,184,166,0.2)'}}>
            <span className="card-title" style={{color: '#0f766e'}}>👥 Importar Reporte - MARTÍN</span>
            <span style={{fontSize:11, color:'var(--text-muted)'}}>Formato Vigente (POLIZAS / SINIESTROS)</span>
          </div>
          <div style={{padding: 24}}>
            <div
              style={{
                border: `2px dashed ${dragOverM ? '#0f766e' : 'var(--border)'}`,
                borderRadius: 'var(--radius-lg)', padding: '30px 16px', textAlign: 'center', cursor: 'pointer',
                background: dragOverM ? 'rgba(20,184,166,0.05)' : 'rgba(255,255,255,0.02)'
              }}
              onClick={() => importing !== 'MARTIN' && fileRefM.current?.click()}
              onDragOver={e => { e.preventDefault(); setDragOverM(true); }}
              onDragLeave={() => setDragOverM(false)}
              onDrop={e => handleFileDrop(e, 'MARTIN')}
            >
              {importing === 'MARTIN' ? <p>Procesando...</p> : (
                <div>
                  <span style={{fontSize: 32}}>📂</span>
                  <p style={{fontWeight: 600, marginTop: 8}}>Sube el Excel de Martín aquí</p>
                </div>
              )}
            </div>
            <input ref={fileRefM} type="file" style={{display: 'none'}} accept=".xlsx,.xls,.csv" onChange={e => { handleFileDrop(e, 'MARTIN'); e.target.value=''; }} />
          </div>
        </div>

        {/* Importador DANIEL */}
        <div className="card" style={{flex: 1}}>
          <div className="card-header" style={{background: 'rgba(99,102,241,0.05)', borderBottom: '1px solid rgba(99,102,241,0.2)'}}>
            <span className="card-title" style={{color: '#3730a3'}}>👤 Importar Reporte - DANIEL</span>
            <span style={{fontSize:11, color:'var(--text-muted)'}}>Prima Devengada y Siniestralidad</span>
          </div>
          <div style={{padding: 24}}>
            <div
              style={{
                border: `2px dashed ${dragOverD ? '#3730a3' : 'var(--border)'}`,
                borderRadius: 'var(--radius-lg)', padding: '30px 16px', textAlign: 'center', cursor: 'pointer',
                background: dragOverD ? 'rgba(99,102,241,0.05)' : 'rgba(255,255,255,0.02)'
              }}
              onClick={() => importing !== 'DANIEL' && fileRefD.current?.click()}
              onDragOver={e => { e.preventDefault(); setDragOverD(true); }}
              onDragLeave={() => setDragOverD(false)}
              onDrop={e => handleFileDrop(e, 'DANIEL')}
            >
              {importing === 'DANIEL' ? <p>Procesando...</p> : (
                <div>
                  <span style={{fontSize: 32}}>📈</span>
                  <p style={{fontWeight: 600, marginTop: 8}}>Sube el Excel de Daniel aquí</p>
                </div>
              )}
            </div>
            <input ref={fileRefD} type="file" style={{display: 'none'}} accept=".xlsx,.xls,.csv" onChange={e => { handleFileDrop(e, 'DANIEL'); e.target.value=''; }} />
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-header">
          <span className="card-title">🛡️ Seguimiento de Siniestros y Reservas ({siniestros.length})</span>
        </div>
        <div className="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>Agente</th>
                <th>Asegurado / Póliza</th>
                <th>Siniestro / Causa</th>
                <th>Costo Total</th>
                <th>Reservas</th>
                <th>Seguimiento</th>
              </tr>
            </thead>
            <tbody>
              {siniestros.length === 0 ? (
                <tr><td colSpan="6" style={{textAlign: 'center', padding: 40, color: 'var(--text-muted)'}}>No hay siniestros importados</td></tr>
              ) : (
                siniestros.map(s => (
                  <tr key={s.id}>
                    <td><AgentBadge agente={s.agente} /></td>
                    <td>
                      <div style={{fontWeight: 600, fontSize: 13}}>{s.asegurado || '—'}</div>
                      <div style={{fontSize: 11, color: 'var(--text-muted)'}}>{s.poliza} • {s.vehiculo || '—'}</div>
                    </td>
                    <td>
                      <div style={{fontSize: 13}}>{s.tipo || 'No especificado'}</div>
                      <div style={{fontSize: 11, color: 'var(--text-muted)'}}>{s.causa || '—'}</div>
                    </td>
                    <td style={{fontWeight: 600, color: 'var(--accent-red)'}}>{formatMoney(s.costo)}</td>
                    <td style={{fontWeight: 600, color: 'var(--accent-yellow)'}}>{formatMoney(s.reservas)}</td>
                    <td>
                      <div className="flex gap-2" style={{alignItems: 'center'}}>
                        <select className="select" style={{fontSize: 11, padding: '4px 8px'}} value={s.estatus} onChange={e => onUpdateEstatus(s.id, e.target.value)}>
                          <option value="PENDIENTE">🔴 PENDIENTE</option>
                          <option value="EN PROCESO">🟡 EN PROCESO</option>
                          <option value="CERRADO">🟢 CERRADO</option>
                        </select>
                        <button className="btn btn-ghost btn-sm" onClick={() => setMsgModal(s)} title="Generar Solicitud" style={{padding: '4px 8px'}}>
                          <Icons.Templates />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {msgModal && (
        <SiniestroMessageModal siniestro={msgModal} onClose={() => setMsgModal(null)} />
      )}
    </div>
  );
}

function SiniestroMessageModal({ siniestro, onClose }) {
  useEscapeKey(onClose);
  const [copied, setCopied] = useState(null);

  const t1 = `Estimado ejecutivo, por medio de la presente solicito su apoyo con el estatus y/o generación de pase a corralón para la unidad del asegurado ${siniestro.asegurado || '[Nombre del Asegurado]'}, correspondiente a la Póliza ${siniestro.poliza || '[Número de Póliza]'}, Vehículo ${siniestro.vehiculo || '[Descripción del Vehículo / Serie]'}, con reporte de siniestro ${siniestro.cvestro || '[CVESTRO]'}. Quedo atento a sus comentarios. Saludos cordiales.`;

  const t2 = `Estimado ejecutivo, solicitamos su apoyo para verificar si es posible realizar una propuesta de pago de daños para un tercero afectado en el siniestro de la póliza ${siniestro.poliza || '[Número de Póliza]'} del asegurado ${siniestro.asegurado || '[Nombre del Asegurado]'}, buscando posteriormente que dicho tercero se asegure con nosotros. Agradezco de antemano su atención.`;

  const copy = (txt, id) => {
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(txt);
    } else {
      // Fallback
      let textArea = document.createElement("textarea");
      textArea.value = txt;
      textArea.style.position = "fixed";
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      try {
        document.execCommand('copy');
      } catch (err) {
        console.error('Fallback: Oops, unable to copy', err);
      }
      document.body.removeChild(textArea);
    }
    setCopied(id);
    setTimeout(() => setCopied(null), 2000);
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{maxWidth: 650, width: '90%'}} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span className="modal-title">Generar Solicitud - Siniestros</span>
          <button className="btn btn-ghost" onClick={onClose}><Icons.Close /></button>
        </div>
        <div className="modal-body" style={{padding: 24, display: 'flex', flexDirection: 'column', gap: 24}}>
          
          <div>
            <div className="flex gap-2" style={{justifyContent: 'space-between', marginBottom: 8, alignItems: 'center'}}>
              <span style={{fontWeight: 600, color: 'var(--accent-blue)'}}>1. Estatus / Pase a Corralón</span>
              <button className="btn btn-primary btn-sm" onClick={() => copy(t1, 1)}>
                {copied === 1 ? '✅ Copiado' : '📄 Copiar'}
              </button>
            </div>
            <textarea className="input" rows={5} readOnly value={t1} style={{fontSize: 13, lineHeight: 1.5, background: 'rgba(255,255,255,0.02)', resize: 'none'}} />
          </div>

          <div>
            <div className="flex gap-2" style={{justifyContent: 'space-between', marginBottom: 8, alignItems: 'center'}}>
              <span style={{fontWeight: 600, color: 'var(--accent-blue)'}}>2. Pago a Terceros</span>
              <button className="btn btn-primary btn-sm" onClick={() => copy(t2, 2)}>
                {copied === 2 ? '✅ Copiado' : '📄 Copiar'}
              </button>
            </div>
            <textarea className="input" rows={4} readOnly value={t2} style={{fontSize: 13, lineHeight: 1.5, background: 'rgba(255,255,255,0.02)', resize: 'none'}} />
          </div>

        </div>
      </div>
    </div>
  );
}

// ─── App Principal ────────────────────────────────────────────
function App() {
  const [page, setPage] = useState('dashboard');
  const [defaultEstatus, setDefaultEstatus] = useState('TODOS');
  const [policies, setPolicies] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_policies');
      if (!stored) return [];
      const parsed = JSON.parse(stored);
      const hasDemo = parsed.some(p => p.poliza === 'POL-2024-001' || p.poliza === 'POL-2024-002');
      if (hasDemo) {
        localStorage.removeItem('sc_policies');
        return [];
      }
      return parsed;
    } catch { return []; }
  });
  const [siniestros, setSiniestros] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_siniestros');
      if (!stored) return [];
      const parsed = JSON.parse(stored);
      // Limpiar datos de ejemplo
      const filtered = parsed.filter(s => s.poliza !== 'POL-123' || s.asegurado !== 'Ejemplo Asegurado');
      if (filtered.length !== parsed.length) {
        localStorage.setItem('sc_siniestros', JSON.stringify(filtered));
      }
      return filtered;
    } catch { return []; }
  });
  const [cotizaciones, setCotizaciones] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_cotizaciones');
      return stored ? JSON.parse(stored) : [];
    } catch { return []; }
  });
  const [templates, setTemplates] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_templates');
      return stored ? JSON.parse(stored) : DEFAULT_TEMPLATES;
    } catch { return DEFAULT_TEMPLATES; }
  });

  const [modalNew, setModalNew] = useState(false);
  const [modalEdit, setModalEdit] = useState(null);
  const [modalPaid, setModalPaid] = useState(null);
  const [modalContact, setModalContact] = useState(null); // { policy, type }
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [dailyModalDate, setDailyModalDate] = useState(null);
  const [showCalendarPicker, setShowCalendarPicker] = useState(false);

  useEscapeKey(deleteConfirm ? () => setDeleteConfirm(null) : null);

  const { toasts, toast } = useToast();

  const [caroPolicies, setCaroPolicies] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_caro_policies');
      return stored ? JSON.parse(stored) : [];
    } catch { return []; }
  });

  const [gmmPolicies, setGmmPolicies] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_gmm_policies');
      return stored ? JSON.parse(stored) : [];
    } catch { return []; }
  });

  const [autosPolicies, setAutosPolicies] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_autos_policies');
      return stored ? JSON.parse(stored) : [];
    } catch { return []; }
  });

  const [vidaPolicies, setVidaPolicies] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_vida_policies');
      return stored ? JSON.parse(stored) : [];
    } catch { return []; }
  });

  const [danosPolicies, setDanosPolicies] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_danos_policies');
      return stored ? JSON.parse(stored) : [];
    } catch { return []; }
  });

  const [hogarPolicies, setHogarPolicies] = useState(() => {
    try {
      const stored = localStorage.getItem('sc_hogar_policies');
      return stored ? JSON.parse(stored) : [];
    } catch { return []; }
  });

  const [dbConnected, setDbConnected] = useState(false);
  const FIREBASE_REST_URL = 'https://pre-pro-consultores-gestion-default-rtdb.firebaseio.com/app_data';

  const parseList = (node) => {
    if (!node) return [];
    if (Array.isArray(node)) return node.filter(Boolean);
    if (typeof node === 'object') return Object.values(node).filter(Boolean);
    return [];
  };

  const applyCloudData = useCallback((data) => {
    if (!data) return;
    setDbConnected(true);

    const pList = parseList(data.policies);
    setPolicies(pList);
    localStorage.setItem('sc_policies', JSON.stringify(pList));

    const cList = parseList(data.caroPolicies);
    setCaroPolicies(cList);
    localStorage.setItem('sc_caro_policies', JSON.stringify(cList));

    const gList = parseList(data.gmmPolicies);
    setGmmPolicies(gList);
    localStorage.setItem('sc_gmm_policies', JSON.stringify(gList));

    const aList = parseList(data.autosPolicies);
    setAutosPolicies(aList);
    localStorage.setItem('sc_autos_policies', JSON.stringify(aList));

    const vList = parseList(data.vidaPolicies);
    setVidaPolicies(vList);
    localStorage.setItem('sc_vida_policies', JSON.stringify(vList));

    const dList = parseList(data.danosPolicies);
    setDanosPolicies(dList);
    localStorage.setItem('sc_danos_policies', JSON.stringify(dList));

    const hList = parseList(data.hogarPolicies);
    setHogarPolicies(hList);
    localStorage.setItem('sc_hogar_policies', JSON.stringify(hList));

    const sList = parseList(data.siniestros);
    setSiniestros(sList);
    localStorage.setItem('sc_siniestros', JSON.stringify(sList));

    const cotList = parseList(data.cotizaciones);
    setCotizaciones(cotList);
    localStorage.setItem('sc_cotizaciones', JSON.stringify(cotList));

    if (data.templates) {
      setTemplates(data.templates);
      localStorage.setItem('sc_templates', JSON.stringify(data.templates));
    }
  }, []);

  const syncCategoryToCloud = useCallback((category, list) => {
    const cleanList = list || [];
    // 1. Firebase Realtime SDK update
    if (window.db) {
      try { window.db.ref(`app_data/${category}`).set(cleanList); } catch(e) {}
    }
    // 2. HTTP REST update (Inquebrantable para cualquier navegador/red)
    fetch(`${FIREBASE_REST_URL}/${category}.json`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(cleanList)
    }).catch(err => console.error('Cloud REST PUT error:', err));
  }, []);

  // ☁️ Sincronización Híbrida en tiempo real (REST + WebSockets) con Firebase
  useEffect(() => {
    // 1. Carga REST Inmediata (funciona en TODOS los perfiles de Chrome, teléfonos e Incógnito)
    fetch(`${FIREBASE_REST_URL}.json`)
      .then(res => res.json())
      .then(data => {
        if (data) applyCloudData(data);
      })
      .catch(err => console.error('Cloud REST GET error:', err));

    // 2. Escuchador Realtime WebSocket (para cambios en vivo instantáneos)
    if (!window.db) return;

    const connectedRef = window.db.ref('.info/connected');
    const onConnected = (snap) => {
      if (snap.val() === true) setDbConnected(true);
    };
    connectedRef.on('value', onConnected);

    const dbRef = window.db.ref('app_data');
    const handleValue = (snapshot) => {
      const data = snapshot.val();
      if (data) applyCloudData(data);
    };
    dbRef.on('value', handleValue);

    return () => {
      connectedRef.off('value', onConnected);
      dbRef.off('value', handleValue);
    };
  }, [applyCloudData]);

  // Función para forzar la actualización/descarga desde la nube por REST + WebSockets
  const syncNowFromCloud = useCallback(() => {
    fetch(`${FIREBASE_REST_URL}.json`)
      .then(res => res.json())
      .then(data => {
        if (data) {
          applyCloudData(data);
          toast('¡Sincronizado con la Nube con éxito! ☁️✅', 'success');
        } else {
          toast('Nube vacía o sin respuesta', 'warning');
        }
      })
      .catch(err => {
        toast('Error al descargar de la nube: ' + err.message, 'error');
      });
  }, [applyCloudData, toast]);

  // Función para forzar la subida de datos locales a la nube por REST + WebSockets
  const uploadLocalToCloud = useCallback(() => {
    const localPols = JSON.parse(localStorage.getItem('sc_policies') || '[]');
    const localCaro = JSON.parse(localStorage.getItem('sc_caro_policies') || '[]');
    const localGmm = JSON.parse(localStorage.getItem('sc_gmm_policies') || '[]');
    const localAutos = JSON.parse(localStorage.getItem('sc_autos_policies') || '[]');
    const localVida = JSON.parse(localStorage.getItem('sc_vida_policies') || '[]');
    const localDanos = JSON.parse(localStorage.getItem('sc_danos_policies') || '[]');
    const localHogar = JSON.parse(localStorage.getItem('sc_hogar_policies') || '[]');
    const localSini = JSON.parse(localStorage.getItem('sc_siniestros') || '[]');
    const localCoti = JSON.parse(localStorage.getItem('sc_cotizaciones') || '[]');
    const localTpls = JSON.parse(localStorage.getItem('sc_templates') || 'null') || DEFAULT_TEMPLATES;

    const payload = {
      policies: localPols,
      caroPolicies: localCaro,
      gmmPolicies: localGmm,
      autosPolicies: localAutos,
      vidaPolicies: localVida,
      danosPolicies: localDanos,
      hogarPolicies: localHogar,
      siniestros: localSini,
      cotizaciones: localCoti,
      templates: localTpls
    };

    if (window.db) {
      try { window.db.ref('app_data').set(payload); } catch(e) {}
    }

    fetch(`${FIREBASE_REST_URL}.json`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(() => {
      toast('¡Datos subidos a la Nube con éxito! ☁️✅', 'success');
    }).catch(err => {
      toast('Error al subir a la nube: ' + err.message, 'error');
    });
  }, [toast]);

  // Guardar en localStorage cuando el usuario modifica los datos
  useEffect(() => { localStorage.setItem('sc_policies', JSON.stringify(policies)); }, [policies]);
  useEffect(() => { localStorage.setItem('sc_caro_policies', JSON.stringify(caroPolicies)); }, [caroPolicies]);
  useEffect(() => { localStorage.setItem('sc_gmm_policies', JSON.stringify(gmmPolicies)); }, [gmmPolicies]);
  useEffect(() => { localStorage.setItem('sc_autos_policies', JSON.stringify(autosPolicies)); }, [autosPolicies]);
  useEffect(() => { localStorage.setItem('sc_vida_policies', JSON.stringify(vidaPolicies)); }, [vidaPolicies]);
  useEffect(() => { localStorage.setItem('sc_danos_policies', JSON.stringify(danosPolicies)); }, [danosPolicies]);
  useEffect(() => { localStorage.setItem('sc_hogar_policies', JSON.stringify(hogarPolicies)); }, [hogarPolicies]);
  useEffect(() => { localStorage.setItem('sc_siniestros', JSON.stringify(siniestros)); }, [siniestros]);
  useEffect(() => { localStorage.setItem('sc_cotizaciones', JSON.stringify(cotizaciones)); }, [cotizaciones]);
  useEffect(() => { localStorage.setItem('sc_templates', JSON.stringify(templates)); }, [templates]);

  const urgentCount = useMemo(() => policies.filter(p => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    return isUpcomingReminder(p) || isExpiredEffective(p);
  }).length, [policies]);

  const caroUrgentCount = useMemo(() => caroPolicies.filter(p => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    return isUpcomingReminder(p) || isExpiredEffective(p);
  }).length, [caroPolicies]);

  const gmmUrgentCount = useMemo(() => gmmPolicies.filter(p => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    return isUpcomingReminder(p) || isExpiredEffective(p);
  }).length, [gmmPolicies]);

  const autosUrgentCount = useMemo(() => autosPolicies.filter(p => {
    if (p.estatus === 'PAGADO' || p.estatus === 'CANCELADO' || p.estatus === 'LIQUIDADO') return false;
    return isUpcomingReminder(p) || isExpiredEffective(p);
  }).length, [autosPolicies]);

  // CRUD
  const savePolicy = useCallback((p) => {
    const policyToSave = { ...p, id: p.id || generateId() };
    setPolicies(prev => {
      const exists = prev.find(x => x.id === policyToSave.id);
      const next = exists ? prev.map(x => x.id === policyToSave.id ? policyToSave : x) : [...prev, policyToSave];
      localStorage.setItem('sc_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('policies', next), 0);
      return next;
    });
  }, []);

  const purgePolicyFromAllCategories = useCallback((id) => {
    setPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('policies', next), 0);
      return next;
    });
    setCaroPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_caro_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('caroPolicies', next), 0);
      return next;
    });
    setGmmPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_gmm_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('gmmPolicies', next), 0);
      return next;
    });
    setAutosPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_autos_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('autosPolicies', next), 0);
      return next;
    });
    setVidaPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_vida_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('vidaPolicies', next), 0);
      return next;
    });
    setDanosPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_danos_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('danosPolicies', next), 0);
      return next;
    });
    setHogarPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_hogar_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('hogarPolicies', next), 0);
      return next;
    });
    setDeleteConfirm(null);
  }, [syncCategoryToCloud]);

  const deletePolicy = useCallback((id) => {
    purgePolicyFromAllCategories(id);
    toast('Póliza eliminada', 'warning');
  }, [purgePolicyFromAllCategories, toast]);

  // Marcar como pagado → re-agendar
  const markPaid = useCallback((policy, nextDate, comprobante, isLastPayment = false, nextMonto) => {
    setPolicies(prev => {
      const next = prev.map(p => {
        if (p.id !== policy.id) return p;
        const basePolicy = { 
          ...p, 
          comprobante: comprobante || p.comprobante,
          fechaPagoAnterior: p.fechaPago,
          fechaUltimoPago: new Date().toISOString().split('T')[0],
          periodoGracia: '',
          monto: (nextMonto !== undefined && nextMonto !== '') ? nextMonto : (p.montoSubsecuente || p.monto)
        };
        if (policy.formaPago === 'CONTADO' || isLastPayment) {
          return { ...basePolicy, estatus: 'LIQUIDADO', fechaPago: p.fechaPago };
        }
        return { ...basePolicy, estatus: 'PENDIENTE', fechaPago: nextDate || p.fechaPago };
      });
      localStorage.setItem('sc_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('policies', next), 0);
      return next;
    });
    toast('Pago confirmado', 'success');
  }, [toast]);

  const importToCategory = useCallback((data, mode, categoryKey) => {
    let setter, storageKey;
    switch (categoryKey) {
      case 'caroPolicies': setter = setCaroPolicies; storageKey = 'sc_caro_policies'; break;
      case 'gmmPolicies': setter = setGmmPolicies; storageKey = 'sc_gmm_policies'; break;
      case 'autosPolicies': setter = setAutosPolicies; storageKey = 'sc_autos_policies'; break;
      case 'vidaPolicies': setter = setVidaPolicies; storageKey = 'sc_vida_policies'; break;
      case 'danosPolicies': setter = setDanosPolicies; storageKey = 'sc_danos_policies'; break;
      case 'hogarPolicies': setter = setHogarPolicies; storageKey = 'sc_hogar_policies'; break;
      default: setter = setPolicies; storageKey = 'sc_policies'; categoryKey = 'policies'; break;
    }

    setter(prev => {
      const next = mode === 'reemplazar' ? data : [...prev, ...data];
      localStorage.setItem(storageKey, JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud(categoryKey, next), 0);
      return next;
    });
  }, []);

  const importSiniestros = useCallback((incomingData) => {
    setSiniestros(prev => {
      const next = [...prev];
      let added = 0;
      let updated = 0;

      incomingData.forEach(inc => {
        const existingIdx = next.findIndex(s => s.poliza === inc.poliza);
        if (existingIdx >= 0) {
          const existing = next[existingIdx];
          next[existingIdx] = {
            ...inc,
            id: existing.id,
            estatus: existing.estatus || 'PENDIENTE'
          };
          updated++;
        } else {
          next.push({
            ...inc,
            id: generateId(),
            estatus: 'PENDIENTE'
          });
          added++;
        }
      });
      
      localStorage.setItem('sc_siniestros', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('siniestros', next), 0);
      toast(`Importación completada: ${added} nuevos, ${updated} actualizados.`, 'success');
      return next;
    });
  }, [toast]);

  const saveCaroPolicy = useCallback((p) => {
    const policyToSave = { ...p, id: p.id || generateId() };
    setCaroPolicies(prev => {
      const exists = prev.find(x => x.id === policyToSave.id);
      const next = exists ? prev.map(x => x.id === policyToSave.id ? policyToSave : x) : [...prev, policyToSave];
      localStorage.setItem('sc_caro_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('caroPolicies', next), 0);
      return next;
    });
  }, []);

  const deleteCaroPolicy = useCallback((id) => {
    purgePolicyFromAllCategories(id);
    toast('Póliza eliminada', 'warning');
  }, [purgePolicyFromAllCategories, toast]);

  const markCaroPaid = useCallback((policy, nextDate, comprobante, isLastPayment = false, nextMonto) => {
    setCaroPolicies(prev => {
      const next = prev.map(p => {
        if (p.id !== policy.id) return p;
        const basePolicy = { 
          ...p, 
          comprobante: comprobante || p.comprobante,
          fechaPagoAnterior: p.fechaPago,
          fechaUltimoPago: new Date().toISOString().split('T')[0],
          periodoGracia: '',
          monto: (nextMonto !== undefined && nextMonto !== '') ? nextMonto : (p.montoSubsecuente || p.monto)
        };
        if (policy.formaPago === 'CONTADO' || isLastPayment) {
          return { ...basePolicy, estatus: 'LIQUIDADO', fechaPago: p.fechaPago };
        }
        return { ...basePolicy, estatus: 'PENDIENTE', fechaPago: nextDate || p.fechaPago };
      });
      localStorage.setItem('sc_caro_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('caroPolicies', next), 0);
      return next;
    });
    toast('Pago confirmado', 'success');
  }, [toast]);

  const saveGmmPolicy = useCallback((p) => {
    const policyToSave = { ...p, id: p.id || generateId() };
    setGmmPolicies(prev => {
      const exists = prev.find(x => x.id === policyToSave.id);
      const next = exists ? prev.map(x => x.id === policyToSave.id ? policyToSave : x) : [...prev, policyToSave];
      localStorage.setItem('sc_gmm_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('gmmPolicies', next), 0);
      return next;
    });
  }, []);

  const deleteGmmPolicy = useCallback((id) => {
    purgePolicyFromAllCategories(id);
    toast('Póliza GMM eliminada', 'warning');
  }, [purgePolicyFromAllCategories, toast]);

  const markGmmPaid = useCallback((policy, nextDate, comprobante, isLastPayment = false, nextMonto) => {
    setGmmPolicies(prev => {
      const next = prev.map(p => {
        if (p.id !== policy.id) return p;
        const basePolicy = { 
          ...p, 
          comprobante: comprobante || p.comprobante,
          fechaPagoAnterior: p.fechaPago,
          fechaUltimoPago: new Date().toISOString().split('T')[0],
          periodoGracia: '',
          monto: (nextMonto !== undefined && nextMonto !== '') ? nextMonto : (p.montoSubsecuente || p.monto)
        };
        if (policy.formaPago === 'CONTADO' || isLastPayment) {
          return { ...basePolicy, estatus: 'LIQUIDADO', fechaPago: p.fechaPago };
        }
        return { ...basePolicy, estatus: 'PENDIENTE', fechaPago: nextDate || p.fechaPago };
      });
      localStorage.setItem('sc_gmm_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('gmmPolicies', next), 0);
      return next;
    });
    toast('Pago confirmado', 'success');
  }, [toast]);

  const saveAutosPolicy = useCallback((p) => {
    const policyToSave = { ...p, id: p.id || generateId() };
    setAutosPolicies(prev => {
      const exists = prev.find(x => x.id === policyToSave.id);
      const next = exists ? prev.map(x => x.id === policyToSave.id ? policyToSave : x) : [...prev, policyToSave];
      localStorage.setItem('sc_autos_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('autosPolicies', next), 0);
      return next;
    });
  }, []);

  const deleteAutosPolicy = useCallback((id) => {
    purgePolicyFromAllCategories(id);
    toast('Póliza de Autos eliminada', 'warning');
  }, [purgePolicyFromAllCategories, toast]);

  const markAutosPaid = useCallback((policy, nextDate, comprobante, isLastPayment = false, nextMonto) => {
    setAutosPolicies(prev => {
      const next = prev.map(p => {
        if (p.id !== policy.id) return p;
        const basePolicy = { 
          ...p, 
          comprobante: comprobante || p.comprobante,
          fechaPagoAnterior: p.fechaPago,
          fechaUltimoPago: new Date().toISOString().split('T')[0],
          periodoGracia: '',
          monto: (nextMonto !== undefined && nextMonto !== '') ? nextMonto : (p.montoSubsecuente || p.monto)
        };
        if (policy.formaPago === 'CONTADO' || isLastPayment) {
          return { ...basePolicy, estatus: 'LIQUIDADO', fechaPago: p.fechaPago };
        }
        return { ...basePolicy, estatus: 'PENDIENTE', fechaPago: nextDate || p.fechaPago };
      });
      localStorage.setItem('sc_autos_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('autosPolicies', next), 0);
      return next;
    });
    toast('Pago confirmado', 'success');
  }, [toast]);

  const saveVidaPolicy = useCallback((p) => {
    const policyToSave = { ...p, id: p.id || generateId() };
    setVidaPolicies(prev => {
      const exists = prev.find(x => x.id === policyToSave.id);
      const next = exists ? prev.map(x => x.id === policyToSave.id ? policyToSave : x) : [...prev, policyToSave];
      localStorage.setItem('sc_vida_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('vidaPolicies', next), 0);
      return next;
    });
  }, []);

  const deleteVidaPolicy = useCallback((id) => {
    purgePolicyFromAllCategories(id);
    toast('Póliza de Vida eliminada', 'warning');
  }, [purgePolicyFromAllCategories, toast]);

  const markVidaPaid = useCallback((policy, nextDate, comprobante, isLastPayment = false, nextMonto) => {
    setVidaPolicies(prev => {
      const next = prev.map(p => {
        if (p.id !== policy.id) return p;
        const basePolicy = { 
          ...p, 
          comprobante: comprobante || p.comprobante,
          fechaPagoAnterior: p.fechaPago,
          fechaUltimoPago: todayISO(),
          periodoGracia: '',
          monto: (nextMonto !== undefined && nextMonto !== '') ? nextMonto : (p.montoSubsecuente || p.monto)
        };
        if (policy.formaPago === 'CONTADO' || isLastPayment) {
          return { ...basePolicy, estatus: 'LIQUIDADO', fechaPago: p.fechaPago };
        }
        return { ...basePolicy, estatus: 'PENDIENTE', fechaPago: nextDate || p.fechaPago };
      });
      localStorage.setItem('sc_vida_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('vidaPolicies', next), 0);
      return next;
    });
    toast('Pago confirmado', 'success');
  }, [toast]);

  const saveDanosPolicy = useCallback((p) => {
    const policyToSave = { ...p, id: p.id || generateId() };
    setDanosPolicies(prev => {
      const exists = prev.find(x => x.id === policyToSave.id);
      const next = exists ? prev.map(x => x.id === policyToSave.id ? policyToSave : x) : [...prev, policyToSave];
      localStorage.setItem('sc_danos_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('danosPolicies', next), 0);
      return next;
    });
  }, []);

  const deleteDanosPolicy = useCallback((id) => {
    purgePolicyFromAllCategories(id);
    toast('Póliza de Daños eliminada', 'warning');
  }, [purgePolicyFromAllCategories, toast]);

  const markDanosPaid = useCallback((policy, nextDate, comprobante, isLastPayment = false, nextMonto) => {
    setDanosPolicies(prev => {
      const next = prev.map(p => {
        if (p.id !== policy.id) return p;
        const basePolicy = { 
          ...p, 
          comprobante: comprobante || p.comprobante,
          fechaPagoAnterior: p.fechaPago,
          fechaUltimoPago: todayISO(),
          periodoGracia: '',
          monto: (nextMonto !== undefined && nextMonto !== '') ? nextMonto : (p.montoSubsecuente || p.monto)
        };
        if (policy.formaPago === 'CONTADO' || isLastPayment) {
          return { ...basePolicy, estatus: 'LIQUIDADO', fechaPago: p.fechaPago };
        }
        return { ...basePolicy, estatus: 'PENDIENTE', fechaPago: nextDate || p.fechaPago };
      });
      localStorage.setItem('sc_danos_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('danosPolicies', next), 0);
      return next;
    });
    toast('Pago confirmado', 'success');
  }, [toast]);

  const saveHogarPolicy = useCallback((p) => {
    const policyToSave = { ...p, id: p.id || generateId() };
    setHogarPolicies(prev => {
      const exists = prev.find(x => x.id === policyToSave.id);
      const next = exists ? prev.map(x => x.id === policyToSave.id ? policyToSave : x) : [...prev, policyToSave];
      localStorage.setItem('sc_hogar_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('hogarPolicies', next), 0);
      return next;
    });
  }, []);

  const deleteHogarPolicy = useCallback((id) => {
    setHogarPolicies(prev => {
      const next = prev.filter(p => p.id !== id);
      localStorage.setItem('sc_hogar_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('hogarPolicies', next), 0);
      return next;
    });
    toast('Póliza de Hogar eliminada', 'warning');
    setDeleteConfirm(null);
  }, [toast]);

  const markHogarPaid = useCallback((policy, nextDate, comprobante, isLastPayment = false, nextMonto) => {
    setHogarPolicies(prev => {
      const next = prev.map(p => {
        if (p.id !== policy.id) return p;
        const basePolicy = { 
          ...p, 
          comprobante: comprobante || p.comprobante,
          fechaPagoAnterior: p.fechaPago,
          fechaUltimoPago: todayISO(),
          periodoGracia: '',
          monto: (nextMonto !== undefined && nextMonto !== '') ? nextMonto : (p.montoSubsecuente || p.monto)
        };
        if (policy.formaPago === 'CONTADO' || isLastPayment) {
          return { ...basePolicy, estatus: 'LIQUIDADO', fechaPago: p.fechaPago };
        }
        return { ...basePolicy, estatus: 'PENDIENTE', fechaPago: nextDate || p.fechaPago };
      });
      localStorage.setItem('sc_hogar_policies', JSON.stringify(next));
      setTimeout(() => syncCategoryToCloud('hogarPolicies', next), 0);
      return next;
    });
    toast('Pago confirmado', 'success');
  }, [toast]);

  const updateSiniestroEstatus = useCallback((id, estatus) => {
    setSiniestros(prev => {
      const next = prev.map(s => s.id === id ? { ...s, estatus } : s);
      localStorage.setItem('sc_siniestros', JSON.stringify(next));
      if (window.db) window.db.ref('app_data/siniestros').set(next);
      return next;
    });
  }, []);

  const saveCotizacion = useCallback((coti) => {
    setCotizaciones(prev => {
      const next = [coti, ...prev];
      localStorage.setItem('sc_cotizaciones', JSON.stringify(next));
      if (window.db) window.db.ref('app_data/cotizaciones').set(next);
      return next;
    });
    toast('Cotización registrada', 'success');
  }, [toast]);

  const updateCotizacionEstatus = useCallback((id, estatus) => {
    setCotizaciones(prev => {
      const next = prev.map(c => c.id === id ? { ...c, estatus } : c);
      localStorage.setItem('sc_cotizaciones', JSON.stringify(next));
      if (window.db) window.db.ref('app_data/cotizaciones').set(next);
      return next;
    });
  }, []);

  const vidaUrgentCount = useMemo(() => vidaPolicies.filter(p => isUpcomingReminder(p) || isExpiredEffective(p)).length, [vidaPolicies]);
  const danosUrgentCount = useMemo(() => danosPolicies.filter(p => isUpcomingReminder(p) || isExpiredEffective(p)).length, [danosPolicies]);
  const hogarUrgentCount = useMemo(() => hogarPolicies.filter(p => isUpcomingReminder(p) || isExpiredEffective(p)).length, [hogarPolicies]);

  const navItems = [
    { id: 'dashboard', label: 'Panel de Control', Icon: Icons.Dashboard },
    { id: 'policies', label: 'Autos Qualitas', Icon: Icons.Policies },
    { id: 'caro_policies', label: 'Autos Qualitas Caro', Icon: Icons.Policies, badge: caroUrgentCount > 0 ? caroUrgentCount : null },
    { id: 'gmm_policies', label: 'GMM', Icon: Icons.Shield, badge: gmmUrgentCount > 0 ? gmmUrgentCount : null },
    { id: 'autos_policies', label: 'Autos (Otras Aseg.)', Icon: Icons.Policies, badge: autosUrgentCount > 0 ? autosUrgentCount : null },
    { id: 'vida_policies', label: 'Vida', Icon: Icons.Heart, badge: vidaUrgentCount > 0 ? vidaUrgentCount : null },
    { id: 'danos_policies', label: 'Daños', Icon: Icons.Briefcase, badge: danosUrgentCount > 0 ? danosUrgentCount : null },
    { id: 'hogar_policies', label: 'Hogar', Icon: Icons.Home, badge: hogarUrgentCount > 0 ? hogarUrgentCount : null },
    { id: 'templates', label: 'Plantillas', Icon: Icons.Templates },
    { id: 'import', label: 'Importar / Exportar', Icon: Icons.Import },
  ];

  const pageTitles = {
    dashboard: 'Panel de Control',
    policies: 'Pólizas Autos Qualitas',
    caro_policies: 'Pólizas Autos Qualitas Caro',
    gmm_policies: 'Pólizas Gastos Médicos Mayores (GMM)',
    autos_policies: 'Autos (Otras Aseguradoras)',
    vida_policies: 'Pólizas de Vida',
    danos_policies: 'Pólizas de Daños',
    hogar_policies: 'Pólizas de Hogar',
    templates: 'Plantillas de Mensajes',
    import: 'Importar / Exportar',
  };

  const allPolicies = useMemo(() => [
    ...policies,
    ...caroPolicies.map(p => ({ ...p, _isCaro: true })),
    ...gmmPolicies.map(p => ({ ...p, _isGmm: true })),
    ...autosPolicies.map(p => ({ ...p, _isAutos: true })),
    ...vidaPolicies.map(p => ({ ...p, _isVida: true })),
    ...danosPolicies.map(p => ({ ...p, _isDanos: true })),
    ...hogarPolicies.map(p => ({ ...p, _isHogar: true }))
  ], [policies, caroPolicies, gmmPolicies, autosPolicies, vidaPolicies, danosPolicies, hogarPolicies]);

  const allProps = {
    policies: allPolicies,
    onEdit: (p) => setModalEdit(p),
    onDelete: (p) => setDeleteConfirm(p),
    onMarkPaid: (p) => setModalPaid(p),
    onWhatsApp: (p) => setModalContact({ policy: p, type: 'whatsapp' }),
    onEmail: (p) => setModalContact({ policy: p, type: 'email' }),
    onUpdatePolicy: (p) => {
      if (p._isCaro) saveCaroPolicy(p);
      else if (p._isGmm) saveGmmPolicy(p);
      else if (p._isAutos) saveAutosPolicy(p);
      else if (p._isVida) saveVidaPolicy(p);
      else if (p._isDanos) saveDanosPolicy(p);
      else if (p._isHogar) saveHogarPolicy(p);
      else savePolicy(p);
    }
  };

  const commonProps = {
    policies,
    onEdit: (p) => setModalEdit(p),
    onDelete: (p) => setDeleteConfirm(p),
    onMarkPaid: (p) => setModalPaid(p),
    onWhatsApp: (p) => setModalContact({ policy: p, type: 'whatsapp' }),
    onEmail: (p) => setModalContact({ policy: p, type: 'email' }),
    onUpdatePolicy: savePolicy,
  };

  return (
    <div className="app-layout">
      {/* Sidebar */}
      <aside className="sidebar">
        <div className="sidebar-logo" style={{ padding: '24px 20px', textAlign: 'center' }}>
          <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'stretch', margin: '0 auto' }}>
            <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'baseline', justifyContent: 'center', fontSize: '28px', fontWeight: '800', lineHeight: 1, fontFamily: 'Times New Roman, serif', letterSpacing: '-0.5px' }}>
              <span style={{ color: '#1771c5' }}>PRE</span>
              <span style={{ color: '#111111', margin: '0 2px' }}>&amp;</span>
              <span style={{ color: '#1ba54b' }}>PRO</span>
            </div>
            <div style={{ width: '100%', height: '2px', background: '#ea7d23', margin: '4px 0' }} />
            <div style={{ fontSize: '11px', color: '#a3a3a3', letterSpacing: '4px', textTransform: 'uppercase', fontFamily: 'Times New Roman, serif', textAlign: 'center' }}>
              C O N S U L T O R E S
            </div>
          </div>
          <span style={{ display: 'block', marginTop: '16px', fontSize: '11px', color: 'var(--text-muted)' }}>Sistema de Cobranza Interna</span>
        </div>
        <nav className="sidebar-nav">
          {navItems.map(({ id, label, Icon, badge }) => (
            <button key={id} className={`nav-item ${page === id ? 'active' : ''}`}
              onClick={() => setPage(id)}>
              <Icon />
              {label}
              {badge && <span className="nav-badge">{badge}</span>}
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11 }}>
            <span style={{ width: 8, height: 8, borderRadius: '50%', background: dbConnected ? '#10b981' : '#f59e0b', display: 'inline-block' }} />
            {dbConnected ? '🟢 Sincronizado en Nube' : '🟡 Conectando Nube...'}
          </p>
          <p style={{marginTop:4, fontSize: 11}}>{allPolicies.length} pólizas registradas</p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4, marginTop: 8 }}>
            <button 
              onClick={syncNowFromCloud}
              className="btn btn-primary btn-sm"
              style={{ fontSize: 10, padding: '4px 6px', width: '100%', textTransform: 'none' }}
              title="Descargar las pólizas más recientes guardadas en la Nube"
            >
              🔄 Sincronizar Nube Ahora
            </button>
            <button 
              onClick={uploadLocalToCloud}
              className="btn btn-ghost btn-sm"
              style={{ fontSize: 10, padding: '3px 6px', width: '100%', textTransform: 'none', border: '1px solid var(--border)' }}
              title="Subir las pólizas guardadas localmente en esta computadora a Firebase"
            >
              ☁️ Subir Datos a Nube
            </button>
          </div>
        </div>
      </aside>

      {/* Main */}
      <div className="main-content">
        <header className="topbar">
          <h2 className="topbar-title">{pageTitles[page]}</h2>
          <div className="topbar-actions">
            <div 
              style={{ position: 'relative', display: 'flex', alignItems: 'center', cursor: 'pointer', background: 'var(--bg-card)', border: '1px solid var(--border)', padding: '6px 14px', borderRadius: 20, gap: 6 }} 
              title="Haz clic para ver el calendario de cobros"
              onClick={() => setShowCalendarPicker(true)}
            >
              <span style={{fontSize:13, color:'var(--text-muted)'}}>
                📅 {new Date().toLocaleDateString('es-MX', { weekday:'long', day:'2-digit', month:'long', year:'numeric' })}
              </span>
            </div>
            {(page === 'dashboard' || page === 'policies') && (
              <button className="btn btn-primary btn-sm" onClick={() => setModalNew(true)}>
                <Icons.Plus /> Nueva Póliza
              </button>
            )}
          </div>
        </header>

        <main className="page-content">
          {page === 'dashboard' && (
            <DashboardPage {...allProps} onNew={() => setModalNew(true)} onStatClick={(estatus) => {
              setDefaultEstatus(estatus);
              setPage('policies');
            }} />
          )}
          {page === 'policies' && (
            <PoliciesPage {...commonProps} defaultEstatus={defaultEstatus} onNew={() => setModalNew(true)} />
          )}
          {page === 'urgent' && (
            <UrgentPage {...allProps} />
          )}
          {page === 'caro_policies' && (
            <CaroPoliciesPage 
              policies={caroPolicies.map(p => ({ ...p, _isCaro: true }))} 
              onSave={saveCaroPolicy} 
              onDelete={deleteCaroPolicy} 
              onMarkPaid={markCaroPaid} 
              onWhatsApp={p => setModalContact({ policy: p, type: 'whatsapp' })}
              onEmail={p => setModalContact({ policy: p, type: 'email' })}
              toast={toast}
            />
          )}
          {page === 'gmm_policies' && (
            <GmmPoliciesPage 
              policies={gmmPolicies.map(p => ({ ...p, _isGmm: true }))} 
              onSave={saveGmmPolicy} 
              onDelete={deleteGmmPolicy} 
              onMarkPaid={markGmmPaid} 
              onWhatsApp={p => setModalContact({ policy: p, type: 'whatsapp' })}
              onEmail={p => setModalContact({ policy: p, type: 'email' })}
              toast={toast}
            />
          )}
          {page === 'autos_policies' && (
            <AutosOtrasPoliciesPage 
              policies={autosPolicies.map(p => ({ ...p, _isAutos: true }))} 
              onSave={saveAutosPolicy} 
              onDelete={deleteAutosPolicy} 
              onMarkPaid={markAutosPaid} 
              onWhatsApp={p => setModalContact({ policy: p, type: 'whatsapp' })}
              onEmail={p => setModalContact({ policy: p, type: 'email' })}
              toast={toast}
            />
          )}
          {page === 'vida_policies' && (
            <SectionPoliciesPage 
              title="Vida" 
              icon="💚" 
              policies={vidaPolicies.map(p => ({ ...p, _isVida: true }))} 
              onSave={saveVidaPolicy} 
              onDelete={deleteVidaPolicy} 
              onMarkPaid={markVidaPaid} 
              onWhatsApp={p => setModalContact({ policy: p, type: 'whatsapp' })}
              onEmail={p => setModalContact({ policy: p, type: 'email' })}
              toast={toast}
              isVida={true}
            />
          )}
          {page === 'danos_policies' && (
            <SectionPoliciesPage 
              title="Daños" 
              icon="🏢" 
              policies={danosPolicies.map(p => ({ ...p, _isDanos: true }))} 
              onSave={saveDanosPolicy} 
              onDelete={deleteDanosPolicy} 
              onMarkPaid={markDanosPaid} 
              onWhatsApp={p => setModalContact({ policy: p, type: 'whatsapp' })}
              onEmail={p => setModalContact({ policy: p, type: 'email' })}
              toast={toast}
              isDanos={true}
            />
          )}
          {page === 'hogar_policies' && (
            <SectionPoliciesPage 
              title="Hogar" 
              icon="🏠" 
              policies={hogarPolicies.map(p => ({ ...p, _isHogar: true }))} 
              onSave={saveHogarPolicy} 
              onDelete={deleteHogarPolicy} 
              onMarkPaid={markHogarPaid} 
              onWhatsApp={p => setModalContact({ policy: p, type: 'whatsapp' })}
              onEmail={p => setModalContact({ policy: p, type: 'email' })}
              toast={toast}
              isHogar={true}
            />
          )}
          {page === 'siniestros' && (
            <SiniestrosPage siniestros={siniestros} onImport={importSiniestros} onUpdateEstatus={updateSiniestroEstatus} />
          )}
          {page === 'cotizaciones' && (
            <CotizacionesPage cotizaciones={cotizaciones} onSave={saveCotizacion} onUpdateEstatus={updateCotizacionEstatus} />
          )}
          {page === 'templates' && (
            <TemplatesPage templates={templates} onSave={setTemplates} toast={toast} />
          )}
          {page === 'comprobantes' && (
            <ComprobantesPage policies={allPolicies} onUpdatePolicy={allProps.onUpdatePolicy} />
          )}
          {page === 'import' && (
            <ImportExportPage 
              policies={policies}
              caroPolicies={caroPolicies}
              gmmPolicies={gmmPolicies}
              autosPolicies={autosPolicies}
              vidaPolicies={vidaPolicies}
              danosPolicies={danosPolicies}
              hogarPolicies={hogarPolicies}
              onImport={importToCategory}
              toast={toast} 
            />
          )}
        </main>
      </div>

      {/* Modales */}
      {modalNew && (
        <PolicyModal policy={null} onSave={savePolicy} onClose={() => setModalNew(false)} toast={toast} />
      )}
      {modalEdit && (
        <PolicyModal 
          policy={modalEdit} 
          isGmm={!!modalEdit._isGmm}
          isAutos={!!modalEdit._isAutos}
          onSave={
            modalEdit._isCaro ? saveCaroPolicy :
            modalEdit._isGmm ? saveGmmPolicy :
            modalEdit._isAutos ? saveAutosPolicy :
            savePolicy
          } 
          onClose={() => setModalEdit(null)} 
          toast={toast} 
        />
      )}
      {modalPaid && (
        <MarkPaidModal 
          policy={modalPaid} 
          onConfirm={
            modalPaid._isCaro ? markCaroPaid :
            modalPaid._isGmm ? markGmmPaid :
            modalPaid._isAutos ? markAutosPaid :
            markPaid
          } 
          onClose={() => setModalPaid(null)} 
          toast={toast} 
        />
      )}
      {modalContact && (
        <ContactModal
          policy={modalContact.policy}
          type={modalContact.type}
          templates={templates}
          onClose={() => setModalContact(null)}
        />
      )}
      {showCalendarPicker && (
        <CustomCalendarPickerModal
          policies={page === 'caro_policies' ? [] : page === 'gmm_policies' ? [] : policies}
          caroPolicies={page === 'caro_policies' ? caroPolicies : page === 'gmm_policies' ? gmmPolicies : []}
          onClose={() => setShowCalendarPicker(false)}
          onSelectDate={(dateStr) => setDailyModalDate(dateStr)}
        />
      )}
      {dailyModalDate && (
        <DailyPaymentsModal 
          dateStr={dailyModalDate} 
          policies={page === 'caro_policies' ? [] : page === 'gmm_policies' ? [] : policies} 
          caroPolicies={page === 'caro_policies' ? caroPolicies : page === 'gmm_policies' ? gmmPolicies : []} 
          onClose={() => setDailyModalDate(null)} 
          onEdit={(p) => setModalEdit(p)}
          onDelete={(p) => setDeleteConfirm(p)}
          onMarkPaid={(p) => setModalPaid(p)}
          onWhatsApp={(p) => setModalContact({ policy: p, type: 'whatsapp' })}
          onEmail={(p) => setModalContact({ policy: p, type: 'email' })}
        />
      )}

      {/* Confirmación de eliminación */}
      {deleteConfirm && (
        <div className="modal-overlay" onClick={() => setDeleteConfirm(null)}>
          <div className="modal" style={{maxWidth:420}} onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>🗑️ Confirmar Eliminación</h2>
              <button className="modal-close" onClick={() => setDeleteConfirm(null)}><Icons.Close /></button>
            </div>
            <div className="modal-body">
              <p style={{fontSize:14, color:'var(--text-secondary)', lineHeight:1.7}}>
                ¿Estás seguro de que deseas eliminar esta póliza? Esta acción <strong>no se puede deshacer</strong>.
              </p>
            </div>
            <div className="modal-footer">
              <button className="btn btn-outline" onClick={() => setDeleteConfirm(null)}>Cancelar</button>
              <button className="btn btn-danger" onClick={() => {
                if (deleteConfirm._isCaro) deleteCaroPolicy(deleteConfirm.id);
                else if (deleteConfirm._isGmm) deleteGmmPolicy(deleteConfirm.id);
                else if (deleteConfirm._isAutos) deleteAutosPolicy(deleteConfirm.id);
                else if (deleteConfirm._isVida) deleteVidaPolicy(deleteConfirm.id);
                else if (deleteConfirm._isDanos) deleteDanosPolicy(deleteConfirm.id);
                else if (deleteConfirm._isHogar) deleteHogarPolicy(deleteConfirm.id);
                else deletePolicy(deleteConfirm.id);
                setDeleteConfirm(null);
              }}>
                🗑️ Eliminar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Toasts */}
      <ToastContainer toasts={toasts} />
    </div>
  );
}

// ─── Mount ────────────────────────────────────────────────────
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);



