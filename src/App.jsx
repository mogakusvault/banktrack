import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'

/* ─── AI Prompt ─── */
const SYSTEM_PROMPT = `Eres un experto en análisis de comprobantes bancarios militares y civiles. Analiza la imagen y extrae los datos.
Responde SOLO con JSON válido (sin markdown, sin backticks):
{
  "nombre": "nombre de pila del remitente/ordenante (quien envía el dinero), o null",
  "apellido": "apellido(s) del remitente/ordenante, o null",
  "fecha": "DD/MM/AAAA",
  "descripcion": "descripción del movimiento",
  "tipo": "ingreso" o "egreso",
  "monto": número (solo valor numérico, sin símbolos),
  "moneda": "ARS" o "MXN" o "USD" o la detectada,
  "banco": "nombre del banco o billetera (Mercado Pago, etc.)",
  "referencia": "número de referencia o CVU o null",
  "cuenta": "últimos 4 dígitos o null",
  "categoria": "Aporte fondo|Transferencia|Pago|Retiro|Depósito|Compra|Otro",
  "notas": "notas relevantes o null"
}
El nombre y apellido son de QUIEN ENVÍA (remitente/ordenante), no del destinatario.
Si no es un comprobante bancario: {"error":"No es un comprobante válido"}`

/* ─── Helpers ─── */
const fmt = (n, cur = 'ARS') => {
  const validCur = ['ARS','MXN','USD','EUR'].includes(cur) ? cur : 'ARS'
  return new Intl.NumberFormat('es-AR', {
    style: 'currency', currency: validCur, minimumFractionDigits: 2,
  }).format(n || 0)
}

const CAT_META = {
  'Aporte fondo': { icon: '🪖', color: '#34d399' },
  Transferencia:  { icon: '↔',  color: '#60a5fa' },
  Pago:           { icon: '📋', color: '#a78bfa' },
  Retiro:         { icon: '🏧', color: '#f87171' },
  Depósito:       { icon: '💵', color: '#34d399' },
  Compra:         { icon: '🛍', color: '#fbbf24' },
  Otro:           { icon: '•',  color: '#94a3b8' },
}
const catMeta = (c) => CAT_META[c] || CAT_META['Otro']

const fileToB64 = (f) =>
  new Promise((res, rej) => {
    const r = new FileReader()
    r.onload = () => res(r.result.split(',')[1])
    r.onerror = () => rej(new Error('Error lectura'))
    r.readAsDataURL(f)
  })

const Icon = ({ d, size = 22, color = 'currentColor', strokeWidth = 1.8 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
    stroke={color} strokeWidth={strokeWidth} strokeLinecap="round" strokeLinejoin="round">
    <path d={d} />
  </svg>
)
const IC = {
  home:   'M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z M9 22V12h6v10',
  add:    'M12 5v14M5 12h14',
  chart:  'M18 20V10M12 20V4M6 20v-6',
  list:   'M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01',
  export: 'M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3',
  trash:  'M3 6h18M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6M10 11v6M14 11v6M9 6V4h6v2',
  close:  'M18 6L6 18M6 6l12 12',
  cam:    'M23 19a2 2 0 01-2 2H3a2 2 0 01-2-2V8a2 2 0 012-2h4l2-3h6l2 3h4a2 2 0 012 2z M12 17a4 4 0 100-8 4 4 0 000 8z',
  search: 'M21 21l-4.35-4.35M17 11A6 6 0 115 11a6 6 0 0112 0z',
  edit:   'M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z',
  user:   'M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2M12 11a4 4 0 100-8 4 4 0 000 8z',
  check:  'M20 6L9 17l-5-5',
}

/* ════════ INPUT FIELD ════════ */
function Field({ label, value, onChange, placeholder, required }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <p style={{ fontSize: 11, color: '#64748b', letterSpacing: 1, marginBottom: 6 }}>
        {label.toUpperCase()} {required && <span style={{ color: '#f87171' }}>*</span>}
      </p>
      <input
        value={value || ''}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        style={{
          width: '100%', background: '#0f172a', border: '1px solid #334155',
          borderRadius: 12, padding: '12px 14px', fontSize: 14,
          color: '#e2e8f0', outline: 'none', fontFamily: 'inherit',
        }}
      />
    </div>
  )
}

/* ════════ TX ROW ════════ */
function TxRow({ tx, onTap, showDelete, onDelete }) {
  const m = catMeta(tx.categoria)
  const fullName = [tx.nombre, tx.apellido].filter(Boolean).join(' ') || '—'
  return (
    <div onClick={() => onTap?.(tx)} style={{
      display: 'flex', alignItems: 'center', gap: 14,
      padding: '13px 16px', borderBottom: '1px solid #1e293b',
      cursor: onTap ? 'pointer' : 'default',
      WebkitTapHighlightColor: 'transparent',
    }}>
      {/* Avatar con iniciales */}
      <div style={{
        width: 44, height: 44, borderRadius: 14, flexShrink: 0,
        background: m.color + '20',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        fontSize: tx.nombre ? 14 : 20, fontWeight: 800,
        color: m.color,
      }}>
        {tx.nombre
          ? (tx.nombre[0] + (tx.apellido?.[0] || '')).toUpperCase()
          : m.icon}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <p style={{
          fontSize: 14, fontWeight: 700, color: '#e2e8f0',
          overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
          marginBottom: 2,
        }}>
          {fullName}
        </p>
        <p style={{ fontSize: 12, color: '#64748b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {tx.fecha || '—'} · {tx.banco || tx.descripcion || '—'}
        </p>
      </div>
      <div style={{ textAlign: 'right', flexShrink: 0, display: 'flex', alignItems: 'center', gap: 8 }}>
        <div>
          <p style={{
            fontSize: 15, fontWeight: 800,
            color: tx.tipo === 'ingreso' ? '#34d399' : '#f87171',
          }}>
            {tx.tipo === 'egreso' ? '−' : '+'}{fmt(tx.monto, tx.moneda)}
          </p>
          <p style={{ fontSize: 10, color: '#475569', textAlign: 'right' }}>{tx.moneda || 'ARS'}</p>
        </div>
        {showDelete && (
          <button onClick={(e) => { e.stopPropagation(); onDelete(tx.id) }} style={{
            background: '#f8717112', border: '1px solid #f8717130',
            borderRadius: 10, padding: '6px 8px', cursor: 'pointer',
            display: 'flex', alignItems: 'center',
            WebkitTapHighlightColor: 'transparent',
          }}>
            <Icon d={IC.trash} size={14} color="#f87171" />
          </button>
        )}
      </div>
    </div>
  )
}

/* ════════ HOME SCREEN ════════ */
function HomeScreen({ transactions, onAddTap }) {
  const total = transactions.reduce(
    (a, t) => ({
      ing: a.ing + (t.tipo === 'ingreso' ? t.monto || 0 : 0),
      egr: a.egr + (t.tipo === 'egreso'  ? t.monto || 0 : 0),
    }),
    { ing: 0, egr: 0 }
  )
  const recent = [...transactions].reverse().slice(0, 6)
  const meta = 153 // aspirantes
  const pagaron = new Set(transactions.filter(t => t.tipo === 'ingreso').map(t => `${t.nombre}${t.apellido}`).filter(Boolean)).size
  const pct = Math.min(100, Math.round((pagaron / meta) * 100))

  return (
    <div style={{ height: '100%', overflowY: 'auto' }}>
      {/* Hero */}
      <div style={{
        margin: '16px 16px 0', borderRadius: 28,
        background: 'linear-gradient(135deg, #0a1628 0%, #0f2044 50%, #0d1b38 100%)',
        padding: '24px 22px',
        position: 'relative', overflow: 'hidden',
        boxShadow: '0 20px 60px #0008',
        border: '1px solid #1e3a5f',
      }}>
        <div style={{ position: 'absolute', top: -40, right: -40, width: 150, height: 150,
          borderRadius: '50%', background: 'radial-gradient(#3b82f633,transparent 70%)' }} />

        {/* Escudo / badge FAA */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 16 }}>
          <div style={{
            width: 36, height: 36, borderRadius: 10,
            background: 'linear-gradient(135deg, #1d4ed8, #1e40af)',
            display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20,
          }}>✈️</div>
          <div>
            <p style={{ fontSize: 12, fontWeight: 700, color: '#93c5fd', letterSpacing: 0.5 }}>
              FONDO DE PROMOCIÓN
            </p>
            <p style={{ fontSize: 10, color: '#475569' }}>Fuerza Aérea Argentina</p>
          </div>
        </div>

        <p style={{ fontSize: 11, color: '#64748b', letterSpacing: 2, marginBottom: 4 }}>TOTAL RECAUDADO</p>
        <p style={{
          fontSize: 36, fontWeight: 800, letterSpacing: -1,
          fontFamily: 'Sora, sans-serif', color: '#e2e8f0', marginBottom: 20,
        }}>
          {fmt(total.ing)}
        </p>

        {/* Progreso aspirantes */}
        <div style={{ marginBottom: 4 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8 }}>
            <span style={{ fontSize: 11, color: '#64748b' }}>Aspirantes que aportaron</span>
            <span style={{ fontSize: 12, fontWeight: 700, color: '#93c5fd' }}>
              {pagaron} / {meta}
            </span>
          </div>
          <div style={{ background: '#0f172a', borderRadius: 6, height: 8, overflow: 'hidden' }}>
            <div style={{
              width: pct + '%', height: '100%', borderRadius: 6,
              background: 'linear-gradient(90deg, #3b82f6, #60a5fa)',
              transition: 'width 1s ease',
            }} />
          </div>
          <p style={{ fontSize: 10, color: '#475569', marginTop: 4, textAlign: 'right' }}>
            {pct}% completado · {meta - pagaron} pendientes
          </p>
        </div>
      </div>

      {/* Stats row */}
      <div style={{ display: 'flex', gap: 10, margin: '12px 16px 0' }}>
        {[
          { label: 'Aportes', val: transactions.filter(t=>t.tipo==='ingreso').length, color: '#34d399', icon: '↑' },
          { label: 'Total $', val: fmt(total.ing), color: '#60a5fa', icon: '💰' },
        ].map((s) => (
          <div key={s.label} style={{
            flex: 1, background: '#1e293b', borderRadius: 16, padding: '14px',
          }}>
            <p style={{ fontSize: 10, color: '#64748b', marginBottom: 4 }}>{s.label}</p>
            <p style={{ fontSize: 16, fontWeight: 800, color: s.color }}>{s.val}</p>
          </div>
        ))}
      </div>

      {/* Add button */}
      <button onClick={onAddTap} style={{
        margin: '12px 16px 0', width: 'calc(100% - 32px)',
        background: 'linear-gradient(135deg, #1d4ed8, #2563eb)',
        border: 'none', borderRadius: 18, padding: '16px',
        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
        cursor: 'pointer', color: '#fff', fontSize: 15, fontWeight: 700,
        boxShadow: '0 8px 24px #1d4ed844',
        WebkitTapHighlightColor: 'transparent',
      }}>
        <Icon d={IC.cam} size={20} />
        Registrar comprobante
      </button>

      <div style={{ padding: '20px 16px 8px' }}>
        <p style={{ fontSize: 12, fontWeight: 700, color: '#64748b', letterSpacing: 1 }}>ÚLTIMOS APORTES</p>
      </div>

      {recent.length === 0 ? (
        <div style={{ textAlign: 'center', padding: '32px 32px', color: '#334155' }}>
          <div style={{ fontSize: 44, marginBottom: 10 }}>✈️</div>
          <p style={{ fontWeight: 600, color: '#475569', fontSize: 15 }}>Sin aportes registrados</p>
          <p style={{ fontSize: 13, color: '#334155', marginTop: 6 }}>
            Subí el primer comprobante para comenzar
          </p>
        </div>
      ) : (
        <div style={{ paddingBottom: 16 }}>
          {recent.map((tx) => <TxRow key={tx.id} tx={tx} />)}
        </div>
      )}
    </div>
  )
}

/* ════════ ADD SCREEN ════════ */
function AddScreen({ onDone, onToast }) {
  const [step, setStep] = useState('upload') // upload | confirm | manual
  const [loading, setLoading] = useState(false)
  const [preview, setPreview] = useState(null)
  const [result, setResult] = useState(null)
  const [err, setErr] = useState(null)
  const [form, setForm] = useState({
    nombre: '', apellido: '', monto: '', fecha: '',
    banco: '', referencia: '', moneda: 'ARS',
    categoria: 'Aporte fondo', tipo: 'ingreso', notas: '',
  })
  const fileRef = useRef()

  const analyze = async (file) => {
    setLoading(true); setErr(null)
    setPreview(URL.createObjectURL(file))
    try {
      const b64 = await fileToB64(file)
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1000,
          system: SYSTEM_PROMPT,
          messages: [{
            role: 'user',
            content: [
              { type: 'image', source: { type: 'base64', media_type: file.type || 'image/png', data: b64 } },
              { type: 'text', text: 'Analiza este comprobante de aporte al fondo de promoción.' },
            ],
          }],
        }),
      })
      const data = await res.json()
      const text = data.content?.map((b) => b.text || '').join('') || ''
      const parsed = JSON.parse(text.replace(/```json|```/g, '').trim())
      if (parsed.error) { setErr(parsed.error); setStep('upload') }
      else {
        setResult(parsed)
        setForm(f => ({
          ...f,
          nombre:    parsed.nombre    || '',
          apellido:  parsed.apellido  || '',
          monto:     parsed.monto     || '',
          fecha:     parsed.fecha     || '',
          banco:     parsed.banco     || '',
          referencia: parsed.referencia || '',
          moneda:    parsed.moneda    || 'ARS',
          categoria: parsed.categoria || 'Aporte fondo',
          tipo:      parsed.tipo      || 'ingreso',
          notas:     parsed.notas     || '',
          descripcion: parsed.descripcion || '',
        }))
        setStep('confirm')
      }
    } catch (e) {
      setErr('No se pudo analizar: ' + e.message)
      setStep('upload')
    }
    setLoading(false)
  }

  const handleFile = (files) => {
    const f = Array.from(files).find((f) => f.type.startsWith('image/'))
    if (f) analyze(f)
    else setErr('Sube una imagen válida (PNG, JPG, WEBP).')
  }

  const reset = () => {
    setStep('upload'); setPreview(null); setResult(null); setErr(null)
    setForm({ nombre:'', apellido:'', monto:'', fecha:'', banco:'', referencia:'', moneda:'ARS', categoria:'Aporte fondo', tipo:'ingreso', notas:'' })
  }

  const save = () => {
    if (!form.apellido && !form.nombre) { setErr('Ingresá al menos el nombre o apellido.'); return }
    if (!form.monto) { setErr('El monto es obligatorio.'); return }
    onDone({
      ...form,
      monto: parseFloat(String(form.monto).replace(',', '.')),
      id: Date.now(),
      imagen: preview,
    })
    onToast('✓ Aporte registrado')
    reset()
  }

  /* ── UPLOAD STEP ── */
  if (step === 'upload') return (
    <div style={{ height: '100%', overflowY: 'auto', padding: '8px 16px 32px' }}>
      <div
        onClick={() => !loading && fileRef.current?.click()}
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => { e.preventDefault(); handleFile(e.dataTransfer.files) }}
        style={{
          border: `2px dashed #1d4ed866`,
          borderRadius: 24, minHeight: 220,
          display: 'flex', flexDirection: 'column',
          alignItems: 'center', justifyContent: 'center',
          background: '#080f1e', cursor: loading ? 'default' : 'pointer',
          position: 'relative', marginBottom: 16,
          WebkitTapHighlightColor: 'transparent',
        }}
      >
        {loading ? (
          <>
            <div style={{
              width: 48, height: 48, border: '3px solid #3b82f6',
              borderTopColor: 'transparent', borderRadius: '50%',
              animation: 'spin 0.8s linear infinite', marginBottom: 14,
            }} />
            <p style={{ fontSize: 14, color: '#94a3b8' }}>Analizando con IA…</p>
            <p style={{ fontSize: 12, color: '#475569', marginTop: 4 }}>Extrayendo nombre y monto</p>
          </>
        ) : (
          <>
            <div style={{
              width: 72, height: 72, borderRadius: 22,
              background: '#1d4ed818',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              marginBottom: 14, fontSize: 32,
            }}>
              📎
            </div>
            <p style={{ fontSize: 15, fontWeight: 700, color: '#94a3b8', marginBottom: 6 }}>
              Subir comprobante
            </p>
            <p style={{ fontSize: 12, color: '#475569', textAlign: 'center', lineHeight: 1.6 }}>
              La IA detecta automáticamente<br />nombre, apellido y monto
            </p>
            <p style={{ fontSize: 11, color: '#334155', marginTop: 10 }}>PNG · JPG · WEBP</p>
          </>
        )}
      </div>
      <input ref={fileRef} type="file" accept="image/*" style={{ display: 'none' }}
        onChange={(e) => handleFile(e.target.files)} />

      {err && (
        <div style={{
          background: '#f8717118', border: '1px solid #f8717140',
          borderRadius: 14, padding: '12px 16px', marginBottom: 14,
          fontSize: 13, color: '#fca5a5', lineHeight: 1.5,
        }}>⚠ {err}</div>
      )}

      {/* Carga manual */}
      <div style={{ textAlign: 'center', margin: '8px 0 4px' }}>
        <p style={{ fontSize: 12, color: '#475569' }}>¿Sin comprobante?</p>
      </div>
      <button onClick={() => setStep('manual')} style={{
        width: '100%', background: '#1e293b', border: '1px solid #334155',
        borderRadius: 16, padding: '14px', color: '#94a3b8',
        fontSize: 14, fontWeight: 600, cursor: 'pointer',
        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
        WebkitTapHighlightColor: 'transparent',
      }}>
        <Icon d={IC.edit} size={16} color="#94a3b8" />
        Ingresar datos manualmente
      </button>
    </div>
  )

  /* ── CONFIRM STEP ── */
  if (step === 'confirm') return (
    <div style={{ height: '100%', overflowY: 'auto', padding: '8px 16px 32px' }}>
      {preview && (
        <div style={{ marginBottom: 16 }}>
          <img src={preview} alt="comprobante" style={{
            width: '100%', maxHeight: 200, objectFit: 'contain',
            borderRadius: 16, border: '1px solid #1e293b', background: '#080f1e',
          }} />
        </div>
      )}

      <div style={{
        background: '#1e293b', borderRadius: 20,
        overflow: 'hidden', marginBottom: 16,
        border: '1px solid #334155',
      }}>
        <div style={{
          padding: '14px 18px',
          background: 'linear-gradient(135deg, #0a1628, #0f2044)',
          borderBottom: '1px solid #334155',
          display: 'flex', alignItems: 'center', gap: 8,
        }}>
          <Icon d={IC.check} size={16} color="#34d399" />
          <p style={{ fontSize: 13, fontWeight: 700, color: '#34d399' }}>
            Datos detectados — verificá y editá si es necesario
          </p>
        </div>
        <div style={{ padding: '16px' }}>
          <div style={{ display: 'flex', gap: 10, marginBottom: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Nombre" value={form.nombre}
                onChange={(v) => setForm(f => ({ ...f, nombre: v }))}
                placeholder="Ej: Juan" required />
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Apellido" value={form.apellido}
                onChange={(v) => setForm(f => ({ ...f, apellido: v }))}
                placeholder="Ej: Pérez" required />
            </div>
          </div>
          <Field label="Monto aportado" value={form.monto}
            onChange={(v) => setForm(f => ({ ...f, monto: v }))}
            placeholder="Ej: 15000" required />
          <div style={{ display: 'flex', gap: 10 }}>
            <div style={{ flex: 1 }}>
              <Field label="Fecha" value={form.fecha}
                onChange={(v) => setForm(f => ({ ...f, fecha: v }))}
                placeholder="DD/MM/AAAA" />
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Moneda" value={form.moneda}
                onChange={(v) => setForm(f => ({ ...f, moneda: v }))}
                placeholder="ARS" />
            </div>
          </div>
          <Field label="Banco / Billetera" value={form.banco}
            onChange={(v) => setForm(f => ({ ...f, banco: v }))}
            placeholder="Ej: Mercado Pago, Banco Nación" />
          <Field label="Referencia / CVU" value={form.referencia}
            onChange={(v) => setForm(f => ({ ...f, referencia: v }))}
            placeholder="Número de operación (opcional)" />
          <Field label="Notas" value={form.notas}
            onChange={(v) => setForm(f => ({ ...f, notas: v }))}
            placeholder="Observaciones (opcional)" />
        </div>
      </div>

      {err && (
        <div style={{
          background: '#f8717118', border: '1px solid #f8717140',
          borderRadius: 14, padding: '12px 16px', marginBottom: 14,
          fontSize: 13, color: '#fca5a5',
        }}>⚠ {err}</div>
      )}

      <button onClick={save} style={{
        width: '100%',
        background: 'linear-gradient(135deg, #1d4ed8, #2563eb)',
        border: 'none', borderRadius: 18, padding: '18px',
        color: '#fff', fontSize: 16, fontWeight: 700,
        cursor: 'pointer', marginBottom: 10,
        WebkitTapHighlightColor: 'transparent',
        boxShadow: '0 8px 24px #1d4ed844',
      }}>
        Guardar aporte ✓
      </button>
      <button onClick={reset} style={{
        width: '100%', background: 'transparent', border: '1px solid #334155',
        borderRadius: 16, padding: '14px', color: '#64748b',
        fontSize: 14, cursor: 'pointer', WebkitTapHighlightColor: 'transparent',
      }}>
        Cancelar
      </button>
    </div>
  )

  /* ── MANUAL STEP ── */
  return (
    <div style={{ height: '100%', overflowY: 'auto', padding: '8px 16px 32px' }}>
      <div style={{ background: '#1e293b', borderRadius: 20, padding: '16px', marginBottom: 16 }}>
        <p style={{ fontSize: 13, fontWeight: 700, color: '#94a3b8', marginBottom: 14, display: 'flex', alignItems: 'center', gap: 8 }}>
          <Icon d={IC.user} size={15} color="#94a3b8" /> Datos del aspirante
        </p>
        <div style={{ display: 'flex', gap: 10 }}>
          <div style={{ flex: 1 }}>
            <Field label="Nombre" value={form.nombre}
              onChange={(v) => setForm(f => ({ ...f, nombre: v }))}
              placeholder="Ej: Juan" required />
          </div>
          <div style={{ flex: 1 }}>
            <Field label="Apellido" value={form.apellido}
              onChange={(v) => setForm(f => ({ ...f, apellido: v }))}
              placeholder="Ej: Pérez" required />
          </div>
        </div>
        <Field label="Monto aportado" value={form.monto}
          onChange={(v) => setForm(f => ({ ...f, monto: v }))}
          placeholder="Ej: 15000" required />
        <div style={{ display: 'flex', gap: 10 }}>
          <div style={{ flex: 1 }}>
            <Field label="Fecha" value={form.fecha}
              onChange={(v) => setForm(f => ({ ...f, fecha: v }))}
              placeholder="DD/MM/AAAA" />
          </div>
          <div style={{ flex: 1 }}>
            <Field label="Moneda" value={form.moneda}
              onChange={(v) => setForm(f => ({ ...f, moneda: v }))}
              placeholder="ARS" />
          </div>
        </div>
        <Field label="Banco / Billetera" value={form.banco}
          onChange={(v) => setForm(f => ({ ...f, banco: v }))}
          placeholder="Ej: Mercado Pago" />
        <Field label="Referencia" value={form.referencia}
          onChange={(v) => setForm(f => ({ ...f, referencia: v }))}
          placeholder="Número de operación (opcional)" />
        <Field label="Notas" value={form.notas}
          onChange={(v) => setForm(f => ({ ...f, notas: v }))}
          placeholder="Observaciones (opcional)" />
      </div>

      {err && (
        <div style={{
          background: '#f8717118', border: '1px solid #f8717140',
          borderRadius: 14, padding: '12px 16px', marginBottom: 14,
          fontSize: 13, color: '#fca5a5',
        }}>⚠ {err}</div>
      )}

      <button onClick={save} style={{
        width: '100%',
        background: 'linear-gradient(135deg, #1d4ed8, #2563eb)',
        border: 'none', borderRadius: 18, padding: '18px',
        color: '#fff', fontSize: 16, fontWeight: 700,
        cursor: 'pointer', marginBottom: 10,
        WebkitTapHighlightColor: 'transparent',
        boxShadow: '0 8px 24px #1d4ed844',
      }}>
        Guardar aporte ✓
      </button>
      <button onClick={() => setStep('upload')} style={{
        width: '100%', background: 'transparent', border: '1px solid #334155',
        borderRadius: 16, padding: '14px', color: '#64748b',
        fontSize: 14, cursor: 'pointer', WebkitTapHighlightColor: 'transparent',
      }}>
        ← Volver
      </button>
    </div>
  )
}

/* ════════ MOVEMENTS SCREEN ════════ */
function MovementsScreen({ transactions, onSelect, onDelete }) {
  const [search, setSearch] = useState('')

  const list = [...transactions].reverse().filter((t) =>
    !search ||
    (t.nombre  || '').toLowerCase().includes(search.toLowerCase()) ||
    (t.apellido|| '').toLowerCase().includes(search.toLowerCase()) ||
    (t.banco   || '').toLowerCase().includes(search.toLowerCase())
  )

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{ padding: '12px 16px 8px', flexShrink: 0 }}>
        <div style={{
          display: 'flex', alignItems: 'center', gap: 10,
          background: '#1e293b', borderRadius: 14, padding: '10px 14px',
        }}>
          <Icon d={IC.search} size={16} color="#64748b" />
          <input value={search} onChange={(e) => setSearch(e.target.value)}
            placeholder="Buscar por nombre o apellido..."
            style={{
              flex: 1, background: 'transparent', border: 'none',
              outline: 'none', color: '#e2e8f0', fontSize: 14, fontFamily: 'inherit',
            }}
          />
        </div>
        <p style={{ fontSize: 11, color: '#475569', marginTop: 8, textAlign: 'right' }}>
          {list.length} registro{list.length !== 1 ? 's' : ''}
        </p>
      </div>

      <div style={{ flex: 1, overflowY: 'auto' }}>
        {list.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '60px 32px', color: '#334155' }}>
            <div style={{ fontSize: 44, marginBottom: 12 }}>🔍</div>
            <p style={{ color: '#475569', fontWeight: 600 }}>Sin resultados</p>
          </div>
        ) : (
          list.map((tx) => (
            <TxRow key={tx.id} tx={tx} onTap={onSelect} showDelete onDelete={onDelete} />
          ))
        )}
      </div>
    </div>
  )
}

/* ════════ STATS SCREEN ════════ */
function StatsScreen({ transactions, onExport }) {
  const aportes = transactions.filter(t => t.tipo === 'ingreso')
  const totalIng = aportes.reduce((s, t) => s + (t.monto || 0), 0)
  const meta = 153

  // Quiénes pagaron (por nombre único)
  const pagaron = [...new Map(
    aportes.filter(t => t.apellido || t.nombre)
      .map(t => [`${t.nombre}|${t.apellido}`, t])
  ).values()].sort((a,b) => (a.apellido||'').localeCompare(b.apellido||''))

  return (
    <div style={{ height: '100%', overflowY: 'auto' }}>
      <div style={{ padding: '8px 16px 12px', display: 'flex', flexDirection: 'column', gap: 10 }}>
        {[
          { label: 'Total recaudado',       val: fmt(totalIng),       color: '#34d399' },
          { label: 'Aspirantes que pagaron', val: `${pagaron.length} / ${meta}`, color: '#60a5fa' },
          { label: 'Pendientes',            val: `${meta - pagaron.length}`, color: '#f87171' },
          { label: 'Total aportes',         val: `${aportes.length} registros`, color: '#a78bfa' },
        ].map((s) => (
          <div key={s.label} style={{
            background: '#1e293b', borderRadius: 16, padding: '14px 18px',
            display: 'flex', justifyContent: 'space-between', alignItems: 'center',
          }}>
            <span style={{ fontSize: 13, color: '#64748b' }}>{s.label}</span>
            <span style={{ fontSize: 17, fontWeight: 800, color: s.color, fontFamily: 'Sora, sans-serif' }}>
              {s.val}
            </span>
          </div>
        ))}
      </div>

      {pagaron.length > 0 && (
        <>
          <div style={{ padding: '8px 16px 10px' }}>
            <p style={{ fontSize: 11, color: '#64748b', letterSpacing: 1.5 }}>LISTA DE APORTES</p>
          </div>
          {pagaron.map((t) => (
            <div key={t.id} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              padding: '11px 16px', borderBottom: '1px solid #1e293b',
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <div style={{
                  width: 34, height: 34, borderRadius: 10,
                  background: '#34d39918',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontSize: 12, fontWeight: 800, color: '#34d399',
                }}>
                  {((t.nombre?.[0] || '') + (t.apellido?.[0] || '')).toUpperCase() || '?'}
                </div>
                <div>
                  <p style={{ fontSize: 13, fontWeight: 600, color: '#e2e8f0' }}>
                    {[t.apellido, t.nombre].filter(Boolean).join(', ')}
                  </p>
                  <p style={{ fontSize: 11, color: '#475569' }}>{t.fecha || '—'}</p>
                </div>
              </div>
              <p style={{ fontSize: 14, fontWeight: 700, color: '#34d399' }}>
                {fmt(t.monto, t.moneda)}
              </p>
            </div>
          ))}
        </>
      )}

      <div style={{ padding: '20px 16px 32px' }}>
        <button onClick={onExport} disabled={!transactions.length} style={{
          width: '100%',
          background: transactions.length
            ? 'linear-gradient(135deg, #059669, #047857)' : '#1e293b',
          border: 'none', borderRadius: 18, padding: '18px',
          color: transactions.length ? '#fff' : '#475569',
          fontSize: 15, fontWeight: 700,
          cursor: transactions.length ? 'pointer' : 'not-allowed',
          display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10,
          WebkitTapHighlightColor: 'transparent',
        }}>
          <Icon d={IC.export} size={18} />
          Exportar Excel con todos los datos
        </button>
      </div>
    </div>
  )
}

/* ════════ DETAIL MODAL ════════ */
function DetailModal({ tx, onClose, onDelete }) {
  if (!tx) return null
  const m = catMeta(tx.categoria)
  const fullName = [tx.nombre, tx.apellido].filter(Boolean).join(' ') || '—'
  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 200,
      background: '#000000cc', backdropFilter: 'blur(8px)',
      display: 'flex', alignItems: 'flex-end',
      animation: 'fadeIn 0.2s ease',
    }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: '100%', maxHeight: '88vh',
        background: '#0f172a', borderRadius: '28px 28px 0 0',
        paddingBottom: 'max(32px, env(safe-area-inset-bottom))',
        overflowY: 'auto', animation: 'slideUp 0.3s ease',
      }}>
        <div style={{ display: 'flex', justifyContent: 'center', padding: '12px 0 4px' }}>
          <div style={{ width: 40, height: 4, borderRadius: 2, background: '#334155' }} />
        </div>

        {/* Name header */}
        <div style={{ padding: '12px 20px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{
              width: 52, height: 52, borderRadius: 16,
              background: '#34d39920',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              fontSize: 18, fontWeight: 800, color: '#34d399',
            }}>
              {((tx.nombre?.[0] || '') + (tx.apellido?.[0] || '')).toUpperCase() || '?'}
            </div>
            <div>
              <p style={{ fontSize: 17, fontWeight: 800, color: '#e2e8f0' }}>{fullName}</p>
              <p style={{ fontSize: 12, color: '#64748b' }}>{tx.fecha || '—'} · {tx.banco || '—'}</p>
            </div>
          </div>
          <button onClick={onClose} style={{
            background: '#1e293b', border: 'none', borderRadius: 12,
            padding: 8, cursor: 'pointer', WebkitTapHighlightColor: 'transparent',
          }}>
            <Icon d={IC.close} size={18} color="#64748b" />
          </button>
        </div>

        {/* Monto */}
        <div style={{ margin: '0 20px 16px', background: '#1e293b', borderRadius: 20, padding: '20px', textAlign: 'center' }}>
          <p style={{ fontSize: 11, color: '#64748b', letterSpacing: 1.5, marginBottom: 8 }}>MONTO APORTADO</p>
          <p style={{ fontSize: 40, fontWeight: 800, fontFamily: 'Sora, sans-serif', color: '#34d399' }}>
            {fmt(tx.monto, tx.moneda)}
          </p>
          <p style={{ fontSize: 12, color: '#64748b', marginTop: 4 }}>{tx.moneda || 'ARS'}</p>
        </div>

        {tx.imagen && (
          <div style={{ margin: '0 20px 16px' }}>
            <img src={tx.imagen} alt="comprobante" style={{ width: '100%', borderRadius: 16, border: '1px solid #1e293b' }} />
          </div>
        )}

        <div style={{ margin: '0 20px 20px', background: '#1e293b', borderRadius: 20, overflow: 'hidden' }}>
          {[
            { l: 'Nombre',      v: tx.nombre },
            { l: 'Apellido',    v: tx.apellido },
            { l: 'Banco',       v: tx.banco },
            { l: 'Referencia',  v: tx.referencia },
            { l: 'Descripción', v: tx.descripcion },
            { l: 'Notas',       v: tx.notas },
          ].filter((r) => r.v).map((row, i, arr) => (
            <div key={row.l} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start',
              padding: '13px 18px',
              borderBottom: i < arr.length - 1 ? '1px solid #0f172a' : 'none',
            }}>
              <span style={{ fontSize: 13, color: '#64748b', flexShrink: 0, marginRight: 16 }}>{row.l}</span>
              <span style={{ fontSize: 13, color: '#cbd5e1', fontWeight: 500, textAlign: 'right' }}>{row.v}</span>
            </div>
          ))}
        </div>

        <div style={{ padding: '0 20px' }}>
          <button onClick={() => { onDelete(tx.id); onClose() }} style={{
            width: '100%', background: '#f8717112', border: '1px solid #f8717130',
            borderRadius: 18, padding: '16px', color: '#f87171', fontSize: 15, fontWeight: 700,
            cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
            WebkitTapHighlightColor: 'transparent',
          }}>
            <Icon d={IC.trash} size={16} color="#f87171" />
            Eliminar registro
          </button>
        </div>
      </div>
    </div>
  )
}

/* ════════════════════════════════════════
   ROOT
════════════════════════════════════════ */
const TABS = [
  { id: 'home',  icon: IC.home,  label: 'Inicio'    },
  { id: 'add',   icon: IC.add,   label: 'Registrar' },
  { id: 'list',  icon: IC.list,  label: 'Lista'     },
  { id: 'stats', icon: IC.chart, label: 'Resumen'   },
]
const TITLES = {
  home:  'Fondo de Promoción ✈️',
  add:   'Registrar aporte',
  list:  'Lista de aportes',
  stats: 'Estadísticas',
}

export default function App() {
  const [transactions, setTransactions] = useState([])
  const [tab, setTab] = useState('home')
  const [selectedTx, setSelectedTx] = useState(null)
  const [toast, setToast] = useState(null)

  const showToast = (msg) => {
    setToast(msg)
    setTimeout(() => setToast(null), 2500)
  }

  const addTx = (tx) => {
    setTransactions((p) => [...p, tx])
    setTab('home')
  }

  const deleteTx = (id) => {
    setTransactions((p) => p.filter((t) => t.id !== id))
    showToast('Registro eliminado')
  }

  const exportXLSX = () => {
    if (!transactions.length) return
    const rows = transactions.map((t, i) => ({
      'N°':          i + 1,
      'Apellido':    t.apellido || '',
      'Nombre':      t.nombre   || '',
      'Monto':       t.tipo === 'egreso' ? -(t.monto || 0) : t.monto || 0,
      'Moneda':      t.moneda   || 'ARS',
      'Fecha':       t.fecha    || '',
      'Banco/Billetera': t.banco || '',
      'Referencia':  t.referencia || '',
      'Categoría':   t.categoria  || '',
      'Tipo':        (t.tipo || '').toUpperCase(),
      'Notas':       t.notas || '',
    }))

    const ws = XLSX.utils.json_to_sheet(rows)
    ws['!cols'] = [4,18,16,14,8,12,20,18,15,9,25].map((w) => ({ wch: w }))

    // Estilo de encabezado
    const headerRange = XLSX.utils.decode_range(ws['!ref'])
    for (let C = headerRange.s.c; C <= headerRange.e.c; C++) {
      const cell = XLSX.utils.encode_cell({ r: 0, c: C })
      if (ws[cell]) {
        ws[cell].s = {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { fgColor: { rgb: '1D4ED8' } },
          alignment: { horizontal: 'center' },
        }
      }
    }

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Aportes')

    // Hoja resumen
    const ing = transactions.filter(t=>t.tipo==='ingreso').reduce((s,t)=>s+(t.monto||0),0)
    const ws2 = XLSX.utils.aoa_to_sheet([
      ['FONDO DE PROMOCIÓN — RESUMEN'],
      [],
      ['Total aspirantes',       153],
      ['Aportes registrados',    transactions.filter(t=>t.tipo==='ingreso').length],
      ['Total recaudado (ARS)',   ing],
      ['Fecha exportación',       new Date().toLocaleDateString('es-AR')],
    ])
    ws2['!cols'] = [{ wch: 28 }, { wch: 18 }]
    XLSX.utils.book_append_sheet(wb, ws2, 'Resumen')

    XLSX.writeFile(wb, `fondo_promocion_FAA_${new Date().toLocaleDateString('es-AR').replace(/\//g,'-')}.xlsx`)
    showToast('📊 Excel exportado')
  }

  return (
    <div style={{
      width: '100%', height: '100%',
      display: 'flex', flexDirection: 'column',
      background: '#0f172a', color: '#e2e8f0',
      maxWidth: 500, margin: '0 auto',
      position: 'relative', overflow: 'hidden',
    }}>
      {/* Header */}
      <div style={{
        background: '#080f1e',
        padding: 'max(16px, env(safe-area-inset-top)) 20px 14px',
        flexShrink: 0, borderBottom: '1px solid #1e293b',
      }}>
        <h1 style={{
          fontSize: 20, fontWeight: 800, fontFamily: 'Sora, sans-serif',
          letterSpacing: -0.5, color: '#f1f5f9',
        }}>
          {TITLES[tab]}
        </h1>
        {tab === 'home' && (
          <p style={{ fontSize: 11, color: '#475569', marginTop: 2, letterSpacing: 0.3 }}>
            153 aspirantes · Fuerza Aérea Argentina
          </p>
        )}
      </div>

      <div style={{ flex: 1, overflowY: 'auto', overflowX: 'hidden' }}>
        {tab === 'home'  && <HomeScreen transactions={transactions} onAddTap={() => setTab('add')} />}
        {tab === 'add'   && <AddScreen onDone={addTx} onToast={showToast} />}
        {tab === 'list'  && <MovementsScreen transactions={transactions} onSelect={setSelectedTx} onDelete={deleteTx} />}
        {tab === 'stats' && <StatsScreen transactions={transactions} onExport={exportXLSX} />}
      </div>

      {/* Bottom Nav */}
      <div style={{
        background: '#080f1e', borderTop: '1px solid #1e293b',
        display: 'flex', alignItems: 'center', padding: '6px 8px',
        paddingBottom: 'max(6px, env(safe-area-inset-bottom))',
        flexShrink: 0,
      }}>
        {TABS.map((t) => {
          const active = tab === t.id
          const isAdd = t.id === 'add'
          return (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              flex: 1,
              background: isAdd
                ? active ? 'linear-gradient(135deg,#1d4ed8,#2563eb)' : '#1e293b'
                : 'transparent',
              border: 'none', borderRadius: isAdd ? 18 : 0,
              padding: isAdd ? '10px 8px' : '10px 0',
              margin: isAdd ? '0 4px' : 0,
              display: 'flex', flexDirection: 'column',
              alignItems: 'center', justifyContent: 'center', gap: 4,
              cursor: 'pointer', WebkitTapHighlightColor: 'transparent',
              transition: 'all 0.2s', minHeight: 56,
            }}>
              <Icon d={t.icon} size={22}
                color={active ? (isAdd ? '#fff' : '#60a5fa') : '#334155'}
                strokeWidth={active ? 2.2 : 1.7}
              />
              <span style={{
                fontSize: 10, fontWeight: active ? 700 : 500,
                color: active ? (isAdd ? '#fff' : '#60a5fa') : '#334155',
                letterSpacing: 0.3,
              }}>
                {t.label}
              </span>
            </button>
          )
        })}
      </div>

      {selectedTx && (
        <DetailModal tx={selectedTx} onClose={() => setSelectedTx(null)}
          onDelete={(id) => { deleteTx(id); setSelectedTx(null) }} />
      )}

      {toast && (
        <div style={{
          position: 'absolute', bottom: 90, left: '50%',
          transform: 'translateX(-50%)',
          background: '#1e293b', border: '1px solid #334155',
          borderRadius: 20, padding: '10px 20px',
          fontSize: 13, fontWeight: 600, color: '#e2e8f0',
          whiteSpace: 'nowrap', zIndex: 300,
          boxShadow: '0 8px 32px #00000088',
          animation: 'toastIn 0.3s ease',
        }}>
          {toast}
        </div>
      )}
    </div>
  )
}
