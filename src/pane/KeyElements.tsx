import { useEffect, useRef, useState } from "react";
import {
  IDEA_STATUSES, INSTRUMENTS, SIDES, TRADING_AREAS, cardTitle, isRecommendation, newLeg, underlyingLabel,
  type KeyElements, type Leg,
} from "./keyElementsModel";

const EMEA_LABEL = "Instrument or underlying mainly traded in EMEA or if FX, involves an EMEA currency";
const FACTUAL_LABEL = "This communication is a factual market comment, not an investment recommendation";

function Segmented<T extends string>({ options, value, onChange }: { options: T[]; value: T; onChange: (v: T) => void }) {
  return (
    <div className="seg" role="radiogroup">
      {options.map((o) => (
        <button key={o} type="button" role="radio" aria-checked={o === value} className={o === value ? "on" : ""} onClick={() => onChange(o)}>
          {o}
        </button>
      ))}
    </div>
  );
}

function InstrumentMenu({ label, onPick }: { label: string; onPick: (instrument: string) => void }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  return (
    <div className="menu" ref={ref}>
      <button type="button" className="menu-btn" aria-expanded={open} onClick={() => setOpen((v) => !v)}>
        <span className="plus">+</span> {label} <span className="chev">{open ? "︿" : "﹀"}</span>
      </button>
      {open && (
        <ul className="menu-list" role="listbox">
          {INSTRUMENTS.map((i) => (
            <li key={i.name}>
              <button type="button" role="option" onClick={() => { onPick(i.name); setOpen(false); }}>{i.name}</button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

function LegEditor({ leg, index, count, onChange, onRemove }: { leg: Leg; index: number; count: number; onChange: (patch: Partial<Leg>) => void; onRemove: () => void }) {
  const name = count > 1 ? `Leg ${index + 1}: ${leg.instrument}` : leg.instrument;
  return (
    <div className="leg">
      <div className="leg-head">
        <span>{name}</span>
        <button type="button" className="x" aria-label={`Remove ${name}`} onClick={onRemove}>×</button>
      </div>
      <label className="check">
        <input type="checkbox" checked={leg.emea} onChange={(e) => onChange({ emea: e.target.checked })} />
        <span>{EMEA_LABEL}</span>
      </label>
      {leg.emea && (
        <label className="check">
          <input type="checkbox" checked={leg.factual} onChange={(e) => onChange({ factual: e.target.checked })} />
          <span>{FACTUAL_LABEL}</span>
        </label>
      )}
      {isRecommendation(leg) && (
        <div className="rec">
          <Segmented options={SIDES} value={leg.side} onChange={(side) => onChange({ side })} />
          <label className="row">
            <span>Price</span>
            <input value={leg.price} placeholder="Enter Price…" onChange={(e) => onChange({ price: e.target.value })} />
          </label>
          <label className="row">
            <span>{underlyingLabel(leg.instrument)}</span>
            <input value={leg.underlying} placeholder="Details e.g. BBG" onChange={(e) => onChange({ underlying: e.target.value })} />
          </label>
          <label className="row">
            <span>Time Horizon</span>
            <input value={leg.timeHorizon} onChange={(e) => onChange({ timeHorizon: e.target.value })} />
          </label>
        </div>
      )}
    </div>
  );
}

export function KeyElementsForm({ value, onChange }: { value: KeyElements; onChange: (next: KeyElements) => void }) {
  const set = (patch: Partial<KeyElements>) => onChange({ ...value, ...patch });
  const setLeg = (id: string, patch: Partial<Leg>) => set({ legs: value.legs.map((l) => (l.id === id ? { ...l, ...patch } : l)) });
  const recommendation = value.legs.some(isRecommendation);
  return (
    <div className={`card${recommendation ? " rec-card" : ""}`}>
      <div className="card-title">
        <span>{cardTitle(value)}</span>
        {recommendation && <span className="warn" title="Investment recommendation">!</span>}
      </div>
      <Segmented options={IDEA_STATUSES} value={value.status} onChange={(status) => set({ status })} />
      <label className="field">
        <span>Headline</span>
        <textarea rows={2} value={value.headline} onChange={(e) => set({ headline: e.target.value })} />
      </label>
      <label className="row">
        <span>Trading area</span>
        <select value={value.tradingArea} onChange={(e) => set({ tradingArea: e.target.value })}>
          <option value="">Select</option>
          {TRADING_AREAS.map((a) => (
            <option key={a} value={a}>{a}</option>
          ))}
        </select>
      </label>
      {value.legs.length > 0 && <div className="section">Instruments</div>}
      {value.legs.map((leg, i) => (
        <LegEditor
          key={leg.id}
          leg={leg}
          index={i}
          count={value.legs.length}
          onChange={(patch) => setLeg(leg.id, patch)}
          onRemove={() => set({ legs: value.legs.filter((l) => l.id !== leg.id) })}
        />
      ))}
      <InstrumentMenu label={value.legs.length ? "Add Leg" : "Add instrument"} onPick={(name) => set({ legs: [...value.legs, newLeg(name)] })} />
    </div>
  );
}

export function KeyElementsView({ value }: { value: KeyElements }) {
  return (
    <div className="card view">
      <div className="view-title">{`${value.status} ${cardTitle(value)}`.toUpperCase()}</div>
      {value.headline && <div className="view-headline">{value.headline}</div>}
      <div className="kv"><span>Trading Area:</span> {value.tradingArea || "–"}</div>
      {value.legs.map((leg, i) => (
        <div key={leg.id} className="view-leg">
          <div className="view-leg-head">
            <strong>Leg {i + 1}: {leg.instrument}</strong>
            {isRecommendation(leg) && <span className="chip">{leg.side.toUpperCase()}</span>}
          </div>
          {leg.emea && <div>Instrument mainly traded in EMEA</div>}
          {leg.emea && leg.factual && <div>Factual market comment, not an investment recommendation</div>}
          {isRecommendation(leg) && (
            <>
              <div className="kv"><span>Price:</span> {leg.price || "–"}</div>
              <div className="kv"><span>{underlyingLabel(leg.instrument)}:</span> {leg.underlying || "–"}</div>
              <div className="kv"><span>Time Horizon:</span> {leg.timeHorizon || "–"}</div>
            </>
          )}
        </div>
      ))}
      {value.legs.length === 0 && <div className="muted">No instruments.</div>}
    </div>
  );
}
