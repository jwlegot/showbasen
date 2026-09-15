import React, { useState, useEffect, useRef } from 'react';
import { Search, Plus, X, Trash2, Clock, Save, Library, ListMusic, Briefcase, Printer, GripVertical, MapPin, Calendar, Music, Video, ExternalLink, Filter, Download, Film, Settings, BookOpen, Target, ChevronDown, ChevronUp } from 'lucide-react';
import * as XLSX from 'xlsx';

const DEFAULT_CATEGORIES = [
  { id: 'barnkalas', label: 'Barnkalas' },
  { id: 'mingel', label: 'Mingel' },
  { id: 'foretagsevent', label: 'Företagsevent' },
  { id: 'scenshow', label: 'Scenshow' },
];

const DEFAULT_GENRES = [
  { id: 'korttrick', label: 'Korttrick' },
  { id: 'myntmagi', label: 'Myntmagi' },
  { id: 'mentalism', label: 'Mentalism' },
  { id: 'narmagi', label: 'Närmagi' },
  { id: 'scenillusion', label: 'Scenillusion' },
  { id: 'manipulation', label: 'Manipulation' },
  { id: 'barnmagi', label: 'Barnmagi' },
  { id: 'komedimagi', label: 'Komedimagi' },
  { id: 'teori-historia', label: 'Teori & historia' },
  { id: 'dokumentar-biografi', label: 'Dokumentär & biografi' },
];


const EXPORT_FIELDS = [
  { key: 'name', label: 'Namn' },
  { key: 'date', label: 'Datum' },
  { key: 'type', label: 'Typ' },
  { key: 'location', label: 'Plats' },
  { key: 'audience', label: 'Publik' },
  { key: 'audienceCount', label: 'Antal i publik' },
  { key: 'plannedLengthMin', label: 'Planerad längd (min)' },
  { key: 'venue', label: 'Scen' },
  { key: 'plannedLength', label: 'Bedömd längd (min)' },
  { key: 'actualLength', label: 'Verklig längd' },
  { key: 'price', label: 'Pris' },
  { key: 'filmed', label: 'Filmad' },
  { key: 'rating', label: 'Betyg helhet' },
  { key: 'opening', label: 'Öppning (trick)' },
  { key: 'middle', label: 'Mitten (trick)' },
  { key: 'closing', label: 'Avslutning (trick)' },
  { key: 'comment', label: 'Kommentar' },
  { key: 'good', label: 'Bra' },
  { key: 'improve', label: 'Förbättring' },
  { key: 'customerReview', label: 'Kundrecension' },
  { key: 'customerImprovement', label: 'Kundens förbättringsförslag' },
  { key: 'formResponseUrl', label: 'Länk till formulärsvar' },
];
const allColumnsOn = () => EXPORT_FIELDS.reduce((acc, f) => ({ ...acc, [f.key]: true }), {});

const defaultSettings = () => ({
  defaultEventType: '',
  defaultVenue: 'inne',
  theme: 'light',
  exportProfiles: [{ id: 'standard', name: 'Standard (allt)', columns: allColumnsOn() }],
});

const THEMES = {
  light: {
    '--bg': '#f5f0e6', '--panel': '#faf6ee', '--card': '#ffffff', '--border': '#ddd3bf',
    '--text': '#2b2620', '--text-muted': '#7d7364', '--accent-gold': '#9c7233', '--accent-red': '#a03d2e', '--ink': '#2b2620',
  },
  dark: {
    '--bg': '#171512', '--panel': '#201d19', '--card': '#2a2521', '--border': '#3b342b',
    '--text': '#f2e9d8', '--text-muted': '#a89a83', '--accent-gold': '#b8935a', '--accent-red': '#a85a4a', '--ink': '#2b2620',
  },
};

const STATUS = {
  aktiv: { label: 'Aktiv', color: '#7a9b6e' },
  vilande: { label: 'Vilande', color: '#8a7f6a' },
  utveckling: { label: 'Under utveckling', color: '#b8935a' },
};

const PHASES = [
  { id: 'opening', label: 'Öppning', hint: 'Öppna grupper, bygg momentum' },
  { id: 'middle', label: 'Mitten', hint: 'Jobba rummet' },
  { id: 'closing', label: 'Avslutning', hint: 'Avsluta starkt' },
];

// Derived, never stored: flips automatically once the gig's date has passed.
function gigStatus(g) {
  if (!g.date) return null;
  const today = new Date().toISOString().slice(0, 10);
  return g.date < today ? 'genomford' : 'kommande';
}
const GIG_STATUS_LABEL = { kommande: 'Kommande', genomford: 'Genomförd' };

function rowVariant(row, trick) {
  return trick?.variants?.find(v => v.id === row.variantId) || null;
}
function effectiveName(row, trick) {
  if (!trick) return '(borttaget trick)';
  const v = rowVariant(row, trick);
  return v ? `${trick.name} — ${v.name || 'namnlös variant'}` : trick.name;
}
function effectiveLength(row, trick) {
  if (row.overrideLengthMin !== '' && row.overrideLengthMin != null) return Number(row.overrideLengthMin) || 0;
  return trick ? trick.lengthMin : 0;
}
function effectiveMusicOn(row, trick) {
  return (row.musicOn === null || row.musicOn === undefined) ? !!trick?.usesMusic : row.musicOn;
}
function isMusicCustomized(row, trick) {
  const onDiffers = row.musicOn !== null && row.musicOn !== undefined && row.musicOn !== !!trick?.usesMusic;
  const songChanged = row.overrideSong && row.overrideSong.trim() && row.overrideSong !== (trick?.song || '');
  const artistChanged = row.overrideArtist && row.overrideArtist.trim() && row.overrideArtist !== (trick?.artist || '');
  return onDiffers || songChanged || artistChanged;
}
function effectiveMusic(row, trick) {
  if (!effectiveMusicOn(row, trick)) return null;
  const song = (row.overrideSong && row.overrideSong.trim()) ? row.overrideSong : (trick?.song || '');
  const artist = (row.overrideArtist && row.overrideArtist.trim()) ? row.overrideArtist : (trick?.artist || '');
  if (!song) return null;
  return { song, artist };
}

const emptyGig = () => ({
  id: null,
  name: '',
  eventType: 'mingel',
  date: '',
  location: '',
  audienceType: '',
  audienceCount: '',
  plannedLengthMin: '',
  venue: 'inne',
  priceSEK: '',
  backgroundNotes: '',
  phaseTricks: { opening: [], middle: [], closing: [] },
  actualLengthMin: '',
  overallRating: '',
  overallComment: '',
  good: '',
  improve: '',
  filmed: false,
  customerReview: '',
  customerImprovement: '',
  formResponseUrl: '',
});

const emptyTrick = () => ({
  name: '', source: '', contexts: [], lengthMin: 3, status: 'aktiv', notes: '', description: '',
  usesMusic: false, song: '', artist: '', altSong: '', readiness: 0, variants: [],
  hasTutorial: false, tutorialUrl: '', tutorialAccess: '',
  mediaSources: [],
});

const seedTricks = () => ([
  { id: crypto.randomUUID(), ...emptyTrick(), name: "Crazy Man's Handcuffs", contexts: ['mingel'], lengthMin: 3, description: 'Åskådarens handbojor låses ihop med utövarens utan möjlighet, ändå separeras de.', notes: 'Fast öppnare i mingelrotationen.', readiness: 9 },
  { id: crypto.randomUUID(), ...emptyTrick(), name: 'Double Cross', source: 'Mark Southworth', contexts: ['mingel'], lengthMin: 4, description: 'Ett X flyttar osynligt mellan två kort inför åskådarens ögon.', notes: 'Öppen fråga: ingen tematisk motivering hittad än för varför X måste flytta till åskådaren, utan att luta sig mot förkunskap eller övernaturlig inramning.', readiness: 6 },
  { id: crypto.randomUUID(), ...emptyTrick(), name: 'Invisible Deck', contexts: ['mingel', 'foretagsevent'], lengthMin: 3, description: 'Åskådaren tänker ett kort, det visar sig vara vänt i en kortlek som varit förseglad hela tiden.', usesMusic: true, song: 'Nuvole Bianche', artist: 'Ludovico Einaudi', readiness: 8,
    variants: [{ id: crypto.randomUUID(), name: 'Barnversion', readiness: 5, notes: 'Enklare presentation, ingen förutsägelse-twist.' }] },
  { id: crypto.randomUUID(), ...emptyTrick(), name: "Professor's Nightmare", contexts: ['mingel'], lengthMin: 4, description: 'Tre olika snören visar sig vara exakt lika långa, om och om igen.', readiness: 7 },
]);

// Migrate old-format gigs (phase arrays of plain trick-id strings) to the current
// row format { entryId, trickId, rating, comment }.
function migrateGig(g) {
  const rowDefaults = { rating: '', comment: '', variantId: '', overrideLengthMin: '', musicOn: null, overrideSong: '', overrideArtist: '', variantNote: '' };
  const fixPhase = (arr) => (arr || []).map(item =>
    typeof item === 'string'
      ? { entryId: crypto.randomUUID(), trickId: item, ...rowDefaults }
      : { ...rowDefaults, ...item }
  );
  return {
    ...emptyGig(),
    ...g,
    phaseTricks: {
      opening: fixPhase(g.phaseTricks?.opening),
      middle: fixPhase(g.phaseTricks?.middle),
      closing: fixPhase(g.phaseTricks?.closing),
    },
  };
}

function useStorage() {
  const [tricks, setTricks] = useState([]);
  const [gigs, setGigs] = useState([]);
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES);
  const [media, setMedia] = useState([]);
  const [genres, setGenres] = useState(DEFAULT_GENRES);
  const [formats, setFormats] = useState(DEFAULT_FORMATS);
  const [devItems, setDevItems] = useState([]);
  const [settings, setSettings] = useState(defaultSettings());
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    (async () => {
      let t = null, s = null, c = null, m = null, st = null, g = null, f = null, d = null;
      try { t = await window.storage.get('repertoar-tricks'); } catch (e) { t = null; }
      try { s = await window.storage.get('repertoar-setlists'); } catch (e) { s = null; }
      try { c = await window.storage.get('repertoar-categories'); } catch (e) { c = null; }
      try { m = await window.storage.get('repertoar-media'); } catch (e) { m = null; }
      try { st = await window.storage.get('repertoar-settings'); } catch (e) { st = null; }
      try { g = await window.storage.get('repertoar-genres'); } catch (e) { g = null; }
      try { f = await window.storage.get('repertoar-formats'); } catch (e) { f = null; }
      try { d = await window.storage.get('repertoar-development'); } catch (e) { d = null; }
      const loadedCategories = c ? JSON.parse(c.value) : DEFAULT_CATEGORIES;
      setTricks(t ? JSON.parse(t.value) : seedTricks());
      setGigs(s ? JSON.parse(s.value).map(migrateGig) : []);
      setCategories(loadedCategories);
      setMedia(m ? JSON.parse(m.value).map(x => ({ ...emptyMedia(), ...x, genres: x.genres || [] })) : []);
      setGenres(g ? JSON.parse(g.value) : DEFAULT_GENRES);
      setFormats(f ? { ...DEFAULT_FORMATS, ...JSON.parse(f.value) } : DEFAULT_FORMATS);
      setDevItems(d ? JSON.parse(d.value).map(x => ({ ...emptyDevItem(), ...x })) : []);
      setSettings(st ? { ...defaultSettings(), ...JSON.parse(st.value) } : defaultSettings());
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-tricks', JSON.stringify(tricks)).catch(() => {});
  }, [tricks, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-setlists', JSON.stringify(gigs)).catch(() => {});
  }, [gigs, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-settings', JSON.stringify(settings)).catch(() => {});
  }, [settings, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-categories', JSON.stringify(categories)).catch(() => {});
  }, [categories, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-media', JSON.stringify(media)).catch(() => {});
  }, [media, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-development', JSON.stringify(devItems)).catch(() => {});
  }, [devItems, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-genres', JSON.stringify(genres)).catch(() => {});
  }, [genres, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set('repertoar-formats', JSON.stringify(formats)).catch(() => {});
  }, [formats, loaded]);

  return { tricks, setTricks, gigs, setGigs, categories, setCategories, media, setMedia, genres, setGenres, formats, setFormats, devItems, setDevItems, settings, setSettings, loaded };
}

function ReadinessBar({ value, onChange, compact }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      {onChange ? (
        <input type="range" min="0" max="10" value={value} onChange={e => onChange(Number(e.target.value))} style={{ flex: 1, accentColor: 'var(--accent-gold)' }} />
      ) : (
        <div style={{ flex: 1, height: compact ? 5 : 7, background: 'var(--card)', border: '1px solid var(--border)', position: 'relative' }}>
          <div style={{ position: 'absolute', inset: 0, width: `${(value / 10) * 100}%`, background: readinessColor(value) }} />
        </div>
      )}
      <span style={{ fontSize: compact ? 11 : 12.5, color: 'var(--text-muted)', width: 30, textAlign: 'right', flexShrink: 0 }}>{value}/10</span>
    </div>
  );
}
function readinessColor(v) {
  if (v >= 8) return '#7a9b6e';
  if (v >= 4) return '#b8935a';
  return '#a85a4a';
}

function TrickForm({ initial, onSave, onCancel, categories, media }) {
  const [form, setForm] = useState({ ...emptyTrick(), ...(initial || {}) });
  const toggleContext = (c) => {
    setForm(f => ({ ...f, contexts: f.contexts.includes(c) ? f.contexts.filter(x => x !== c) : [...f.contexts, c] }));
  };
  const addVariant = () => setForm(f => ({ ...f, variants: [...f.variants, { id: crypto.randomUUID(), name: '', readiness: 0, notes: '' }] }));
  const updateVariant = (id, patch) => setForm(f => ({ ...f, variants: f.variants.map(v => v.id === id ? { ...v, ...patch } : v) }));
  const removeVariant = (id) => setForm(f => ({ ...f, variants: f.variants.filter(v => v.id !== id) }));

  const addMediaSource = () => setForm(f => ({ ...f, mediaSources: [...f.mediaSources, { id: crypto.randomUUID(), mediaId: '', locator: '' }] }));
  const updateMediaSource = (id, patch) => setForm(f => ({ ...f, mediaSources: f.mediaSources.map(s => s.id === id ? { ...s, ...patch } : s) }));
  const removeMediaSource = (id) => setForm(f => ({ ...f, mediaSources: f.mediaSources.filter(s => s.id !== id) }));

  return (
    <div style={{ border: '1px solid var(--border)', background: 'var(--panel)', padding: 20, display: 'flex', flexDirection: 'column', gap: 14 }}>
      <input placeholder="Namn på trick" value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} style={inputStyle} />
      <input placeholder="Källa (t.ex. Mark Southworth) — valfritt" value={form.source} onChange={e => setForm(f => ({ ...f, source: e.target.value }))} style={inputStyle} />
      <div>
        <div style={labelStyle}>Kort beskrivning</div>
        <input placeholder="Vad tricket går ut på, i en mening" value={form.description} onChange={e => setForm(f => ({ ...f, description: e.target.value }))} style={inputStyle} />
      </div>
      <div>
        <div style={labelStyle}>Kontext</div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {categories.map(c => (
            <button key={c.id} onClick={() => toggleContext(c.id)}
              style={{ ...chipStyle, background: form.contexts.includes(c.id) ? 'var(--accent-gold)' : 'transparent', color: form.contexts.includes(c.id) ? 'var(--ink)' : 'var(--text-muted)', borderColor: form.contexts.includes(c.id) ? 'var(--accent-gold)' : 'var(--border)' }}>
              {c.label}
            </button>
          ))}
        </div>
      </div>
      <div style={{ display: 'flex', gap: 14 }}>
        <div style={{ flex: 1 }}>
          <div style={labelStyle}>Längd (min)</div>
          <input type="number" min="1" value={form.lengthMin} onChange={e => setForm(f => ({ ...f, lengthMin: Number(e.target.value) || 1 }))} style={inputStyle} />
        </div>
        <div style={{ flex: 1 }}>
          <div style={labelStyle}>Status</div>
          <select value={form.status} onChange={e => setForm(f => ({ ...f, status: e.target.value }))} style={inputStyle}>
            {Object.entries(STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
          </select>
        </div>
      </div>

      <div>
        <div style={labelStyle}>Inövad / redo (0–10)</div>
        <ReadinessBar value={form.readiness} onChange={v => setForm(f => ({ ...f, readiness: v }))} />
      </div>

      <div style={{ borderTop: '1px solid var(--border)', paddingTop: 14 }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13.5, cursor: 'pointer', marginBottom: form.usesMusic ? 12 : 0 }}>
          <input type="checkbox" checked={form.usesMusic} onChange={e => setForm(f => ({ ...f, usesMusic: e.target.checked }))} />
          Har musik till tricket
        </label>
        {form.usesMusic && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div style={{ display: 'flex', gap: 10 }}>
              <input placeholder="Låt" value={form.song} onChange={e => setForm(f => ({ ...f, song: e.target.value }))} style={{ ...inputStyle, flex: 1 }} />
              <input placeholder="Artist" value={form.artist} onChange={e => setForm(f => ({ ...f, artist: e.target.value }))} style={{ ...inputStyle, flex: 1 }} />
            </div>
            <input placeholder="Alternativ låt — valfritt" value={form.altSong} onChange={e => setForm(f => ({ ...f, altSong: e.target.value }))} style={inputStyle} />
          </div>
        )}
      </div>

      <div style={{ borderTop: '1px solid var(--border)', paddingTop: 14 }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13.5, cursor: 'pointer', marginBottom: form.hasTutorial ? 12 : 0 }}>
          <input type="checkbox" checked={form.hasTutorial} onChange={e => setForm(f => ({ ...f, hasTutorial: e.target.checked }))} />
          Har online-tutorial
        </label>
        {form.hasTutorial && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <input placeholder="Länk till tutorial" value={form.tutorialUrl} onChange={e => setForm(f => ({ ...f, tutorialUrl: e.target.value }))} style={inputStyle} />
            <div>
              <input placeholder="Åtkomstinfo, t.ex. lösenord — valfritt" value={form.tutorialAccess} onChange={e => setForm(f => ({ ...f, tutorialAccess: e.target.value }))} style={inputStyle} />
              <div style={{ fontSize: 11, color: 'var(--text-muted)', marginTop: 4, fontStyle: 'italic' }}>Sparas i klartext i appens lagring. Bra för minneslappar, inte för känsliga lösenord du återanvänder på andra ställen.</div>
            </div>
          </div>
        )}
      </div>

      <div>
        <div style={labelStyle}>Anteckningar — vad funkade, vad floppade, öppna frågor</div>
        <textarea rows={3} value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
      </div>

      <div style={{ borderTop: '1px solid var(--border)', paddingTop: 14 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
          <div style={labelStyle}>Källa i bok/film</div>
          <button onClick={addMediaSource} style={{ ...ghostButtonStyle, padding: '5px 10px', fontSize: 12 }}><Plus size={12} /> Lägg till källa</button>
        </div>
        {media.length === 0 && <div style={{ fontSize: 12.5, color: 'var(--text-muted)' }}>Inga titlar i Film & böcker än — lägg till där först.</div>}
        {media.length > 0 && form.mediaSources.length === 0 && <div style={{ fontSize: 12.5, color: 'var(--text-muted)' }}>Koppla tricket till en bok eller film, med sida eller tid/avsnitt, för att lätt hitta tillbaka.</div>}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {form.mediaSources.map(s => (
            <div key={s.id} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <select value={s.mediaId} onChange={e => updateMediaSource(s.id, { mediaId: e.target.value })} style={{ ...inputStyle, flex: 2 }}>
                <option value="">Välj titel…</option>
                {media.map(m => <option key={m.id} value={m.id}>{m.title}</option>)}
              </select>
              <input placeholder="Sida / tid / avsnitt" value={s.locator} onChange={e => updateMediaSource(s.id, { locator: e.target.value })} style={{ ...inputStyle, flex: 1 }} />
              <button onClick={() => removeMediaSource(s.id)} style={iconButtonStyle}><X size={13} /></button>
            </div>
          ))}
        </div>
      </div>

      <div style={{ borderTop: '1px solid var(--border)', paddingTop: 14 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
          <div style={labelStyle}>Varianter / framföranden</div>
          <button onClick={addVariant} style={{ ...ghostButtonStyle, padding: '5px 10px', fontSize: 12 }}><Plus size={12} /> Lägg till variant</button>
        </div>
        {form.variants.length === 0 && <div style={{ fontSize: 12.5, color: 'var(--text-muted)' }}>Inga varianter. Lägg till t.ex. "Barnversion" eller "Kort mingelversion" om du framför tricket på olika sätt.</div>}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {form.variants.map(v => (
            <div key={v.id} style={{ border: '1px solid var(--border)', background: 'var(--card)', padding: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <input placeholder="Namn på variant" value={v.name} onChange={e => updateVariant(v.id, { name: e.target.value })} style={{ ...inputStyle, flex: 1 }} />
                <button onClick={() => removeVariant(v.id)} style={iconButtonStyle}><X size={13} /></button>
              </div>
              <ReadinessBar value={v.readiness} onChange={val => updateVariant(v.id, { readiness: val })} compact />
              <input placeholder="Anteckning om varianten" value={v.notes} onChange={e => updateVariant(v.id, { notes: e.target.value })} style={{ ...inputStyle, fontSize: 12.5 }} />
            </div>
          ))}
        </div>
      </div>

      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
        <button onClick={onCancel} style={ghostButtonStyle}>Avbryt</button>
        <button onClick={() => form.name.trim() && onSave(form)} style={primaryButtonStyle}><Save size={14} /> Spara</button>
      </div>
    </div>
  );
}

function lastUsedGig(trickId, gigs) {
  const matches = gigs.filter(g => PHASES.some(p => g.phaseTricks[p.id].some(r => r.trickId === trickId)));
  if (!matches.length) return null;
  return [...matches].sort((a, b) => (b.date || '').localeCompare(a.date || ''))[0];
}

function TrickCard({ trick, onEdit, onDelete, categories, gigs, onOpenGig, media }) {
  const [expanded, setExpanded] = useState(false);
  const hasSong = trick.usesMusic && trick.song.trim();
  const lastGig = lastUsedGig(trick.id, gigs);
  return (
    <div style={cardStyle}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 10 }}>
        <div>
          <div style={{ fontFamily: 'var(--font-display)', fontSize: 18, color: 'var(--text)' }}>{trick.name}</div>
          {trick.source && <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginTop: 2 }}>{trick.source}</div>}
        </div>
        <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
          <button onClick={() => onEdit(trick)} style={iconButtonStyle} aria-label="Redigera">✎</button>
          <button onClick={() => onDelete(trick.id)} style={iconButtonStyle} aria-label="Ta bort"><Trash2 size={13} /></button>
          <button onClick={() => setExpanded(v => !v)} style={iconButtonStyle} aria-label={expanded ? 'Minska' : 'Öka'}>
            {expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          </button>
        </div>
      </div>
      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 12 }}>
        {trick.contexts.map(c => <span key={c} style={tagStyle}>{categories.find(x => x.id === c)?.label}</span>)}
      </div>

      {expanded && (
        <>
          {trick.description && <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginTop: 10, lineHeight: 1.4 }}>{trick.description}</div>}
          <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginTop: 12, fontSize: 12.5, color: 'var(--text-muted)', flexWrap: 'wrap' }}>
            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><Clock size={12} /> {trick.lengthMin} min</span>
            <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
              <span style={{ width: 6, height: 6, borderRadius: '50%', background: STATUS[trick.status].color }} />
              {STATUS[trick.status].label}
            </span>
            {trick.usesMusic && (
              <span title={hasSong ? 'Musik vald' : 'Har musik, låt ej vald än'} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <span style={{ width: 7, height: 7, borderRadius: '50%', background: hasSong ? '#7a9b6e' : '#c9a63c', display: 'inline-block' }} />
                <Music size={12} /> Musik
              </span>
            )}
            {trick.hasTutorial && trick.tutorialUrl && (
              <a href={trick.tutorialUrl} target="_blank" rel="noopener noreferrer" style={{ display: 'flex', alignItems: 'center', gap: 4, color: 'var(--accent-gold)', textDecoration: 'none' }}>
                <ExternalLink size={12} /> Tutorial
              </a>
            )}
          </div>
          {trick.hasTutorial && trick.tutorialAccess && (
            <div style={{ fontSize: 11.5, color: 'var(--text-muted)', marginTop: 6, fontStyle: 'italic' }}>Åtkomst: {trick.tutorialAccess}</div>
          )}
          {hasSong && (
            <div style={{ marginTop: 10, fontSize: 12.5, color: 'var(--text-muted)' }}>
              {trick.song}{trick.artist ? ` — ${trick.artist}` : ''}
              {trick.altSong && <div style={{ marginTop: 2 }}>Alt: {trick.altSong}</div>}
            </div>
          )}
          <div style={{ marginTop: 12 }}>
            <ReadinessBar value={trick.readiness ?? 0} compact />
          </div>
          {trick.variants && trick.variants.length > 0 && (
            <div style={{ marginTop: 12, borderTop: '1px solid var(--border)', paddingTop: 10, display: 'flex', flexDirection: 'column', gap: 8 }}>
              {trick.variants.map(v => (
                <div key={v.id}>
                  <div style={{ fontSize: 12.5, color: 'var(--text)', marginBottom: 3 }}>{v.name || '(namnlös variant)'}</div>
                  <ReadinessBar value={v.readiness ?? 0} compact />
                </div>
              ))}
            </div>
          )}
          {trick.notes && <div style={{ marginTop: 10, fontSize: 13, color: 'var(--text-muted)', borderTop: '1px solid var(--border)', paddingTop: 10, lineHeight: 1.5 }}>{trick.notes}</div>}
          {trick.mediaSources && trick.mediaSources.filter(s => s.mediaId).length > 0 && (
            <div style={{ marginTop: 10, borderTop: '1px solid var(--border)', paddingTop: 10, fontSize: 12.5, color: 'var(--text-muted)', display: 'flex', flexDirection: 'column', gap: 4 }}>
              {trick.mediaSources.filter(s => s.mediaId).map(s => {
                const m = media.find(x => x.id === s.mediaId);
                if (!m) return null;
                return (
                  <span key={s.id} style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                    {m.type === 'bok' ? <BookOpen size={12} /> : <Film size={12} />}
                    {m.title}{s.locator ? ` — ${s.locator}` : ''}
                  </span>
                );
              })}
            </div>
          )}
          <div style={{ marginTop: 10, borderTop: '1px solid var(--border)', paddingTop: 10, fontSize: 12.5, color: 'var(--text-muted)' }}>
            {lastGig ? (
              <span>Senast använd: {lastGig.name || '(namnlöst uppdrag)'}{lastGig.date ? `, ${lastGig.date}` : ''}
                {' '}<button onClick={() => onOpenGig(lastGig)} style={{ background: 'none', border: 'none', padding: 0, color: 'var(--accent-gold)', cursor: 'pointer', font: 'inherit', textDecoration: 'underline' }}>Visa uppdrag</button>
              </span>
            ) : 'Inte använt i något sparat uppdrag ännu'}
          </div>
        </>
      )}
    </div>
  );
}

function LibraryView({ tricks, setTricks, categories, gigs, onOpenGig, media }) {
  const [showFilters, setShowFilters] = useState(false);
  const [search, setSearch] = useState('');
  const [contextFilters, setContextFilters] = useState([]);
  const [musicFilter, setMusicFilter] = useState(false);
  const [statusFilter, setStatusFilter] = useState(null);
  const [editing, setEditing] = useState(null);
  const [adding, setAdding] = useState(false);

  const toggleContextFilter = (id) => setContextFilters(cf => cf.includes(id) ? cf.filter(x => x !== id) : [...cf, id]);

  const filtered = tricks.filter(t =>
    (t.name.toLowerCase().includes(search.toLowerCase()) || t.notes.toLowerCase().includes(search.toLowerCase())) &&
    (contextFilters.length === 0 || t.contexts.some(c => contextFilters.includes(c))) &&
    (!statusFilter || t.status === statusFilter) &&
    (!musicFilter || (t.usesMusic && t.song))
  );
  const anyFilterActive = contextFilters.length > 0 || musicFilter || statusFilter || search;

  const saveTrick = (form) => {
    if (editing && editing.id) {
      setTricks(ts => ts.map(t => t.id === editing.id ? { ...form, id: editing.id } : t));
    } else {
      setTricks(ts => [...ts, { ...form, id: crypto.randomUUID() }]);
    }
    setEditing(null);
    setAdding(false);
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, flexWrap: 'wrap' }}>
        <button onClick={() => setShowFilters(v => !v)} style={{ ...ghostButtonStyle, background: showFilters ? 'var(--card)' : 'transparent' }}>
          <Filter size={14} /> Filter{anyFilterActive ? ' •' : ''}
        </button>
        <button onClick={() => { setAdding(true); setEditing({}); }} style={primaryButtonStyle}><Plus size={14} /> Nytt trick</button>
      </div>

      {showFilters && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginBottom: 18 }}>
          <div>
            <div style={labelStyle}>Sök</div>
            <div style={{ position: 'relative', maxWidth: 320 }}>
              <Search size={14} style={{ position: 'absolute', left: 12, top: 12, color: 'var(--text-muted)' }} />
              <input placeholder="Sök trick eller anteckning" value={search} onChange={e => setSearch(e.target.value)} style={{ ...inputStyle, paddingLeft: 34 }} />
            </div>
          </div>
          <div>
            <div style={labelStyle}>Kategori</div>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button onClick={() => setContextFilters([])} style={{ ...chipStyle, background: contextFilters.length === 0 ? 'var(--accent-gold)' : 'transparent', color: contextFilters.length === 0 ? 'var(--ink)' : 'var(--text-muted)' }}>Alla typer</button>
              {categories.map(c => (
                <button key={c.id} onClick={() => toggleContextFilter(c.id)} style={{ ...chipStyle, background: contextFilters.includes(c.id) ? 'var(--accent-gold)' : 'transparent', color: contextFilters.includes(c.id) ? 'var(--ink)' : 'var(--text-muted)' }}>{c.label}</button>
              ))}
            </div>
          </div>
          <div>
            <div style={labelStyle}>Teknik</div>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button onClick={() => setMusicFilter(v => !v)} style={{ ...chipStyle, display: 'flex', alignItems: 'center', gap: 6, background: musicFilter ? 'var(--accent-gold)' : 'transparent', color: musicFilter ? 'var(--ink)' : 'var(--text-muted)' }}><Music size={12} /> Har musik</button>
            </div>
          </div>
          <div>
            <div style={labelStyle}>Utveckling</div>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              {Object.entries(STATUS).map(([k, v]) => (
                <button key={k} onClick={() => setStatusFilter(statusFilter === k ? null : k)} style={{ ...chipStyle, background: statusFilter === k ? 'var(--accent-gold)' : 'transparent', color: statusFilter === k ? 'var(--ink)' : 'var(--text-muted)' }}>{v.label}</button>
              ))}
            </div>
          </div>
        </div>
      )}

      {(adding || editing) && (
        <div style={{ marginBottom: 18 }}>
          <TrickForm initial={editing?.id ? editing : null} onSave={saveTrick} onCancel={() => { setEditing(null); setAdding(false); }} categories={categories} media={media} />
        </div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {filtered.map(t => (
          <TrickCard key={t.id} trick={t} onEdit={(tr) => { setEditing(tr); setAdding(false); }} onDelete={(id) => setTricks(ts => ts.filter(x => x.id !== id))} categories={categories} gigs={gigs} onOpenGig={onOpenGig} media={media} />
        ))}
      </div>
      {filtered.length === 0 && <div style={{ color: 'var(--text-muted)', padding: 30, textAlign: 'center', fontSize: 14 }}>Inga trick hittades.</div>}
    </div>
  );
}

function ratingLabel(r) { return r ? `${r}/5` : '–'; }


function Builder({ tricks, gigs, setGigs, activeGig, setActiveGig, onPrint, categories }) {
  const gig = activeGig;
  const [pickerFor, setPickerFor] = useState(null);
  const [expandedRows, setExpandedRows] = useState({});
  const dragSource = useRef(null);

  const setGig = (patch) => setActiveGig(g => ({ ...g, ...patch }));

  const totalLength = PHASES.reduce((sum, p) => sum + gig.phaseTricks[p.id].reduce((s, row) => {
    const t = tricks.find(x => x.id === row.trickId);
    return s + effectiveLength(row, t);
  }, 0), 0);
  const totalTricks = PHASES.reduce((sum, p) => sum + gig.phaseTricks[p.id].length, 0);
  const phaseStart = (phaseId) => {
    let n = 0;
    for (const p of PHASES) { if (p.id === phaseId) break; n += gig.phaseTricks[p.id].length; }
    return n;
  };

  const addToPhase = (phaseId, trickId) => {
    setGig({ phaseTricks: { ...gig.phaseTricks, [phaseId]: [...gig.phaseTricks[phaseId], {
      entryId: crypto.randomUUID(), trickId, rating: '', comment: '',
      variantId: '', overrideLengthMin: '', musicOn: null, overrideSong: '', overrideArtist: '', variantNote: '',
    }] } });
    setPickerFor(null);
  };

  const toggleExpanded = (entryId) => setExpandedRows(r => ({ ...r, [entryId]: !r[entryId] }));

  const updateRow = (phaseId, entryId, patch) => {
    setGig({ phaseTricks: { ...gig.phaseTricks, [phaseId]: gig.phaseTricks[phaseId].map(r => r.entryId === entryId ? { ...r, ...patch } : r) } });
  };

  const removeRow = (phaseId, entryId) => {
    setGig({ phaseTricks: { ...gig.phaseTricks, [phaseId]: gig.phaseTricks[phaseId].filter(r => r.entryId !== entryId) } });
  };

  const moveItem = (source, dest) => {
    const next = { opening: [...gig.phaseTricks.opening], middle: [...gig.phaseTricks.middle], closing: [...gig.phaseTricks.closing] };
    const [moved] = next[source.phaseId].splice(source.index, 1);
    let destIndex = dest.index;
    if (source.phaseId === dest.phaseId && source.index < dest.index) destIndex -= 1;
    next[dest.phaseId].splice(destIndex, 0, moved);
    setGig({ phaseTricks: next });
  };

  const handleDragStart = (phaseId, index) => (e) => {
    dragSource.current = { phaseId, index };
    e.dataTransfer.effectAllowed = 'move';
  };
  const handleDropOnRow = (phaseId, index) => (e) => {
    e.preventDefault();
    if (dragSource.current) moveItem(dragSource.current, { phaseId, index });
    dragSource.current = null;
  };
  const handleDropOnColumn = (phaseId) => (e) => {
    e.preventDefault();
    if (dragSource.current) moveItem(dragSource.current, { phaseId, index: gig.phaseTricks[phaseId].length });
    dragSource.current = null;
  };

  const [justSaved, setJustSaved] = useState(false);
  const save = () => {
    if (!gig.name.trim()) return;
    if (gig.id) {
      setGigs(gs => gs.map(g => g.id === gig.id ? gig : g));
    } else {
      const withId = { ...gig, id: crypto.randomUUID() };
      setGigs(gs => [...gs, withId]);
      setActiveGig(withId);
    }
    setJustSaved(true);
    setTimeout(() => setJustSaved(false), 2500);
  };

  const relevant = tricks.filter(t => t.contexts.includes(gig.eventType));
  const [showAddFilter, setShowAddFilter] = useState(false);
  const [addFilter, setAddFilter] = useState([]);
  const toggleAddFilter = (id) => setAddFilter(af => af.includes(id) ? af.filter(x => x !== id) : [...af, id]);
  const pickerTricks = addFilter.length > 0
    ? tricks.filter(t => t.contexts.some(c => addFilter.includes(c)))
    : (relevant.length ? relevant : tricks);

  return (
    <div>
      <div style={{ ...panelBox, marginBottom: 18 }}>
        <div style={fieldGrid}>
          <div>
            <div style={labelStyle}>Namn på uppdrag</div>
            <input value={gig.name} onChange={e => setGig({ name: e.target.value })} placeholder="t.ex. Barnkalas Follingbo" style={inputStyle} />
          </div>
          <div>
            <div style={labelStyle}>Uppdragstyp</div>
            <select value={gig.eventType} onChange={e => setGig({ eventType: e.target.value })} style={inputStyle}>
              {categories.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
            </select>
          </div>
          <div>
            <div style={labelStyle}>Datum</div>
            <input type="date" value={gig.date} onChange={e => setGig({ date: e.target.value })} style={inputStyle} />
          </div>
          <div>
            <div style={labelStyle}>Plats</div>
            <input value={gig.location} onChange={e => setGig({ location: e.target.value })} placeholder="t.ex. Vokstugan, Visby" style={inputStyle} />
          </div>
          <div>
            <div style={labelStyle}>Publik</div>
            <input value={gig.audienceType} onChange={e => setGig({ audienceType: e.target.value })} placeholder="t.ex. Barnkalas 5-8 år" style={inputStyle} />
          </div>
          <div>
            <div style={labelStyle}>Antal i publik</div>
            <input type="number" min="0" value={gig.audienceCount} onChange={e => setGig({ audienceCount: e.target.value })} style={inputStyle} />
          </div>
          <div>
            <div style={labelStyle}>Planerad längd (min) <span style={{ color: 'var(--text-muted)', fontStyle: 'italic' }}>— det du sagt till kunden</span></div>
            <input type="number" min="0" value={gig.plannedLengthMin} onChange={e => setGig({ plannedLengthMin: e.target.value })} style={inputStyle} />
          </div>
          <div>
            <div style={labelStyle}>Scen</div>
            <select value={gig.venue} onChange={e => setGig({ venue: e.target.value })} style={inputStyle}>
              <option value="inne">Inomhus</option>
              <option value="ute">Utomhus</option>
            </select>
          </div>
          <div>
            <div style={labelStyle}>Pris (kr) <span style={printHintStyle}>— visas ej vid utskrift</span></div>
            <input type="number" min="0" value={gig.priceSEK} onChange={e => setGig({ priceSEK: e.target.value })} style={inputStyle} />
          </div>
        </div>
        <div style={{ marginTop: 14 }}>
          <div style={labelStyle}>Bakgrundsinfo <span style={printHintStyle}>— visas ej vid utskrift</span></div>
          <textarea rows={2} value={gig.backgroundNotes} onChange={e => setGig({ backgroundNotes: e.target.value })}
            placeholder="Kontaktperson, särskilda önskemål, allt du behöver komma ihåg men inte ha på papperet på scen"
            style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
        </div>
      </div>

      <div style={{ marginBottom: 14 }}>
        <button onClick={() => setShowAddFilter(v => !v)} style={{ ...ghostButtonStyle, background: showAddFilter ? 'var(--card)' : 'transparent' }}>
          <Filter size={14} /> Filtrera trickval{addFilter.length > 0 ? ' •' : ''}
        </button>
        {showAddFilter && (
          <div style={{ marginTop: 10 }}>
            <div style={labelStyle}>Kategori</div>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              {categories.map(c => (
                <button key={c.id} onClick={() => toggleAddFilter(c.id)}
                  style={{ ...chipStyle, background: addFilter.includes(c.id) ? 'var(--accent-gold)' : 'transparent', color: addFilter.includes(c.id) ? 'var(--ink)' : 'var(--text-muted)' }}>
                  {c.label}
                </button>
              ))}
              {addFilter.length > 0 && <button onClick={() => setAddFilter([])} style={ghostButtonStyle}>Rensa</button>}
            </div>
          </div>
        )}
        <div style={{ fontSize: 11.5, color: 'var(--text-muted)', marginTop: 6 }}>
          {addFilter.length > 0 ? 'Visar bara valda kategorier i "Lägg till trick" nedan.' : `Utan filter visas trick taggade "${categories.find(c => c.id === gig.eventType)?.label || gig.eventType}" i förslagen, eller alla om inga matchar.`}
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 16 }}>
        {PHASES.map(phase => (
          <div key={phase.id} style={{ border: '1px solid var(--border)', background: 'var(--panel)', padding: 16 }}
            onDragOver={(e) => e.preventDefault()} onDrop={handleDropOnColumn(phase.id)}>
            <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)' }}>{phase.label}</div>
            <div style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 12 }}>{phase.hint}</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, minHeight: 20 }}>
              {gig.phaseTricks[phase.id].map((row, idx) => {
                const t = tricks.find(x => x.id === row.trickId);
                const num = phaseStart(phase.id) + idx + 1;
                const music = effectiveMusic(row, t);
                const isExpanded = !!expandedRows[row.entryId];
                const isCustomized = row.variantId || row.overrideLengthMin !== '' || isMusicCustomized(row, t) || row.variantNote;
                return (
                  <div key={row.entryId} draggable onDragStart={handleDragStart(phase.id, idx)}
                    onDragOver={(e) => e.preventDefault()} onDrop={handleDropOnRow(phase.id, idx)}
                    style={{ background: 'var(--card)', border: '1px solid var(--border)', padding: '8px 10px', cursor: 'grab' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
                      <span style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13.5 }}>
                        <GripVertical size={13} style={{ color: 'var(--text-muted)', flexShrink: 0 }} />
                        <span style={{ color: 'var(--text-muted)', fontSize: 12, minWidth: 14 }}>{num}.</span>
                        {effectiveName(row, t)}
                        <span style={{ color: 'var(--text-muted)', fontSize: 11.5 }}>({effectiveLength(row, t)} min)</span>
                      </span>
                      <span style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
                        <button onClick={() => toggleExpanded(row.entryId)}
                          style={{ ...ghostButtonStyle, padding: '3px 8px', fontSize: 11, borderColor: isCustomized ? 'var(--accent-gold)' : 'var(--border)', color: isCustomized ? 'var(--accent-gold)' : 'var(--text-muted)' }}>
                          {isCustomized ? 'Anpassad' : 'Anpassa'}
                        </button>
                        <button onClick={() => removeRow(phase.id, row.entryId)} style={iconButtonStyle}><X size={13} /></button>
                      </span>
                    </div>
                    {music && (
                      <div style={{ fontSize: 11.5, color: 'var(--text-muted)', marginTop: 4, paddingLeft: 19, display: 'flex', alignItems: 'center', gap: 4 }}>
                        <Music size={11} /> {music.song}{music.artist ? ` — ${music.artist}` : ''}
                      </div>
                    )}

                    {isExpanded && (
                      <div style={{ marginTop: 8, paddingLeft: 19, display: 'flex', flexDirection: 'column', gap: 8, borderTop: '1px solid var(--border)', paddingTop: 8 }}>
                        {t?.variants?.length > 0 && (
                          <select value={row.variantId} onChange={e => updateRow(phase.id, row.entryId, { variantId: e.target.value })}
                            style={{ ...inputStyle, padding: '5px 8px', fontSize: 12 }}>
                            <option value="">Standardversion</option>
                            {t.variants.map(v => <option key={v.id} value={v.id}>{v.name || 'namnlös variant'}</option>)}
                          </select>
                        )}
                        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                          <input type="number" min="0" placeholder={`Längd, standard ${t?.lengthMin ?? ''} min`} value={row.overrideLengthMin}
                            onChange={e => updateRow(phase.id, row.entryId, { overrideLengthMin: e.target.value })}
                            style={{ ...inputStyle, padding: '5px 8px', fontSize: 12, width: 160 }} />
                          <span style={{ fontSize: 11, color: 'var(--text-muted)' }}>min för denna show</span>
                        </div>
                        <div>
                          <button onClick={() => updateRow(phase.id, row.entryId, { musicOn: !effectiveMusicOn(row, t) })}
                            style={{ ...chipStyle, display: 'inline-flex', alignItems: 'center', gap: 6, background: effectiveMusicOn(row, t) ? 'var(--accent-gold)' : 'transparent', color: effectiveMusicOn(row, t) ? 'var(--ink)' : 'var(--text-muted)' }}>
                            <Music size={12} /> Musik {effectiveMusicOn(row, t) ? 'på' : 'av'}
                          </button>
                        </div>
                        {effectiveMusicOn(row, t) && (
                          <div style={{ display: 'flex', gap: 8 }}>
                            <input placeholder="Låt" value={row.overrideSong || t?.song || ''} onChange={e => updateRow(phase.id, row.entryId, { overrideSong: e.target.value })}
                              style={{ ...inputStyle, padding: '5px 8px', fontSize: 12, flex: 1 }} />
                            <input placeholder="Artist" value={row.overrideArtist || t?.artist || ''} onChange={e => updateRow(phase.id, row.entryId, { overrideArtist: e.target.value })}
                              style={{ ...inputStyle, padding: '5px 8px', fontSize: 12, flex: 1 }} />
                          </div>
                        )}
                        <input placeholder="Övrig anpassning för denna show — valfritt" value={row.variantNote} onChange={e => updateRow(phase.id, row.entryId, { variantNote: e.target.value })}
                          style={{ ...inputStyle, padding: '5px 8px', fontSize: 12 }} />
                      </div>
                    )}

                    <div style={{ display: 'flex', gap: 8, marginTop: 8, paddingLeft: 19 }}>
                      <select value={row.rating} onChange={e => updateRow(phase.id, row.entryId, { rating: e.target.value })}
                        style={{ ...inputStyle, padding: '5px 8px', fontSize: 12, width: 68 }}>
                        <option value="">Betyg</option>
                        {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}/5</option>)}
                      </select>
                      <input placeholder="Kommentar" value={row.comment} onChange={e => updateRow(phase.id, row.entryId, { comment: e.target.value })}
                        style={{ ...inputStyle, padding: '5px 8px', fontSize: 12, flex: 1 }} />
                    </div>
                  </div>
                );
              })}
            </div>
            {pickerFor === phase.id ? (
              <select autoFocus onChange={e => e.target.value && addToPhase(phase.id, e.target.value)} onBlur={() => setPickerFor(null)} style={{ ...inputStyle, marginTop: 10 }} defaultValue="">
                <option value="" disabled>Välj trick…</option>
                {pickerTricks.map(t => <option key={t.id} value={t.id}>{t.name} ({t.lengthMin} min)</option>)}
              </select>
            ) : (
              <button onClick={() => setPickerFor(phase.id)} style={{ ...ghostButtonStyle, marginTop: 10, width: '100%' }}><Plus size={13} /> Lägg till trick</button>
            )}
          </div>
        ))}
      </div>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', margin: '16px 0 20px', flexWrap: 'wrap', gap: 10 }}>
        <div style={{ fontSize: 13, color: 'var(--text-muted)' }}>Bedömd längd: <strong style={{ color: 'var(--text)' }}>{totalLength} min</strong> &nbsp;·&nbsp; Antal trick: <strong style={{ color: 'var(--text)' }}>{totalTricks}</strong></div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <button onClick={() => onPrint(gig)} style={ghostButtonStyle}><Printer size={14} /> Skriv ut</button>
          <button onClick={save} style={primaryButtonStyle}><Save size={14} /> Spara uppdrag</button>
          {justSaved && <span style={{ color: '#7a9b6e', fontSize: 13, display: 'flex', alignItems: 'center', gap: 4 }}>✓ Sparad</span>}
        </div>
      </div>

      <div style={{ ...panelBox, marginTop: 20 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)' }}>Efter föreställningen</div>
          <button onClick={() => setGig({ filmed: !gig.filmed })}
            style={{ ...chipStyle, display: 'flex', alignItems: 'center', gap: 6, background: gig.filmed ? 'var(--accent-gold)' : 'transparent', color: gig.filmed ? 'var(--ink)' : 'var(--text-muted)', borderColor: gig.filmed ? 'var(--accent-gold)' : 'var(--border)' }}>
            <Video size={13} /> {gig.filmed ? 'Filmad' : 'Inte filmad'}
          </button>
        </div>
        <div style={fieldGrid}>
          <div>
            <div style={labelStyle}>Verklig showlängd (min)</div>
            <input type="number" min="0" value={gig.actualLengthMin} onChange={e => setGig({ actualLengthMin: e.target.value })} style={inputStyle} />
          </div>
          <div>
            <div style={labelStyle}>Betyg hela föreställningen</div>
            <select value={gig.overallRating} onChange={e => setGig({ overallRating: e.target.value })} style={inputStyle}>
              <option value="">–</option>
              {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}/5</option>)}
            </select>
          </div>
        </div>
        <div style={{ marginTop: 14 }}>
          <div style={labelStyle}>Kommentar</div>
          <textarea rows={2} value={gig.overallComment} onChange={e => setGig({ overallComment: e.target.value })} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
        </div>
        <div style={{ display: 'flex', gap: 14, marginTop: 14 }}>
          <div style={{ flex: 1 }}>
            <div style={labelStyle}>Bra</div>
            <textarea rows={2} value={gig.good} onChange={e => setGig({ good: e.target.value })} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
          </div>
          <div style={{ flex: 1 }}>
            <div style={labelStyle}>Förbättring</div>
            <textarea rows={2} value={gig.improve} onChange={e => setGig({ improve: e.target.value })} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
          </div>
        </div>

        <div style={{ borderTop: '1px solid var(--border)', marginTop: 18, paddingTop: 14 }}>
          <div style={{ fontSize: 13, color: 'var(--text-muted)', marginBottom: 10 }}>Kundenkät — vad kunden själv svarat, inte din egen reflektion</div>
          <div>
            <div style={labelStyle}>Kundrecension ("skriv gärna några ord om framförandet")</div>
            <textarea rows={2} value={gig.customerReview} onChange={e => setGig({ customerReview: e.target.value })} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
          </div>
          <div style={{ marginTop: 10 }}>
            <div style={labelStyle}>Kundens förbättringsförslag ("finns något jag kan göra ännu bättre?")</div>
            <textarea rows={2} value={gig.customerImprovement} onChange={e => setGig({ customerImprovement: e.target.value })} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
          </div>
          <div style={{ marginTop: 10 }}>
            <div style={labelStyle}>Länk till formulärsvar — valfritt</div>
            <input value={gig.formResponseUrl} onChange={e => setGig({ formResponseUrl: e.target.value })} placeholder="Länk till Microsoft Forms-svaret eller exporten" style={inputStyle} />
          </div>
        </div>
      </div>
    </div>
  );
}

function GigsView({ gigs, setGigs, tricks, onOpen, onPrint, categories, exportProfiles }) {
  const [showFilters, setShowFilters] = useState(false);
  const [filmedOnly, setFilmedOnly] = useState(false);
  const [stageFilter, setStageFilter] = useState(null); // null | 'kommande' | 'genomford'
  const [search, setSearch] = useState('');
  const [typeFilters, setTypeFilters] = useState([]);
  const [year, setYear] = useState('');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [showExportMenu, setShowExportMenu] = useState(false);

  const toggleTypeFilter = (id) => setTypeFilters(tf => tf.includes(id) ? tf.filter(x => x !== id) : [...tf, id]);

  const years = [...new Set(gigs.map(g => (g.date || '').slice(0, 4)).filter(Boolean))].sort().reverse();

  const filtered = gigs.filter(g =>
    (g.name || '').toLowerCase().includes(search.toLowerCase()) &&
    (!filmedOnly || g.filmed) &&
    (!stageFilter || gigStatus(g) === stageFilter) &&
    (typeFilters.length === 0 || typeFilters.includes(g.eventType)) &&
    (!year || (g.date || '').startsWith(year)) &&
    (!dateFrom || (g.date && g.date >= dateFrom)) &&
    (!dateTo || (g.date && g.date <= dateTo))
  );
  const sorted = [...filtered].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  const anyFilterActive = filmedOnly || stageFilter || typeFilters.length > 0 || year || dateFrom || dateTo || search;

  const phaseNames = (g, phaseId) => g.phaseTricks[phaseId].map(row => effectiveName(row, tricks.find(x => x.id === row.trickId))).filter(Boolean).join(', ');

  const exportExcel = (chosenProfileId) => {
    const profile = exportProfiles.find(p => p.id === chosenProfileId) || exportProfiles[0];
    const cols = profile?.columns || allColumnsOn();
    const valueFor = (g, key) => {
      switch (key) {
        case 'name': return g.name;
        case 'date': return g.date;
        case 'type': return categories.find(c => c.id === g.eventType)?.label || g.eventType;
        case 'location': return g.location;
        case 'audience': return g.audienceType;
        case 'audienceCount': return g.audienceCount;
        case 'plannedLengthMin': return g.plannedLengthMin;
        case 'venue': return g.venue === 'ute' ? 'Utomhus' : 'Inomhus';
        case 'plannedLength': return PHASES.reduce((s, p) => s + g.phaseTricks[p.id].reduce((s2, row) => s2 + effectiveLength(row, tricks.find(x => x.id === row.trickId)), 0), 0);
        case 'actualLength': return g.actualLengthMin;
        case 'price': return g.priceSEK;
        case 'filmed': return g.filmed ? 'Ja' : 'Nej';
        case 'rating': return g.overallRating;
        case 'opening': return phaseNames(g, 'opening');
        case 'middle': return phaseNames(g, 'middle');
        case 'closing': return phaseNames(g, 'closing');
        case 'comment': return g.overallComment;
        case 'good': return g.good;
        case 'improve': return g.improve;
        case 'customerReview': return g.customerReview;
        case 'customerImprovement': return g.customerImprovement;
        case 'formResponseUrl': return g.formResponseUrl;
        default: return '';
      }
    };
    const activeFields = EXPORT_FIELDS.filter(f => cols[f.key]);
    const rows = sorted.map(g => {
      const row = {};
      activeFields.forEach(f => { row[f.label] = valueFor(g, f.key); });
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Uppdrag');
    XLSX.writeFile(wb, `uppdrag-${profile?.name?.toLowerCase().replace(/\s+/g, '-') || 'export'}.xlsx`);
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        <button onClick={() => setShowFilters(v => !v)} style={{ ...ghostButtonStyle, background: showFilters ? 'var(--card)' : 'transparent' }}>
          <Filter size={14} /> Filter{anyFilterActive ? ' •' : ''}
        </button>
        <button onClick={() => setStageFilter(stageFilter === 'kommande' ? null : 'kommande')}
          style={{ ...chipStyle, background: stageFilter === 'kommande' ? 'var(--accent-gold)' : 'transparent', color: stageFilter === 'kommande' ? 'var(--ink)' : 'var(--text-muted)' }}>Kommande</button>
        <button onClick={() => setStageFilter(stageFilter === 'genomford' ? null : 'genomford')}
          style={{ ...chipStyle, background: stageFilter === 'genomford' ? 'var(--accent-gold)' : 'transparent', color: stageFilter === 'genomford' ? 'var(--ink)' : 'var(--text-muted)' }}>Genomförda</button>
        <div style={{ position: 'relative' }}>
          <button onClick={() => exportProfiles.length > 1 ? setShowExportMenu(v => !v) : exportExcel(exportProfiles[0]?.id)} style={ghostButtonStyle}>
            <Download size={14} /> Exportera (Excel)
          </button>
          {showExportMenu && (
            <div style={{ position: 'absolute', top: '110%', left: 0, background: 'var(--panel)', border: '1px solid var(--border)', zIndex: 10, minWidth: 180 }}>
              {exportProfiles.map(p => (
                <button key={p.id} onClick={() => { exportExcel(p.id); setShowExportMenu(false); }}
                  style={{ display: 'block', width: '100%', textAlign: 'left', padding: '9px 12px', background: 'transparent', border: 'none', color: 'var(--text)', fontSize: 13, cursor: 'pointer' }}>
                  {p.name}
                </button>
              ))}
            </div>
          )}
        </div>
      </div>

      {showFilters && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginBottom: 18 }}>
      <div>
        <div style={labelStyle}>Sök</div>
        <div style={{ position: 'relative', maxWidth: 320 }}>
          <Search size={14} style={{ position: 'absolute', left: 12, top: 12, color: 'var(--text-muted)' }} />
          <input placeholder="Sök uppdrag" value={search} onChange={e => setSearch(e.target.value)} style={{ ...inputStyle, paddingLeft: 34 }} />
        </div>
      </div>
      <div>
        <div style={labelStyle}>Kategori</div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button onClick={() => setTypeFilters([])} style={{ ...chipStyle, background: typeFilters.length === 0 ? 'var(--accent-gold)' : 'transparent', color: typeFilters.length === 0 ? 'var(--ink)' : 'var(--text-muted)' }}>Alla typer</button>
          {categories.map(c => (
            <button key={c.id} onClick={() => toggleTypeFilter(c.id)}
              style={{ ...chipStyle, background: typeFilters.includes(c.id) ? 'var(--accent-gold)' : 'transparent', color: typeFilters.includes(c.id) ? 'var(--ink)' : 'var(--text-muted)' }}>{c.label}</button>
          ))}
        </div>
      </div>
      <div>
        <div style={labelStyle}>Filmad</div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button onClick={() => setFilmedOnly(v => !v)}
            style={{ ...chipStyle, display: 'flex', alignItems: 'center', gap: 6, background: filmedOnly ? '#5a7a94' : 'transparent', color: filmedOnly ? '#fff' : 'var(--text-muted)', borderColor: filmedOnly ? '#5a7a94' : 'var(--border)' }}>
            <Video size={13} /> Filmade
          </button>
        </div>
      </div>
      <div>
        <div style={labelStyle}>Period</div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'flex-end' }}>
        <div>
          <div style={{ ...labelStyle, fontSize: 11 }}>År</div>
          <select value={year} onChange={e => setYear(e.target.value)} style={inputStyle}>
            <option value="">Alla år</option>
            {years.map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        </div>
        <div>
          <div style={{ ...labelStyle, fontSize: 11 }}>Från datum</div>
          <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} style={inputStyle} />
        </div>
        <div>
          <div style={{ ...labelStyle, fontSize: 11 }}>Till datum</div>
          <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} style={inputStyle} />
        </div>
        {anyFilterActive && (
          <button onClick={() => { setSearch(''); setTypeFilters([]); setStageFilter(null); setYear(''); setDateFrom(''); setDateTo(''); setFilmedOnly(false); }} style={ghostButtonStyle}>Rensa filter</button>
        )}
        </div>
      </div>
      </div>
      )}
      {sorted.length > 0 && (
        <div style={{ fontSize: 13, color: 'var(--text-muted)', marginBottom: 14 }}>
          {sorted.length} uppdrag &nbsp;·&nbsp; Totalt pris: <strong style={{ color: 'var(--text)' }}>{sorted.reduce((s, g) => s + (Number(g.priceSEK) || 0), 0).toLocaleString('sv-SE')} kr</strong>
        </div>
      )}
      {sorted.length === 0 && (
        <div style={{ color: 'var(--text-muted)', padding: 30, textAlign: 'center', fontSize: 14 }}>
          {anyFilterActive ? 'Inga uppdrag matchar filtret.' : 'Inga sparade uppdrag ännu. Bygg ett set och spara det så dyker det upp här.'}
        </div>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        {sorted.map(g => {
          const totalLength = PHASES.reduce((sum, p) => sum + g.phaseTricks[p.id].reduce((s, row) => {
            const t = tricks.find(x => x.id === row.trickId);
            return s + effectiveLength(row, t);
          }, 0), 0);
          const status = gigStatus(g);
          return (
            <div key={g.id} style={{ border: '1px solid var(--border)', background: 'var(--panel)', padding: '14px 16px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: 10 }}>
                <div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                    <div style={{ fontFamily: 'var(--font-display)', fontSize: 17, color: 'var(--text)' }}>{g.name || '(namnlöst uppdrag)'}</div>
                    <span style={tagStyle}>{categories.find(c => c.id === g.eventType)?.label || g.eventType}</span>
                    {status && <span style={{ ...tagStyle, color: status === 'kommande' ? '#7a9b6e' : 'var(--text-muted)', borderColor: status === 'kommande' ? '#7a9b6e' : 'var(--border)' }}>{GIG_STATUS_LABEL[status]}</span>}
                    {g.filmed && <span style={{ ...tagStyle, display: 'flex', alignItems: 'center', gap: 4, color: '#5a7a94', borderColor: '#5a7a94' }}><Video size={11} /> Filmad</span>}
                  </div>
                  <div style={{ display: 'flex', gap: 14, marginTop: 8, fontSize: 12.5, color: 'var(--text-muted)', flexWrap: 'wrap' }}>
                    {g.date && <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><Calendar size={12} /> {g.date}</span>}
                    {g.location && <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><MapPin size={12} /> {g.location}</span>}
                    {g.audienceType && <span>{g.audienceType}{g.audienceCount ? ` (${g.audienceCount})` : ''}</span>}
                    {g.plannedLengthMin !== '' && g.plannedLengthMin != null && <span>Planerad: {g.plannedLengthMin} min</span>}
                    {g.actualLengthMin !== '' && g.actualLengthMin != null
                      ? <span>Verklig: {g.actualLengthMin} min</span>
                      : <span>Bedömd: {totalLength} min</span>}
                    {g.priceSEK !== '' && g.priceSEK != null && <span>{g.priceSEK} kr</span>}
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <button onClick={() => onPrint(g)} style={ghostButtonStyle}><Printer size={13} /> Skriv ut</button>
                  <button onClick={() => onOpen(g)} style={ghostButtonStyle}>Öppna</button>
                  <button onClick={() => setGigs(gs => gs.filter(x => x.id !== g.id))} style={iconButtonStyle}><Trash2 size={13} /></button>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function buildPrintHtml(gig, tricks) {
  const meta = [gig.date, gig.location, gig.audienceType ? `${gig.audienceType}${gig.audienceCount ? ` (${gig.audienceCount})` : ''}` : '', gig.venue === 'ute' ? 'Utomhus' : 'Inomhus']
    .filter(Boolean).map(escapeHtml).join(' &nbsp;·&nbsp; ');

  let runningNumber = 0;
  const phaseHtml = PHASES.map(phase => {
    const rows = gig.phaseTricks[phase.id];
    if (!rows.length) return '';
    const items = rows.map(row => {
      runningNumber += 1;
      const t = tricks.find(x => x.id === row.trickId);
      const name = escapeHtml(effectiveName(row, t));
      const len = ` <span class="len">${effectiveLength(row, t)} min</span>`;
      const m = effectiveMusic(row, t);
      const music = m && m.song ? `<div class="song">${escapeHtml(m.song)}${m.artist ? ` — ${escapeHtml(m.artist)}` : ''}</div>` : '';
      const note = row.variantNote ? `<div class="song">${escapeHtml(row.variantNote)}</div>` : '';
      return `<li value="${runningNumber}">${name}${len}${music}${note}</li>`;
    }).join('');
    return `<h2>${escapeHtml(phase.label)}</h2><ol>${items}</ol>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="sv"><head><meta charset="utf-8"><title>${escapeHtml(gig.name || 'Uppdrag')}</title>
<style>
  @page { size: A5; margin: 14mm; }
  * { box-sizing: border-box; }
  body { font-family: Georgia, 'Iowan Old Style', serif; color: #1a1a1a; margin: 0; padding: 18px; }
  h1 { font-size: 15px; font-weight: normal; color: #555; margin: 0; font-family: Arial, sans-serif; }
  .meta { font-size: 10.5px; color: #777; margin: 3px 0 20px; font-family: Arial, sans-serif; }
  h2 { font-size: 13px; text-transform: uppercase; letter-spacing: 0.08em; color: #555; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin: 22px 0 8px; font-family: Arial, sans-serif; }
  h2:first-of-type { margin-top: 0; }
  ol { margin: 0; padding-left: 24px; }
  li { font-size: 19px; line-height: 2.2; }
  .len { font-size: 12px; color: #888; font-family: Arial, sans-serif; }
  .song { font-size: 12px; color: #888; font-family: Arial, sans-serif; margin-top: -4px; line-height: 1.3; }
  .note { margin-top: 30px; font-size: 10px; color: #aaa; font-family: Arial, sans-serif; }
</style></head>
<body>
  <h1>${escapeHtml(gig.name || 'Uppdrag')}</h1>
  <div class="meta">${meta}</div>
  ${phaseHtml}
  <div class="note">A5. Välj A4 i skrivardialogen om du vill.</div>
  <script>window.onload = function() { window.print(); };</script>
</body></html>`;
}

function openPrintWindow(gig, tricks) {
  const w = window.open('', '_blank');
  if (!w) {
    alert('Webbläsaren blockerade den nya fliken. Tillåt popup-fönster för den här sidan och försök igen.');
    return;
  }
  w.document.open();
  w.document.write(buildPrintHtml(gig, tricks));
  w.document.close();
}

const emptyDevItem = () => ({ id: null, title: '', trickId: '', notes: '', progress: 0, comment: '', done: false });

function DevelopmentView({ devItems, setDevItems, tricks }) {
  const [adding, setAdding] = useState(false);
  const [editing, setEditing] = useState(null);
  const [showDone, setShowDone] = useState(false);
  const dragSource = useRef(null);

  const active = devItems.filter(d => !d.done);
  const done = devItems.filter(d => d.done);

  const save = (form) => {
    if (form.id) {
      setDevItems(ds => ds.map(d => d.id === form.id ? form : d));
    } else {
      setDevItems(ds => [...ds, { ...form, id: crypto.randomUUID() }]);
    }
    setEditing(null);
    setAdding(false);
  };

  const updateItem = (id, patch) => setDevItems(ds => ds.map(d => d.id === id ? { ...d, ...patch } : d));
  const removeItem = (id) => setDevItems(ds => ds.filter(d => d.id !== id));
  const markDone = (id) => updateItem(id, { done: true });
  const reopen = (id) => setDevItems(ds => [...ds.filter(d => d.id !== id), { ...ds.find(d => d.id === id), done: false }]);

  const handleDragStart = (id) => () => { dragSource.current = id; };
  const handleDrop = (targetId) => (e) => {
    e.preventDefault();
    const fromId = dragSource.current;
    if (!fromId || fromId === targetId) return;
    setDevItems(ds => {
      const activeIds = ds.filter(d => !d.done).map(d => d.id);
      const doneOnes = ds.filter(d => d.done);
      const fromIdx = activeIds.indexOf(fromId);
      const toIdx = activeIds.indexOf(targetId);
      activeIds.splice(fromIdx, 1);
      activeIds.splice(toIdx, 0, fromId);
      const byId = Object.fromEntries(ds.map(d => [d.id, d]));
      return [...activeIds.map(id => byId[id]), ...doneOnes];
    });
    dragSource.current = null;
  };

  const labelFor = (idx) => idx === 0 ? 'Idag' : idx === 1 ? 'Näst på tur' : null;

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)' }}>Dra i handtaget för att ändra prioritet. Överst är det du gör idag.</div>
        <button onClick={() => { setAdding(true); setEditing(null); }} style={primaryButtonStyle}><Plus size={14} /> Nytt mål</button>
      </div>

      {(adding || editing) && (
        <div style={{ marginBottom: 18 }}>
          <DevItemForm initial={editing} tricks={tricks} onSave={save} onCancel={() => { setEditing(null); setAdding(false); }} />
        </div>
      )}

      {active.length === 0 && !adding && (
        <div style={{ color: 'var(--text-muted)', padding: 30, textAlign: 'center', fontSize: 14 }}>Inget att utveckla just nu. Lägg till ett mål, kopplat till ett trick eller helt eget.</div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        {active.map((item, idx) => {
          const trick = tricks.find(t => t.id === item.trickId);
          const label = labelFor(idx);
          return (
            <div key={item.id} draggable onDragStart={handleDragStart(item.id)} onDragOver={(e) => e.preventDefault()} onDrop={handleDrop(item.id)}
              style={{ border: '1px solid var(--border)', background: 'var(--panel)', padding: '14px 16px', cursor: 'grab' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 10 }}>
                <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, flex: 1 }}>
                  <GripVertical size={15} style={{ color: 'var(--text-muted)', flexShrink: 0, marginTop: 3 }} />
                  <div style={{ flex: 1 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                      {label && <span style={{ ...tagStyle, background: 'var(--accent-gold)', color: 'var(--ink)', border: 'none' }}>{label}</span>}
                      <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)' }}>{item.title || trick?.name || '(namnlöst mål)'}</div>
                      {trick && <span style={tagStyle}>{trick.name}</span>}
                    </div>
                    {item.notes && <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginTop: 4 }}>{item.notes}</div>}
                    <div style={{ marginTop: 10, maxWidth: 320 }}>
                      <ReadinessBar value={item.progress} onChange={v => updateItem(item.id, { progress: v })} compact />
                    </div>
                    <input placeholder="Kommentar om läget" value={item.comment} onChange={e => updateItem(item.id, { comment: e.target.value })}
                      style={{ ...inputStyle, marginTop: 8, fontSize: 12.5, padding: '6px 10px' }} />
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                  <button onClick={() => { setEditing(item); setAdding(false); }} style={iconButtonStyle}>✎</button>
                  <button onClick={() => markDone(item.id)} style={{ ...ghostButtonStyle, padding: '5px 10px', fontSize: 12 }}>Klart</button>
                  <button onClick={() => removeItem(item.id)} style={iconButtonStyle}><Trash2 size={13} /></button>
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {done.length > 0 && (
        <div style={{ marginTop: 22 }}>
          <button onClick={() => setShowDone(v => !v)} style={ghostButtonStyle}>{showDone ? 'Dölj' : 'Visa'} avklarat ({done.length})</button>
          {showDone && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 12 }}>
              {done.map(item => {
                const trick = tricks.find(t => t.id === item.trickId);
                return (
                  <div key={item.id} style={{ border: '1px solid var(--border)', background: 'var(--panel)', padding: '10px 14px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: 13.5, color: 'var(--text-muted)' }}>{item.title || trick?.name || '(namnlöst mål)'}</span>
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button onClick={() => reopen(item.id)} style={ghostButtonStyle}>Återuppta</button>
                      <button onClick={() => removeItem(item.id)} style={iconButtonStyle}><Trash2 size={13} /></button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function DevItemForm({ initial, tricks, onSave, onCancel }) {
  const [form, setForm] = useState({ ...emptyDevItem(), ...(initial || {}) });
  return (
    <div style={{ border: '1px solid var(--border)', background: 'var(--panel)', padding: 20, display: 'flex', flexDirection: 'column', gap: 14 }}>
      <div>
        <div style={labelStyle}>Länka till trick från Bibliotek — valfritt</div>
        <select value={form.trickId} onChange={e => setForm(f => ({ ...f, trickId: e.target.value }))} style={inputStyle}>
          <option value="">Inget — eget mål</option>
          {tricks.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
        </select>
      </div>
      <div>
        <div style={labelStyle}>Vad specifikt ska utvecklas? — lämna tomt för att bara visa tricknamnet</div>
        <input placeholder="T.ex. Förbättra presentationen, eller Skriv patter" value={form.title} onChange={e => setForm(f => ({ ...f, title: e.target.value }))} style={inputStyle} />
      </div>
      <textarea rows={2} placeholder="Anteckning — vad ska göras, varför" value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
      <div>
        <div style={labelStyle}>Hur långt kommen (0–10)</div>
        <ReadinessBar value={form.progress} onChange={v => setForm(f => ({ ...f, progress: v }))} />
      </div>
      <input placeholder="Kommentar om läget — valfritt" value={form.comment} onChange={e => setForm(f => ({ ...f, comment: e.target.value }))} style={inputStyle} />
      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
        <button onClick={onCancel} style={ghostButtonStyle}>Avbryt</button>
        <button onClick={() => (form.title.trim() || form.trickId) && onSave(form)} style={primaryButtonStyle}><Save size={14} /> Spara</button>
      </div>
    </div>
  );
}

function SettingsView({ categories, setCategories, genres, setGenres, formats, setFormats, settings, setSettings, onResetGigs, onResetMedia, backupData, onImportBackup }) {
  const updateLabel = (id, label) => setCategories(cs => cs.map(c => c.id === id ? { ...c, label } : c));
  const removeCategory = (id) => setCategories(cs => cs.filter(c => c.id !== id));
  const addCategory = () => setCategories(cs => [...cs, { id: crypto.randomUUID(), label: '' }]);

  const updateGenreLabel = (id, label) => setGenres(gs => gs.map(g => g.id === id ? { ...g, label } : g));
  const removeGenre = (id) => setGenres(gs => gs.filter(g => g.id !== id));
  const addGenre = () => setGenres(gs => [...gs, { id: crypto.randomUUID(), label: '' }]);
  const sortGenres = () => setGenres(gs => [...gs].sort((a, b) => a.label.localeCompare(b.label, 'sv')));
  const genreDragSource = useRef(null);
  const handleGenreDragStart = (index) => () => { genreDragSource.current = index; };
  const handleGenreDrop = (index) => (e) => {
    e.preventDefault();
    const from = genreDragSource.current;
    if (from === null || from === index) return;
    setGenres(gs => {
      const next = [...gs];
      const [moved] = next.splice(from, 1);
      next.splice(index, 0, moved);
      return next;
    });
    genreDragSource.current = null;
  };

  const updateFormat = (type, index, value) => setFormats(f => ({ ...f, [type]: f[type].map((v, i) => i === index ? value : v) }));
  const removeFormat = (type, index) => setFormats(f => ({ ...f, [type]: f[type].filter((_, i) => i !== index) }));
  const addFormat = (type) => setFormats(f => ({ ...f, [type]: [...f[type], ''] }));

  const updateProfileName = (id, name) => setSettings(s => ({ ...s, exportProfiles: s.exportProfiles.map(p => p.id === id ? { ...p, name } : p) }));
  const toggleProfileColumn = (id, key) => setSettings(s => ({ ...s, exportProfiles: s.exportProfiles.map(p => p.id === id ? { ...p, columns: { ...p.columns, [key]: !p.columns[key] } } : p) }));
  const removeProfile = (id) => setSettings(s => ({ ...s, exportProfiles: s.exportProfiles.filter(p => p.id !== id) }));
  const addProfile = () => setSettings(s => ({ ...s, exportProfiles: [...s.exportProfiles, { id: crypto.randomUUID(), name: 'Ny profil', columns: allColumnsOn() }] }));

  const exportBackup = () => {
    const payload = { version: 1, exportedAt: new Date().toISOString(), ...backupData };
    const blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `showbasen-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const importInputRef = useRef(null);
  const importModeRef = useRef('replace');
  const handleImportFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const mode = importModeRef.current;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        if (!data || (!data.tricks && !data.gigs && !data.media)) {
          alert('Filen ser inte ut som en backup från den här appen.');
          return;
        }
        if (mode === 'replace') {
          if (window.confirm('Detta ersätter all nuvarande data med innehållet i filen. Fortsätt?')) {
            onImportBackup(data, 'replace');
          }
        } else {
          if (window.confirm('Lägger till trick, uppdrag och titlar från filen till din nuvarande data. Fortsätt?')) {
            onImportBackup(data, 'merge');
          }
        }
      } catch (err) {
        alert('Kunde inte läsa filen. Är det en giltig backup-fil?');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };
  const startImport = (mode) => { importModeRef.current = mode; importInputRef.current?.click(); };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      <div style={panelBox}>
        <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)', marginBottom: 4 }}>Kategorier / uppdragstyper</div>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>Dessa styr taggarna på trick och typ av uppdrag i hela appen. Byt namn, ta bort eller lägg till egna.</div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {categories.map(c => (
            <div key={c.id} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input value={c.label} onChange={e => updateLabel(c.id, e.target.value)} style={inputStyle} placeholder="Namn på kategori" />
              <button onClick={() => removeCategory(c.id)} style={iconButtonStyle}><Trash2 size={13} /></button>
            </div>
          ))}
        </div>
        <button onClick={addCategory} style={{ ...ghostButtonStyle, marginTop: 12 }}><Plus size={13} /> Lägg till kategori</button>
        {categories.length === 0 && <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginTop: 10 }}>Inga kategorier kvar. Trick och uppdrag går fortfarande att spara, men filtreringen på typ blir tom tills du lägger till minst en.</div>}
      </div>

      <div style={panelBox}>
        <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)', marginBottom: 4 }}>Genrer</div>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>Styr genre-knapparna i Film & böcker, i den ordning du ser dem här. Byt namn, ta bort, lägg till, eller dra i handtaget för att ändra ordning.</div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {genres.map((g, i) => (
            <div key={g.id} draggable onDragStart={handleGenreDragStart(i)} onDragOver={(e) => e.preventDefault()} onDrop={handleGenreDrop(i)}
              style={{ display: 'flex', gap: 8, alignItems: 'center', cursor: 'grab' }}>
              <GripVertical size={14} style={{ color: 'var(--text-muted)', flexShrink: 0 }} />
              <input value={g.label} onChange={e => updateGenreLabel(g.id, e.target.value)} style={inputStyle} placeholder="Namn på genre" />
              <button onClick={() => removeGenre(g.id)} style={iconButtonStyle}><Trash2 size={13} /></button>
            </div>
          ))}
        </div>
        <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
          <button onClick={addGenre} style={ghostButtonStyle}><Plus size={13} /> Lägg till genre</button>
          {genres.length > 1 && <button onClick={sortGenres} style={ghostButtonStyle}>Sortera A-Ö</button>}
        </div>
        {genres.length === 0 && <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginTop: 10 }}>Inga genrer kvar. Titlar går fortfarande att spara, men genre-knapparna blir tomma tills du lägger till minst en.</div>}
      </div>

      <div style={panelBox}>
        <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)', marginBottom: 4 }}>Format</div>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>Styr format-listan i Film & böcker, separat för Film och Bok. Byt namn, ta bort eller lägg till egna.</div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 20 }}>
          {[{ key: 'film', label: 'Filmformat' }, { key: 'bok', label: 'Bokformat' }].map(({ key, label }) => (
            <div key={key}>
              <div style={labelStyle}>{label}</div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {formats[key].map((val, i) => (
                  <div key={i} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <input value={val} onChange={e => updateFormat(key, i, e.target.value)} style={inputStyle} placeholder="Namn på format" />
                    <button onClick={() => removeFormat(key, i)} style={iconButtonStyle}><Trash2 size={13} /></button>
                  </div>
                ))}
              </div>
              <button onClick={() => addFormat(key)} style={{ ...ghostButtonStyle, marginTop: 10 }}><Plus size={13} /> Lägg till</button>
            </div>
          ))}
        </div>
      </div>

      <div style={panelBox}>
        <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)', marginBottom: 4 }}>Utseende</div>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>Byt mellan ljust och mörkt tema när som helst.</div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={() => setSettings(s => ({ ...s, theme: 'light' }))}
            style={{ ...chipStyle, background: settings.theme === 'light' ? 'var(--accent-gold)' : 'transparent', color: settings.theme === 'light' ? 'var(--ink)' : 'var(--text-muted)' }}>Ljust</button>
          <button onClick={() => setSettings(s => ({ ...s, theme: 'dark' }))}
            style={{ ...chipStyle, background: settings.theme === 'dark' ? 'var(--accent-gold)' : 'transparent', color: settings.theme === 'dark' ? 'var(--ink)' : 'var(--text-muted)' }}>Mörkt</button>
        </div>
      </div>

      <div style={panelBox}>
        <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)', marginBottom: 4 }}>Standardvärden för nya set</div>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>Förifyllt när du trycker "Bygg set" — mest till för det du bokar oftast, ändra alltid i efterhand.</div>
        <div style={fieldGrid}>
          <div>
            <div style={labelStyle}>Uppdragstyp</div>
            <select value={settings.defaultEventType} onChange={e => setSettings(s => ({ ...s, defaultEventType: e.target.value }))} style={inputStyle}>
              <option value="">Ingen förvald</option>
              {categories.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
            </select>
          </div>
          <div>
            <div style={labelStyle}>Scen</div>
            <select value={settings.defaultVenue} onChange={e => setSettings(s => ({ ...s, defaultVenue: e.target.value }))} style={inputStyle}>
              <option value="inne">Inomhus</option>
              <option value="ute">Utomhus</option>
            </select>
          </div>
        </div>
      </div>

      <div style={panelBox}>
        <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)', marginBottom: 4 }}>Backup</div>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>
          Exportera all data (trick, uppdrag, film och böcker, kategorier, genrer, inställningar) som en fil, t.ex. innan du delar appen med någon eller byter version. "Ersätt allt" byter ut all nuvarande data mot filen. "Lägg till" lägger filens trick, uppdrag och titlar till det du redan har, t.ex. bra för att föra in gamla pappersuppdrag utan att röra det som redan finns.
        </div>
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          <button onClick={exportBackup} style={ghostButtonStyle}><Download size={14} /> Exportera backup</button>
          <button onClick={() => startImport('replace')} style={ghostButtonStyle}><Plus size={14} /> Importera (ersätt allt)</button>
          <button onClick={() => startImport('merge')} style={ghostButtonStyle}><Plus size={14} /> Importera (lägg till)</button>
          <input ref={importInputRef} type="file" accept="application/json" onChange={handleImportFile} style={{ display: 'none' }} />
        </div>
      </div>

      <div style={panelBox}>
        <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--text)', marginBottom: 4 }}>Exportprofiler</div>
        <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>Skapa olika kolumnval för olika syften, t.ex. en enkel för bokföring och en fullständig för egen uppföljning. Väljs vid export i Uppdrag.</div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          {settings.exportProfiles.map(p => (
            <div key={p.id} style={{ border: '1px solid var(--border)', background: 'var(--card)', padding: 14 }}>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
                <input value={p.name} onChange={e => updateProfileName(p.id, e.target.value)} style={{ ...inputStyle, flex: 1 }} />
                {settings.exportProfiles.length > 1 && <button onClick={() => removeProfile(p.id)} style={iconButtonStyle}><Trash2 size={13} /></button>}
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                {EXPORT_FIELDS.map(f => (
                  <label key={f.key} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 12, color: 'var(--text-muted)', cursor: 'pointer' }}>
                    <input type="checkbox" checked={!!p.columns[f.key]} onChange={() => toggleProfileColumn(p.id, f.key)} />
                    {f.label}
                  </label>
                ))}
              </div>
            </div>
          ))}
        </div>
        <button onClick={addProfile} style={{ ...ghostButtonStyle, marginTop: 12 }}><Plus size={13} /> Ny exportprofil</button>
      </div>

      <DangerReset
        title="Rensa uppdrag"
        description="Tar bort alla sparade uppdrag (set-listor, betyg, kundenkätsvar). Trickbibliotek, kategorier och film/böcker påverkas inte."
        phrase="RADERA UPPDRAG"
        confirmLabel="Radera uppdrag nu"
        onConfirm={onResetGigs}
      />

      <DangerReset
        title="Rensa film & böcker"
        description="Tar bort alla sparade filmer och böcker. Resten av appen påverkas inte."
        phrase="RADERA FILM"
        confirmLabel="Radera film & böcker nu"
        onConfirm={onResetMedia}
      />
    </div>
  );
}

function DangerReset({ title, description, phrase, confirmLabel, onConfirm }) {
  const [step, setStep] = useState(0);
  const [text, setText] = useState('');
  const cancel = () => { setStep(0); setText(''); };
  const confirm = () => { onConfirm(); setStep(0); setText(''); };

  return (
    <div style={{ ...panelBox, borderColor: 'var(--accent-red)' }}>
      <div style={{ fontFamily: 'var(--font-display)', fontSize: 16, color: 'var(--accent-red)', marginBottom: 4 }}>{title}</div>
      <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginBottom: 16 }}>{description} Går inte att ångra. Tre steg med flit, så det inte händer av misstag.</div>

      {step === 0 && (
        <button onClick={() => setStep(1)} style={{ ...ghostButtonStyle, borderColor: 'var(--accent-red)', color: 'var(--accent-red)' }}>
          <Trash2 size={13} /> {title}
        </button>
      )}

      {step === 1 && (
        <div style={{ border: '1px solid var(--accent-red)', background: 'var(--card)', padding: 14 }}>
          <div style={{ fontSize: 13.5, marginBottom: 12 }}>Steg 1 av 3: Detta raderar datan permanent. Är du säker?</div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={cancel} style={ghostButtonStyle}>Avbryt</button>
            <button onClick={() => setStep(2)} style={{ ...ghostButtonStyle, borderColor: 'var(--accent-red)', color: 'var(--accent-red)' }}>Fortsätt</button>
          </div>
        </div>
      )}

      {step === 2 && (
        <div style={{ border: '1px solid var(--accent-red)', background: 'var(--card)', padding: 14 }}>
          <div style={{ fontSize: 13.5, marginBottom: 10 }}>Steg 2 av 3: Skriv <strong>{phrase}</strong> nedan för att bekräfta.</div>
          <input value={text} onChange={e => setText(e.target.value)} style={inputStyle} placeholder={phrase} />
          <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
            <button onClick={cancel} style={ghostButtonStyle}>Avbryt</button>
            <button onClick={() => text.trim() === phrase && setStep(3)} disabled={text.trim() !== phrase}
              style={{ ...ghostButtonStyle, borderColor: 'var(--accent-red)', color: 'var(--accent-red)', opacity: text.trim() === phrase ? 1 : 0.4, cursor: text.trim() === phrase ? 'pointer' : 'not-allowed' }}>
              Fortsätt
            </button>
          </div>
        </div>
      )}

      {step === 3 && (
        <div style={{ border: '1px solid var(--accent-red)', background: 'var(--card)', padding: 14 }}>
          <div style={{ fontSize: 13.5, marginBottom: 12 }}>Steg 3 av 3: Sista chansen. Detta går inte att ångra.</div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={cancel} style={ghostButtonStyle}>Avbryt</button>
            <button onClick={confirm} style={{ ...primaryButtonStyle, background: 'var(--accent-red)' }}>{confirmLabel}</button>
          </div>
        </div>
      )}
    </div>
  );
}

const emptyMedia = () => ({ id: null, title: '', type: 'film', creator: '', genres: [], format: '', status: 'ovald', rating: '', lentTo: '', synopsis: '', notes: '' });
const MEDIA_STATUS = {
  ovald: { label: 'Ej sedd/läst', color: '#8a7f6a' },
  pagar: { label: 'Pågår', color: '#b8935a' },
  klar: { label: 'Sedd/läst', color: '#7a9b6e' },
};
const DEFAULT_FORMATS = {
  film: ['DVD', 'Blu-ray', 'Digital/Nedladdning', 'Streaming'],
  bok: ['Bok (fysisk)', 'E-bok', 'Ljudbok', 'Lecture notes', 'Häfte'],
};

function MediaForm({ initial, onSave, onCancel, genres, formats }) {
  const [form, setForm] = useState({ ...emptyMedia(), ...(initial || {}) });
  const toggleGenre = (id) => setForm(f => ({ ...f, genres: f.genres.includes(id) ? f.genres.filter(x => x !== id) : [...f.genres, id] }));
  return (
    <div style={{ border: '1px solid var(--border)', background: 'var(--panel)', padding: 20, display: 'flex', flexDirection: 'column', gap: 14 }}>
      <input placeholder="Titel" value={form.title} onChange={e => setForm(f => ({ ...f, title: e.target.value }))} style={inputStyle} />
      <input placeholder="Kort om vad den handlar om — valfritt" value={form.synopsis} onChange={e => setForm(f => ({ ...f, synopsis: e.target.value }))} style={inputStyle} />
      <div style={{ display: 'flex', gap: 14 }}>
        <div style={{ flex: 1 }}>
          <div style={labelStyle}>Typ</div>
          <select value={form.type} onChange={e => setForm(f => ({ ...f, type: e.target.value, format: '' }))} style={inputStyle}>
            <option value="film">Film</option>
            <option value="bok">Bok</option>
          </select>
        </div>
        <div style={{ flex: 1 }}>
          <div style={labelStyle}>Status</div>
          <select value={form.status} onChange={e => setForm(f => ({ ...f, status: e.target.value }))} style={inputStyle}>
            {Object.entries(MEDIA_STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
          </select>
        </div>
      </div>
      <input placeholder={form.type === 'bok' ? 'Författare' : 'Regissör'} value={form.creator} onChange={e => setForm(f => ({ ...f, creator: e.target.value }))} style={inputStyle} />
      <div>
        <div style={labelStyle}>Genre</div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {genres.length === 0 && <div style={{ fontSize: 12.5, color: 'var(--text-muted)' }}>Inga genrer ännu — lägg till under Inställningar.</div>}
          {genres.map(g => (
            <button key={g.id} type="button" onClick={() => toggleGenre(g.id)}
              style={{ ...chipStyle, background: form.genres.includes(g.id) ? 'var(--accent-gold)' : 'transparent', color: form.genres.includes(g.id) ? 'var(--ink)' : 'var(--text-muted)', borderColor: form.genres.includes(g.id) ? 'var(--accent-gold)' : 'var(--border)' }}>
              {g.label}
            </button>
          ))}
        </div>
      </div>
      <div>
        <div style={labelStyle}>Format</div>
        <select value={form.format} onChange={e => setForm(f => ({ ...f, format: e.target.value }))} style={inputStyle}>
          <option value="">Välj format</option>
          {formats[form.type].map(f => <option key={f} value={f}>{f}</option>)}
        </select>
      </div>
      <input placeholder="Lånad ut till — valfritt" value={form.lentTo} onChange={e => setForm(f => ({ ...f, lentTo: e.target.value }))} style={inputStyle} />
      <div>
        <div style={labelStyle}>Betyg (0–5)</div>
        <select value={form.rating} onChange={e => setForm(f => ({ ...f, rating: e.target.value }))} style={inputStyle}>
          <option value="">–</option>
          {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}/5</option>)}
        </select>
      </div>
      <textarea rows={2} placeholder="Anteckningar" value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} style={{ ...inputStyle, resize: 'vertical', fontFamily: 'inherit' }} />
      <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
        <button onClick={onCancel} style={ghostButtonStyle}>Avbryt</button>
        <button onClick={() => form.title.trim() && onSave(form)} style={primaryButtonStyle}><Save size={14} /> Spara</button>
      </div>
    </div>
  );
}

function MediaView({ media, setMedia, genres, formats, tricks }) {
  const [showFilters, setShowFilters] = useState(false);
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState(null);
  const [genreFilters, setGenreFilters] = useState([]);
  const [formatFilter, setFormatFilter] = useState('');
  const [creatorFilter, setCreatorFilter] = useState('');
  const [editing, setEditing] = useState(null);
  const [adding, setAdding] = useState(false);

  const toggleGenreFilter = (id) => setGenreFilters(gf => gf.includes(id) ? gf.filter(x => x !== id) : [...gf, id]);

  const relevantFormats = typeFilter ? formats[typeFilter] : [...new Set([...formats.film, ...formats.bok])];
  const anyFilterActive = search || typeFilter || genreFilters.length > 0 || formatFilter || creatorFilter;

  const relevantCreators = [...new Set(
    media.filter(m => !typeFilter || m.type === typeFilter).map(m => m.creator).filter(Boolean)
  )].sort((a, b) => a.localeCompare(b, 'sv'));

  useEffect(() => {
    if (formatFilter && !relevantFormats.includes(formatFilter)) setFormatFilter('');
  }, [typeFilter]);

  useEffect(() => {
    if (creatorFilter && !relevantCreators.includes(creatorFilter)) setCreatorFilter('');
  }, [typeFilter]);

  const filtered = media.filter(m =>
    (m.title.toLowerCase().includes(search.toLowerCase()) || (m.creator || '').toLowerCase().includes(search.toLowerCase())) &&
    (!typeFilter || m.type === typeFilter) &&
    (genreFilters.length === 0 || (m.genres || []).some(g => genreFilters.includes(g))) &&
    (!formatFilter || m.format === formatFilter) &&
    (!creatorFilter || m.creator === creatorFilter)
  );

  const save = (form) => {
    if (editing && editing.id) {
      setMedia(ms => ms.map(m => m.id === editing.id ? { ...form, id: editing.id } : m));
    } else {
      setMedia(ms => [...ms, { ...form, id: crypto.randomUUID() }]);
    }
    setEditing(null);
    setAdding(false);
  };

  const exportExcel = () => {
    const rows = filtered.map(m => ({
      Titel: m.title,
      Typ: m.type === 'bok' ? 'Bok' : 'Film',
      'Regissör/Författare': m.creator,
      Genre: (m.genres || []).map(gid => genres.find(g => g.id === gid)?.label).filter(Boolean).join(', '),
      Format: m.format,
      Status: MEDIA_STATUS[m.status]?.label || m.status,
      Betyg: m.rating,
      'Lånad ut till': m.lentTo,
      'Om den handlar om': m.synopsis,
      Anteckningar: m.notes,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Film & böcker');
    XLSX.writeFile(wb, 'film-och-bocker.xlsx');
  };

  return (
    <div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        <button onClick={() => setShowFilters(v => !v)} style={{ ...ghostButtonStyle, background: showFilters ? 'var(--card)' : 'transparent' }}>
          <Filter size={14} /> Filter{(search || genreFilters.length > 0 || formatFilter || creatorFilter) ? ' •' : ''}
        </button>
        <button onClick={() => setTypeFilter(null)} style={{ ...chipStyle, background: !typeFilter ? 'var(--accent-gold)' : 'transparent', color: !typeFilter ? 'var(--ink)' : 'var(--text-muted)' }}>Alla</button>
        <button onClick={() => setTypeFilter('film')} style={{ ...chipStyle, display: 'flex', alignItems: 'center', gap: 6, background: typeFilter === 'film' ? 'var(--accent-gold)' : 'transparent', color: typeFilter === 'film' ? 'var(--ink)' : 'var(--text-muted)' }}><Film size={13} /> Film</button>
        <button onClick={() => setTypeFilter('bok')} style={{ ...chipStyle, display: 'flex', alignItems: 'center', gap: 6, background: typeFilter === 'bok' ? 'var(--accent-gold)' : 'transparent', color: typeFilter === 'bok' ? 'var(--ink)' : 'var(--text-muted)' }}><BookOpen size={13} /> Bok</button>
        <button onClick={exportExcel} style={ghostButtonStyle}><Download size={14} /> Exportera (Excel)</button>
        <button onClick={() => { setAdding(true); setEditing({}); }} style={primaryButtonStyle}><Plus size={14} /> Ny titel</button>
      </div>

      {showFilters && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14, marginBottom: 18 }}>
          <div>
            <div style={labelStyle}>Sök</div>
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
            <div style={{ position: 'relative', flex: '1 1 200px' }}>
              <Search size={14} style={{ position: 'absolute', left: 12, top: 12, color: 'var(--text-muted)' }} />
              <input placeholder="Sök titel eller regissör/författare" value={search} onChange={e => setSearch(e.target.value)} style={{ ...inputStyle, paddingLeft: 34 }} />
            </div>
            <select value={creatorFilter} onChange={e => setCreatorFilter(e.target.value)} style={{ ...inputStyle, width: 'auto' }}>
              <option value="">Alla regissörer/författare</option>
              {relevantCreators.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            </div>
          </div>

          <div>
            <div style={labelStyle}>Format</div>
            <select value={formatFilter} onChange={e => setFormatFilter(e.target.value)} style={{ ...inputStyle, width: 'auto' }}>
              <option value="">Alla format</option>
              {relevantFormats.map(f => <option key={f} value={f}>{f}</option>)}
            </select>
          </div>

          <div>
            <div style={labelStyle}>Genre</div>
            <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
            {genres.length > 0 && (
              <>
                <button onClick={() => setGenreFilters([])} style={{ ...chipStyle, background: genreFilters.length === 0 ? 'var(--accent-gold)' : 'transparent', color: genreFilters.length === 0 ? 'var(--ink)' : 'var(--text-muted)' }}>Alla genrer</button>
                {genres.map(g => (
                  <button key={g.id} onClick={() => toggleGenreFilter(g.id)}
                    style={{ ...chipStyle, background: genreFilters.includes(g.id) ? 'var(--accent-gold)' : 'transparent', color: genreFilters.includes(g.id) ? 'var(--ink)' : 'var(--text-muted)' }}>{g.label}</button>
                ))}
              </>
            )}
            {anyFilterActive && (
              <button onClick={() => { setSearch(''); setTypeFilter(null); setGenreFilters([]); setFormatFilter(''); setCreatorFilter(''); }} style={ghostButtonStyle}>Rensa filter</button>
            )}
            </div>
          </div>
        </div>
      )}

      {(adding || editing) && (
        <div style={{ marginBottom: 18 }}>
          <MediaForm initial={editing?.id ? editing : null} onSave={save} onCancel={() => { setEditing(null); setAdding(false); }} genres={genres} formats={formats} />
        </div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {filtered.map(m => (
          <MediaCard key={m.id} media={m} genres={genres} tricks={tricks}
            onEdit={(mm) => { setEditing(mm); setAdding(false); }}
            onDelete={(id) => setMedia(ms => ms.filter(x => x.id !== id))} />
        ))}
      </div>
      {filtered.length === 0 && <div style={{ color: 'var(--text-muted)', padding: 30, textAlign: 'center', fontSize: 14 }}>Inget hittades. Lägg till din första titel.</div>}
    </div>
  );
}

function MediaCard({ media: m, genres, tricks, onEdit, onDelete }) {
  const [expanded, setExpanded] = useState(false);
  const referencingTricks = tricks.filter(t => (t.mediaSources || []).some(s => s.mediaId === m.id));
  return (
    <div style={cardStyle}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 12 }}>
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 7, minWidth: 0, flex: 1 }}>
          {m.type === 'bok' ? <BookOpen size={14} style={{ color: 'var(--text-muted)', flexShrink: 0, marginTop: 3 }} /> : <Film size={14} style={{ color: 'var(--text-muted)', flexShrink: 0, marginTop: 3 }} />}
          <div style={{ minWidth: 0 }}>
            <div style={{ fontFamily: 'var(--font-display)', fontSize: 17, color: 'var(--text)' }}>{m.title}</div>
            {m.creator && <div style={{ fontSize: 12.5, color: 'var(--text-muted)', marginTop: 2 }}>{m.creator}</div>}
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, fontSize: 12.5, color: 'var(--text-muted)', flexShrink: 0 }}>
          <span style={{ display: 'flex', alignItems: 'center', gap: 5, whiteSpace: 'nowrap' }}>
            <span style={{ width: 6, height: 6, borderRadius: '50%', background: MEDIA_STATUS[m.status].color, flexShrink: 0 }} />
            {MEDIA_STATUS[m.status].label}
          </span>
          {m.rating && <span style={{ whiteSpace: 'nowrap' }}>{m.rating}/5</span>}
          <button onClick={() => onEdit(m)} style={iconButtonStyle}>✎</button>
          <button onClick={() => onDelete(m.id)} style={iconButtonStyle}><Trash2 size={13} /></button>
          <button onClick={() => setExpanded(v => !v)} style={iconButtonStyle} aria-label={expanded ? 'Minska' : 'Öka'}>
            {expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          </button>
        </div>
      </div>
      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 10 }}>
        {(m.genres || []).map(gid => {
          const g = genres.find(x => x.id === gid);
          return g ? <span key={gid} style={tagStyle}>{g.label}</span> : null;
        })}
        {m.format && <span style={tagStyle}>{m.format}</span>}
      </div>

      {expanded && (
        <>
          {m.synopsis && <div style={{ marginTop: 10, fontSize: 12.5, color: 'var(--text-muted)', lineHeight: 1.4 }}>{m.synopsis}</div>}
          {m.lentTo && <div style={{ marginTop: 8, fontSize: 12.5, color: 'var(--accent-gold)' }}>Utlånad till {m.lentTo}</div>}
          {m.notes && <div style={{ marginTop: 10, fontSize: 13, color: 'var(--text-muted)', borderTop: '1px solid var(--border)', paddingTop: 10, lineHeight: 1.5 }}>{m.notes}</div>}
          {referencingTricks.length > 0 && (
            <div style={{ marginTop: 10, borderTop: '1px solid var(--border)', paddingTop: 10, fontSize: 12.5, color: 'var(--text-muted)' }}>
              Källa till: {referencingTricks.map(t => t.name).join(', ')}
            </div>
          )}
        </>
      )}
    </div>
  );
}

const inputStyle = { width: '100%', background: 'var(--card)', border: '1px solid var(--border)', color: 'var(--text)', padding: '9px 12px', fontSize: 14, outline: 'none', boxSizing: 'border-box' };
const labelStyle = { fontSize: 11.5, color: 'var(--text-muted)', marginBottom: 6, letterSpacing: '0.02em' };
const printHintStyle = { color: 'var(--accent-red)', fontStyle: 'italic' };
const chipStyle = { border: '1px solid var(--border)', padding: '6px 12px', fontSize: 12.5, background: 'transparent', color: 'var(--text-muted)', cursor: 'pointer', whiteSpace: 'nowrap' };
const tagStyle = { fontSize: 11, color: 'var(--accent-gold)', border: '1px solid var(--accent-gold)', padding: '2px 8px' };
const cardStyle = { border: '1px solid var(--border)', background: 'var(--panel)', padding: 16, position: 'relative' };
const primaryButtonStyle = { display: 'flex', alignItems: 'center', gap: 6, background: 'var(--accent-gold)', color: 'var(--ink)', border: 'none', padding: '9px 16px', fontSize: 13.5, cursor: 'pointer', fontWeight: 600 };
const ghostButtonStyle = { display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: 'transparent', color: 'var(--text-muted)', border: '1px solid var(--border)', padding: '8px 14px', fontSize: 13, cursor: 'pointer' };
const iconButtonStyle = { background: 'transparent', border: '1px solid var(--border)', color: 'var(--text-muted)', width: 26, height: 26, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', fontSize: 12 };
const panelBox = { border: '1px solid var(--border)', background: 'var(--panel)', padding: 18 };
const fieldGrid = { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 14 };
const tabStyle = { display: 'flex', alignItems: 'center', gap: 6, background: 'transparent', border: '1px solid', borderBottom: '2px solid', padding: '10px 18px', fontSize: 13.5, cursor: 'pointer' };

export default function RepertoarApp() {
  const { tricks, setTricks, gigs, setGigs, categories, setCategories, media, setMedia, genres, setGenres, formats, setFormats, devItems, setDevItems, settings, setSettings, loaded } = useStorage();
  const [view, setView] = useState('library');
  const [activeGig, setActiveGig] = useState(emptyGig());

  const handlePrint = (gig) => openPrintWindow(gig, tricks);

  const handleResetGigs = () => {
    setGigs([]);
    setActiveGig(emptyGig());
  };

  const handleResetMedia = () => {
    setMedia([]);
  };

  if (!loaded) {
    return <div style={{ padding: 40, color: '#b8ab95', fontFamily: 'Georgia, serif' }}>Öppnar biblioteket…</div>;
  }

  const openGig = (g) => { setActiveGig(g); setView('builder'); };
  const newGig = () => { setActiveGig({ ...emptyGig(), eventType: settings.defaultEventType || 'mingel', venue: settings.defaultVenue || 'inne' }); setView('builder'); };

  return (
    <div style={{
      ...(THEMES[settings.theme] || THEMES.light),
      '--font-display': "'Iowan Old Style', 'Palatino Linotype', Georgia, serif",
      background: 'var(--bg)', minHeight: '100vh', color: 'var(--text)',
      fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      padding: '28px 20px 60px',
    }}>
      <div style={{ maxWidth: 980, margin: '0 auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 26, borderBottom: '1px solid var(--border)', paddingBottom: 20 }}>
          <div>
            <div style={{ fontFamily: 'var(--font-display)', fontSize: 30, letterSpacing: '0.01em' }}>Showbasen</div>
            <div style={{ color: 'var(--text-muted)', fontSize: 13.5, marginTop: 4 }}>Show & trolleridatabasen — trickdatabas, set-listbyggare, uppdrag och samling</div>
          </div>
          <button onClick={() => setView('settings')} aria-label="Inställningar"
            style={{ background: 'transparent', border: '1px solid', borderColor: view === 'settings' ? 'var(--accent-gold)' : 'var(--border)', color: view === 'settings' ? 'var(--accent-gold)' : 'var(--text-muted)', width: 36, height: 36, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', flexShrink: 0 }}>
            <Settings size={16} />
          </button>
        </div>

        <div style={{ display: 'flex', gap: 4, marginBottom: 24, flexWrap: 'wrap' }}>
          <button onClick={() => setView('library')} style={{ ...tabStyle, borderColor: view === 'library' ? 'var(--accent-gold)' : 'var(--border)', color: view === 'library' ? 'var(--text)' : 'var(--text-muted)' }}><Library size={14} /> Bibliotek</button>
          <button onClick={newGig} style={{ ...tabStyle, borderColor: view === 'builder' ? 'var(--accent-gold)' : 'var(--border)', color: view === 'builder' ? 'var(--text)' : 'var(--text-muted)' }}><ListMusic size={14} /> Bygg set</button>
          <button onClick={() => setView('gigs')} style={{ ...tabStyle, borderColor: view === 'gigs' ? 'var(--accent-gold)' : 'var(--border)', color: view === 'gigs' ? 'var(--text)' : 'var(--text-muted)' }}><Briefcase size={14} /> Uppdrag</button>
          <button onClick={() => setView('media')} style={{ ...tabStyle, borderColor: view === 'media' ? 'var(--accent-gold)' : 'var(--border)', color: view === 'media' ? 'var(--text)' : 'var(--text-muted)' }}><Film size={14} /> Film &amp; böcker</button>
          <button onClick={() => setView('development')} style={{ ...tabStyle, borderColor: view === 'development' ? 'var(--accent-gold)' : 'var(--border)', color: view === 'development' ? 'var(--text)' : 'var(--text-muted)' }}><Target size={14} /> Utveckling</button>
        </div>

        {view === 'library' && <LibraryView tricks={tricks} setTricks={setTricks} categories={categories} gigs={gigs} onOpenGig={openGig} media={media} />}
        {view === 'builder' && <Builder tricks={tricks} gigs={gigs} setGigs={setGigs} activeGig={activeGig} setActiveGig={setActiveGig} onPrint={handlePrint} categories={categories} />}
        {view === 'gigs' && <GigsView gigs={gigs} setGigs={setGigs} tricks={tricks} onOpen={openGig} onPrint={handlePrint} categories={categories} exportProfiles={settings.exportProfiles} />}
        {view === 'media' && <MediaView media={media} setMedia={setMedia} genres={genres} formats={formats} tricks={tricks} />}
        {view === 'development' && <DevelopmentView devItems={devItems} setDevItems={setDevItems} tricks={tricks} />}
        {view === 'settings' && <SettingsView categories={categories} setCategories={setCategories} genres={genres} setGenres={setGenres} formats={formats} setFormats={setFormats} settings={settings} setSettings={setSettings} onResetGigs={handleResetGigs} onResetMedia={handleResetMedia}
          backupData={{ tricks, gigs, categories, media, genres, formats, devItems, settings }}
          onImportBackup={(data, mode) => {
            if (mode === 'merge') {
              const importedTricks = data.tricks || [];
              const idMap = {};
              setTricks(ts => {
                const byName = new Map(ts.map(t => [t.name.trim().toLowerCase(), t.id]));
                const newOnes = [];
                importedTricks.forEach(t => {
                  const key = (t.name || '').trim().toLowerCase();
                  if (byName.has(key)) {
                    idMap[t.id] = byName.get(key);
                  } else {
                    const newId = crypto.randomUUID();
                    idMap[t.id] = newId;
                    newOnes.push({ ...t, id: newId });
                  }
                });
                return [...ts, ...newOnes];
              });
              const remapPhaseTricks = (pt) => {
                const remap = (arr) => (arr || []).map(row => ({ ...row, entryId: crypto.randomUUID(), trickId: idMap[row.trickId] || row.trickId }));
                return { opening: remap(pt?.opening), middle: remap(pt?.middle), closing: remap(pt?.closing) };
              };
              setGigs(gs => [...gs, ...((data.gigs || []).map(g => migrateGig({ ...g, id: crypto.randomUUID(), phaseTricks: remapPhaseTricks(g.phaseTricks) })))]);
              setMedia(ms => [...ms, ...((data.media || []).map(m => ({ ...emptyMedia(), ...m, id: crypto.randomUUID(), genres: m.genres || [] })))]);
              setDevItems(ds => [...ds, ...((data.devItems || []).map(d => ({ ...emptyDevItem(), ...d, id: crypto.randomUUID(), trickId: idMap[d.trickId] || d.trickId })))]);
            } else {
              setTricks(data.tricks || []);
              setGigs((data.gigs || []).map(migrateGig));
              setCategories(data.categories || DEFAULT_CATEGORIES);
              setMedia((data.media || []).map(x => ({ ...emptyMedia(), ...x, genres: x.genres || [] })));
              setDevItems((data.devItems || []).map(x => ({ ...emptyDevItem(), ...x })));
              setGenres(data.genres || DEFAULT_GENRES);
              setFormats(data.formats || DEFAULT_FORMATS);
              setSettings({ ...defaultSettings(), ...(data.settings || {}) });
            }
          }}
        />}
      </div>
    </div>
  );
}
