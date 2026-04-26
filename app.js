if (!window.Tone) {
  throw new Error("Tone.js failed to load. Check your internet connection and reload.");
}

const audioContext = Tone.getContext();
audioContext.lookAhead = 0.12;
audioContext.updateInterval = 0.03;

const MAX_STEPS = 64;
const NOTE_COUNT = 24;
const TOP_MIDI = 83; // B5
const STORAGE_KEY = "orchestrion-studio-project-v1";
const INTEGRATION_STORAGE_KEY = "orchestrion-studio-integrations-v1";
const DEFAULT_NOTE_VELOCITY = 0.72;
const SEQUENCER_VELOCITY_SCALE = 0.74;
const LIVE_VELOCITY_SCALE = 0.82;
const MASTER_HEADROOM_GAIN = 0.76;
const SOUNDFONT_BASE_URL = "https://gleitz.github.io/midi-js-soundfonts/FluidR3_GM";
const SOUNDFONT_SAMPLE_NOTES = [
  "C2",
  "D#2",
  "F#2",
  "A2",
  "C3",
  "D#3",
  "F#3",
  "A3",
  "C4",
  "D#4",
  "F#4",
  "A4",
  "C5",
  "D#5",
  "F#5",
  "A5",
  "C6",
];

const NOTE_NAMES = Array.from({ length: NOTE_COUNT }, (_, index) =>
  Tone.Frequency(TOP_MIDI - index, "midi").toNote()
);

const CHORD_DRAG_MIME = "application/x-orchestrion-chord";

const PITCH_CLASS_NAMES = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"];

const CHORD_ROOTS = PITCH_CLASS_NAMES.map((name, pitchClass) => ({
  id: name.replace("#", "-sharp").toLowerCase(),
  name,
  pitchClass,
}));

const CHORD_QUALITIES = [
  { id: "major", name: "Major", intervals: [0, 4, 7, 12] },
  { id: "minor", name: "Minor", intervals: [0, 3, 7, 12] },
  { id: "diminished", name: "Diminished", intervals: [0, 3, 6, 12] },
  { id: "augmented", name: "Augmented", intervals: [0, 4, 8, 12] },
  { id: "sus2", name: "Sus2", intervals: [0, 2, 7, 12] },
  { id: "sus4", name: "Sus4", intervals: [0, 5, 7, 12] },
  { id: "maj7", name: "Major 7", intervals: [0, 4, 7, 11] },
  { id: "min7", name: "Minor 7", intervals: [0, 3, 7, 10] },
  { id: "dom7", name: "Dominant 7", intervals: [0, 4, 7, 10] },
  { id: "add9", name: "Add9", intervals: [0, 4, 7, 14] },
  { id: "minadd9", name: "Minor Add9", intervals: [0, 3, 7, 14] },
  { id: "dim7", name: "Diminished 7", intervals: [0, 3, 6, 9] },
];

function buildChordDescription(rootPitchClass, intervals) {
  return intervals
    .map((interval) => PITCH_CLASS_NAMES[(rootPitchClass + interval) % 12])
    .join(" ");
}

function buildChordEvents(intervals) {
  const velocityMap = [0.84, 0.8, 0.76, 0.72, 0.68, 0.64];
  return intervals.map((interval, index) => ({
    step: 0,
    semitone: interval,
    velocity: velocityMap[Math.min(index, velocityMap.length - 1)],
  }));
}

const CHORD_LIBRARY = CHORD_ROOTS.flatMap((root) => {
  return CHORD_QUALITIES.map((quality) => ({
    id: `${root.id}-${quality.id}`,
    name: `${root.name} ${quality.name}`,
    description: buildChordDescription(root.pitchClass, quality.intervals),
    rootPitchClass: root.pitchClass,
    events: buildChordEvents(quality.intervals),
  }));
});

const TRACK_COLORS = ["#bf5f2f", "#1a7c79", "#a2413e", "#5f6fbb", "#b27a2d", "#3f8a52", "#9e3651", "#6e54a8"];

  Strings: { oscillator: "sawtooth", envelope: { attack: 0.12, decay: 0.2, sustain: 0.62, release: 1.3 }, volume: -8, reverbSend: 0.28, delaySend: 0.07 },
  "Electric Guitar": { oscillator: "sawtooth", envelope: { attack: 0.01, decay: 0.5, sustain: 0.2, release: 1 }, volume: -8, reverbSend: 0.1, delaySend: 0.05 },
  "Acoustic Guitar": { oscillator: "triangle", envelope: { attack: 0.005, decay: 0.8, sustain: 0.1, release: 1 }, volume: -9, reverbSend: 0.2, delaySend: 0 },
  "Electric Bass": { oscillator: "square", envelope: { attack: 0.01, decay: 0.4, sustain: 0.3, release: 0.5 }, volume: -7, reverbSend: 0.02, delaySend: 0 },
  "Acoustic Bass": { oscillator: "triangle", envelope: { attack: 0.02, decay: 0.5, sustain: 0.2, release: 0.8 }, volume: -8, reverbSend: 0.05, delaySend: 0 },
  Brass: { oscillator: "square", envelope: { attack: 0.04, decay: 0.16, sustain: 0.5, release: 0.65 }, volume: -10, reverbSend: 0.17, delaySend: 0.04 },
  Woodwinds: { oscillator: "triangle", envelope: { attack: 0.05, decay: 0.12, sustain: 0.46, release: 0.56 }, volume: -9, reverbSend: 0.14, delaySend: 0.1 },
  Choir: { oscillator: "sine", envelope: { attack: 0.18, decay: 0.24, sustain: 0.74, release: 1.75 }, volume: -11, reverbSend: 0.33, delaySend: 0.06 },
  Percussion: { oscillator: "square", envelope: { attack: 0.001, decay: 0.11, sustain: 0.02, release: 0.18 }, volume: -8, reverbSend: 0.08, delaySend: 0.02 },
  Piano: { oscillator: "triangle", envelope: { attack: 0.01, decay: 0.25, sustain: 0.2, release: 1.4 }, volume: -9, reverbSend: 0.25, delaySend: 0.05 },
  "Drum Machine": { oscillator: "square", envelope: { attack: 0.001, decay: 0.09, sustain: 0.01, release: 0.12 }, volume: -8, reverbSend: 0.04, delaySend: 0 },
  "Kick Drum": { oscillator: "sine", envelope: { attack: 0.001, decay: 0.15, sustain: 0, release: 0.1 }, volume: -6, reverbSend: 0.01, delaySend: 0 },
  "Snare Drum": { oscillator: "square", envelope: { attack: 0.001, decay: 0.1, sustain: 0, release: 0.1 }, volume: -7, reverbSend: 0.1, delaySend: 0 },
  "Hi-Hat": { oscillator: "sawtooth", envelope: { attack: 0.001, decay: 0.05, sustain: 0, release: 0.05 }, volume: -10, reverbSend: 0.02, delaySend: 0 },
  "Toms": { oscillator: "triangle", envelope: { attack: 0.005, decay: 0.3, sustain: 0, release: 0.2 }, volume: -8, reverbSend: 0.08, delaySend: 0 },
  "Cinematic Percussion": { oscillator: "triangle", envelope: { attack: 0.002, decay: 0.14, sustain: 0.02, release: 0.2 }, volume: -9, reverbSend: 0.16, delaySend: 0.02 },
  "Mallets": { oscillator: "sine", envelope: { attack: 0.002, decay: 0.2, sustain: 0.02, release: 0.5 }, volume: -10, reverbSend: 0.2, delaySend: 0.04 },
  "Woodblock": { oscillator: "square", envelope: { attack: 0.001, decay: 0.05, sustain: 0, release: 0.1 }, volume: -8, reverbSend: 0.05, delaySend: 0 },
  "Steel Drums": { oscillator: "sine", envelope: { attack: 0.005, decay: 0.3, sustain: 0.1, release: 0.8 }, volume: -9, reverbSend: 0.15, delaySend: 0.05 },
  "Melodic Tom": { oscillator: "triangle", envelope: { attack: 0.005, decay: 0.4, sustain: 0.05, release: 0.5 }, volume: -8, reverbSend: 0.1, delaySend: 0 },
  "Reverse Cymbal": { oscillator: "noise", envelope: { attack: 0.5, decay: 0.1, sustain: 0, release: 0.1 }, volume: -10, reverbSend: 0.2, delaySend: 0.1 },
  "Sub Bass": { oscillator: "sine", envelope: { attack: 0.01, decay: 0.12, sustain: 0.7, release: 0.25 }, volume: -7, reverbSend: 0.01, delaySend: 0.01 },
  "EDM Bass": { oscillator: "sawtooth", envelope: { attack: 0.005, decay: 0.12, sustain: 0.48, release: 0.2 }, volume: -8, reverbSend: 0.03, delaySend: 0.03 },
  "EDM Lead": { oscillator: "square", envelope: { attack: 0.01, decay: 0.11, sustain: 0.36, release: 0.2 }, volume: -9, reverbSend: 0.12, delaySend: 0.18 },
  "EDM Pluck": { oscillator: "triangle", envelope: { attack: 0.001, decay: 0.18, sustain: 0.02, release: 0.16 }, volume: -8, reverbSend: 0.11, delaySend: 0.14 },
  "Synth Pad": { oscillator: "sawtooth", envelope: { attack: 0.24, decay: 0.22, sustain: 0.78, release: 1.9 }, volume: -11, reverbSend: 0.35, delaySend: 0.08 },
  "Techno Stab": { oscillator: "square", envelope: { attack: 0.001, decay: 0.13, sustain: 0.08, release: 0.12 }, volume: -9, reverbSend: 0.07, delaySend: 0.06 },
  Arp: { oscillator: "triangle", envelope: { attack: 0.002, decay: 0.14, sustain: 0.22, release: 0.18 }, volume: -9, reverbSend: 0.09, delaySend: 0.2 },
  "FX Atmos": { oscillator: "sine", envelope: { attack: 0.4, decay: 0.3, sustain: 0.74, release: 2.2 }, volume: -12, reverbSend: 0.42, delaySend: 0.2 },
};

const SOUNDFONT_INSTRUMENT_MAP = {
  Strings: "string_ensemble_1",
  "Electric Guitar": "electric_guitar_clean",
  "Acoustic Guitar": "acoustic_guitar_nylon",
  "Electric Bass": "electric_bass_finger",
  "Acoustic Bass": "acoustic_bass",
  Brass: "brass_section",
  Woodwinds: "oboe",
  Choir: "choir_aahs",
  Percussion: "timpani",
  Piano: "acoustic_grand_piano",
  "Drum Machine": "synth_drum",
  "Kick Drum": "synth_drum",
  "Snare Drum": "synth_drum",
  "Hi-Hat": "synth_drum",
  "Toms": "melodic_tom",
  "Cinematic Percussion": "taiko_drum",
  "Mallets": "marimba",
  "Woodblock": "woodblock",
  "Steel Drums": "steel_drums",
  "Melodic Tom": "melodic_tom",
  "Reverse Cymbal": "reverse_cymbal",
  "Sub Bass": "synth_bass_1",
  "EDM Bass": "synth_bass_2",
  "EDM Lead": "lead_2_sawtooth",
  "EDM Pluck": "lead_1_square",
  "Synth Pad": "pad_3_polysynth",
  "Techno Stab": "synth_brass_1",
  Arp: "clavinet",
  "FX Atmos": "fx_4_atmosphere",
};

const INSTRUMENT_MIDI_PROGRAMS = {
  Strings: 49,
  "Electric Guitar": 28,
  "Acoustic Guitar": 25,
  "Electric Bass": 34,
  "Acoustic Bass": 33,
  Brass: 62,
  Woodwinds: 69,
  Choir: 53,
  Percussion: 48,
  Piano: 1,
  "Drum Machine": 119,
  "Kick Drum": 119,
  "Snare Drum": 119,
  "Hi-Hat": 119,
  "Toms": 118,
  "Cinematic Percussion": 117,
  "Mallets": 13,
  "Woodblock": 116,
  "Steel Drums": 115,
  "Melodic Tom": 118,
  "Reverse Cymbal": 120,
  "Sub Bass": 39,
  "EDM Bass": 40,
  "EDM Lead": 82,
  "EDM Pluck": 81,
  "Synth Pad": 90,
  "Techno Stab": 63,
  Arp: 8,
  "FX Atmos": 100,
};

const INSTRUMENT_GROUPS = {
  "String Instruments (Chordophones)": ["Strings", "Electric Guitar", "Acoustic Guitar", "Electric Bass", "Acoustic Bass"],
  "Wind Instruments (Aerophones)": ["Brass", "Woodwinds"],
  "Keyboard Instruments": ["Piano"],
  "Percussion Instruments (Idiophones/Membranophones)": [
    "Percussion",
    "Drum Machine",
    "Kick Drum",
    "Snare Drum",
    "Hi-Hat",
    "Toms",
    "Cinematic Percussion",
    "Mallets",
    "Woodblock",
    "Steel Drums",
    "Melodic Tom",
    "Reverse Cymbal"
  ],
  "Electronic Instruments (Electrophones)": [
    "Choir",
    "Sub Bass",
    "EDM Bass",
    "EDM Lead",
    "EDM Pluck",
    "Synth Pad",
    "Techno Stab",
    "Arp",
    "FX Atmos"
  ]
};

const ui = {
  composer: document.querySelector(".composer"),
  rollScroll: document.querySelector(".roll-scroll"),
  composerFullscreenBtn: document.getElementById("composerFullscreenBtn"),
  composerCloseBtn: document.getElementById("composerCloseBtn"),
  composerModalBackdrop: document.getElementById("composerModalBackdrop"),
  chordCompactToggleBtn: document.getElementById("chordCompactToggleBtn"),
  playBtn: document.getElementById("playBtn"),
  stopBtn: document.getElementById("stopBtn"),
  metronomeBtn: document.getElementById("metronomeBtn"),
  countInBtn: document.getElementById("countInBtn"),
  tempoSlider: document.getElementById("tempoSlider"),
  tempoValue: document.getElementById("tempoValue"),
  swingSlider: document.getElementById("swingSlider"),
  swingValue: document.getElementById("swingValue"),
  loopSelect: document.getElementById("loopSelect"),
  trackNameInput: document.getElementById("trackNameInput"),
  instrumentSelect: document.getElementById("instrumentSelect"),
  addTrackBtn: document.getElementById("addTrackBtn"),
  clearTrackBtn: document.getElementById("clearTrackBtn"),
  trackList: document.getElementById("trackList"),
  phraseLibrary: document.getElementById("phraseLibrary"),
  pianoRoll: document.getElementById("pianoRoll"),
  mixerStrips: document.getElementById("mixerStrips"),
  masterMeterFill: document.getElementById("masterMeterFill"),
  masterPeak: document.getElementById("masterPeak"),
  undoBtn: document.getElementById("undoBtn"),
  redoBtn: document.getElementById("redoBtn"),
  saveLocalBtn: document.getElementById("saveLocalBtn"),
  loadLocalBtn: document.getElementById("loadLocalBtn"),
  exportBtn: document.getElementById("exportBtn"),
  exportSheetsBtn: document.getElementById("exportSheetsBtn"),
  importBtn: document.getElementById("importBtn"),
  importFileInput: document.getElementById("importFileInput"),
  templateSelect: document.getElementById("templateSelect"),
  applyTemplateBtn: document.getElementById("applyTemplateBtn"),
  duplicatePhraseBtn: document.getElementById("duplicatePhraseBtn"),
  humanizeBtn: document.getElementById("humanizeBtn"),
  nudgeLeftBtn: document.getElementById("nudgeLeftBtn"),
  nudgeRightBtn: document.getElementById("nudgeRightBtn"),
  midiConnectBtn: document.getElementById("midiConnectBtn"),
  midiInputSelect: document.getElementById("midiInputSelect"),
  midiChannelSelect: document.getElementById("midiChannelSelect"),
  midiCaptureBtn: document.getElementById("midiCaptureBtn"),
  midiStatus: document.getElementById("midiStatus"),
  yamahaUrlInput: document.getElementById("yamahaUrlInput"),
  yamahaAuthHeaderInput: document.getElementById("yamahaAuthHeaderInput"),
  yamahaAuthValueInput: document.getElementById("yamahaAuthValueInput"),
  yamahaAudioInput: document.getElementById("yamahaAudioInput"),
  yamahaAnalyzeBtn: document.getElementById("yamahaAnalyzeBtn"),
  yamahaApplyTempoBtn: document.getElementById("yamahaApplyTempoBtn"),
  yamahaSummary: document.getElementById("yamahaSummary"),
  yamahaResult: document.getElementById("yamahaResult"),
  statusBar: document.getElementById("statusBar"),
};

const fx = {
  reverb: new Tone.Reverb({ decay: 4.2, preDelay: 0.02, wet: 1 }),
  delay: new Tone.FeedbackDelay({ delayTime: "8n", feedback: 0.26, wet: 1 }),
  masterBus: new Tone.Gain(MASTER_HEADROOM_GAIN),
  compressor: new Tone.Compressor({ threshold: -22, ratio: 3, attack: 0.01, release: 0.2 }),
  limiter: new Tone.Limiter(-1),
  meter: new Tone.Meter({ normalRange: false, smoothing: 0.83 }),
  metronomeHigh: new Tone.MembraneSynth({
    pitchDecay: 0.015,
    octaves: 2,
    oscillator: { type: "triangle" },
    envelope: { attack: 0.001, decay: 0.06, sustain: 0, release: 0.03 },
  }),
  metronomeLow: new Tone.MembraneSynth({
    pitchDecay: 0.02,
    octaves: 1,
    oscillator: { type: "sine" },
    envelope: { attack: 0.001, decay: 0.045, sustain: 0, release: 0.03 },
  }),
};

fx.reverb.connect(fx.masterBus);
fx.delay.connect(fx.masterBus);
fx.masterBus.connect(fx.compressor);
fx.compressor.connect(fx.limiter);
fx.limiter.toDestination();
fx.limiter.connect(fx.meter);
fx.metronomeHigh.connect(fx.masterBus);
fx.metronomeLow.connect(fx.masterBus);
fx.metronomeHigh.volume.value = -12;
fx.metronomeLow.volume.value = -17;

let tracks = [];
let selectedTrackId = null;
let loopSteps = Number(ui.loopSelect.value);
let currentStep = 0;
let playheadStep = 0;
let sequenceEventId = null;
let colorCursor = 0;
let audioReady = false;
let pointerPaintMode = null;
let shownSoundfontFallbackWarning = false;
let transportTick = 0;
let metronomeEnabled = false;
let countInEnabled = false;
let countInStepsRemaining = 0;
let didPaintInCurrentStroke = false;

const HISTORY_LIMIT = 80;
const historyState = {
  undoStack: [],
  redoStack: [],
};

const pointerState = {
  isDown: false,
  value: 0,
};

const midiState = {
  access: null,
  selectedInputId: "",
  activeInput: null,
  channel: "all",
  captureToGrid: false,
  activeNotes: new Map(),
};

let yamahaLastAnalysis = null;
let composerModalOpen = false;
let chordPaletteCompact = false;

function clamp(number, min, max) {
  return Math.min(max, Math.max(min, number));
}

function clamp01(number) {
  return clamp(number, 0, 1);
}

function asNumber(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function generateId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }

  return `track-${Date.now()}-${Math.round(Math.random() * 1e7)}`;
}

function nextColor() {
  const color = TRACK_COLORS[colorCursor % TRACK_COLORS.length];
  colorCursor += 1;
  return color;
}

function isBlackKey(noteName) {
  return noteName.includes("#");
}

function getTrackById(trackId) {
  return tracks.find((track) => track.id === trackId);
}

function getSelectedTrack() {
  return getTrackById(selectedTrackId) || tracks[0] || null;
}

function getChordById(chordId) {
  return CHORD_LIBRARY.find((chord) => chord.id === chordId) || null;
}

function nearestMidiForPitchClass(anchorMidi, pitchClass) {
  const normalizedAnchorClass = ((anchorMidi % 12) + 12) % 12;
  let delta = pitchClass - normalizedAnchorClass;

  while (delta > 6) {
    delta -= 12;
  }

  while (delta < -6) {
    delta += 12;
  }

  return anchorMidi + delta;
}

function createPattern() {
  return Array.from({ length: NOTE_COUNT }, () => Array(MAX_STEPS).fill(0));
}

function normalizePattern(pattern) {
  const normalized = createPattern();

  if (!Array.isArray(pattern)) {
    return normalized;
  }

  for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
    const sourceRow = Array.isArray(pattern[noteIndex]) ? pattern[noteIndex] : [];
    for (let stepIndex = 0; stepIndex < MAX_STEPS; stepIndex += 1) {
      normalized[noteIndex][stepIndex] = clamp01(asNumber(sourceRow[stepIndex], 0));
    }
  }

  return normalized;
}

function cloneSnapshot(snapshot) {
  return JSON.parse(JSON.stringify(snapshot));
}

function updateHistoryButtons() {
  if (ui.undoBtn) {
    ui.undoBtn.disabled = historyState.undoStack.length === 0;
  }
  if (ui.redoBtn) {
    ui.redoBtn.disabled = historyState.redoStack.length === 0;
  }
}

function pushHistory(reason = "edit") {
  historyState.undoStack.push(cloneSnapshot(projectSnapshot()));
  if (historyState.undoStack.length > HISTORY_LIMIT) {
    historyState.undoStack.shift();
  }
  historyState.redoStack = [];
  updateHistoryButtons();
  if (reason) {
    setStatus(`Saved undo point for ${reason}.`);
  }
}

function pushHistorySilent() {
  historyState.undoStack.push(cloneSnapshot(projectSnapshot()));
  if (historyState.undoStack.length > HISTORY_LIMIT) {
    historyState.undoStack.shift();
  }
  historyState.redoStack = [];
  updateHistoryButtons();
}

function restoreFromSnapshot(snapshot, sourceLabel = "history") {
  loadProject(snapshot, sourceLabel, { skipHistoryReset: true });
}

function undoHistory() {
  if (historyState.undoStack.length === 0) {
    setStatus("Nothing to undo.", true);
    return;
  }

  historyState.redoStack.push(cloneSnapshot(projectSnapshot()));
  const previous = historyState.undoStack.pop();
  restoreFromSnapshot(previous, "undo snapshot");
  updateHistoryButtons();
  setStatus("Undo complete.");
}

function redoHistory() {
  if (historyState.redoStack.length === 0) {
    setStatus("Nothing to redo.", true);
    return;
  }

  historyState.undoStack.push(cloneSnapshot(projectSnapshot()));
  const next = historyState.redoStack.pop();
  restoreFromSnapshot(next, "redo snapshot");
  updateHistoryButtons();
  setStatus("Redo complete.");
}

function shiftTrackPattern(track, direction) {
  if (!track || !["left", "right"].includes(direction)) {
    return;
  }

  for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
    const row = track.pattern[noteIndex].slice(0, loopSteps);
    const shifted = Array(loopSteps).fill(0);
    for (let stepIndex = 0; stepIndex < loopSteps; stepIndex += 1) {
      const sourceIndex = direction === "left" ? (stepIndex + 1) % loopSteps : (stepIndex - 1 + loopSteps) % loopSteps;
      shifted[stepIndex] = row[sourceIndex];
    }
    for (let stepIndex = 0; stepIndex < loopSteps; stepIndex += 1) {
      track.pattern[noteIndex][stepIndex] = shifted[stepIndex];
    }
  }
}

function nudgeSelectedTrack(direction) {
  const selected = getSelectedTrack();
  if (!selected) {
    return;
  }

  pushHistorySilent();
  shiftTrackPattern(selected, direction);
  renderPianoRoll();
  setStatus(`Nudged ${selected.name} ${direction}.`);
}

function humanizeSelectedTrack() {
  const selected = getSelectedTrack();
  if (!selected) {
    return;
  }

  pushHistorySilent();
  let touchedNotes = 0;

  for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
    for (let stepIndex = 0; stepIndex < loopSteps; stepIndex += 1) {
      const value = selected.pattern[noteIndex][stepIndex];
      if (value <= 0) {
        continue;
      }

      const delta = (Math.random() - 0.5) * 0.18;
      selected.pattern[noteIndex][stepIndex] = clamp01(value + delta);
      touchedNotes += 1;
    }
  }

  renderPianoRoll();
  setStatus(`Humanized ${touchedNotes} notes on ${selected.name}.`);
}

function duplicatePhraseAcrossLoop() {
  const selected = getSelectedTrack();
  if (!selected) {
    return;
  }

  const phraseLength = Math.min(16, loopSteps);
  if (phraseLength <= 0) {
    return;
  }

  pushHistorySilent();

  for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
    const phrase = selected.pattern[noteIndex].slice(0, phraseLength);
    for (let stepIndex = phraseLength; stepIndex < loopSteps; stepIndex += 1) {
      selected.pattern[noteIndex][stepIndex] = phrase[stepIndex % phraseLength];
    }
  }

  renderPianoRoll();
  setStatus(`Duplicated first ${phraseLength} steps across the loop on ${selected.name}.`);
}

function createSynthForInstrument(instrumentName) {
  const preset = INSTRUMENTS[instrumentName] || INSTRUMENTS.Strings;
  const envelope = {
    ...preset.envelope,
    attack: Math.max(0.005, asNumber(preset.envelope.attack, 0.01)),
  };

  const synth = new Tone.PolySynth(Tone.Synth, { maxPolyphony: 24 });
  synth.set({
    oscillator: { type: preset.oscillator },
    envelope,
  });
  return synth;
}

function noteNameToSoundfontFile(noteName) {
  return `${noteName.replace("#", "s")}.mp3`;
}

function createSamplerForInstrument(instrumentName, onLoad, onError) {
  const gmInstrument = SOUNDFONT_INSTRUMENT_MAP[instrumentName] || "acoustic_grand_piano";
  const urls = {};

  for (const noteName of SOUNDFONT_SAMPLE_NOTES) {
    urls[noteName] = noteNameToSoundfontFile(noteName);
  }

  return new Tone.Sampler({
    urls,
    baseUrl: `${SOUNDFONT_BASE_URL}/${gmInstrument}-mp3/`,
    attack: 0.01,
    release: 0.65,
    onload: onLoad,
    onerror: onError,
  });
}

function initializeTrackSampler(track) {
  let errorNotified = false;

  track.samplerReady = false;
  track.samplerFailed = false;
  track.samplerArmed = false;
  track.sampler = createSamplerForInstrument(
    track.instrument,
    () => {
      track.samplerReady = true;
      if (Tone.Transport.state !== "started") {
        track.samplerArmed = true;
      }
    },
    () => {
      track.samplerFailed = true;
      track.samplerArmed = false;

      if (!errorNotified) {
        errorNotified = true;

        if (!shownSoundfontFallbackWarning) {
          shownSoundfontFallbackWarning = true;
          setStatus("Some online soundfont assets failed to load. Using synth fallback where needed.", true);
        }
      }
    }
  );

  track.sampler.connect(track.channel);
}

function shouldUseSampler(track) {
  if (!track || !track.sampler || track.samplerFailed || !track.samplerReady) {
    return false;
  }

  if (Tone.Transport.state === "started") {
    return Boolean(track.samplerArmed);
  }

  return true;
}

function triggerTrackAttackRelease(track, noteName, duration, time, velocity, velocityScale = 1) {
  const safeVelocity = clamp01(velocity * velocityScale);

  if (shouldUseSampler(track)) {
    track.sampler.triggerAttackRelease(noteName, duration, time, safeVelocity);
    return;
  }

  track.synth.triggerAttackRelease(noteName, duration, time, safeVelocity);
}

function triggerTrackAttack(track, noteName, time, velocity, velocityScale = 1) {
  const safeVelocity = clamp01(velocity * velocityScale);

  if (shouldUseSampler(track)) {
    track.sampler.triggerAttack(noteName, time, safeVelocity);
    return;
  }

  track.synth.triggerAttack(noteName, time, safeVelocity);
}

function triggerTrackRelease(track, noteName, time) {
  if (shouldUseSampler(track)) {
    track.sampler.triggerRelease(noteName, time);
    return;
  }

  track.synth.triggerRelease(noteName, time);
}

function createTrack(options = {}) {
  const instrumentName = INSTRUMENTS[options.instrument] ? options.instrument : "Strings";
  const instrumentPreset = INSTRUMENTS[instrumentName];

  const track = {
    id: options.id || generateId(),
    name: (options.name || `${instrumentName} Layer`).slice(0, 24),
    instrument: instrumentName,
    color: options.color || nextColor(),
    pattern: normalizePattern(options.pattern),
    volume: clamp(asNumber(options.volume, instrumentPreset.volume), -36, 6),
    pan: clamp(asNumber(options.pan, 0), -1, 1),
    mute: Boolean(options.mute),
    solo: Boolean(options.solo),
    reverbSend: clamp01(asNumber(options.reverbSend, instrumentPreset.reverbSend)),
    delaySend: clamp01(asNumber(options.delaySend, instrumentPreset.delaySend)),
    synth: null,
    sampler: null,
    samplerReady: false,
    samplerFailed: false,
    samplerArmed: false,
    channel: null,
    reverbSendNode: null,
    delaySendNode: null,
  };

  track.channel = new Tone.Channel({ volume: track.volume, pan: track.pan });
  track.channel.connect(fx.masterBus);

  track.reverbSendNode = new Tone.Gain(track.reverbSend);
  track.delaySendNode = new Tone.Gain(track.delaySend);
  track.channel.connect(track.reverbSendNode);
  track.channel.connect(track.delaySendNode);
  track.reverbSendNode.connect(fx.reverb);
  track.delaySendNode.connect(fx.delay);

  track.synth = createSynthForInstrument(track.instrument);
  track.synth.connect(track.channel);
  initializeTrackSampler(track);

  return track;
}

function disposeTrack(track) {
  if (track.synth) {
    track.synth.dispose();
  }
  if (track.reverbSendNode) {
    track.reverbSendNode.dispose();
  }
  if (track.delaySendNode) {
    track.delaySendNode.dispose();
  }
  if (track.sampler) {
    track.sampler.dispose();
  }
  if (track.channel) {
    track.channel.dispose();
  }
}

function applySoloState() {
  const soloExists = tracks.some((track) => track.solo);
  for (const track of tracks) {
    track.channel.mute = track.mute || (soloExists && !track.solo);
  }
}

function setStatus(message, isError = false) {
  ui.statusBar.textContent = message;
  ui.statusBar.classList.toggle("status-error", isError);
}

function setIntegrationFeedback(node, message, isError = false) {
  if (!node) {
    return;
  }

  node.textContent = message;
  node.classList.toggle("is-error", isError);
}

function setMidiFeedback(message, isError = false) {
  setIntegrationFeedback(ui.midiStatus, message, isError);
}

function setYamahaFeedback(message, isError = false) {
  setIntegrationFeedback(ui.yamahaSummary, message, isError);
}

function setComposerModalState(isOpen) {
  composerModalOpen = Boolean(isOpen);
  document.body.classList.toggle("composer-modal-open", composerModalOpen);

  if (ui.composerModalBackdrop) {
    ui.composerModalBackdrop.hidden = !composerModalOpen;
  }

  if (ui.composerFullscreenBtn) {
    ui.composerFullscreenBtn.textContent = composerModalOpen ? "Exit Fullscreen" : "Open Fullscreen";
    ui.composerFullscreenBtn.setAttribute("aria-expanded", composerModalOpen ? "true" : "false");
  }
}

function toggleComposerModal() {
  setComposerModalState(!composerModalOpen);
}

function setChordPaletteCompactState(isCompact) {
  chordPaletteCompact = Boolean(isCompact);

  if (ui.phraseLibrary) {
    ui.phraseLibrary.classList.toggle("compact", chordPaletteCompact);
  }

  if (ui.chordCompactToggleBtn) {
    ui.chordCompactToggleBtn.textContent = chordPaletteCompact ? "Compact: On" : "Compact: Off";
    ui.chordCompactToggleBtn.setAttribute("aria-pressed", chordPaletteCompact ? "true" : "false");
  }
}

function toggleChordPaletteCompact() {
  setChordPaletteCompactState(!chordPaletteCompact);
}

function onRollScrollWheel(event) {
  if (!(event.currentTarget instanceof HTMLElement)) {
    return;
  }

  if (!event.shiftKey || Math.abs(event.deltaY) <= Math.abs(event.deltaX)) {
    return;
  }

  event.preventDefault();
  event.currentTarget.scrollLeft += event.deltaY;
}

function supportsWebMidi() {
  return typeof navigator !== "undefined" && typeof navigator.requestMIDIAccess === "function";
}

function midiNoteToPatternIndex(noteNumber) {
  if (!Number.isFinite(noteNumber)) {
    return -1;
  }

  const bottomMidi = TOP_MIDI - (NOTE_COUNT - 1);
  if (noteNumber < bottomMidi || noteNumber > TOP_MIDI) {
    return -1;
  }

  return TOP_MIDI - noteNumber;
}

function getCaptureStep() {
  if (Tone.Transport.state === "started") {
    const ticksPerStep = Tone.Time("16n").toTicks();
    const rawStep = Math.floor(Tone.Transport.ticks / ticksPerStep);
    return ((rawStep % loopSteps) + loopSteps) % loopSteps;
  }

  return clamp(asNumber(playheadStep, 0), 0, loopSteps - 1);
}

function writePatternValue(noteIndex, stepIndex, value) {
  const selected = getSelectedTrack();
  if (!selected) {
    return;
  }

  selected.pattern[noteIndex][stepIndex] = value;

  const cell = ui.pianoRoll.querySelector(`.step-cell[data-note="${noteIndex}"][data-step="${stepIndex}"]`);
  if (cell instanceof HTMLElement) {
    paintGridCell(cell, value);
  }
}

function captureMidiNoteToGrid(noteNumber, velocity) {
  const noteIndex = midiNoteToPatternIndex(noteNumber);
  if (noteIndex < 0) {
    return;
  }

  const step = getCaptureStep();
  writePatternValue(noteIndex, step, clamp01(velocity));
}

function releaseMidiNoteByKey(noteKey) {
  const active = midiState.activeNotes.get(noteKey);
  if (!active) {
    return;
  }

  const track = getTrackById(active.trackId);
  if (track) {
    triggerTrackRelease(track, active.noteName);
  }

  midiState.activeNotes.delete(noteKey);
}

function releaseAllMidiNotes() {
  for (const noteKey of midiState.activeNotes.keys()) {
    releaseMidiNoteByKey(noteKey);
  }
}

function onMidiMessage(event) {
  if (!event || !event.data || event.data.length < 2) {
    return;
  }

  const [statusByte, data1, data2 = 0] = event.data;
  const type = statusByte & 0xf0;
  const channel = (statusByte & 0x0f) + 1;

  if (midiState.channel !== "all" && Number(midiState.channel) !== channel) {
    return;
  }

  const noteKey = `${channel}:${data1}`;

  if (type === 0x90 && data2 > 0) {
    const selected = getSelectedTrack();
    if (!selected) {
      return;
    }

    const noteIndex = midiNoteToPatternIndex(data1);
    if (noteIndex < 0) {
      return;
    }

    const normalizedVelocity = clamp(data2 / 127, 0.05, 1);
    const noteName = NOTE_NAMES[noteIndex];
    triggerTrackAttack(selected, noteName, undefined, normalizedVelocity, LIVE_VELOCITY_SCALE);
    midiState.activeNotes.set(noteKey, { trackId: selected.id, noteName });

    if (midiState.captureToGrid) {
      captureMidiNoteToGrid(data1, normalizedVelocity);
    }

    return;
  }

  if (type === 0x80 || (type === 0x90 && data2 === 0)) {
    releaseMidiNoteByKey(noteKey);
  }
}

function detachActiveMidiInput() {
  if (midiState.activeInput) {
    midiState.activeInput.onmidimessage = null;
  }

  midiState.activeInput = null;
}

function attachMidiInputById(inputId) {
  if (!midiState.access) {
    return;
  }

  detachActiveMidiInput();
  const input = midiState.access.inputs.get(inputId);

  if (!input) {
    setMidiFeedback("No MIDI input selected.", true);
    return;
  }

  input.onmidimessage = onMidiMessage;
  midiState.activeInput = input;
  midiState.selectedInputId = input.id;
  ui.midiInputSelect.value = input.id;
  setMidiFeedback(`Listening to ${input.name || "MIDI input"}.`);
}

function refreshMidiInputOptions() {
  if (!ui.midiInputSelect) {
    return;
  }

  if (!midiState.access) {
    ui.midiInputSelect.innerHTML = '<option value="">No device</option>';
    ui.midiInputSelect.disabled = true;
    return;
  }

  const inputs = Array.from(midiState.access.inputs.values());
  ui.midiInputSelect.disabled = inputs.length === 0;

  if (inputs.length === 0) {
    detachActiveMidiInput();
    ui.midiInputSelect.innerHTML = '<option value="">No device</option>';
    setMidiFeedback("No MIDI inputs detected. Plug in a keyboard and reconnect.", true);
    return;
  }

  ui.midiInputSelect.innerHTML = inputs
    .map((input) => `<option value="${escapeHtml(input.id)}">${escapeHtml(input.name || input.id)}</option>`)
    .join("");

  if (!midiState.selectedInputId || !midiState.access.inputs.has(midiState.selectedInputId)) {
    midiState.selectedInputId = inputs[0].id;
  }

  attachMidiInputById(midiState.selectedInputId);
}

function updateMidiCaptureButton() {
  if (!ui.midiCaptureBtn) {
    return;
  }

  ui.midiCaptureBtn.classList.toggle("capture-on", midiState.captureToGrid);
  ui.midiCaptureBtn.textContent = midiState.captureToGrid ? "Capture: On" : "Capture: Off";
}

function disconnectMidi() {
  releaseAllMidiNotes();
  detachActiveMidiInput();

  if (midiState.access) {
    midiState.access.onstatechange = null;
  }

  midiState.access = null;
  midiState.selectedInputId = "";

  if (ui.midiConnectBtn) {
    ui.midiConnectBtn.textContent = "Connect MIDI";
  }

  refreshMidiInputOptions();
  setMidiFeedback("MIDI disconnected.");
}

async function connectMidi() {
  if (!supportsWebMidi()) {
    setMidiFeedback("Web MIDI is unavailable in this browser.", true);
    return;
  }

  if (!window.isSecureContext) {
    setMidiFeedback("Web MIDI requires HTTPS or localhost.", true);
    return;
  }

  try {
    await ensureAudioReady();
    midiState.access = await navigator.requestMIDIAccess({ sysex: false });
    midiState.access.onstatechange = () => {
      refreshMidiInputOptions();
    };

    if (ui.midiConnectBtn) {
      ui.midiConnectBtn.textContent = "Disconnect MIDI";
    }

    refreshMidiInputOptions();
    saveIntegrationSettings();
  } catch (error) {
    setMidiFeedback(`MIDI connection failed: ${error.message}`, true);
  }
}

async function toggleMidiConnection() {
  if (midiState.access) {
    disconnectMidi();
    return;
  }

  await connectMidi();
}

function saveIntegrationSettings() {
  try {
    const settings = {
      midi: {
        channel: midiState.channel,
        captureToGrid: midiState.captureToGrid,
        selectedInputId: midiState.selectedInputId,
      },
      yamaha: {
        endpointUrl: ui.yamahaUrlInput ? ui.yamahaUrlInput.value.trim() : "",
        authHeader: ui.yamahaAuthHeaderInput ? ui.yamahaAuthHeaderInput.value.trim() : "",
      },
    };

    localStorage.setItem(INTEGRATION_STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // Persisting integration settings is optional.
  }
}

function loadIntegrationSettings() {
  try {
    const raw = localStorage.getItem(INTEGRATION_STORAGE_KEY);
    if (!raw) {
      return;
    }

    const parsed = JSON.parse(raw);

    if (parsed.midi) {
      midiState.channel = parsed.midi.channel || "all";
      midiState.captureToGrid = Boolean(parsed.midi.captureToGrid);
      midiState.selectedInputId = parsed.midi.selectedInputId || "";
      if (ui.midiChannelSelect) {
        ui.midiChannelSelect.value = midiState.channel;
      }
    }

    if (parsed.yamaha) {
      if (ui.yamahaUrlInput) {
        ui.yamahaUrlInput.value = parsed.yamaha.endpointUrl || "";
      }
      if (ui.yamahaAuthHeaderInput) {
        ui.yamahaAuthHeaderInput.value = parsed.yamaha.authHeader || "";
      }
    }
  } catch {
    // Ignore invalid settings snapshots.
  }
}

function renderYamahaPayload(payload) {
  if (!ui.yamahaResult) {
    return;
  }

  ui.yamahaResult.textContent = JSON.stringify(payload, null, 2);
}

function readValueAtPath(objectValue, path) {
  let current = objectValue;

  for (const segment of path) {
    if (!current || typeof current !== "object" || !(segment in current)) {
      return null;
    }

    current = current[segment];
  }

  return current;
}

function pickFirstNumericValue(objectValue, candidatePaths) {
  for (const path of candidatePaths) {
    const value = readValueAtPath(objectValue, path);
    const numeric = asNumber(value, Number.NaN);
    if (Number.isFinite(numeric) && numeric > 0) {
      return numeric;
    }
  }

  return null;
}

function pickFirstStringValue(objectValue, candidatePaths) {
  for (const path of candidatePaths) {
    const value = readValueAtPath(objectValue, path);
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }

  return null;
}

function extractYamahaInsights(payload) {
  const bpm = pickFirstNumericValue(payload, [
    ["bpm"],
    ["tempo"],
    ["analysis", "bpm"],
    ["analysis", "tempo"],
    ["result", "bpm"],
    ["music", "bpm"],
    ["metadata", "bpm"],
  ]);

  const key = pickFirstStringValue(payload, [
    ["key"],
    ["analysis", "key"],
    ["result", "key"],
    ["metadata", "key"],
  ]);

  let sectionsCount = null;
  const sections =
    readValueAtPath(payload, ["sections"]) ||
    readValueAtPath(payload, ["analysis", "sections"]) ||
    readValueAtPath(payload, ["result", "sections"]);

  if (Array.isArray(sections)) {
    sectionsCount = sections.length;
  }

  return { bpm, key, sectionsCount };
}

function applyYamahaInsights(payload) {
  const insights = extractYamahaInsights(payload);
  const parts = [];

  if (Number.isFinite(insights.bpm)) {
    parts.push(`BPM ${insights.bpm.toFixed(1)}`);
  }

  if (insights.key) {
    parts.push(`Key ${insights.key}`);
  }

  if (Number.isFinite(insights.sectionsCount)) {
    parts.push(`${insights.sectionsCount} sections`);
  }

  if (parts.length === 0) {
    setYamahaFeedback("Analysis returned data, but no BPM/key fields were detected.");
  } else {
    setYamahaFeedback(`Detected: ${parts.join(" | ")}.`);
  }

  if (ui.yamahaApplyTempoBtn) {
    ui.yamahaApplyTempoBtn.disabled = !Number.isFinite(insights.bpm);
  }
}

function applyYamahaTempoFromLastAnalysis() {
  if (!yamahaLastAnalysis) {
    setStatus("No Yamaha analysis to apply yet.", true);
    return;
  }

  const insights = extractYamahaInsights(yamahaLastAnalysis);
  if (!Number.isFinite(insights.bpm)) {
    setStatus("No BPM value found in the last analysis response.", true);
    return;
  }

  const clampedBpm = clamp(insights.bpm, 50, 180);
  pushHistorySilent();
  ui.tempoSlider.value = String(Math.round(clampedBpm));
  applyTransportControls();
  setStatus(`Applied BPM ${clampedBpm.toFixed(1)} from Yamaha analysis.`);
}

async function analyzeWithYamaha() {
  if (!ui.yamahaUrlInput || !ui.yamahaAudioInput) {
    return;
  }

  const endpointUrl = ui.yamahaUrlInput.value.trim();
  const authHeader = ui.yamahaAuthHeaderInput ? ui.yamahaAuthHeaderInput.value.trim() : "";
  const authValue = ui.yamahaAuthValueInput ? ui.yamahaAuthValueInput.value.trim() : "";
  const file = ui.yamahaAudioInput.files ? ui.yamahaAudioInput.files[0] : null;

  if (!endpointUrl) {
    setYamahaFeedback("Enter a Yamaha-compatible analysis endpoint URL.", true);
    return;
  }

  if (!file) {
    setYamahaFeedback("Choose an audio file before analyzing.", true);
    return;
  }

  try {
    new URL(endpointUrl);
  } catch {
    setYamahaFeedback("Endpoint URL is invalid.", true);
    return;
  }

  const originalLabel = ui.yamahaAnalyzeBtn ? ui.yamahaAnalyzeBtn.textContent : "Analyze Audio";

  try {
    if (ui.yamahaAnalyzeBtn) {
      ui.yamahaAnalyzeBtn.disabled = true;
      ui.yamahaAnalyzeBtn.textContent = "Analyzing...";
    }

    setYamahaFeedback("Uploading audio for Yamaha analysis...");
    saveIntegrationSettings();

    const formData = new FormData();
    formData.append("file", file, file.name);
    formData.append("filename", file.name);
    formData.append("tempo_hint", ui.tempoSlider.value);

    const headers = {};
    if (authHeader && authValue) {
      headers[authHeader] = authValue;
    }

    const response = await fetch(endpointUrl, {
      method: "POST",
      headers,
      body: formData,
    });

    const contentType = response.headers.get("content-type") || "";
    const payload = contentType.includes("application/json") ? await response.json() : { raw: await response.text() };

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${JSON.stringify(payload).slice(0, 180)}`);
    }

    yamahaLastAnalysis = payload;
    renderYamahaPayload(payload);
    applyYamahaInsights(payload);
    setStatus("Yamaha analysis complete.");
  } catch (error) {
    const isNetworkLike = error instanceof TypeError;
    const suffix = isNetworkLike ? " Check CORS, endpoint access, and API auth." : "";
    setYamahaFeedback(`Analysis failed: ${error.message}.${suffix}`, true);
    setStatus(`Yamaha analysis failed: ${error.message}`, true);
  } finally {
    if (ui.yamahaAnalyzeBtn) {
      ui.yamahaAnalyzeBtn.disabled = false;
      ui.yamahaAnalyzeBtn.textContent = originalLabel;
    }
  }
}

function replaceTrackSynth(track, instrumentName) {
  const preset = INSTRUMENTS[instrumentName];
  if (!preset || !track) {
    return;
  }

  track.synth.dispose();
  if (track.sampler) {
    track.sampler.dispose();
  }
  track.instrument = instrumentName;
  track.synth = createSynthForInstrument(instrumentName);
  track.synth.connect(track.channel);
  initializeTrackSampler(track);
}

function setNoteOnPattern(track, noteName, stepIndex, velocity) {
  const noteIndex = NOTE_NAMES.indexOf(noteName);
  if (noteIndex < 0 || stepIndex < 0 || stepIndex >= MAX_STEPS) {
    return;
  }

  track.pattern[noteIndex][stepIndex] = clamp01(velocity);
}

function seedStarterPattern() {
  const strings = getTrackById(selectedTrackId);
  const brass = tracks[1];
  const woodwinds = tracks[2];
  const choir = tracks[3];
  const percussion = tracks[4];

  if (!strings || !brass || !woodwinds || !choir || !percussion) {
    return;
  }

  const progression = [
    ["C4", "E4", "G4"],
    ["A4", "C5", "E5"],
    ["F4", "A4", "C5"],
    ["G4", "B4", "D5"],
  ];

  progression.forEach((chord, phraseIndex) => {
    const baseStep = phraseIndex * 8;
    for (let offset = 0; offset < 8; offset += 2) {
      const velocity = 0.55 + (offset % 4 === 0 ? 0.08 : 0);
      chord.forEach((note) => {
        setNoteOnPattern(strings, note, baseStep + offset, velocity);
      });
    }
  });

  ["C5", "A4", "F4", "G4"].forEach((note, phraseIndex) => {
    const baseStep = phraseIndex * 8;
    setNoteOnPattern(brass, note, baseStep, 0.9);
    setNoteOnPattern(brass, note, baseStep + 4, 0.72);
  });

  ["E5", "G5", "A5", "G5", "E5", "D5", "C5", "D5"].forEach((note, index) => {
    const step = index * 2;
    setNoteOnPattern(woodwinds, note, step, 0.68);
    setNoteOnPattern(woodwinds, note, step + 16, 0.68);
  });

  ["G4", "A4", "E5", "D5"].forEach((note, phraseIndex) => {
    const baseStep = phraseIndex * 8;
    for (let offset = 0; offset < 8; offset += 4) {
      setNoteOnPattern(choir, note, baseStep + offset, 0.63);
    }
  });

  for (let step = 0; step < 32; step += 2) {
    const isDownbeat = step % 8 === 0;
    setNoteOnPattern(percussion, isDownbeat ? "C4" : "G4", step, isDownbeat ? 0.95 : 0.72);
  }
}

function applyTemplatePattern(templateId) {
  const template = String(templateId || "cinematic").toLowerCase();
  const strings = tracks[0];
  const brass = tracks[1];
  const woodwinds = tracks[2];
  const choir = tracks[3];
  const percussion = tracks[4];

  if (!strings || !brass || !woodwinds || !choir || !percussion) {
    return;
  }

  for (const track of tracks) {
    track.pattern = createPattern();
  }

  switch (template) {
    case "pop": {
      const chords = ["C4", "G4", "A4", "F4"];
      for (let bar = 0; bar < loopSteps / 4; bar += 1) {
        const step = bar * 4;
        const note = chords[bar % chords.length];
        setNoteOnPattern(strings, note, step, 0.78);
        setNoteOnPattern(strings, Tone.Frequency(note).transpose(7).toNote(), step + 2, 0.62);
      }
      for (let step = 0; step < loopSteps; step += 2) {
        setNoteOnPattern(percussion, step % 8 === 0 ? "C4" : "G4", step, 0.86);
      }
      for (let step = 2; step < loopSteps; step += 8) {
        setNoteOnPattern(woodwinds, "E5", step, 0.58);
      }
      break;
    }
    case "lofi": {
      for (let step = 0; step < loopSteps; step += 4) {
        setNoteOnPattern(strings, "C4", step, 0.46);
        setNoteOnPattern(strings, "E4", step, 0.42);
        setNoteOnPattern(strings, "G4", step, 0.4);
      }
      for (let step = 0; step < loopSteps; step += 4) {
        setNoteOnPattern(percussion, "C4", step, 0.74);
      }
      for (let step = 2; step < loopSteps; step += 4) {
        setNoteOnPattern(percussion, "G4", step, 0.58);
      }
      for (let step = 3; step < loopSteps; step += 8) {
        setNoteOnPattern(choir, "D5", step, 0.48);
      }
      break;
    }
    case "drill": {
      for (let step = 0; step < loopSteps; step += 4) {
        const bass = step % 8 === 0 ? "C4" : "A3";
        setNoteOnPattern(brass, bass, step, 0.9);
      }
      for (let step = 0; step < loopSteps; step += 2) {
        setNoteOnPattern(percussion, step % 8 === 0 ? "C4" : "G4", step, step % 8 === 0 ? 0.92 : 0.66);
      }
      for (let step = 1; step < loopSteps; step += 4) {
        setNoteOnPattern(woodwinds, "E5", step, 0.6);
      }
      for (let step = 7; step < loopSteps; step += 8) {
        setNoteOnPattern(woodwinds, "D5", step, 0.67);
      }
      break;
    }
    case "cinematic":
    default:
      seedStarterPattern();
      break;
  }
}

function updateTransportOptionButtons() {
  if (ui.metronomeBtn) {
    ui.metronomeBtn.textContent = `Metronome: ${metronomeEnabled ? "On" : "Off"}`;
    ui.metronomeBtn.classList.toggle("btn-primary", metronomeEnabled);
  }

  if (ui.countInBtn) {
    ui.countInBtn.textContent = `Count-In: ${countInEnabled ? "On" : "Off"}`;
    ui.countInBtn.classList.toggle("btn-primary", countInEnabled);
  }
}

function toggleMetronome() {
  metronomeEnabled = !metronomeEnabled;
  updateTransportOptionButtons();
  setStatus(`Metronome ${metronomeEnabled ? "enabled" : "disabled"}.`);
}

function toggleCountIn() {
  countInEnabled = !countInEnabled;
  updateTransportOptionButtons();
  setStatus(`Count-in ${countInEnabled ? "enabled" : "disabled"}.`);
}

function applyTemplateFromSelect() {
  const template = ui.templateSelect ? ui.templateSelect.value : "cinematic";
  pushHistorySilent();
  applyTemplatePattern(template);
  renderAll();
  setStatus(`Applied ${template} template.`);
}

function createStarterProject() {
  tracks.forEach((track) => disposeTrack(track));

  tracks = [
    createTrack({ name: "Strings Bed", instrument: "Strings" }),
    createTrack({ name: "Brass Swell", instrument: "Brass" }),
    createTrack({ name: "Winds Motif", instrument: "Woodwinds" }),
    createTrack({ name: "Choir Air", instrument: "Choir" }),
    createTrack({ name: "Timpani Pulse", instrument: "Percussion" }),
  ];

  selectedTrackId = tracks[0].id;
  seedStarterPattern();
  applySoloState();
}

function renderTrackList() {
  ui.trackList.innerHTML = "";

  for (const track of tracks) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `track-chip${track.id === selectedTrackId ? " selected" : ""}`;
    button.dataset.trackId = track.id;
    button.style.setProperty("--track-color", track.color);
    button.textContent = `${track.name} (${track.instrument})`;
    ui.trackList.append(button);
  }
}

function renderPhraseLibrary() {
  if (!ui.phraseLibrary) {
    return;
  }

  ui.phraseLibrary.innerHTML = CHORD_LIBRARY.map((chord) => {
    return `
      <article class="phrase-card" draggable="true" data-chord-id="${chord.id}" aria-label="${escapeHtml(chord.name)}">
        <div class="phrase-card-title">${escapeHtml(chord.name)}</div>
        <div class="phrase-card-note">${escapeHtml(chord.description)}</div>
        <button class="btn btn-subtle" type="button" data-chord-action="preview" data-chord-id="${chord.id}">Preview</button>
      </article>
    `;
  }).join("");

  ui.phraseLibrary.classList.toggle("compact", chordPaletteCompact);
}

async function previewChord(chordId) {
  const chord = getChordById(chordId);
  const track = getSelectedTrack();

  if (!chord || !track) {
    return;
  }

  await ensureAudioReady();

  const baseMidi = 60;
  const rootMidi = baseMidi + asNumber(chord.rootPitchClass, 0);
  const stepSeconds = Tone.Time("16n").toSeconds();
  const startTime = Tone.now() + 0.05;
  const minMidi = TOP_MIDI - (NOTE_COUNT - 1);
  const maxMidi = TOP_MIDI;

  for (const event of chord.events) {
    const targetMidi = clamp(rootMidi + asNumber(event.semitone, 0), minMidi, maxMidi);
    const noteName = Tone.Frequency(targetMidi, "midi").toNote();
    const velocity = clamp01(asNumber(event.velocity, DEFAULT_NOTE_VELOCITY));
    triggerTrackAttackRelease(
      track,
      noteName,
      stepSeconds * 0.95,
      startTime + asNumber(event.step, 0) * stepSeconds,
      velocity,
      LIVE_VELOCITY_SCALE
    );
  }

  setStatus(`Previewing chord: ${chord.name}.`);
}

function insertChordAtCell(chordId, anchorNoteIndex, anchorStep) {
  const chord = getChordById(chordId);
  const track = getSelectedTrack();

  if (!chord || !track) {
    return;
  }

  const anchorNoteName = NOTE_NAMES[anchorNoteIndex];
  if (!anchorNoteName) {
    return;
  }

  pushHistorySilent();

  const anchorMidi = Math.round(Tone.Frequency(anchorNoteName).toMidi());
  const rootMidi = nearestMidiForPitchClass(anchorMidi, asNumber(chord.rootPitchClass, 0));
  let written = 0;

  for (const event of chord.events) {
    const step = (anchorStep + asNumber(event.step, 0)) % loopSteps;
    const targetMidi = Math.round(rootMidi + asNumber(event.semitone, 0));
    const targetNoteIndex = midiNoteToPatternIndex(targetMidi);

    if (targetNoteIndex < 0) {
      continue;
    }

    track.pattern[targetNoteIndex][step] = clamp01(asNumber(event.velocity, DEFAULT_NOTE_VELOCITY));
    written += 1;
  }

  renderPianoRoll();

  if (written > 0) {
    setStatus(`Inserted chord ${chord.name} into ${track.name}.`);
  } else {
    setStatus(`Chord ${chord.name} could not fit in this note range.`, true);
  }
}

function onPhraseLibraryDragStart(event) {
  const card = event.target instanceof HTMLElement ? event.target.closest(".phrase-card") : null;
  if (!(card instanceof HTMLElement) || !event.dataTransfer) {
    return;
  }

  const chordId = card.dataset.chordId;
  if (!chordId) {
    return;
  }

  event.dataTransfer.setData(CHORD_DRAG_MIME, chordId);
  event.dataTransfer.setData("text/plain", chordId);
  event.dataTransfer.effectAllowed = "copy";
}

function onPhraseLibraryClick(event) {
  const button = event.target instanceof HTMLElement ? event.target.closest("button[data-chord-action='preview']") : null;
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }

  const chordId = button.dataset.chordId;
  if (!chordId) {
    return;
  }

  previewChord(chordId);
}

function onPianoRollDragOver(event) {
  const cell = event.target instanceof HTMLElement ? event.target.closest(".step-cell") : null;
  if (!cell || !event.dataTransfer) {
    return;
  }

  const transferTypes = Array.from(event.dataTransfer.types);
  const hasChordPayload = transferTypes.includes(CHORD_DRAG_MIME) || transferTypes.includes("text/plain");
  if (hasChordPayload) {
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  }
}

function onPianoRollDrop(event) {
  const cell = event.target instanceof HTMLElement ? event.target.closest(".step-cell") : null;
  if (!(cell instanceof HTMLElement) || !event.dataTransfer) {
    return;
  }

  const chordId = event.dataTransfer.getData(CHORD_DRAG_MIME) || event.dataTransfer.getData("text/plain");
  if (!chordId) {
    return;
  }

  const noteIndex = asNumber(cell.dataset.note, -1);
  const stepIndex = asNumber(cell.dataset.step, -1);
  if (noteIndex < 0 || stepIndex < 0) {
    return;
  }

  event.preventDefault();
  insertChordAtCell(chordId, noteIndex, stepIndex);
}

function renderPianoRoll() {
  const selected = getSelectedTrack();
  if (!selected) {
    ui.pianoRoll.innerHTML = "";
    return;
  }

  const rows = [];

  rows.push('<div class="timeline-row">');
  rows.push('<div class="timeline-label">Bars</div>');
  rows.push(`<div class="timeline-steps" style="--steps:${loopSteps}">`);

  for (let step = 0; step < loopSteps; step += 1) {
    const markerClass = `timeline-step${(step + 1) % 4 === 0 ? " bar" : ""}${step === playheadStep ? " is-playhead" : ""}`;
    const label = step % 4 === 0 ? Math.floor(step / 4) + 1 : "";
    rows.push(`<div class="${markerClass}" data-step="${step}">${label}</div>`);
  }

  rows.push("</div>");
  rows.push("</div>");

  for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
    const noteName = NOTE_NAMES[noteIndex];
    const rowClass = isBlackKey(noteName) ? "note-label black" : "note-label";
    rows.push('<div class="note-row">');
    rows.push(`<div class="${rowClass}"><span>${noteName}</span></div>`);
    rows.push(`<div class="note-steps" style="--steps:${loopSteps}">`);

    for (let step = 0; step < loopSteps; step += 1) {
      const velocity = selected.pattern[noteIndex][step];
      
      let isGhost = false;
      if (velocity === 0) {
        for (const track of tracks) {
          if (track.id !== selected.id && track.pattern[noteIndex][step] > 0) {
            isGhost = true;
            break;
          }
        }
      }

      const classNames = ["step-cell"];
      if ((step + 1) % 4 === 0) {
        classNames.push("bar");
      }
      if (velocity > 0) {
        classNames.push("active");
      } else if (isGhost) {
        classNames.push("ghost");
      }
      
      if (velocity >= 0.95) {
        classNames.push("accent");
      }
      if (step === playheadStep) {
        classNames.push("is-playhead");
      }

      rows.push(
        `<div class="${classNames.join(" ")}" data-note="${noteIndex}" data-step="${step}" title="${noteName} / step ${step + 1}"></div>`
      );
    }

    rows.push("</div>");
    rows.push("</div>");
  }

  ui.pianoRoll.innerHTML = rows.join("");
}

function buildInstrumentSelectOptions(selectedInstrument) {
  let markup = "";
  for (const [groupName, instruments] of Object.entries(INSTRUMENT_GROUPS)) {
    markup += `<optgroup label="${groupName}">`;
    for (const instrumentName of instruments) {
      if (INSTRUMENTS[instrumentName]) {
        const selectedFlag = instrumentName === selectedInstrument ? " selected" : "";
        markup += `<option value="${instrumentName}"${selectedFlag}>${instrumentName}</option>`;
      }
    }
    markup += `</optgroup>`;
  }
  return markup;
}

function renderMixer() {
  const markup = tracks
    .map((track) => {
      const volumeText = `${track.volume.toFixed(1)} dB`;
      const panText = `${Math.round(track.pan * 100)}%`;
      const reverbText = `${Math.round(track.reverbSend * 100)}%`;
      const delayText = `${Math.round(track.delaySend * 100)}%`;
      const selectedClass = track.id === selectedTrackId ? " selected-track" : "";

      return `
        <article class="mixer-strip${selectedClass}" data-track-id="${track.id}">
          <div class="strip-head">
            <span class="strip-name"><span class="strip-dot" style="--track-color:${track.color}"></span>${escapeHtml(track.name)}</span>
          </div>
          <div class="strip-controls">
            <label>Instrument</label>
            <select data-action="instrument">${buildInstrumentSelectOptions(track.instrument)}</select>

            <label>Volume <output data-readout="volume">${volumeText}</output></label>
            <input type="range" min="-36" max="6" step="0.5" value="${track.volume}" data-action="volume" />

            <label>Pan <output data-readout="pan">${panText}</output></label>
            <input type="range" min="-1" max="1" step="0.01" value="${track.pan}" data-action="pan" />

            <label>Reverb Send <output data-readout="reverb">${reverbText}</output></label>
            <input type="range" min="0" max="1" step="0.01" value="${track.reverbSend}" data-action="reverb" />

            <label>Delay Send <output data-readout="delay">${delayText}</output></label>
            <input type="range" min="0" max="1" step="0.01" value="${track.delaySend}" data-action="delay" />

            <div class="strip-toggle-row">
              <button class="toggle${track.mute ? " active" : ""}" type="button" data-action="mute">Mute</button>
              <button class="toggle solo${track.solo ? " active" : ""}" type="button" data-action="solo">Solo</button>
            </div>

            <div class="strip-actions">
              <button class="btn btn-subtle" type="button" data-action="select">Edit</button>
              <button class="btn btn-danger" type="button" data-action="remove">Remove</button>
            </div>
          </div>
        </article>
      `;
    })
    .join("");

  ui.mixerStrips.innerHTML = markup;
}

function renderAll() {
  renderTrackList();
  renderPhraseLibrary();
  renderPianoRoll();
  renderMixer();
}

function applyTransportControls() {
  const bpm = clamp(asNumber(ui.tempoSlider.value, 102), 50, 180);
  const swing = clamp(asNumber(ui.swingSlider.value, 0), 0, 0.6);

  Tone.Transport.bpm.rampTo(bpm, 0.05);
  Tone.Transport.swing = swing;
  Tone.Transport.swingSubdivision = "8n";

  ui.tempoValue.textContent = `${Math.round(bpm)} BPM`;
  ui.swingValue.textContent = `${Math.round(swing * 100)}%`;
}

function refreshPlayButton() {
  ui.playBtn.textContent = Tone.Transport.state === "started" ? "Pause" : "Play";
}

function scheduleSequencer() {
  if (sequenceEventId !== null) {
    Tone.Transport.clear(sequenceEventId);
  }

  sequenceEventId = Tone.Transport.scheduleRepeat((time) => {
    const isQuarterNote = transportTick % 4 === 0;
    const beatInBar = Math.floor((transportTick % 16) / 4);

    if (metronomeEnabled && isQuarterNote) {
      if (beatInBar === 0) {
        fx.metronomeHigh.triggerAttackRelease("C5", "32n", time, 0.9);
      } else {
        fx.metronomeLow.triggerAttackRelease("G4", "32n", time, 0.68);
      }
    }

    if (countInStepsRemaining > 0) {
      countInStepsRemaining -= 1;
      transportTick += 1;

      if (countInStepsRemaining === 0) {
        setStatus("Count-in complete. Playback started.");
        setPlayhead(0);
      } else {
        setPlayhead(null);
      }

      return;
    }

    const step = currentStep % loopSteps;

    for (const track of tracks) {
      for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
        const velocity = track.pattern[noteIndex][step];
        if (velocity > 0) {
          const noteDuration = track.instrument === "Percussion" ? "32n" : "16n";
          triggerTrackAttackRelease(track, NOTE_NAMES[noteIndex], noteDuration, time, velocity, SEQUENCER_VELOCITY_SCALE);
        }
      }
    }

    Tone.Draw.schedule(() => {
      setPlayhead(step);
    }, time);

    currentStep = (currentStep + 1) % loopSteps;
    transportTick += 1;
  }, "16n");
}

function setPlayhead(step) {
  if (playheadStep !== null) {
    const previous = ui.pianoRoll.querySelectorAll(`[data-step="${playheadStep}"]`);
    for (const node of previous) {
      node.classList.remove("is-playhead");
    }
  }

  playheadStep = step;

  if (playheadStep === null) {
    return;
  }

  const current = ui.pianoRoll.querySelectorAll(`[data-step="${playheadStep}"]`);
  for (const node of current) {
    node.classList.add("is-playhead");
  }
}

function paintGridCell(cell, value) {
  cell.classList.toggle("active", value > 0);
  cell.classList.toggle("accent", value >= 0.95);
}

function updatePatternAtCell(cell, value, audition = false) {
  const selected = getSelectedTrack();
  if (!selected) {
    return false;
  }

  const noteIndex = asNumber(cell.dataset.note, -1);
  const stepIndex = asNumber(cell.dataset.step, -1);

  if (noteIndex < 0 || stepIndex < 0 || stepIndex >= loopSteps) {
    return false;
  }

  if (selected.pattern[noteIndex][stepIndex] === value) {
    return false;
  }

  selected.pattern[noteIndex][stepIndex] = value;
  paintGridCell(cell, value);

  if (audition && value > 0 && audioReady) {
    triggerTrackAttackRelease(selected, NOTE_NAMES[noteIndex], "16n", undefined, value, LIVE_VELOCITY_SCALE);
  }

  return true;
}

function clearSelectedTrackPattern() {
  const selected = getSelectedTrack();
  if (!selected) {
    return;
  }

  pushHistorySilent();
  selected.pattern = createPattern();
  renderPianoRoll();
  setStatus(`Cleared notes from ${selected.name}.`);
}

function addTrack() {
  pushHistorySilent();

  const requestedName = ui.trackNameInput.value.trim();
  const instrument = ui.instrumentSelect.value;
  const generatedName = requestedName || `${instrument} ${tracks.length + 1}`;

  const track = createTrack({
    name: generatedName,
    instrument,
    color: nextColor(),
  });

  tracks.push(track);
  selectedTrackId = track.id;
  applySoloState();
  renderAll();

  setStatus(`Added ${generatedName}.`);
}

function removeTrack(trackId) {
  if (tracks.length <= 1) {
    setStatus("At least one track must remain in the project.", true);
    return;
  }

  const index = tracks.findIndex((track) => track.id === trackId);
  if (index < 0) {
    return;
  }

  pushHistorySilent();
  const [removed] = tracks.splice(index, 1);
  releaseAllMidiNotes();
  disposeTrack(removed);

  if (selectedTrackId === trackId) {
    selectedTrackId = tracks[Math.max(0, index - 1)].id;
  }

  applySoloState();
  renderAll();
  setStatus(`Removed ${removed.name}.`);
}

function projectSnapshot() {
  return {
    version: 1,
    bpm: clamp(asNumber(ui.tempoSlider.value, 102), 50, 180),
    swing: clamp(asNumber(ui.swingSlider.value, 0), 0, 0.6),
    loopSteps,
    selectedTrackId,
    metronomeEnabled,
    countInEnabled,
    tracks: tracks.map((track) => ({
      id: track.id,
      name: track.name,
      instrument: track.instrument,
      color: track.color,
      volume: track.volume,
      pan: track.pan,
      mute: track.mute,
      solo: track.solo,
      reverbSend: track.reverbSend,
      delaySend: track.delaySend,
      pattern: track.pattern.map((row) => row.slice(0, MAX_STEPS)),
    })),
  };
}

function loadProject(project, sourceLabel = "project", options = {}) {
  if (!project || !Array.isArray(project.tracks) || project.tracks.length === 0) {
    throw new Error("Project format is invalid.");
  }

  const skipHistoryReset = Boolean(options.skipHistoryReset);

  Tone.Transport.stop();
  releaseAllMidiNotes();
  refreshPlayButton();
  currentStep = 0;
  transportTick = 0;
  countInStepsRemaining = 0;

  for (const track of tracks) {
    disposeTrack(track);
  }

  tracks = project.tracks.map((trackData, index) =>
    createTrack({
      ...trackData,
      color: trackData.color || TRACK_COLORS[index % TRACK_COLORS.length],
    })
  );

  selectedTrackId = getTrackById(project.selectedTrackId) ? project.selectedTrackId : tracks[0].id;

  const allowedLoopSteps = new Set([16, 32, 48, 64]);
  const requestedLoop = asNumber(project.loopSteps, 32);
  loopSteps = allowedLoopSteps.has(requestedLoop) ? requestedLoop : 32;
  ui.loopSelect.value = String(loopSteps);

  metronomeEnabled = Boolean(project.metronomeEnabled);
  countInEnabled = Boolean(project.countInEnabled);
  updateTransportOptionButtons();

  ui.tempoSlider.value = String(clamp(asNumber(project.bpm, 102), 50, 180));
  ui.swingSlider.value = String(clamp(asNumber(project.swing, 0), 0, 0.6));

  applyTransportControls();
  applySoloState();
  scheduleSequencer();
  renderAll();
  setPlayhead(0);

  if (!skipHistoryReset) {
    historyState.undoStack = [];
    historyState.redoStack = [];
    updateHistoryButtons();
  }

  setStatus(`Loaded ${sourceLabel}.`);
}

function saveLocalProject() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(projectSnapshot()));
    setStatus("Saved project to local storage.");
  } catch (error) {
    setStatus(`Local save failed: ${error.message}`, true);
  }
}

function loadLocalProject() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      setStatus("No local project found yet.", true);
      return;
    }

    loadProject(JSON.parse(raw), "local project");
  } catch (error) {
    setStatus(`Local load failed: ${error.message}`, true);
  }
}

function exportProject() {
  try {
    const snapshot = projectSnapshot();
    const blob = new Blob([JSON.stringify(snapshot, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    const stamp = new Date().toISOString().slice(0, 10);
    link.href = url;
    link.download = `orchestrion-score-${stamp}.json`;
    link.click();

    URL.revokeObjectURL(url);
    setStatus("Exported project JSON.");
  } catch (error) {
    setStatus(`Export failed: ${error.message}`, true);
  }
}

function sanitizeFileName(value) {
  return String(value || "track")
    .trim()
    .replace(/[^a-zA-Z0-9-_ ]/g, "")
    .replace(/\s+/g, "-")
    .slice(0, 48) || "track";
}

function hasTrackNotes(track) {
  if (!track || !Array.isArray(track.pattern)) {
    return false;
  }

  for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
    for (let stepIndex = 0; stepIndex < loopSteps; stepIndex += 1) {
      if (asNumber(track.pattern[noteIndex][stepIndex], 0) > 0) {
        return true;
      }
    }
  }

  return false;
}

function noteNameToMusicXmlPitch(noteName) {
  const midi = Math.round(Tone.Frequency(noteName).toMidi());
  const pitchIndex = ((midi % 12) + 12) % 12;
  const octave = Math.floor(midi / 12) - 1;

  const pitchMap = {
    0: { step: "C", alter: 0 },
    1: { step: "C", alter: 1 },
    2: { step: "D", alter: 0 },
    3: { step: "D", alter: 1 },
    4: { step: "E", alter: 0 },
    5: { step: "F", alter: 0 },
    6: { step: "F", alter: 1 },
    7: { step: "G", alter: 0 },
    8: { step: "G", alter: 1 },
    9: { step: "A", alter: 0 },
    10: { step: "A", alter: 1 },
    11: { step: "B", alter: 0 },
  };

  return {
    ...pitchMap[pitchIndex],
    octave,
  };
}

function buildMusicXmlForTrack(track, bpm = 102) {
  const stepsPerMeasure = 16;
  const totalMeasures = Math.max(1, Math.ceil(loopSteps / stepsPerMeasure));
  const program = INSTRUMENT_MIDI_PROGRAMS[track.instrument] || 1;
  const scorePartName = escapeHtml(track.name || track.instrument || "Track");
  const instrumentName = escapeHtml(track.instrument || "Instrument");

  const measures = [];

  for (let measureIndex = 0; measureIndex < totalMeasures; measureIndex += 1) {
    const measureStart = measureIndex * stepsPerMeasure;
    const measureEnd = Math.min(loopSteps, measureStart + stepsPerMeasure);
    const noteLines = [];

    if (measureIndex === 0) {
      noteLines.push(
        "<attributes>",
        "<divisions>4</divisions>",
        "<key><fifths>0</fifths></key>",
        "<time><beats>4</beats><beat-type>4</beat-type></time>",
        "<clef><sign>G</sign><line>2</line></clef>",
        "</attributes>",
        "<direction placement=\"above\">",
        "<direction-type><metronome><beat-unit>quarter</beat-unit><per-minute>" + Math.round(bpm) + "</per-minute></metronome></direction-type>",
        "<sound tempo=\"" + Math.round(bpm) + "\"/>",
        "</direction>"
      );
    }

    for (let stepIndex = measureStart; stepIndex < measureEnd; stepIndex += 1) {
      const notesAtStep = [];

      for (let noteIndex = 0; noteIndex < NOTE_COUNT; noteIndex += 1) {
        if (asNumber(track.pattern[noteIndex][stepIndex], 0) > 0) {
          notesAtStep.push(NOTE_NAMES[noteIndex]);
        }
      }

      if (notesAtStep.length === 0) {
        noteLines.push("<note><rest/><duration>1</duration><voice>1</voice><type>16th</type></note>");
        continue;
      }

      notesAtStep.forEach((noteName, notePosition) => {
        const pitch = noteNameToMusicXmlPitch(noteName);
        noteLines.push("<note>");
        if (notePosition > 0) {
          noteLines.push("<chord/>");
        }
        noteLines.push("<pitch>");
        noteLines.push(`<step>${pitch.step}</step>`);
        if (pitch.alter !== 0) {
          noteLines.push(`<alter>${pitch.alter}</alter>`);
        }
        noteLines.push(`<octave>${pitch.octave}</octave>`);
        noteLines.push("</pitch>");
        noteLines.push("<duration>1</duration>");
        noteLines.push("<voice>1</voice>");
        noteLines.push("<type>16th</type>");
        noteLines.push("</note>");
      });
    }

    measures.push(`<measure number=\"${measureIndex + 1}\">${noteLines.join("")}</measure>`);
  }

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>',
    '<!DOCTYPE score-partwise PUBLIC "-//Recordare//DTD MusicXML 3.1 Partwise//EN" "http://www.musicxml.org/dtds/partwise.dtd">',
    '<score-partwise version="3.1">',
    "<part-list>",
    '<score-part id="P1">',
    `<part-name>${scorePartName}</part-name>`,
    '<score-instrument id="P1-I1">',
    `<instrument-name>${instrumentName}</instrument-name>`,
    "</score-instrument>",
    '<midi-instrument id="P1-I1">',
    "<midi-channel>1</midi-channel>",
    `<midi-program>${program}</midi-program>`,
    "</midi-instrument>",
    "</score-part>",
    "</part-list>",
    '<part id="P1">',
    measures.join(""),
    "</part>",
    "</score-partwise>",
  ].join("\n");
}

function downloadTextFile(content, fileName, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function exportSheetMusicPerTrack() {
  const activeTracks = tracks.filter((track) => hasTrackNotes(track));

  if (activeTracks.length === 0) {
    setStatus("No notes found. Add notes before exporting sheet music.", true);
    return;
  }

  const bpm = clamp(asNumber(ui.tempoSlider.value, 102), 50, 180);
  const stamp = new Date().toISOString().slice(0, 10);

  activeTracks.forEach((track) => {
    const xml = buildMusicXmlForTrack(track, bpm);
    const safeTrackName = sanitizeFileName(track.name || track.instrument || "track");
    downloadTextFile(xml, `${safeTrackName}-${stamp}.musicxml`, "application/vnd.recordare.musicxml+xml");
  });

  setStatus(`Exported ${activeTracks.length} sheet file(s). MusicXML files are compatible with Flat and Noteflight import APIs.`);
}

async function importProjectFile(file) {
  if (!file) {
    return;
  }

  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    loadProject(parsed, file.name);
  } catch (error) {
    setStatus(`Import failed: ${error.message}`, true);
  }
}

async function ensureAudioReady() {
  if (audioReady) {
    return;
  }

  await Tone.start();
  if (typeof fx.reverb.generate === "function") {
    await fx.reverb.generate();
  }

  audioReady = true;
  setStatus("Audio engine armed. Press Play to perform your arrangement.");
}

async function togglePlayback() {
  try {
    await ensureAudioReady();

    if (Tone.Transport.state === "started") {
      Tone.Transport.pause();
      countInStepsRemaining = 0;
      setStatus("Playback paused.");
    } else {
      for (const track of tracks) {
        track.samplerArmed = track.samplerReady && !track.samplerFailed;
      }

      if (countInEnabled) {
        countInStepsRemaining = 16;
        setPlayhead(null);
        setStatus("Count-in started. Playback will enter after one bar.");
      } else {
        countInStepsRemaining = 0;
      }

      Tone.Transport.start();
      if (!countInEnabled) {
        setStatus("Playback started.");
      }
    }

    refreshPlayButton();
  } catch (error) {
    setStatus(`Playback failed: ${error.message}`, true);
  }
}

function stopPlayback() {
  Tone.Transport.stop();
  currentStep = 0;
  transportTick = 0;
  countInStepsRemaining = 0;
  setPlayhead(0);
  releaseAllMidiNotes();
  for (const track of tracks) {
    track.samplerArmed = track.samplerReady && !track.samplerFailed;
    track.synth.releaseAll();
    if (track.sampler && track.samplerReady && typeof track.sampler.releaseAll === "function") {
      track.sampler.releaseAll();
    }
  }
  refreshPlayButton();
  setStatus("Transport stopped.");
}

function updateMasterMeter() {
  const meterValue = asNumber(fx.meter.getValue(), Number.NEGATIVE_INFINITY);
  const normalized = Number.isFinite(meterValue) ? clamp((meterValue + 48) / 48, 0, 1) : 0;
  ui.masterMeterFill.style.width = `${(normalized * 100).toFixed(1)}%`;
  ui.masterPeak.textContent = Number.isFinite(meterValue) ? `${meterValue.toFixed(1)} dB` : "-∞ dB";
  requestAnimationFrame(updateMasterMeter);
}

function isTextInputElement(target) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }

  return target.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName);
}

function onPianoRollPointerDown(event) {
  const target = event.target instanceof HTMLElement ? event.target.closest(".step-cell") : null;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  event.preventDefault();
  const noteIndex = asNumber(target.dataset.note, -1);
  const stepIndex = asNumber(target.dataset.step, -1);
  const selected = getSelectedTrack();

  if (!selected || noteIndex < 0 || stepIndex < 0 || stepIndex >= loopSteps) {
    return;
  }

  const currentValue = selected.pattern[noteIndex][stepIndex];
  pointerState.isDown = true;
  didPaintInCurrentStroke = false;
  pointerState.value = currentValue > 0 ? 0 : event.shiftKey ? 1 : DEFAULT_NOTE_VELOCITY;
  pointerPaintMode = pointerState.value > 0 ? "add" : "erase";

  pushHistorySilent();

  didPaintInCurrentStroke = updatePatternAtCell(target, pointerState.value, true) || didPaintInCurrentStroke;
}

function onPianoRollPointerOver(event) {
  if (!pointerState.isDown) {
    return;
  }

  const target = event.target instanceof HTMLElement ? event.target.closest(".step-cell") : null;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  didPaintInCurrentStroke = updatePatternAtCell(target, pointerState.value, false) || didPaintInCurrentStroke;
}

function onPianoRollPointerUp() {
  if (pointerState.isDown && !didPaintInCurrentStroke && historyState.undoStack.length > 0) {
    historyState.undoStack.pop();
    updateHistoryButtons();
  }

  pointerState.isDown = false;
  didPaintInCurrentStroke = false;
  pointerPaintMode = null;
}

function onTrackListClick(event) {
  const chip = event.target instanceof HTMLElement ? event.target.closest(".track-chip") : null;
  if (!(chip instanceof HTMLElement)) {
    return;
  }

  const trackId = chip.dataset.trackId;
  if (!trackId) {
    return;
  }

  selectedTrackId = trackId;
  renderAll();
}

function onMixerInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }

  const strip = target.closest(".mixer-strip");
  if (!(strip instanceof HTMLElement)) {
    return;
  }

  const trackId = strip.dataset.trackId;
  const track = getTrackById(trackId);
  if (!track) {
    return;
  }

  const action = target.dataset.action;
  const value = Number(target.value);

  switch (action) {
    case "volume":
      track.volume = clamp(value, -36, 6);
      track.channel.volume.rampTo(track.volume, 0.05);
      strip.querySelector('[data-readout="volume"]').textContent = `${track.volume.toFixed(1)} dB`;
      break;
    case "pan":
      track.pan = clamp(value, -1, 1);
      track.channel.pan.rampTo(track.pan, 0.05);
      strip.querySelector('[data-readout="pan"]').textContent = `${Math.round(track.pan * 100)}%`;
      break;
    case "reverb":
      track.reverbSend = clamp01(value);
      track.reverbSendNode.gain.rampTo(track.reverbSend, 0.05);
      strip.querySelector('[data-readout="reverb"]').textContent = `${Math.round(track.reverbSend * 100)}%`;
      break;
    case "delay":
      track.delaySend = clamp01(value);
      track.delaySendNode.gain.rampTo(track.delaySend, 0.05);
      strip.querySelector('[data-readout="delay"]').textContent = `${Math.round(track.delaySend * 100)}%`;
      break;
    default:
      break;
  }
}

function onMixerChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) {
    return;
  }

  const strip = target.closest(".mixer-strip");
  if (!(strip instanceof HTMLElement)) {
    return;
  }

  const trackId = strip.dataset.trackId;
  const track = getTrackById(trackId);
  if (!track) {
    return;
  }

  if (target.dataset.action === "instrument") {
    pushHistorySilent();
    replaceTrackSynth(track, target.value);
    renderTrackList();
    setStatus(`Changed ${track.name} to ${target.value}.`);
  }
}

function onMixerClick(event) {
  const button = event.target;
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }

  const strip = button.closest(".mixer-strip");
  if (!(strip instanceof HTMLElement)) {
    return;
  }

  const trackId = strip.dataset.trackId;
  const track = getTrackById(trackId);
  if (!track) {
    return;
  }

  const action = button.dataset.action;

  switch (action) {
    case "mute":
      pushHistorySilent();
      track.mute = !track.mute;
      applySoloState();
      renderMixer();
      break;
    case "solo":
      pushHistorySilent();
      track.solo = !track.solo;
      applySoloState();
      renderMixer();
      break;
    case "select":
      selectedTrackId = track.id;
      renderAll();
      break;
    case "remove":
      removeTrack(track.id);
      break;
    default:
      break;
  }
}

function populateInstrumentPicker() {
  ui.instrumentSelect.innerHTML = Object.keys(INSTRUMENTS)
    .map((instrumentName) => `<option value="${instrumentName}">${instrumentName}</option>`)
    .join("");
}

function tryLoadBootProject() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return false;
  }

  try {
    const parsed = JSON.parse(raw);
    loadProject(parsed, "saved local project");
    return true;
  } catch (error) {
    setStatus(`Saved project was invalid, starter arrangement restored. (${error.message})`, true);
    return false;
  }
}

function initializeIntegrationPanels() {
  loadIntegrationSettings();
  updateMidiCaptureButton();

  if (ui.yamahaApplyTempoBtn) {
    ui.yamahaApplyTempoBtn.disabled = true;
  }

  renderYamahaPayload({});

  if (!supportsWebMidi()) {
    if (ui.midiConnectBtn) {
      ui.midiConnectBtn.disabled = true;
    }
    if (ui.midiInputSelect) {
      ui.midiInputSelect.disabled = true;
    }
    setMidiFeedback("Web MIDI is not supported in this browser. Use Chromium/Edge on desktop.", true);
    return;
  }

  if (!window.isSecureContext) {
    setMidiFeedback("Web MIDI requires HTTPS or localhost.", true);
    return;
  }

  setMidiFeedback("MIDI ready. Click Connect MIDI to authorize your keyboard.");
}

function bindEvents() {
  ui.playBtn.addEventListener("click", () => {
    togglePlayback();
  });

  ui.stopBtn.addEventListener("click", () => {
    stopPlayback();
  });

  if (ui.metronomeBtn) {
    ui.metronomeBtn.addEventListener("click", () => {
      toggleMetronome();
    });
  }

  if (ui.countInBtn) {
    ui.countInBtn.addEventListener("click", () => {
      toggleCountIn();
    });
  }

  ui.tempoSlider.addEventListener("input", () => {
    applyTransportControls();
  });

  ui.swingSlider.addEventListener("input", () => {
    applyTransportControls();
  });

  ui.loopSelect.addEventListener("change", () => {
    pushHistorySilent();
    const nextValue = Number(ui.loopSelect.value);
    loopSteps = [16, 32, 48, 64].includes(nextValue) ? nextValue : 32;
    currentStep %= loopSteps;
    renderPianoRoll();
    setPlayhead(currentStep);
    setStatus(`Loop length set to ${loopSteps} steps.`);
  });

  ui.addTrackBtn.addEventListener("click", () => {
    addTrack();
  });

  ui.clearTrackBtn.addEventListener("click", () => {
    clearSelectedTrackPattern();
  });

  ui.trackNameInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      addTrack();
    }
  });

  ui.trackList.addEventListener("click", onTrackListClick);

  if (ui.composerFullscreenBtn) {
    ui.composerFullscreenBtn.addEventListener("click", () => {
      toggleComposerModal();
    });
  }

  if (ui.composerCloseBtn) {
    ui.composerCloseBtn.addEventListener("click", () => {
      setComposerModalState(false);
    });
  }

  if (ui.composerModalBackdrop) {
    ui.composerModalBackdrop.addEventListener("click", () => {
      setComposerModalState(false);
    });
  }

  if (ui.chordCompactToggleBtn) {
    ui.chordCompactToggleBtn.addEventListener("click", () => {
      toggleChordPaletteCompact();
    });
  }

  if (ui.phraseLibrary) {
    ui.phraseLibrary.addEventListener("dragstart", onPhraseLibraryDragStart);
    ui.phraseLibrary.addEventListener("click", onPhraseLibraryClick);
  }

  ui.pianoRoll.addEventListener("pointerdown", onPianoRollPointerDown);
  ui.pianoRoll.addEventListener("pointerover", onPianoRollPointerOver);
  ui.pianoRoll.addEventListener("dragover", onPianoRollDragOver);
  ui.pianoRoll.addEventListener("drop", onPianoRollDrop);
  window.addEventListener("pointerup", onPianoRollPointerUp);

  if (ui.rollScroll) {
    ui.rollScroll.addEventListener("wheel", onRollScrollWheel, { passive: false });
  }

  ui.mixerStrips.addEventListener("input", onMixerInput);
  ui.mixerStrips.addEventListener("change", onMixerChange);
  ui.mixerStrips.addEventListener("click", onMixerClick);

  if (ui.undoBtn) {
    ui.undoBtn.addEventListener("click", undoHistory);
  }

  if (ui.redoBtn) {
    ui.redoBtn.addEventListener("click", redoHistory);
  }

  if (ui.applyTemplateBtn) {
    ui.applyTemplateBtn.addEventListener("click", () => {
      applyTemplateFromSelect();
    });
  }

  if (ui.duplicatePhraseBtn) {
    ui.duplicatePhraseBtn.addEventListener("click", () => {
      duplicatePhraseAcrossLoop();
    });
  }

  if (ui.humanizeBtn) {
    ui.humanizeBtn.addEventListener("click", () => {
      humanizeSelectedTrack();
    });
  }

  if (ui.nudgeLeftBtn) {
    ui.nudgeLeftBtn.addEventListener("click", () => {
      nudgeSelectedTrack("left");
    });
  }

  if (ui.nudgeRightBtn) {
    ui.nudgeRightBtn.addEventListener("click", () => {
      nudgeSelectedTrack("right");
    });
  }

  ui.saveLocalBtn.addEventListener("click", saveLocalProject);
  ui.loadLocalBtn.addEventListener("click", loadLocalProject);
  ui.exportBtn.addEventListener("click", exportProject);

  if (ui.exportSheetsBtn) {
    ui.exportSheetsBtn.addEventListener("click", exportSheetMusicPerTrack);
  }

  ui.importBtn.addEventListener("click", () => {
    ui.importFileInput.click();
  });

  ui.importFileInput.addEventListener("change", async () => {
    const file = ui.importFileInput.files ? ui.importFileInput.files[0] : null;
    await importProjectFile(file);
    ui.importFileInput.value = "";
  });

  if (ui.midiConnectBtn) {
    ui.midiConnectBtn.addEventListener("click", () => {
      toggleMidiConnection();
    });
  }

  if (ui.midiInputSelect) {
    ui.midiInputSelect.addEventListener("change", () => {
      const inputId = ui.midiInputSelect.value;
      if (inputId) {
        attachMidiInputById(inputId);
        saveIntegrationSettings();
      }
    });
  }

  if (ui.midiChannelSelect) {
    ui.midiChannelSelect.addEventListener("change", () => {
      midiState.channel = ui.midiChannelSelect.value;
      saveIntegrationSettings();
      setMidiFeedback(`MIDI channel filter set to ${midiState.channel === "all" ? "all channels" : `channel ${midiState.channel}`}.`);
    });
  }

  if (ui.midiCaptureBtn) {
    ui.midiCaptureBtn.addEventListener("click", () => {
      midiState.captureToGrid = !midiState.captureToGrid;
      updateMidiCaptureButton();
      saveIntegrationSettings();
      setMidiFeedback(
        midiState.captureToGrid
          ? "Capture to grid is ON. Incoming MIDI notes are written to the active track."
          : "Capture to grid is OFF. MIDI only auditions notes."
      );
    });
  }

  if (ui.yamahaAnalyzeBtn) {
    ui.yamahaAnalyzeBtn.addEventListener("click", () => {
      analyzeWithYamaha();
    });
  }

  if (ui.yamahaApplyTempoBtn) {
    ui.yamahaApplyTempoBtn.addEventListener("click", () => {
      applyYamahaTempoFromLastAnalysis();
    });
  }

  if (ui.yamahaUrlInput) {
    ui.yamahaUrlInput.addEventListener("change", () => {
      saveIntegrationSettings();
    });
  }

  if (ui.yamahaAuthHeaderInput) {
    ui.yamahaAuthHeaderInput.addEventListener("change", () => {
      saveIntegrationSettings();
    });
  }

  document.addEventListener("keydown", (event) => {
    const key = event.key.toLowerCase();

    if (event.key === "Escape" && composerModalOpen) {
      event.preventDefault();
      setComposerModalState(false);
      return;
    }

    if (event.code === "Space" && !isTextInputElement(event.target)) {
      event.preventDefault();
      togglePlayback();
      return;
    }

    if (!isTextInputElement(event.target) && key === "m") {
      event.preventDefault();
      toggleMetronome();
      return;
    }

    if (!isTextInputElement(event.target) && key === "c") {
      event.preventDefault();
      toggleCountIn();
      return;
    }

    const isSave = (event.ctrlKey || event.metaKey) && key === "s";
    if (isSave) {
      event.preventDefault();
      saveLocalProject();
      return;
    }

    const isUndo = (event.ctrlKey || event.metaKey) && !event.shiftKey && key === "z";
    if (isUndo) {
      event.preventDefault();
      undoHistory();
      return;
    }

    const isRedo =
      ((event.ctrlKey || event.metaKey) && event.shiftKey && key === "z") ||
      (event.ctrlKey && !event.shiftKey && key === "y");
    if (isRedo) {
      event.preventDefault();
      redoHistory();
    }
  });
}

function initialize() {
  initializeIntegrationPanels();
  populateInstrumentPicker();
  bindEvents();
  setChordPaletteCompactState(false);
  applyTransportControls();
  updateTransportOptionButtons();
  updateHistoryButtons();

  if (!tryLoadBootProject()) {
    createStarterProject();
    renderAll();
    setPlayhead(0);
    setStatus("Starter orchestral sketch loaded. Edit notes and press Play.");
  }

  scheduleSequencer();
  updateMasterMeter();
  refreshPlayButton();
}

// Wire up the HTML modal close/open buttons
const _modalOpenBtn = document.getElementById("openFullscreen");
const _modalCloseBtn = document.getElementById("closeFullscreen");
if (_modalOpenBtn) {
  _modalOpenBtn.addEventListener("click", () => setComposerModalState(true));
}
if (_modalCloseBtn) {
  _modalCloseBtn.addEventListener("click", () => setComposerModalState(false));
}

initialize();
