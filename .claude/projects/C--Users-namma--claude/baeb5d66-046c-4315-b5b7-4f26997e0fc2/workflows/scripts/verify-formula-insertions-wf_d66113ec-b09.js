export const meta = {
  name: 'verify-formula-insertions',
  description: 'Adversarially verify RF/antenna/ADC formulas and completeness before inserting into interview-prep docx',
  phases: [
    { title: 'Verify', detail: '3 lenses: formula correctness, numeric examples, completeness vs source doc' },
  ],
}

const SPEC = 'C:\\Users\\namma\\.claude\\moacom_work\\formula_spec.json'
const DUMP = 'C:\\Users\\namma\\.claude\\moacom_work\\dump_interview.txt'

const FINDINGS = {
  type: 'object',
  properties: {
    findings: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          blockId: { type: 'string', description: 'id of the insert block, or "GLOBAL"/"MISSING" for doc-wide items' },
          severity: { type: 'string', enum: ['error', 'warning', 'ok-note'] },
          issue: { type: 'string', description: 'what is wrong or missing, one or two sentences' },
          correction: { type: 'string', description: 'exact corrected formula/number/Korean wording to use, empty if none' }
        },
        required: ['blockId', 'severity', 'issue', 'correction']
      }
    },
    overallVerdict: { type: 'string', description: 'one-line summary' }
  },
  required: ['findings', 'overallVerdict']
}

phase('Verify')
const results = await parallel([
  () => agent(
    `You are an RF/microwave/antenna/mixed-signal engineering reviewer. Read the JSON file at ${SPEC}. It contains Korean text blocks with embedded engineering formulas that will be inserted into a technical-interview prep document for a defense/satcom HW design role (antenna control & signal processing modules, FPGA/CPU/ADC-DAC, link budget, ESA/phased array, repeaters, 5G RU).

Your ONLY job: try to REFUTE each formula on physics/convention grounds. Check every equation for: correct form per standard textbooks (Pozar, Friis, 3GPP), sign conventions, dB vs linear mixing errors, unit consistency (dBm/dBW/dBi/dB-Hz/dB/K), variable definitions matching the surrounding Korean prose, and misleading approximations (state validity limits if a rule-of-thumb is presented as exact). Specifically scrutinize: quantization SNR 6.02N+1.76; jitter SNR -20log10(2π·fin·tj); Friis cascade; sensitivity -174+10logB+NF+SNRreq; FSPL 32.44+20logf(MHz)+20logd(km); EIRP; G/T = Gr - 10log10(Tsys); T=(F-1)·290K; C/N0 = EIRP + G/T - L + 228.6; Eb/N0 = C/N0 - 10log10(Rb); repeater stability G ≤ I - 15 dB; EVM(%)=100/√SNRlin; Δφ=(2πd/λ)sinθ; grating-lobe d ≤ λ/(1+|sinθmax|); scan loss 10log10(cosθ); pointing loss 12(θerr/θ3dB)²; AR=20log10(Emaj/Emin); PLF=|ê_rx·ê_tx|²; IM3=3Pin-2·IIP3; OIP3=IIP3+Gain; OP1dB≈OIP3-9.6dB; SFDR; mismatch loss -10log10(1-|Γ|²); Z0=√(L/C); BW≈0.35/tr; oversampling gain 10log10(Fs/2B); dBW=dBm-30. Report every genuine defect as severity=error with the exact corrected wording (in Korean matching the block style) in 'correction'. Use 'warning' for approximations needing a caveat. If a formula is correct, do not report it (or use ok-note only when a subtle point deserves a note). Do NOT invent problems.`,
    { label: 'verify:physics', phase: 'Verify', schema: FINDINGS }
  ),
  () => agent(
    `You are a numerical checker. Read the JSON file at ${SPEC} (Korean engineering text blocks with formulas AND worked numeric examples). Recompute EVERY numeric claim from scratch and flag any that are wrong or imprecise (off by more than rounding). Claims to recheck include: 12-bit ideal SNR ≈ 74 dB; jitter SNR for fin=1 GHz, tj=1 ps ≈ 44 dB; gain chain +20-3-7+30=+40 dB; rise time 1 ns → BW ≈ 350 MHz; FSPL at 12 GHz / 1200 km ≈ 175.6 dB; EIRP 10 dBW + 33 dBi - 2 dB = 41 dBW; NF 1 dB → ≈75 K; C/N0 -3 dB → halve Rb for same Eb/N0; VSWR 2.0 → |Γ|=0.33, ~11% reflected, mismatch loss ≈ 0.5 dB; ±60° steering → d ≤ 0.54λ; pointing error = θ3dB/4 → 0.75 dB and θ3dB/2 → 3 dB; SNR 30 dB → EVM ≈ 3.2%; 3GPP 256QAM EVM limit 3.5%; Boltzmann constant 1.38×10^-23 J/K ↔ -228.6 dBW/K/Hz. Show your arithmetic mentally, and report ONLY discrepancies as severity=error (with corrected number in 'correction', Korean wording). If all numbers check out, return findings=[] and say so in overallVerdict.`,
    { label: 'verify:numbers', phase: 'Verify', schema: FINDINGS }
  ),
  () => agent(
    `You are a completeness critic for a Korean technical-interview prep document. Read BOTH files: (1) ${DUMP} — full paragraph dump of the existing document (main body sections 1-10 = theory & interview strategy; appendices A-J = formula reference), and (2) ${SPEC} — planned formula insertions. The user's requirement: "꼭 해당 이론들에 수식 계산식이 있으면 누락하지 말고 넣어주세요" — every theory/concept DISCUSSED in the document that has a standard textbook formula must end up with that formula present somewhere (existing appendix OR planned insert). Sweep the dump for every technical concept mentioned (e.g. Flash/boot, CPU/FPGA, SPI/DDR/PCIe/Ethernet, gain plan, NF, sensitivity, linearity, phase noise, beam steering, beamforming, calibration, scan loss, pointing error, RHCP/axial ratio, ESA tracking, Link Budget, EIRP, G/T, C/N0, FSPL, repeater, 5G RU, EVM, ACLR, power integrity, signal integrity, VSWR, EMI/EMC, thermal, FMEA/RPN, verification coverage, defect closure...). For each concept that (a) has a well-known standard formula AND (b) that formula is in NEITHER the existing appendix NOR the planned inserts, report severity=warning with blockId="MISSING", the concept in 'issue', and in 'correction' give the formula plus ONE natural Korean sentence in the document's style (polite -입니다 register) ready to insert. Only propose formulas that are genuinely standard and interview-relevant for a HW design 부장급 role; do not pad. Also flag (severity=error) if any planned insert ANCHOR string does not appear in the dump or appears in more than one paragraph (ambiguous anchor) — check each of the 26 anchors literally against the dump text.`,
    { label: 'verify:completeness', phase: 'Verify', schema: FINDINGS }
  ),
])

const all = results.filter(Boolean)
return {
  physics: all[0] ?? null,
  numbers: all[1] ?? null,
  completeness: all[2] ?? null,
}