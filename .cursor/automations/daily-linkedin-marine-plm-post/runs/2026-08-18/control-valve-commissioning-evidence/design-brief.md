# Design Brief — control-valve-commissioning-evidence

Reference style: commissioning-gates solution overview

## Header

- Title: CONTROL VALVE COMMISSIONING EVIDENCE GATES
- Subtitle: HART loop proof and fail-safe evidence before handover
- Badge: SOLUTION OVERVIEW

## Seven commissioning gates

1. **Loop Baseline**
   - P&ID tag matched to installed serial
   - Positioner firmware recorded
   - Accountable engineer assigned
2. **HART Configuration**
   - Parameter set uploaded and diffed
   - Approved template referenced
   - Deviations dispositioned
3. **Calibration Proof**
   - Certificates traced to tag
   - Bench and in-situ results linked
   - Validity dates verified
4. **Stroke Verification**
   - Open/close times witnessed
   - Travel and linearity inside spec
   - Response to setpoint steps logged
5. **Fail-Safe Trip**
   - Air-loss action demonstrated
   - Signal-loss action demonstrated
   - Measured travel and time recorded
6. **System Integration**
   - Automation demand tested end-to-end
   - Alarm and interlock behavior checked
   - Control-loop response measured
7. **Handover Evidence**
   - Loop records linked to baseline
   - Exceptions with owner and status
   - Class-witness sign-off completed
- Amber HOLD / ESCALATE path applies to missing certificates or failed fail-safe trips.

## Center engineering scene

- Premium photorealistic engine-room valve commissioning scene: commissioning engineer with HART communicator at a globe control valve with smart positioner, air lines and gauges visible, piping and cable trays, newbuild vessel context.
- Cyan HUD nodes: P&ID TAG → HART POSITIONER → LOOP CHECK → EVIDENCE STORE.
- Solid cyan TRACED LINK terminates at the visible positioner.
- Amber dashed HOLD / ESCALATE path appears only for missing certificate or failed fail-safe evidence.

## Left sidebar — LOOP EVENTS

- Tag-to-serial verified
- HART set uploaded
- Certificates traced
- Stroke times witnessed
- Fail-safe proven

## Right sidebar — PYTHON AUTOMATION

- Diff HART parameters
- Flag missing certificates
- Compare stroke limits
- Build loop evidence index
- Raise handover holds

## Bottom value pillars

1. **Traceable Loops** — Every tag, serial, and parameter set is bound to one reviewable record.
2. **Proven Fail-Safe** — Air-loss and signal-loss actions are measured on the installed valve, not assumed.
3. **Evidence-Based Handover** — Certificates, logs, and exceptions are linked to the class-witness baseline.

## Footer takeaway

AI can draft the completeness checks. Engineers still own acceptance limits, fail-safe proof, and the handover decision.

## Visual system

- Canvas: 1080 × 1350 px, portrait.
- Palette: deep navy `#081423`, panel navy `#10243A`, cyan `#18D7F3`, white `#F4F8FC`, muted blue-gray `#A8B8C8`, amber `#F7A928`.
- Typography: Inter with bold condensed hierarchy, all visible text in English.
- Composition: 12-column grid, high-density executive engineering layout, balanced whitespace, thin technical borders, subtle glow, no childish icons or generic poster treatment.
