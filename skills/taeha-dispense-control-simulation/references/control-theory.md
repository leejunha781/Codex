# Control theory for material dispensing

## Multi-rate hierarchy

Treat the plant as several loops with different bandwidths.

- Motor current or torque loop: fastest; usually implemented by the drive or FOC firmware.
- Motor speed loop: slower than current; rejects mechanical load changes.
- Position or delivered-angle loop: optional for syringe, screw, piston, or servo mechanisms.
- Pressure/flow loop: limited by fluid compliance, sensor bandwidth, valve and transport delay.
- Shot mass loop: often batch/run-to-run because stable weight is available only after cutoff and settling.
- Supervisory ratio/safety loop: validates A/B, limits, profiles, confidence, state, and faults.

Do not tune these loops at the same bandwidth. Keep inner loops materially faster than outer loops and verify actual separation.

## Feed-forward

Create a profile-specific inverse map:

[
u_{ff}=f^{-1}(m^*, T, P, material, lot, nozzle, mixer, pump, gear, actuator)
]

The map may output speed, angle/steps, valve timing, precharge, cutoff, suck-back, and settle time. Use tables, piecewise interpolation, or a fitted physical/empirical model. Preserve units, validity ranges, data provenance, residuals, and uncertainty.

Feed-forward is the primary action for a short shot. Feedback corrects model and disturbance error; it should not be expected to create the entire trajectory after a delayed sensor finally responds.

## Cascaded PI/PID

Use PI rather than PID unless derivative action produces a measured benefit. A production implementation requires:

- output and rate saturation;
- anti-windup by conditional integration or back calculation;
- derivative-on-measurement and filtering when D is used;
- setpoint weighting or two-degree-of-freedom structure;
- bumpless manual/automatic and profile transfer;
- timestamp/freshness checks;
- fixed-point range and sample-time proof;
- safe behavior on sensor, drive, or communication faults.

Identify a first-order-plus-dead-time model

[
G(s)=\frac{K e^{-Ls}}{\tau s+1}
]

and use a robust model-based method such as IMC/SIMC as a starting point. Validate gain and phase margins and nonlinear saturation by simulation and experiment. Ziegler–Nichols may be useful for teaching, but aggressive closed-loop cycling is a poor default on a hazardous dispenser.

## Delay compensation

Use measured total delay:

[
L=L_{transport}+L_{mechanical}+L_{ADC/filter}+L_{task}+L_{communication}
]

If delay is large relative to the dominant time constant, lower feedback bandwidth or use prediction. A Smith predictor can remove nominal delay from the internal feedback model, but delay/gain mismatch can remove the expected benefit. Bound its validity, monitor prediction residual, and fall back to a conservative controller.

For a load-cell Σ-Δ ADC, output data rate is not equal to zero-latency response. Include digital-filter group delay and full settling after steps, channel changes, gain changes, standby, or filter resets.

## Shot-to-shot or run-to-run control

For stable post-shot mass (y_k), target (r_k), and recipe multiplier (g_k), use a bounded update such as

[
e_k=r_k-y_k
]
[
g_{k+1}=clip(rateLimit(g_k(1+\alpha e_k/r_k)),g_{min},g_{max})
]

or update the model intercept with EWMA. Tune (alpha) against noise, drift, profile frequency, and model uncertainty. A higher gain converges faster but amplifies noise and alternating errors.

Update only when all gates pass:

- correct material/component/lot/nozzle/mixer/pump/gear/temperature band;
- stable, fresh, calibrated, nonsaturated measurement;
- normal production shot, not prime/purge/manual/rework;
- no timeout, motor slip, valve fault, air/cavitation, blockage, interlock, or ratio fault;
- target above validated minimum;
- sufficient confidence and sample history.

Use separate control threads/profiles for high-mix operation. Never share a learned state merely because products look similar.

## Iterative learning control

ILC learns a time-indexed command waveform for a repeated finite trajectory. It can improve pressure/flow tracking when each shot follows the same time base and a meaningful within-shot error waveform exists. It is not automatically equivalent to scalar shot-to-shot gain correction.

Before ILC, prove:

- repeatable initial state or explicit reset;
- consistent horizon and sampling;
- stable learning operator and robustness to nonrepeating disturbance;
- bounded command and learning filters;
- profile isolation and rollback.

## Ratio control

For desired mass ratio (R=m_A/m_B), compute independent component targets:

[
m_A^*=m^*\frac{R}{R+1},\quad m_B^*=m^*\frac{1}{R+1}
]

Supervise both component measurements. Useful architectures include independent feed-forward plus trim, master–follower ratio control, or constrained multivariable control. Account for density versus temperature, component-specific delay, valve minimum dose, startup lag, purge, pot life, and mixer pressure.

Total mixed mass provides one equation for two unknown component masses and cannot identify ratio without additional measurement or a separately validated model.

## Estimation and advanced control

Use observers only for unmeasured states that affect a requirement.

- Kalman filter: linear stochastic state estimation with justified noise models.
- Extended/unscented Kalman filter: nonlinear estimation; validate consistency and divergence handling.
- RLS/EWMA: slow gain/intercept/drift estimation; apply forgetting and parameter bounds.
- MHE: constrained estimation when compute budget and observability justify it.
- MPC: use for coupled A/B, pressure, actuator and constraint management when a validated predictive model and solver deadline are available.

Advanced control does not remove the need for independent hard limits. On STM32, prove worst-case execution, memory, numerical conditioning, warm-start/failure behavior, fallback controller, and code/model equivalence.

## Recommended progression

1. Safety state machine and hard limits.
2. Profile feed-forward.
3. Post-shot stability detector and bounded EWMA/integral correction.
4. Separate A/B correction and ratio supervision.
5. Inner motor PI/FOC tuning using exact drive/motor evidence.
6. Optional pressure/flow PI if sensor delay permits.
7. Smith predictor, ILC, observer, or MPC only after a measured gap remains.
