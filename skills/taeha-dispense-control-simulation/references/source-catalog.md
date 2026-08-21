# Source catalog

Retrieved 2026-08-18. Recheck current versions, licenses, and product dependencies before adoption.

## Open simulation and control

- Python Control Systems Library, optimization-based control and MPC: https://python-control.readthedocs.io/en/stable/optimal.html
- SciPy state-space and optimization documentation: https://docs.scipy.org/doc/scipy/reference/generated/scipy.signal.StateSpace.html and https://docs.scipy.org/doc/scipy/tutorial/optimize.html
- CasADi documentation, optimal control and C code generation: https://web.casadi.org/docs/
- OpenModelica User Guide: https://openmodelica.org/doc/OpenModelicaUsersGuide/latest/
- OpenModelica FMI support: https://openmodelica.org/doc/OpenModelicaUsersGuide/latest/fmitlm.html
- Modelica language and standard libraries: https://modelica.org/language/ and https://modelica.org/libraries/
- Scilab/Xcos dynamic-system simulation: https://www.scilab.org/software/xcos
- GNU Octave control package: https://gnu-octave.github.io/packages/control/
- Renode documentation and peripheral modeling: https://renode.readthedocs.io/ and https://renode.readthedocs.io/en/latest/advanced/writing-peripherals.html

## MATLAB/Simulink

- Simscape Fluids: https://www.mathworks.com/help/hydro/index.html
- Simscape isothermal/thermal liquid modeling: https://www.mathworks.com/help/simscape/ug/modeling-isothermal-liquid-systems.html
- Simulink Control Design: https://www.mathworks.com/help/simulink/control-design.html
- PID Tuner: https://www.mathworks.com/help/control/ref/pidtuner-app.html
- Smith predictor example: https://www.mathworks.com/help/control/ug/control-of-processes-with-long-dead-time-the-smith-predictor.html
- System Identification Toolbox requirements: https://www.mathworks.com/support/requirements/system-identification-toolbox.html
- Simscape real-time HIL: https://www.mathworks.com/help/simscape/hardware-in-the-loop-simulation-with-simulink-real-time.html
- Verification and validation product workflow: https://www.mathworks.com/help/slcoverage/verification-and-validation.html
- Simulink Design Verifier: https://www.mathworks.com/help/sldv/getting-started-with-simulink-design-verifier.html

## STM32 and motor control

- ST X-CUBE-MCSDK: https://www.st.com/en/embedded-software/x-cube-mcsdk.html
- STM32 motor control ecosystem: https://www.st.com/content/st_com/en/ecosystems/stm32-motor-control-ecosystem.html
- STM32 MC Workbench manual UM3027: https://www.st.com/resource/en/user_manual/um3027-how-to-use-stm32-motor-control-sdk-v60-workbench-stmicroelectronics.pdf
- STM32 AC induction motor MCSDK guidance: https://wiki.st.com/stm32mcu/wiki/STM32MotorControl:SDK_AC_induction_motor

## Control theory

- Bristow, Tharayil, and Alleyne, Survey of Iterative Learning Control, DOI 10.1109/MCS.2006.1636313: https://experts.illinois.edu/en/publications/survey-of-iterative-learning-control-a-learning-based-method-for-/
- Skogestad, Simple analytic rules for model reduction and PID controller tuning: https://web01.usn.no/~davidr/iia1117/control/theory/pensum/tuningshort.pdf
- Kothare et al., unified anti-windup framework, DOI 10.1016/0005-1098(94)90048-5: https://www.sciencedirect.com/science/article/pii/0005109894900485
- Liu et al., survey of run-to-run control for batch processes, DOI 10.1016/j.isatra.2018.09.005: https://www.sciencedirect.com/science/article/pii/S0019057818303355
- Wang et al., EWMA gain versus intercept adaptation, DOI 10.1016/j.jprocont.2009.06.002: https://www.sciencedirect.com/science/article/pii/S0959152409001164

## Dispensing and measurement evidence

- Graco ProMix 2KE operation manual with ratio tolerance, pot-life and calibration concepts: https://www.graco.com/content/dam/graco/tech_documents/manuals/3A0/3A0869/3A0869EN-P.pdf
- Graco ProMix PD positive-displacement proportioner: https://www.graco.com/us/en/in-plant-manufacturing/products/liquid-coating/meter-mix/plural-component-mixing-equipment/promix-pd.html
- Nordson EFD 797PCP-2K: https://www.nordson.com/en/products/efd-products/797pcp-2k-progressive-cavity-pumps
- ViscoTec progressive-cavity technology: https://www.viscotec.de/en/technology/
- Analog Devices AN-0979 digital-filter settling: https://www.analog.com/en/resources/app-notes/an-0979.html
- Analog Devices CN0600 load-cell signal chain: https://www.analog.com/en/resources/reference-designs/circuits-from-the-lab/cn0600.html
- TI ADS1261 design calculator and CRC/filter timing: https://www.ti.com/tool/ADS1261-EXCEL-CALC-TOOL

## Process validation context

For regulated materials or processes, treat controller changes as process changes requiring the applicable quality system. Useful general references include FDA Process Validation and PAT guidance:

- https://www.fda.gov/media/71021/download
- https://www.fda.gov/regulatory-information/search-fda-guidance-documents/pat-framework-innovative-pharmaceutical-development-manufacturing-and-quality-assurance

These do not make a dispenser compliant by themselves. Apply the actual product, jurisdiction, safety, machinery, electrical, chemical, and quality requirements.
