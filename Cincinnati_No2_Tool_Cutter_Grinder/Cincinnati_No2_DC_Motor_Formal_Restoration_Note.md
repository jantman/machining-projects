# Restoration Note  
## Cincinnati No. 2 Tool & Cutter Grinder  
### DC Power Feed Motor — Diagnosis, Rewiring, and Final Configuration

**Machine:** Cincinnati No. 2 Tool & Cutter Grinder  
**Approx. Year:** 1952  
**Subsystem:** Table Power Feed  
**Motor:** General Electric DC Motor, Type 5BC44AB2139  
**Prepared by:** Restoration notes compiled during troubleshooting and repair  
**Purpose:** Formal restoration record and technical reference

---

## 1. Scope and Intent

This document serves as a **formal restoration note** for the DC power feed system of a Cincinnati No. 2 Tool & Cutter Grinder.  
It records the *actual as‑found condition*, the investigative process, electrical measurements, internal motor topology, corrective actions taken, and the final verified wiring configuration.

The intent is to:
- Preserve technical knowledge that is not present in factory documentation
- Prevent re‑introduction of known faults
- Aid future owners, restorers, and technicians
- Provide a factual record of deviations between schematic expectation and physical reality

---

## 2. Factory Design Intent (Summary)

The Cincinnati factory wiring diagram indicates the machine is designed to drive a:

**Separately excited shunt DC motor**

Key characteristics:
- Armature (A1/A2): variable, reversible DC
- Field (F1/F2): fixed DC, constant polarity
- Speed control via Powerstat varying armature voltage
- Direction reversal accomplished by reversing armature polarity only

This architecture is electrically sound **only** for a true shunt motor.

---

## 3. As‑Found Symptoms

Prior to intervention, the following symptoms were observed:

- Feed motor operated in only one direction
- Severe speed difference forward vs reverse
- Speed control behaved non‑linearly and inconsistently
- Selenium rectifiers ran excessively hot
- Transformer input fuses (2 A time‑delay) blew under some conditions
- Motor wiring did not correlate cleanly with schematic expectations
- Motor presented five external leads, not four

---

## 4. Motor Identification

**Nameplate Data:**
- Manufacturer: General Electric
- Type: 5BC44AB2139
- Rating: 1/6 HP
- Voltage: 115 V DC
- Current: 1.8 A
- Speed: 1140 RPM
- Duty: Continuous
- Winding: **Compound**

The compound winding classification proved critical.

---

## 5. Motor Internal Construction (As Discovered)

Upon careful disassembly:

### 5.1 Mechanical Layout
- Two carbon brushes, 180° apart
- Three field coils mounted to the housing:
  - Two large coils at 0° and 180°
  - One smaller coil at 90°

### 5.2 Electrical Topology
- The 90° coil is a **series field**
- The 0° and 180° coils form the **shunt field**
- One brush was originally routed through the series field
- Extensive cloth insulation concealed internal connections

---

## 6. Shunt Field Topology (Critical Discovery)

Resistance mapping revealed the shunt field is wired as follows:

```
              Shunt Coil (~485 Ω)
Shunt A o─────/////─────.
                                                     +──── Shunt CT
                          /
Shunt B o─────/////─────'
```

- Shunt A and Shunt B are tied together with a very low resistance strap (~0.78 Ω)
- The true shunt field resistance (~485 Ω) exists between the A/B node and the center tap
- This configuration is typical of GE compound motors of this era but is **not obvious externally**

---

## 7. Resistance Mapping Results

| Measurement | Typical Value |
|------------|---------------|
| Brush ↔ Brush | ~14.4 Ω |
| Series A ↔ Series B | ~1.86 Ω |
| Shunt A ↔ Shunt B | ~0.78 Ω |
| Shunt A ↔ Shunt CT | ~485–495 Ω |
| Shunt B ↔ Shunt CT | ~484–488 Ω |
| Any lead ↔ frame | OL |

These measurements conclusively identified each winding.

---

## 8. Root Cause of Failure

The machine reverses **only the armature**, while the motor was compound wound with the series field in series with the armature.

This resulted in:
- Additive field in one direction
- Subtractive field in the opposite direction
- Torque and speed asymmetry
- Excess current draw
- Rectifier overheating
- Fuse failures (made obvious after selenium rectifier replacement)

---

## 9. Corrective Strategy

Several solutions were evaluated:
- Modify machine controls
- Replace motor
- Install modern DC drive
- **Internally convert motor to pure shunt operation**

The chosen solution preserved:
- Mechanical fit
- Original motor iron
- Machine control philosophy

---

## 10. Internal Motor Rewiring

Actions taken:
- Series field disconnected from armature circuit
- Both brushes brought out as independent armature leads
- Series field leads insulated and parked
- Shunt field left electrically intact
- Shunt center tap preserved

No windings were removed or altered.

---

## 11. Rectifier Modernization

Original selenium rectifiers were replaced with:
- **KBPC5010 silicon bridge rectifiers**

Result:
- Lower voltage drop
- Improved reliability
- Latent wiring errors became immediately detectable

---

## 12. Final Verified Wiring (Authoritative)

### Armature
- A1 → Brush A
- A2 → Brush B

### Shunt Field (IMPORTANT)
- Tie Shunt A and Shunt B together
- F1 → Shunt A + B node
- F2 → Shunt Center Tap

### Not Connected
- Series field leads
- Any unused taps

---

## 13. Final Results

- Smooth speed control via Powerstat
- Proper reversing
- Equal performance forward and reverse
- No rectifier overheating
- No fuse failures
- System behaves exactly as Cincinnati intended

---

## 14. Lessons Learned

1. Do not assume motor topology from external leads
2. Resistance mapping is definitive
3. Compound motors hide non‑obvious connections
4. Silicon rectifiers reveal faults masked by selenium
5. GE shunt fields often use outer‑end strapping

---

## Appendix A — Resistance Mapping Worksheet (Reference)

| From \ To | Brush A | Brush B | Series A | Series B | Shunt A | Shunt B | Shunt CT |
|-----------|---------|---------|----------|----------|---------|---------|----------|
| Brush A | — | | | | | | |
| Brush B | | — | | | | | |
| Series A | | | — | | | | |
| Series B | | | | — | | | |
| Shunt A | | | | | — | | |
| Shunt B | | | | | | — | |
| Shunt CT | | | | | | | — |

---

## Appendix B — Safety and Verification Checklist

- All unused leads individually insulated
- No continuity from any lead to frame
- ~485 Ω measured across F1–F2
- ~14 Ω measured across A1–A2
- Field energized before armature
- Initial test at minimum Powerstat setting

---

## Closing Note

This restoration resolved a subtle but fundamental mismatch between a compound DC motor and a shunt‑motor control system.  
The final configuration is electrically correct, mechanically original, and now fully documented.

**Future guidance:**  
If symptoms reappear, repeat resistance mapping before replacing parts.

---

**End of Restoration Note**
