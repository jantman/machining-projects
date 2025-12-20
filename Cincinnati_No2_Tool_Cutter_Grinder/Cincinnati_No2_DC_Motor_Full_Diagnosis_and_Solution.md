# Cincinnati No. 2 Tool & Cutter Grinder  
## DC Power Feed Motor — Diagnosis, Rewiring, and Resolution

**Machine:** Cincinnati No. 2 Tool & Cutter Grinder (circa 1952)  
**Subsystem:** Table power feed DC motor and controls  
**Motor:** General Electric 5BC44AB2139  
**Purpose:** Permanent technical record of diagnosis and solution

---

## 1. Original Symptoms

- Power feed motor ran only in one direction or inconsistently
- Large speed difference forward vs reverse
- Powerstat speed control behaved erratically
- Selenium rectifiers overheated
- Transformer input fuses blew under load
- Motor lead count and behavior did not match schematic

---

## 2. What the Machine Expects

The Cincinnati schematic expects a **separately excited shunt DC motor**:

- Armature: A1 / A2 (reversible DC)
- Field: F1 / F2 (fixed DC)
- Speed via armature voltage
- Reversal via armature polarity only

---

## 3. Motor Nameplate

- GE DC Motor 5BC44AB2139
- 1/6 HP, 115 V DC, 1.8 A
- 1140 RPM
- **Compound wound**
- Continuous duty

The compound nature is the root cause of incompatibility.

---

## 4. Internal Motor Construction

- Two brushes (180° apart)
- Three field coils:
  - Two large coils at 0° and 180° (shunt field)
  - One smaller coil at 90° (series field)
- Extensive cloth insulation and varnish

### Internal wiring discovered

- One brush directly to external lead
- Other brush routed through series field
- Shunt coils:
  - Inner ends tied together (center tap)
  - Outer ends tied together with heavy strap

---

## 5. Resistance Mapping (Key Data)

| Measurement | Value |
|-----------|------|
| Brush–Brush | ~14.4 Ω |
| Series–Series | ~1.86 Ω |
| Shunt A–B | ~0.78 Ω |
| Shunt A–CT | ~485–495 Ω |
| Shunt B–CT | ~484–488 Ω |

This proves Shunt A and B are the same electrical node.

---

## 6. Why It Could Not Work Originally

- Series field polarity reversed with armature
- Field strength added in one direction, subtracted in the other
- Resulted in asymmetric torque, speed, and rectifier stress

---

## 7. Chosen Solution

Convert the motor internally to **pure shunt operation**:

- Remove series field from armature circuit
- Bring both brushes out independently
- Preserve all windings

---

## 8. Rectifier Upgrade

- Selenium rectifiers replaced with KBPC5010 silicon bridges
- Lower voltage drop revealed wiring errors immediately
- No further issues once field wired correctly

---

## 9. Final Correct Wiring

### Armature
- A1 → Brush A
- A2 → Brush B

### Shunt Field (CRITICAL)
- Tie Shunt A and Shunt B together
- F1 → Shunt A+B
- F2 → Shunt Center Tap

### Not connected
- Series field leads
- Any other unused taps

---

## 10. Final Result

- Smooth speed control
- Proper reversing
- Equal forward/reverse performance
- No fuse blowing
- No rectifier heating
- System behaves as original design intended

---

## 11. Lessons Learned

1. Never assume motor topology from external leads
2. Resistance mapping is definitive
3. Compound motors hide complexity
4. Silicon rectifiers expose latent faults
5. GE shunt fields often use outer-end strapping

---

## 12. Closing

This repair required careful measurement, theory, and patience.
The final solution is electrically correct, stable, and documented.

If future work is required: **map first, wire second, power last**.
