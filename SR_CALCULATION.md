# ⭐ Star Rating (SR) calculation methodology

## 🔎 Sources of SR confirmation

SR values are confirmed from the following two categories:

1. **Primary (real-world) sources**
   - In-game screenshots (osu! client)
   - YouTube thumbnails / descriptions
   - Other publicly verifiable materials

2. **Recalculated SR values**
   - Independently recalculated by the author using period-accurate tools and formulas  
   - Used to supplement or validate records when primary sources are incomplete or ambiguous


## 📅 SR calculation methods by era
---

### 🔁 Standard → Taiko conversion

For **standard-converted taiko maps**, SR is **underestimated** if calculated directly in standard mode.  
Therefore, the following workflow is used:

- Standard beatmaps are first converted into taiko beatmaps using `Std2Taiko`
- SR is then calculated in **taiko ruleset**

### 1️⃣ Before 2020-09-16 (ppv2 era)

- SR is calculated using `osu-trainer`, which imitates the legacy ppv2 SR formula.
- Please note that this calculation formula may not be accurate.

### 2️⃣ On and after 2020-09-16 (ppv3 era and later)

- SR is calculated using the **official osu!lazer SR formulas**
- The formulas are taken from period-accurate branches of the `osu!lazer codebase`
- Calculations are performed via `osu-tools / PerformanceCalculator`

Relevant repositories:

- **Std2Taiko**  
  https://github.com/calmeel/Std2Taiko

- **osu-trainer**  
  https://github.com/FunOrange/osu-trainer

- **osu!lazer codebase (formula reference)**  
  https://github.com/Piotrekol/osu

- **osu-tools / PerformanceCalculator (SR calculation tool)**  
  https://github.com/ppy/osu-tools

[Back to README](README.md)
