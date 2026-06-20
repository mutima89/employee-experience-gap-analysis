# Employee Experience Gap Analysis & Retention Risk Assessment

A professionally designed Excel workbook for measuring the gap between **organizational promises** and **employee actual experience**, with scientifically defensible scoring for engagement risk, retention risk, and cultural alignment.

## Overview

This workbook implements a **6-dimension gap analysis methodology** drawing on organizational psychology and employee engagement research. Each dimension is weighted by its proven impact on retention and engagement outcomes.

### Dimensions

| Dimension | Weight | What It Measures |
|-----------|--------|-----------------|
| **Job Fit** | 20% | Role clarity, skill utilization, career growth, workload, meaningfulness |
| **Manager Effectiveness** | 20% | Coaching, feedback, fairness, communication, trust |
| **Team Collaboration** | 15% | Cooperation, psychological safety, knowledge sharing, conflict resolution |
| **Workplace Culture** | 15% | Inclusion, respect, transparency, recognition, values alignment |
| **Innovation & Empowerment** | 15% | Autonomy, decision authority, idea acceptance, learning, experimentation |
| **Rewards & Growth** | 15% | Compensation, promotion, recognition, training, career progression |

## Sheet Structure

| Sheet | Description |
|-------|-------------|
| **Survey** | 30-item question bank with dual-scale (Promise 1–10, Experience 1–10) |
| **Raw Data** | Respondent entry sheet with data validation and 50-sample dataset |
| **Scoring Engine** | Native Excel formulas: gaps, health scores (0–100), Overall Experience Index, Retention Risk |
| **Dashboard** | Executive KPIs, Radar Chart (Promise vs Experience), Heat Map, Risk Matrix, filterable department comparison |
| **Reliability Analysis** | Cronbach's Alpha (α ≥ 0.70 target) for each dimension's internal consistency |
| **Driver Analysis** | CORREL() analysis ranking each dimension's relationship to the Overall Index |
| **Turnover Prediction** | Turnover probability model with per-respondent and aggregate risk classification |

## Key Formulas

- **Gap** = Promise Score − Experience Score
- **Dimension Health** = 100 − (Gap × 20), clamped to 0–100
- **Overall Experience Index** = Weighted average of all 6 dimension health scores
- **Retention Risk** = (Manager×25%) + (Culture×20%) + (Growth×20%) + (Job Fit×15%) + (Team×10%) + (Innovation×10%)
- **Turnover Probability** = (Retention Risk × 0.5) + (Manager Gap × 0.3) + (Culture Gap × 0.2)

## Usage

1. **Add respondents**: Enter employee survey data in the `Raw Data` sheet (columns P01–E30, 1–10 scale)
2. **Review scores**: The `Scoring Engine` sheet auto-calculates all dimension scores and gaps
3. **Explore the Dashboard**: Filter by department/location using dropdowns to compare groups
4. **Assess reliability**: Check Cronbach's Alpha in the `Reliability Analysis` sheet (target ≥ 0.70)
5. **Identify drivers**: The `Driver Analysis` sheet shows which dimensions most impact overall experience
6. **Predict turnover**: The `Turnover Prediction` sheet estimates retention risk per respondent

## Technical Notes

- All calculations use **native Excel formulas** — no VBA, macros, or external dependencies
- Compatible with Excel 2016+, Excel for Microsoft 365, and Excel for Mac
- Supports up to **500 respondents** (extend row references in formulas if exceeding this)
- Sample data included for 50 respondents across 6 departments, 5 locations, and 4 business units
- Professional corporate design with Blue/Navy/Green/Amber/Red color-coding

## Regenerating the Workbook

```bash
python generate_workbook.py
```

Requires Python 3.8+ and openpyxl 3.1+.

## License

MIT
