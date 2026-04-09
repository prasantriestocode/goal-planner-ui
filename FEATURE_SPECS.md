# Feature Specifications: What-If Scenarios & AI Recommendations

## 1. WHAT-IF SCENARIOS

### How It Works:

**Current State:**
- User builds a plan: 60-year retirement, ₹50L corpus needed, ₹10K SIP
- Dashboard shows: "On track" ✓

**What-If Feature:**
- User clicks "What if I retire at 58 instead of 60?"
- System creates a COPY of the plan
- Recalculates everything: Goals, corpus, SIP required
- Shows side-by-side comparison

### User Flow:

```
Current Plan          vs.        Scenario: Retire at 58
─────────────────────────────────────────────────────
Retirement Age: 60               Retirement Age: 58 ✎ (editable)
Goal Corpus: ₹1.5Cr              Goal Corpus: ₹1.7Cr (+13%)
Current SIP: ₹10K                Required SIP: ₹13.5K (+35%)
Status: On track ✓               Status: Behind ⚠️ (shortfall ₹2L)

Action: [Save as Plan B] [Close]
```

### What Can Be Varied:

| Parameter | Current | What-If Example | Impact |
|-----------|---------|-----------------|--------|
| Retirement Age | 60 | 58, 62, 65 | Changes goal corpus & required SIP |
| Monthly SIP | ₹10K | ₹15K, ₹20K | Shows achievable goals |
| Inflation Rate | 6% | 5%, 7%, 8% | Changes future values |
| Return Rate | 12% | 10%, 15% | Portfolio growth impact |
| Life Expectancy | 85 | 80, 90 | Retirement duration |
| Child Education Cost | ₹20L | ₹15L, ₹30L | Goal corpus change |
| Home Loan Amount | ₹50L | ₹30L, ₹75L | Liability impact |

### Implementation:

```javascript
// Pseudo-code
function createScenario(originalPlan, changes) {
  // Clone the plan
  const scenario = JSON.parse(JSON.stringify(originalPlan));
  
  // Apply changes
  scenario.model.retirementAge = changes.retirementAge || originalPlan.model.retirementAge;
  scenario.model.inflationRate = changes.inflationRate || originalPlan.model.inflationRate;
  // ... more parameters
  
  // Recalculate everything
  scenario.goals = recalculateGoals(scenario);
  scenario.cashflow = recalculateCashflow(scenario);
  scenario.networth = recalculateNetworth(scenario);
  
  // Return both for comparison
  return {
    original: originalPlan,
    scenario: scenario,
    delta: {
      sipIncrease: scenario.requiredSip - originalPlan.requiredSip,
      corpusIncrease: scenario.totalGoalCorpus - originalPlan.totalGoalCorpus,
      status: getStatus(scenario)
    }
  };
}
```

### UI Flow:

```
Dashboard
├── Main Plan Card (Current)
│   └── [What-If Scenarios] button
│
└── Scenarios Modal
    ├── Slider: Retirement Age (55-75)
    ├── Slider: Monthly SIP (₹1K-₹1L)
    ├── Slider: Inflation Rate (2%-10%)
    ├── Slider: Return Rate (6%-18%)
    │
    ├── Side-by-Side Comparison:
    │   ├── Current Plan | Scenario A | Scenario B | Scenario C
    │   ├── ──────────────────────────────────────────
    │   ├── Retirement Age: 60 | 58 | 62 | 65
    │   ├── Required SIP: ₹10K | ₹13.5K | ₹7.8K | ₹5K
    │   ├── Goal Corpus: ₹1.5Cr | ₹1.7Cr | ₹1.3Cr | ₹0.9Cr
    │   ├── Status: On track ✓ | Behind ⚠️ | On track ✓ | Ahead ✓
    │
    ├── [Save Scenario as Plan B]
    └── [Close]
```

### Data Storage:

```javascript
// Store in Firestore
{
  plans: {
    currentPlan: { /* full plan data */ },
    scenarios: [
      {
        id: "scenario-1",
        name: "Retire at 58",
        baseParams: {
          retirementAge: 58,
          // ... other overrides
        },
        createdAt: timestamp,
        results: { /* cached results */ }
      },
      {
        id: "scenario-2",
        name: "10K to 15K SIP",
        baseParams: { currentSipPm: 15000 },
        createdAt: timestamp,
        results: { /* cached results */ }
      }
    ]
  }
}
```

### Benefits:
- ✅ Investors explore without anxiety
- ✅ Shows trade-offs visually
- ✅ Helps with decision-making
- ✅ "What do I need to retire at 58?" → Answer: "₹13.5K SIP"
- ✅ AMFI-compliant (just showing numbers, not advising)

---

## 2. AI-POWERED RECOMMENDATIONS

### How It Works:

**Scenario 1: Gap Detection**
```
System analyzes plan and finds:
- Goal X is 30% underfunded (₹10L gap)
- Current SIP will fall short by ₹2L/year
- Market downturn would push gap to ₹15L

AI Recommendation:
"⚠️ Goal: Son's Education
  Current gap: ₹10L (30% underfunded)
  Action: Increase monthly SIP by ₹5K to stay on track
  Impact: Closes gap in 24 months"
```

**Scenario 2: Behavior Flagging**
```
System checks historical data:
- User hasn't reviewed plan in 18 months
- Inflation rose 1% since last review
- Real goal costs increased by ₹5L
- Actual portfolio returns were 10% vs. assumed 12%

AI Recommendation:
"📊 Plan Review Needed
  Your assumptions may be outdated:
  • Inflation: 6% → 7% (impacts goal costs +₹5L)
  • Your returns: 10% vs. assumed 12%
  • Action: Update assumptions & recalculate"
```

**Scenario 3: Market-Triggered Rebalancing**
```
System monitors:
- Nifty 50 down 15% (market correction)
- User's portfolio likely down ~12%
- Emergency fund not affected
- Goal timeline still intact

AI Recommendation:
"📉 Market Correction Detected
  Your plan status: Still ON TRACK ✓
  Why? Emergency fund covers 18 months
  Suggested action: Stay invested, review in 3 months"
```

### Implementation Architecture:

```
┌─────────────────────────────────────────────────────────┐
│                    ANALYTICS ENGINE                       │
├─────────────────────────────────────────────────────────┤
│                                                          │
│  1. DATA COLLECTION                                     │
│     ├── Real-time: Monthly SIP activity                 │
│     ├── Quarterly: Portfolio performance                │
│     ├── Annual: Goal progress vs. corpus                │
│     ├── Market: Nifty index, inflation rates            │
│     └── User: Plan updates, assumption changes          │
│                                                          │
│  2. ANALYSIS RULES (No ML needed yet!)                  │
│     ├── Gap Analysis: (corpusNeeded - currentValue) / corpusNeeded
│     ├── Burn Rate: (monthlyExpense - monthlyIncome) / portfolio
│     ├── Risk Score: volatility * (yearsToRetire)
│     ├── Status Check: achievedSoFar vs. expectedAtThisDate
│     └── Drift Detection: actualReturns vs. assumptions
│                                                          │
│  3. RECOMMENDATION ENGINE                               │
│     ├── IF gap > 20% → "Increase SIP"                   │
│     ├── IF no update in 12mo → "Review plan"           │
│     ├── IF market -15% & fund intact → "Stay invested"  │
│     ├── IF actualReturns < assumed → "Adjust returns"   │
│     └── IF lifeExpectancy risk → "Increase corpus"      │
│                                                          │
│  4. PRIORITIZATION (Score high→low priority)            │
│     ├── Life-threatening gaps (retirement at risk)      │
│     ├── Medium-term gaps (needs 12mo action)            │
│     ├── Informational (nice-to-knows)                   │
│     └── Celebratory (you're ahead! ✓)                   │
│                                                          │
└─────────────────────────────────────────────────────────┘
```

### Rule-Based Recommendations (Start Here):

```javascript
function generateRecommendations(plan) {
  const recs = [];
  
  // Rule 1: Gap Detection
  for (let goal of plan.goals) {
    const gap = goal.corpus - goal.currentProvision;
    const gapPercent = gap / goal.corpus;
    
    if (gapPercent > 0.3) {
      recs.push({
        priority: "HIGH",
        type: "gap",
        goal: goal.name,
        message: `${goal.name} is ${(gapPercent*100).toFixed(0)}% underfunded (₹${formatRs(gap)})`,
        action: `Increase SIP by ₹${suggestSipIncrease(gap, goal.years)}`,
        impact: "Closes gap in 24 months"
      });
    } else if (gapPercent > 0.1) {
      recs.push({
        priority: "MEDIUM",
        message: `${goal.name} is slightly underfunded`,
        action: `Consider lump sum of ₹${formatRs(gap)}`
      });
    }
  }
  
  // Rule 2: Stale Plan Detection
  const daysSinceUpdate = (Date.now() - plan.lastUpdated) / (1000 * 60 * 60 * 24);
  if (daysSinceUpdate > 365) {
    recs.push({
      priority: "MEDIUM",
      type: "review",
      message: `Your plan hasn't been reviewed in ${Math.floor(daysSinceUpdate/365)} years`,
      action: "Review assumptions (inflation, returns, life expectancy)",
      impact: "Could impact goal achievability by 10-20%"
    });
  }
  
  // Rule 3: Market Event Response
  const niftyChange = getCurrentNiftyChange(); // -15% in last 3mo
  if (niftyChange < -10 && plan.portfolio.emergencyFund > 18 * plan.monthlyExpense) {
    recs.push({
      priority: "INFO",
      type: "market",
      message: "Market is down 15%, but your plan is safe",
      action: "Stay invested, review in 3 months",
      impact: "Historical: Markets recover within 18 months"
    });
  }
  
  // Rule 4: Assumption Drift
  const actualReturns = calculateActualReturns(plan);
  if (actualReturns < plan.preRetRate * 0.9) { // 10% below assumed
    recs.push({
      priority: "HIGH",
      type: "assumption",
      message: `Your returns (${actualReturns.toFixed(1)}%) are below assumption (${plan.preRetRate}%)`,
      action: "Update assumption or increase SIP by 10%",
      impact: "Will impact retirement corpus by ₹15L+"
    });
  }
  
  // Sort by priority
  return recs.sort(byPriority);
}
```

### AI Evolution (Phase 2+):

```
Phase 1 (Now): Rule-based
├── IF/THEN rules
├── Fixed thresholds
├── No machine learning
└── Easy to explain & AMFI-compliant

Phase 2 (6 months): Predictive
├── Predict market corrections
├── Estimate SIP defaults (who might miss payments)
├── Forecast retirement needs (based on lifestyle)
└── Personalize thresholds (your profile)

Phase 3 (12 months): Learning
├── Analyze all user plans (anonymized)
├── Find patterns ("people like you increase SIP 2x")
├── Predict likelihood of goal achievement
└── Recommend optimal allocation by profile
```

### UI for AI Recommendations:

```
Dashboard
│
├── Recommendation Center (Top)
│   │
│   ├── 🔴 HIGH PRIORITY
│   │   └── Son's Education Gap: 30% underfunded
│   │       Action: Increase SIP by ₹5K
│   │       Impact: Closes gap in 24 months
│   │       [View Details] [Dismiss]
│   │
│   ├── 🟡 MEDIUM PRIORITY
│   │   └── Plan not reviewed in 18 months
│   │       Action: Update assumptions
│   │       [Review Now]
│   │
│   └── 🟢 INFO
│       └── Market is down, but you're safe ✓
│           Stay invested
│
├── Main Dashboard (as before)
│
└── Recommendation History
    ├── Last 3 months
    ├── Dismissed: 2
    ├── Acted on: 1 (increased SIP)
    └── Pending: 1
```

### Data Flow:

```
User's Plan
    ↓
Daily/Weekly Analysis
    ↓
Check 20+ Rules
    ↓
Generate Scores
    ↓
Rank by Priority
    ↓
Show Top 3-5 on Dashboard
    ↓
User Acts / Dismisses
    ↓
Log Action in Audit Trail (for AMFI compliance)
```

---

## AMFI COMPLIANCE FOR AI RECOMMENDATIONS

### ✅ SAFE (COMPLIANT):
- "Your goal X is 20% underfunded" (FACT)
- "Increasing SIP by ₹5K closes gap in 24 months" (MATH)
- "Your assumptions may be outdated" (NOTIFICATION)
- "Market is down 15%, but emergency fund covers 18 months" (ANALYSIS)

### ❌ UNSAFE (VIOLATION):
- "BUY this fund now" (SPECIFIC ADVICE)
- "Your allocation is wrong, SWITCH immediately" (DIRECTIVE)
- "Interest rates will rise, reduce bonds" (PREDICTION)
- "Your portfolio will earn 14% next year" (GUARANTEE)

### Compliance Rules:
1. Every recommendation ends with: "Consult your advisor before acting"
2. Show calculations (transparent, auditable)
3. Log every recommendation (5-year audit trail)
4. Allow users to dismiss/opt-out
5. Disclose conflict of interest (you're MF distributor)

---

## IMPLEMENTATION TIMELINE

### Week 1-2: What-If Scenarios
- Add "Scenario" button to Dashboard
- Sliders for: Retirement Age, SIP, Inflation, Returns
- Side-by-side comparison UI
- Save scenarios to Firestore

### Week 3-4: AI Recommendations - Phase 1
- Implement gap detection
- Implement plan review detection
- Implement market event handler
- Show top 3 recommendations on dashboard
- Add audit logging

### Month 2: Refinement
- User feedback on scenarios
- Improve recommendation wording
- Add "Dismissed" tracking
- Mobile responsiveness

### Month 3+: Phase 2
- Predictive models
- Behavior-based thresholds
- Personalization

