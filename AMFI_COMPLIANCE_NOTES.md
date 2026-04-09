# AMFI Compliance Guidelines for Goal Planner

## Critical Rules for MFD (Mutual Fund Distributor)

### 1. NO Investment Advice (Unless SEBI-registered as Advisor)
- ❌ Cannot say "Buy this fund" or "Sell that fund"
- ❌ Cannot recommend specific fund switches
- ✅ CAN say "You're 20% short on equity allocation"
- ✅ CAN say "Consider equity MFs to close gap"
- Must add: "Consult your financial advisor before investing"

### 2. Conflict of Interest Disclosure (MANDATORY)
- MUST disclose you're an MF distributor
- MUST reveal you earn commission on MF sales
- MUST show disclaimer on every recommendation
- Example: "I earn X% commission if you invest through me"

### 3. Portfolio Integration Restrictions
- ✅ User uploads their portfolio manually = OK
- ✅ User authorizes OAuth to read broker data = OK
- ❌ Scrape without consent = ILLEGAL
- ❌ Store plain-text credentials = ILLEGAL
- ❌ Share portfolio data with competitors = ILLEGAL

### 4. Data Security (DPDP Act 2023 + SEBI Rules)
- MUST encrypt data at rest (AES-256)
- MUST encrypt in transit (TLS 1.2+)
- MUST have 5-year audit trail
- MUST notify users within 72 hours of any breach
- MUST have privacy policy + user consent

### 5. Suitable Recommendations Only
- CANNOT push high-risk products to conservative investors
- MUST assess risk profile before recommending
- CANNOT ignore past complaints
- MUST keep documentation of suitability

### 6. No Product Bundling
- ❌ "Use this tool only if you invest in our funds"
- ❌ "Get free analysis if you invest ₹10L"
- ✅ Tool is free/separate from sales

### 7. Fair Representation
- NO misleading performance claims
- NO "guaranteed returns" language
- MUST show: "Past performance ≠ future results"
- NO manipulation of assumptions (inflation, returns)

### 8. Record Keeping
- Keep all user interactions for 5 years
- Document every recommendation made
- Log portfolio access, changes, deletions
- Annual compliance audit required

### 9. Consumer Grievance Handling
- MUST have complaint resolution process
- Response within 7 days (AMFI standard)
- Escalation path to AMFI ombudsman
- No blocking users from filing complaints

### 10. Prohibited Practices
- ❌ Cold-calling using tool data
- ❌ Selling data to other advisors
- ❌ Auto-investing without explicit instruction
- ❌ Switching funds without documented approval
- ❌ Charging undisclosed fees

---

## Implementation Checklist for Goal Planner

### Before Launch:
- [ ] Legal review of Terms of Service (₹15-20K)
- [ ] Privacy policy addressing DPDP Act
- [ ] Conflict of interest disclosure template
- [ ] Compliance approval from AMFI nodal officer
- [ ] Data security audit
- [ ] Audit trail logging implemented

### During Operation:
- [ ] Every recommendation has disclaimer
- [ ] User consent forms for data access
- [ ] Monthly compliance audit
- [ ] Quarterly training for team
- [ ] Annual external audit

### Features to AVOID:
- Auto-rebalancing (looks like advice)
- Push notifications with buy/sell signals
- Performance tracking linked to commission
- Forced fund switches
- Hidden fees

### Features SAFE to Build:
- Portfolio vs. plan comparison
- Allocation gap analysis
- Goal progress tracking
- Asset class recommendations (not specific funds)
- Educational content on MFs

---

## Key Contacts:
- **AMFI**: www.amfiindia.com (Compliance section)
- **SEBI**: https://www.sebi.gov.in (Investor Protection)
- **DPDP Compliance**: Check Data Protection Board guidelines

**Last Updated:** April 2026
