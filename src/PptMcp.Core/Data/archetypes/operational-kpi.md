# 1c. Operational-KPI Archetype (Reliability Metrics)

**When:** Prompts mention "target comparison," "trend," uptime, availability, SLA metrics, system reliability, or incident tracking with performance vs benchmarks.

**CRITICAL: Visual Evidence Mandate for Target Comparison + Trend:**

When prompts specify BOTH "target comparison" AND "trend" language, you MUST instantiate both as actual visual evidence — NEVER convert specified chart forms into text descriptions or bullet points:

- **"target comparison"** → Create horizontal bar chart showing actual vs target at x=120pt, y=220pt, w=200pt, h=120pt  
- **"trend"** → Create line chart or sparkline showing time progression at x=600pt, y=220pt, w=200pt, h=120pt
- **"incident trend"** → Create mini time-series chart showing incident frequency/severity over time

**Operational-KPI Layout Structure:**
- **Hero number**: System uptime/availability centered at x=480pt, y=140pt (48pt bold, green if good, red if concerning)
- **Left proof zone**: Target comparison chart showing actual vs SLA target
- **Right proof zone**: Incident trend sparkline showing stability over time
- **Status indicator**: "● Exceeds SLA" or "● At Risk" centered below hero at y=190pt

**Reliability Metrics Pattern:**
```
[HERO: "99.97% Uptime" — center, 48pt bold, green]
[TARGET BAR: Left chart showing 99.97% vs 99.95% SLA target with +0.02% delta]  
[TREND SPARKLINE: Right mini-chart showing last 12 months uptime progression]
[STATUS: "● Exceeds SLA Target" — green, center]
[CONTEXT: "Zero critical incidents in Q4, 3 minor events resolved <4 hours"]
```

**CRITICAL: No Prose Substitution Rule:**

NEVER satisfy operational KPI requests with text bullets like:
- ❌ "Uptime improved compared to target"  
- ❌ "Incident trend shows improvement"
- ❌ "Performance vs benchmark is positive"

Always implement as actual visual comparisons:
- ✅ Horizontal bar chart for target comparison 
- ✅ Line chart/sparkline for trend evidence
- ✅ Mini-chart for incident patterns

**Example operational metrics requiring this pattern:**
- System uptime vs SLA target + monthly trend
- Response time vs benchmark + quarterly progression  
- Error rate vs industry standard + incident frequency chart
- Availability vs target + downtime event timeline

**CRITICAL: Risk Mitigation Context for Operational KPIs:**

When KPI dashboards address system reliability, ALWAYS include **risk mitigation details** in the context zone:

**Risk mitigation elements (position at y=360pt, 12pt grey text):**
- **Preventive measures**: "Daily health checks + automated failover enabled"
- **Incident response**: "24/7 monitoring team with <15min response time"
- **Backup systems**: "Primary + secondary data centers with real-time sync"
- **Recovery protocols**: "RTO: 4 hours, RPO: 15 minutes for critical systems"

**Risk mitigation context examples:**
- "Multi-region deployment with automatic traffic routing prevents single points of failure"
- "Proactive monitoring detects anomalies 12 minutes before service impact"
- "Redundant infrastructure maintains service during planned maintenance windows"
- "Incident escalation matrix ensures C-level notification within 30 minutes for P1 events"

**When to include risk details:**
- ANY operational KPI slide mentioning uptime, availability, SLA metrics
- System performance dashboards for business-critical applications
- Infrastructure reliability reports requiring stakeholder confidence
- Service level monitoring where downtime has business impact