# 10. Executive Summary

**When:** Opening or closing a section with key findings.

**Layout:**
```
Title: top (e.g., "Three priorities will drive growth in 2027")
Bullets: 3-5 key points, each with:
  - Bold lead-in phrase + supporting sentence
  - Icon or number marker on left
  - 18pt font, 24pt line spacing
```

**INCIDENT POSTMORTEM COVER Pattern:**

When prompts mention **"incident postmortem," "incident report," or "executive incident summary"** with quantified business impact, use this specialized executive summary structure:

**Layout for Incident Executive Cover:**
```
Title: Incident summary with containment status 
  - "Database outage contained in 47 minutes — customer impact limited to checkout flow"
  
Quantified Impact Block: (position: x=240pt, y=180pt, w=480pt)
  - Revenue impact: "$23K in delayed transactions (0.8% of daily volume)"
  - Customer impact: "1,247 checkout attempts failed during 14:23-15:10 window"
  - System impact: "Payment processing restored, no data loss"

Containment Status Badge: (position: x=680pt, y=140pt)
  - "RESOLVED" in green box or "CONTAINED" in orange box
  
Next Steps/Approval Frame: (position: x=240pt, y=340pt, w=480pt)
  - "Immediate: Deploy hotfix to staging (2hrs) → Requires VP approval for production"
  - "This week: Complete root cause analysis → Present findings to architecture review"
  - "Next 30 days: Implement monitoring improvements → Budget impact $45K"

Lightweight Source Footer: (position: x=36pt, y=500pt)
  - "Source: incident management ticket #INC-2024-0892, payment logs 14:00-16:00 EST, Feb 18 2024"
```

**CRITICAL: Source Attribution for Executive Incident Slides:**
- **ALWAYS include** system source + date/time window for any quantified incident impact
- Format: "Source: [system name] [incident ID], [log source], [time window], [date]"
- Essential for executive credibility and audit trails
- Even cover slides require this when presenting quantified business impact

**Rules:**
- Each bullet stands alone as a complete thought
- Bold the first phrase of each bullet
- Use numbered list if order matters, bullets if not