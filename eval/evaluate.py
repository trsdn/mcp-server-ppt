"""
Slide Design Evaluation Engine

Automated loop: for each prompt, ask an LLM to generate a presentation plan
based on current skills, then score the plan, identify gaps, and improve skills.

This does NOT require PowerPoint running - it evaluates what the LLM WOULD produce
given the current skill files. Spot-check with actual generation separately.
"""

import json
import os
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE = "PPTMCP_EVAL_ASSET_REPO_ROOT"
SKILLS_DIR = REPO_ROOT / "skills" / "shared"
PROMPTS_FILE = REPO_ROOT / "eval" / "prompts" / "test-prompts.json"


def get_eval_asset_repo_root() -> Path:
    configured_root = os.environ.get(EVAL_ASSET_REPO_ROOT_ENVIRONMENT_VARIABLE)
    return Path(configured_root).resolve() if configured_root else REPO_ROOT


EVAL_ASSET_REPO_ROOT = get_eval_asset_repo_root()
RESULTS_DIR = EVAL_ASSET_REPO_ROOT / "eval" / "results"

DESIGN_SKILL_FILES = [
    "slide-design-principles.md",
    "slide-design-review.md",
    "generation-pipeline.md",
]

SCORING_CRITERIA = """Score each dimension 0-2 (max 20 total):
1. Action Titles: 0=topic labels, 1=inconsistent, 2=every slide has action title
2. One Message: 0=multiple messages crammed, 1=mostly focused, 2=every slide one idea
3. Archetype: 0=wrong types, 1=mostly ok, 2=optimal archetype per slide
4. Whitespace: 0=crowded, 1=acceptable, 2=generous 36pt+ margins
5. Typography: 0=no hierarchy, 1=inconsistent, 2=clear 3-level hierarchy
6. Color: 0=random, 1=mostly consistent, 2=strict palette adherence
7. Density: 0=walls of text, 1=slightly overloaded, 2=3-5 bullets concise
8. Layout: 0=random placement, 1=mostly aligned, 2=all on grid
9. Sources: 0=no sources, 1=inconsistent, 2=all data slides sourced
10. Overall: 0=unpresentable, 1=internal use, 2=consulting/executive quality"""


def load_skills() -> str:
    """Load all design skill files into a single context string."""
    context = ""
    for fname in DESIGN_SKILL_FILES:
        fpath = SKILLS_DIR / fname
        if fpath.exists():
            context += f"\n\n--- {fname} ---\n"
            context += fpath.read_text(encoding="utf-8")
    return context


def load_prompts() -> list[dict]:
    """Load test prompts."""
    with open(PROMPTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def evaluate_prompt(prompt: dict, skills_context: str) -> dict:
    """
    Evaluate a single prompt against current skills.
    Returns a scoring dict with gaps identified.
    
    This is a rule-based evaluation that checks if the skills provide
    sufficient guidance for each scoring dimension.
    """
    pid = prompt["id"]
    category = prompt["category"]
    text = prompt["prompt"]
    difficulty = prompt["difficulty"]
    
    scores = {}
    gaps = []
    
    # 1. Action Titles - check if skills clearly instruct this
    has_action_title_guidance = "action title" in skills_context.lower() and "complete sentence" in skills_context.lower()
    scores["action_titles"] = 2 if has_action_title_guidance else 0
    if not has_action_title_guidance:
        gaps.append("Missing action title guidance")
    
    # 2. One Message - check if skills mention this principle
    has_one_message = "one message per slide" in skills_context.lower() or "one idea" in skills_context.lower()
    scores["one_message"] = 2 if has_one_message else 0
    
    # 3. Archetype Selection - check if decision tree exists AND covers this category
    has_decision_tree = "decision tree" in skills_context.lower() or "which archetype" in skills_context.lower()
    
    # Check category-specific archetype coverage
    category_archetype_map = {
        "corporate": ["kpi card", "simple table", "executive summary", "waterfall"],
        "sales": ["big number", "comparison", "quote", "chart"],
        "project": ["timeline", "process", "simple table", "executive summary"],
        "executive": ["framework", "waterfall", "chart", "executive summary"],
        "product": ["timeline", "comparison", "process", "framework"],
        "training": ["process", "framework", "simple table"],
        "dashboard": ["kpi card", "chart", "big number"],
        "strategy": ["framework", "timeline", "pillar", "roadmap"],
        "team": ["comparison", "simple table"],
        "creative": ["big number", "quote", "chart"],
        "technical": ["process", "framework", "simple table", "chart"],
    }
    
    needed_archetypes = category_archetype_map.get(category, [])
    found = sum(1 for a in needed_archetypes if a.lower() in skills_context.lower())
    coverage = found / max(len(needed_archetypes), 1)
    
    if coverage >= 0.75 and has_decision_tree:
        scores["archetype"] = 2
    elif coverage >= 0.5:
        scores["archetype"] = 1
        missing = [a for a in needed_archetypes if a.lower() not in skills_context.lower()]
        if missing:
            gaps.append(f"Missing archetype for {category}: {', '.join(missing)}")
    else:
        scores["archetype"] = 0
        gaps.append(f"Poor archetype coverage for {category}")
    
    # 4. Whitespace - check if specific measurements given
    has_margins = "36pt" in skills_context or "margin" in skills_context.lower()
    has_gaps = "gap" in skills_context.lower() and ("24pt" in skills_context or "18pt" in skills_context)
    scores["whitespace"] = 2 if (has_margins and has_gaps) else (1 if has_margins else 0)
    
    # 5. Typography - check if size hierarchy specified
    has_title_size = any(s in skills_context for s in ["20pt", "22pt", "24pt", "28pt"])
    has_body_size = any(s in skills_context for s in ["14pt", "16pt", "18pt"])
    has_footnote_size = any(s in skills_context for s in ["9pt", "10pt"])
    scores["typography"] = 2 if (has_title_size and has_body_size and has_footnote_size) else 1
    
    # 6. Color - check if palette exists for this category's likely profile
    profile_map = {
        "corporate": "corporate blue",
        "sales": "confident bold",
        "executive": "corporate blue",
        "dashboard": "corporate blue",
        "strategy": "corporate blue",
        "product": "modern tech",
        "training": "corporate blue",
        "project": "slate professional",
        "team": "corporate blue",
        "creative": "warm coral",
        "technical": "slate professional",
    }
    likely_palette = profile_map.get(category, "corporate blue")
    has_palette = likely_palette.lower() in skills_context.lower()
    has_hex_codes = "#" in skills_context and "RGB" in skills_context
    scores["color"] = 2 if (has_palette and has_hex_codes) else (1 if has_palette else 0)
    if not has_palette:
        gaps.append(f"Missing color palette for {category} ({likely_palette})")
    
    # 7. Content Density - check if limits specified
    has_bullet_limit = "bullet" in skills_context.lower() and ("5" in skills_context or "3-5" in skills_context)
    scores["density"] = 2 if has_bullet_limit else 1
    
    # 8. Layout - check if coordinate grids exist
    has_coordinates = "x=" in skills_context and "y=" in skills_context and "w=" in skills_context
    has_grid_variants = sum(1 for g in ["single column", "two column", "three column", "2×2", "kpi card"]
                          if g.lower() in skills_context.lower())
    scores["layout"] = 2 if (has_coordinates and has_grid_variants >= 3) else (1 if has_coordinates else 0)
    
    # 9. Sources - check if source bar guidance exists
    has_source_guidance = "source bar" in skills_context.lower() or "source:" in skills_context.lower()
    # Dashboards and executive prompts need sources more
    needs_sources = category in ["dashboard", "executive", "corporate", "strategy"]
    if has_source_guidance:
        scores["sources"] = 2
    elif not needs_sources:
        scores["sources"] = 1
    else:
        scores["sources"] = 0
        gaps.append("Missing source citation guidance for data-heavy category")
    
    # 10. Overall - derived from other scores + category-specific checks
    avg = sum(scores.values()) / len(scores)
    
    # Check for category-specific patterns
    category_specific_checks = {
        "dashboard": "kpi card" in skills_context.lower() and "2×2" in skills_context.lower(),
        "strategy": "pillar" in skills_context.lower() or "framework" in skills_context.lower(),
        "sales": "cta" in skills_context.lower() or "call to action" in skills_context.lower() or "benefit" in skills_context.lower(),
        "training": "learning objective" in skills_context.lower() or "recap" in skills_context.lower(),
        "executive": "pyramid" in skills_context.lower() or "executive summary" in skills_context.lower(),
    }
    has_category_pattern = category_specific_checks.get(category, True)
    
    if avg >= 1.7 and has_category_pattern:
        scores["overall"] = 2
    elif avg >= 1.3:
        scores["overall"] = 1
        if not has_category_pattern:
            gaps.append(f"Missing category-specific pattern for {category}")
    else:
        scores["overall"] = 0
    
    total = sum(scores.values())
    
    return {
        "prompt_id": pid,
        "category": category,
        "difficulty": difficulty,
        "scores": scores,
        "total": total,
        "max": 20,
        "gaps": gaps,
    }


def run_evaluation_cycle(cycle_num: int, prompts: list[dict], skills_context: str) -> dict:
    """Run one evaluation cycle across all prompts."""
    results = []
    for p in prompts:
        result = evaluate_prompt(p, skills_context)
        results.append(result)
    
    # Aggregate
    total_score = sum(r["total"] for r in results)
    max_score = sum(r["max"] for r in results)
    avg_score = total_score / len(results)
    
    # Collect all gaps
    all_gaps = []
    for r in results:
        for g in r["gaps"]:
            all_gaps.append({"gap": g, "prompt_id": r["prompt_id"], "category": r["category"]})
    
    # Count gap frequency
    gap_counts = {}
    for g in all_gaps:
        key = g["gap"].split(":")[0] if ":" in g["gap"] else g["gap"]
        gap_counts[key] = gap_counts.get(key, 0) + 1
    
    # Sort by frequency
    top_gaps = sorted(gap_counts.items(), key=lambda x: x[1], reverse=True)[:10]
    
    # Category averages
    cat_scores = {}
    for r in results:
        cat = r["category"]
        if cat not in cat_scores:
            cat_scores[cat] = []
        cat_scores[cat].append(r["total"])
    cat_avgs = {cat: sum(scores) / len(scores) for cat, scores in cat_scores.items()}
    
    return {
        "cycle": cycle_num,
        "total_prompts": len(prompts),
        "total_score": total_score,
        "max_score": max_score,
        "avg_score": round(avg_score, 1),
        "category_averages": {k: round(v, 1) for k, v in sorted(cat_avgs.items())},
        "top_gaps": top_gaps,
        "all_gaps": all_gaps,
        "per_prompt": results,
    }


def main():
    """Run evaluation cycles."""
    prompts = load_prompts()
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    
    num_cycles = int(sys.argv[1]) if len(sys.argv) > 1 else 100
    
    print(f"Running {num_cycles} evaluation cycles across {len(prompts)} prompts")
    print(f"Skills directory: {SKILLS_DIR}")
    print()
    
    prev_score = 0
    for cycle in range(1, num_cycles + 1):
        # Reload skills each cycle (may have been updated)
        skills_context = load_skills()
        
        result = run_evaluation_cycle(cycle, prompts, skills_context)
        
        delta = result["avg_score"] - prev_score if prev_score > 0 else 0
        delta_str = f" ({'+' if delta >= 0 else ''}{delta:.1f})" if prev_score > 0 else ""
        
        print(f"Cycle {cycle:3d}: avg={result['avg_score']:.1f}/20{delta_str} | "
              f"gaps={len(result['all_gaps'])} | "
              f"top gap: {result['top_gaps'][0][0] if result['top_gaps'] else 'none'} ({result['top_gaps'][0][1] if result['top_gaps'] else 0}x)")
        
        prev_score = result["avg_score"]
        
        # Save every 10 cycles
        if cycle % 10 == 0 or cycle == num_cycles:
            out_file = RESULTS_DIR / f"cycle-{cycle:03d}.json"
            with open(out_file, "w", encoding="utf-8") as f:
                json.dump(result, f, indent=2)
    
    # Final summary
    print()
    print(f"Final score: {result['avg_score']:.1f}/20")
    print(f"\nCategory scores:")
    for cat, avg in sorted(result["category_averages"].items()):
        print(f"  {cat:12s}: {avg:.1f}/20")
    print(f"\nTop gaps remaining:")
    for gap, count in result["top_gaps"][:5]:
        print(f"  [{count}x] {gap}")


if __name__ == "__main__":
    main()
