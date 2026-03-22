import streamlit as st
import anthropic
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import math
import json
import os
from datetime import datetime

st.set_page_config(page_title="Fiction Writer + Detection Scorer", layout="wide")
st.title("Fiction Chapter Writer")

# ──────────────────────────────────────────────
# SOURCE RHYTHM ANALYSIS
# ──────────────────────────────────────────────

def analyze_source_rhythm(text):
    """
    Analyzes source texts (Dare, Quinn, Faulks, etc.) to extract a
    sentence rhythm fingerprint. This becomes the target for revision,
    not an abstract threshold.
    
    Returns a dict with the source's actual sentence length profile.
    """
    if not text or not text.strip():
        return None
    
    # Split into sentences
    sentence_endings = re.split(r'(?<=[.!?])\s+(?=[A-Z"\u201C])', text)
    sentences = [s.strip() for s in sentence_endings if s.strip()]
    lengths = [len(s.split()) for s in sentences]
    
    if len(lengths) < 10:
        return None  # not enough data
    
    total = len(lengths)
    mean_len = sum(lengths) / total
    variance = sum((x - mean_len) ** 2 for x in lengths) / total
    std_dev = math.sqrt(variance)
    cv = std_dev / mean_len if mean_len > 0 else 0
    
    # Distribution buckets
    ultra_short = sum(1 for l in lengths if l <= 5)       # "He sat." / "She went."
    short = sum(1 for l in lengths if 6 <= l <= 12)       # Simple declaratives
    medium = sum(1 for l in lengths if 13 <= l <= 25)     # Standard prose
    long = sum(1 for l in lengths if 26 <= l <= 40)       # Complex sentences
    very_long = sum(1 for l in lengths if l > 40)         # Accumulating clauses
    
    # Transition patterns: what follows what
    short_after_long = 0   # <=5 words after >=25 words (jagged rhythm)
    long_after_short = 0   # >=25 words after <=5 words
    same_band = 0          # consecutive sentences in same length band
    
    for i in range(1, len(lengths)):
        prev, curr = lengths[i-1], lengths[i]
        if prev >= 25 and curr <= 5:
            short_after_long += 1
        if prev <= 5 and curr >= 25:
            long_after_short += 1
        # Same band check (both ultra-short, both medium, etc.)
        prev_band = 0 if prev <= 5 else (1 if prev <= 12 else (2 if prev <= 25 else 3))
        curr_band = 0 if curr <= 5 else (1 if curr <= 12 else (2 if curr <= 25 else 3))
        if prev_band == curr_band:
            same_band += 1
    
    transitions = total - 1 if total > 1 else 1
    
    # Paratactic accumulation: sentences with 3+ "and" 
    paratactic = sum(1 for s in sentences if s.count(' and ') >= 3)
    
    # Longest and shortest
    max_len = max(lengths)
    min_len = min(lengths)
    
    # Range ratio (max / min) — higher = more variation
    range_ratio = max_len / min_len if min_len > 0 else max_len
    
    return {
        "total_sentences": total,
        "word_count": sum(lengths),
        "mean_length": round(mean_len, 1),
        "cv": round(cv, 3),
        "std_dev": round(std_dev, 1),
        "min_length": min_len,
        "max_length": max_len,
        "range_ratio": round(range_ratio, 1),
        "distribution": {
            "ultra_short_pct": round(100 * ultra_short / total, 1),  # <=5 words
            "short_pct": round(100 * short / total, 1),              # 6-12
            "medium_pct": round(100 * medium / total, 1),            # 13-25
            "long_pct": round(100 * long / total, 1),                # 26-40
            "very_long_pct": round(100 * very_long / total, 1),      # >40
        },
        "transitions": {
            "short_after_long_pct": round(100 * short_after_long / transitions, 1),
            "long_after_short_pct": round(100 * long_after_short / transitions, 1),
            "same_band_pct": round(100 * same_band / transitions, 1),
        },
        "paratactic_pct": round(100 * paratactic / total, 1),
    }


def display_source_profile(profile):
    """Renders the source rhythm profile in Streamlit."""
    st.markdown("### Source Text Rhythm Profile")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Sentence CV", profile["cv"], help="Target for chapter rhythm")
    col2.metric("Mean length", f"{profile['mean_length']} words")
    col3.metric("Range", f"{profile['min_length']}–{profile['max_length']} words")
    
    col4, col5, col6 = st.columns(3)
    col4.metric("Short after long", f"{profile['transitions']['short_after_long_pct']}%",
                help="% of times a ≤5-word sentence follows a ≥25-word sentence")
    col5.metric("Same-band consecutive", f"{profile['transitions']['same_band_pct']}%",
                help="Lower = more variation")
    col6.metric("Paratactic (3+ 'and')", f"{profile['paratactic_pct']}%")
    
    d = profile["distribution"]
    st.caption(f"Distribution: ≤5w: {d['ultra_short_pct']}% | 6-12w: {d['short_pct']}% | "
               f"13-25w: {d['medium_pct']}% | 26-40w: {d['long_pct']}% | >40w: {d['very_long_pct']}%")


def build_rhythm_instructions(source_profile, chapter_profile):
    """
    Compares the chapter's rhythm to the source's rhythm and generates
    specific revision instructions for sentence length variation.
    """
    if source_profile is None:
        return ""
    
    instructions = []
    
    src_cv = source_profile["cv"]
    ch_cv = chapter_profile
    
    if ch_cv < src_cv - 0.1:
        gap = round(src_cv - ch_cv, 2)
        instructions.append(
            f"SENTENCE RHYTHM: The source authors have a sentence length CV of {src_cv}. "
            f"This chapter is at {ch_cv} — {gap} points too uniform. To fix this:"
        )
        
        src_d = source_profile["distribution"]
        instructions.append(
            f"- Source distribution: {src_d['ultra_short_pct']}% ultra-short (≤5 words), "
            f"{src_d['very_long_pct']}% very long (>40 words). "
            f"The chapter needs MORE of both extremes — more 2-4 word sentences AND more 35-50 word accumulating sentences."
        )
        
        src_t = source_profile["transitions"]
        if src_t["short_after_long_pct"] > 5:
            instructions.append(
                f"- In the source texts, {src_t['short_after_long_pct']}% of sentences ≤5 words follow sentences ≥25 words. "
                f"After a long descriptive sentence, follow it with something brutally short: 'She went.' / 'I knew.' / 'Small mercies.'"
            )
        
        if source_profile["paratactic_pct"] > 2:
            instructions.append(
                f"- The source texts use paratactic accumulation (3+ 'and' in a sentence) {source_profile['paratactic_pct']}% of the time. "
                f"Occasionally fuse two medium sentences into one long one using 'and...and...and' rhythm: "
                f"'the gate open and the yard swept and everything that needed doing, done.'"
            )
        
        instructions.append(
            "- Do NOT just break long sentences into medium ones — that makes CV worse. "
            "Add ultra-short sentences (2-5 words) after long passages AND let some sentences run long with accumulating clauses."
        )
    
    return "\n".join(instructions)


# ──────────────────────────────────────────────
# DETECTION SCORING ENGINE
# ──────────────────────────────────────────────

def score_chapter(text):
    """
    Scores a chapter against 11 detection metrics derived from
    reverse-engineering Originality.ai's Lite 1.0.2 detector across
    1,925 segments (54,721 words) and 17 separate documents.
    
    Returns a dict with metric scores, overall risk assessment,
    and flagged passages with line-level detail.
    """
    words = text.split()
    word_count = len(words)
    if word_count == 0:
        return None
    
    kw = word_count / 1000  # per-thousand-word normalizer
    
    # ── Split into sentences ──
    sentence_endings = re.split(r'(?<=[.!?])\s+(?=[A-Z"\u201C])', text)
    sentences = [s.strip() for s in sentence_endings if s.strip()]
    sentence_lengths = [len(s.split()) for s in sentences]
    
    # ── 1. Em dash density ──
    em_dashes = text.count('\u2014') + text.count('--')
    em_dash_rate = em_dashes / kw if kw > 0 else 0
    
    # ── 2. "As though" / "as if" density ──
    as_though_count = len(re.findall(r'\bas though\b', text, re.I))
    as_if_count = len(re.findall(r'\bas if\b', text, re.I))
    as_though_rate = (as_though_count + as_if_count) / kw if kw > 0 else 0
    
    # ── 3. "The way he/she/they/I/men/people" density ──
    the_way_count = len(re.findall(r'\bthe way (?:he|she|they|I|it|men|women|people|a man|a woman|soldiers|hungry men)\b', text, re.I))
    the_way_rate = the_way_count / kw if kw > 0 else 0
    
    # ── 4. Negation-leading constructions ──
    negation_patterns = [
        r'\bIt was not\b', r'\bI did not\b', r'\bShe did not\b', r'\bHe did not\b',
        r'\bThat was not\b', r'\bThis was not\b', r'\bI had not\b',
        r'\bI was not\b', r'\bShe was not\b', r'\bHe was not\b',
        r'\bI could not\b', r'\bIt did not\b'
    ]
    negation_count = sum(len(re.findall(p, text)) for p in negation_patterns)
    negation_rate = negation_count / kw if kw > 0 else 0
    
    # ── 5. Period-to-comma ratio ──
    periods = text.count('.')
    commas = text.count(',')
    period_comma_ratio = periods / commas if commas > 0 else periods
    
    # ── 6. Dialogue density ──
    dialogue_matches = re.findall(r'[\u201C"][^"\u201D]*[\u201D"]', text)
    dialogue_words = sum(len(m.split()) for m in dialogue_matches)
    dialogue_pct = (dialogue_words / word_count * 100) if word_count > 0 else 0
    
    # ── 7. Sentence length variation (coefficient of variation) ──
    if len(sentence_lengths) > 1:
        mean_len = sum(sentence_lengths) / len(sentence_lengths)
        variance = sum((x - mean_len) ** 2 for x in sentence_lengths) / len(sentence_lengths)
        std_dev = math.sqrt(variance)
        cv = std_dev / mean_len if mean_len > 0 else 0
    else:
        cv = 0
        mean_len = sentence_lengths[0] if sentence_lengths else 0
    
    # ── 8. Metacognitive verbs ──
    meta_verbs = re.findall(r'\b(?:I\s+)?(?:noted|filed|registered|understood|recognised|recognized|observed)\b', text, re.I)
    meta_rate = len(meta_verbs) / kw if kw > 0 else 0
    
    # ── 9. "The fact that" and "not X but Y" ──
    fact_that = len(re.findall(r'\bthe fact that\b', text, re.I))
    not_x_but_y = len(re.findall(r'\bnot [a-z]+ but [a-z]+\b', text, re.I))
    analytical_frames = fact_that + not_x_but_y
    analytical_rate = analytical_frames / kw if kw > 0 else 0
    
    # ── 10. "Of a man/woman/person who" characterization ──
    of_person_who = len(re.findall(r'\bof a (?:man|woman|person|people) who\b', text, re.I))
    of_person_rate = of_person_who / kw if kw > 0 else 0
    
    # ── 11. Constructed similes and metaphors ──
    # Catches: "as though", "as if", "the way [pronoun]", "like a [noun] that/who/which",
    # "the kind of [noun] that", "[verb]ed the way", "as a [noun] [verb]s",
    # and extended metaphor constructions.
    # Does NOT catch dead metaphors or idioms (those are too varied to regex).
    simile_patterns = [
        r'\bas though\b',
        r'\bas if\b',
        r'\bthe way (?:he|she|they|I|it|men|women|people|a man|a woman|soldiers|hungry)\b',
        r'\blike a [a-z]+ (?:that|who|which)\b',
        r'\blike a [a-z]+ [a-z]+ing\b',           # "like a man carrying..."
        r'\bthe kind of [a-z]+ (?:that|who|which|you)\b',  # "the kind of man who..."
        r'\bthe sort of [a-z]+ (?:that|who|which|you)\b',
        r'\b[a-z]+ed the way [a-z]+\b',            # "moved the way soldiers..."
        r'\bas a [a-z]+ (?:does|would|might|could|who)\b',  # "as a man does", "as a man who"
        r'\bthe particular [a-z]+ (?:of|that)\b',  # "the particular quality of"
        r'\bwith the [a-z]+ of a\b',               # "with the patience of a"
        r'\bhad the [a-z]+ of a\b',                # "had the eyes of a"
        r'\bin the manner of\b',
        r'\bwith the air of\b',
    ]
    simile_count = 0
    simile_matches_all = []
    for pat in simile_patterns:
        found = re.findall(pat, text, re.I)
        simile_count += len(found)
        simile_matches_all.extend(found)
    # Deduplicate: "as though" and "as if" are already counted in metric 2,
    # but we want the combined simile count for the new metric
    simile_rate = simile_count / kw if kw > 0 else 0
    
    # ── Scoring: each metric gets GREEN / YELLOW / RED ──
    metrics = {}
    
    def rate(name, value, green_thresh, yellow_thresh, unit, invert=False):
        if invert:
            if value >= green_thresh:
                level = "GREEN"
            elif value >= yellow_thresh:
                level = "YELLOW"
            else:
                level = "RED"
        else:
            if value <= green_thresh:
                level = "GREEN"
            elif value <= yellow_thresh:
                level = "YELLOW"
            else:
                level = "RED"
        metrics[name] = {"value": round(value, 3), "level": level, "unit": unit}
    
    rate("Em dash density", em_dash_rate, 1.0, 2.5, "/1000w")
    rate("'As though/as if' density", as_though_rate, 0.2, 0.6, "/1000w")
    rate("'The way he/she' density", the_way_rate, 0.5, 1.0, "/1000w")
    rate("Negation-leading density", negation_rate, 2.0, 4.0, "/1000w")
    rate("Period-to-comma ratio", period_comma_ratio, 1.8, 1.5, "ratio", invert=True)
    rate("Dialogue density", dialogue_pct, 15.0, 8.0, "% of words", invert=True)
    rate("Sentence length CV", cv, 1.2, 1.0, "coefficient", invert=True)
    rate("Metacognitive verb density", meta_rate, 0.3, 1.0, "/1000w")
    rate("Analytical frames", analytical_rate, 0.0, 0.2, "/1000w")
    rate("'Of a person who' density", of_person_rate, 0.0, 0.3, "/1000w")
    rate("Constructed simile density", simile_rate, 1.0, 3.0, "/1000w")
    
    # ── Flag specific high-risk passages ──
    flagged = []
    
    for i, sent in enumerate(sentences):
        risks = []
        
        if '\u2014' in sent or '--' in sent:
            risks.append("em_dash")
        
        if re.search(r'\bas though\b|\bas if\b', sent, re.I):
            risks.append("as_though")
        
        if re.search(r'\bthe way (?:he|she|they|I|it|men|women|people|a man|a woman|soldiers|hungry)\b', sent, re.I):
            risks.append("the_way")
        
        if re.search(r'\bof a (?:man|woman|person) who\b', sent, re.I):
            risks.append("of_person_who")
        
        if re.search(r'\b(?:I\s+)?(?:noted|filed|registered|understood|recognised|recognized)\b', sent, re.I):
            risks.append("metacognitive")
        
        if re.search(r'\bthe fact that\b', sent, re.I):
            risks.append("fact_that")
        
        if re.search(r'^(?:It|That|This|I|She|He) (?:was|did|had|could) not\b', sent):
            risks.append("negation_leading")
        
        comma_count = sent.count(',')
        sent_words = len(sent.split())
        if comma_count >= 4 and sent_words >= 35:
            risks.append("long_compound")
        
        if ('\u2014' in sent or '--' in sent) and sent_words > 25:
            if re.search(r'\bas though\b|\bthe way\b|\bwhich (?:was|meant|told)\b', sent, re.I):
                risks.append("obs_interp_coupling")
        
        # Constructed simile detection at sentence level
        simile_sentence_patterns = [
            r'\blike a [a-z]+ (?:that|who|which)\b',
            r'\blike a [a-z]+ [a-z]+ing\b',
            r'\bthe kind of [a-z]+ (?:that|who|which|you)\b',
            r'\bthe sort of [a-z]+ (?:that|who|which|you)\b',
            r'\b[a-z]+ed the way [a-z]+\b',
            r'\bas a [a-z]+ (?:does|would|might|could|who)\b',
            r'\bwith the [a-z]+ of a\b',
            r'\bhad the [a-z]+ of a\b',
            r'\bin the manner of\b',
            r'\bwith the air of\b',
            r'\bthe particular [a-z]+ (?:of|that)\b',
        ]
        for sp in simile_sentence_patterns:
            if re.search(sp, sent, re.I):
                risks.append("constructed_simile")
                break
        
        if risks:
            flagged.append({
                "index": i,
                "sentence": sent,
                "risks": risks,
                "risk_count": len(risks)
            })
    
    # ── Overall risk score ──
    red_count = sum(1 for m in metrics.values() if m["level"] == "RED")
    yellow_count = sum(1 for m in metrics.values() if m["level"] == "YELLOW")
    green_count = sum(1 for m in metrics.values() if m["level"] == "GREEN")
    
    if red_count >= 4:
        overall = "HIGH RISK"
    elif red_count >= 2 or (red_count >= 1 and yellow_count >= 3):
        overall = "MODERATE RISK"
    elif red_count == 0 and yellow_count <= 2:
        overall = "LOW RISK"
    else:
        overall = "MODERATE RISK"
    
    return {
        "word_count": word_count,
        "sentence_count": len(sentences),
        "metrics": metrics,
        "flagged": sorted(flagged, key=lambda x: -x["risk_count"]),
        "overall": overall,
        "red_count": red_count,
        "yellow_count": yellow_count,
        "green_count": green_count,
        "summary": {
            "em_dashes": em_dashes,
            "as_though_total": as_though_count + as_if_count,
            "the_way_total": the_way_count,
            "negation_total": negation_count,
            "simile_total": simile_count,
            "dialogue_word_pct": round(dialogue_pct, 1),
            "mean_sentence_length": round(mean_len, 1),
            "flagged_sentences": len(flagged),
            "total_sentences": len(sentences)
        }
    }


def build_revision_prompt(chapter_text, score_result, source_profile=None):
    """
    Builds a targeted revision prompt that identifies specific flagged
    passages and tells the model how to fix them without destroying
    the writing.
    """
    flagged = score_result["flagged"]
    metrics = score_result["metrics"]
    
    top_flagged = flagged[:15]
    
    risk_descriptions = {
        "em_dash": "Contains em dash — often used for interpretive gloss. Replace with a period and a new sentence, or remove the gloss entirely.",
        "as_though": "Contains 'as though' or 'as if' — appends speculative interpretation to concrete observation. Cut the simile. Let the action stand alone.",
        "the_way": "Contains 'the way he/she/they' — embeds interpretation inside description. Replace with direct physical observation.",
        "of_person_who": "Contains 'of a man/woman who' — analytical characterization. Replace with action or a short direct statement.",
        "metacognitive": "Contains metacognitive verb (noted/filed/registered/understood). Show through action, don't narrate the cognitive process.",
        "fact_that": "Contains 'the fact that' — analytical abstraction. Remove the frame; state the fact directly.",
        "negation_leading": "Negation-leading construction — defines by what something is not. Replace with a direct assertion of what it is.",
        "long_compound": "Long compound sentence with many commas — may read as inventory or process narration. Consider breaking into shorter sentences OR following with a brutally short sentence (2-5 words).",
        "obs_interp_coupling": "Observation-interpretation coupling — concrete detail and interpretive gloss fused in one sentence. Separate them or cut the interpretation.",
        "constructed_simile": "Constructed simile or metaphor — literary comparison that reads as crafted rather than natural. Replace with a flat statement, an idiom, or cut entirely. Only colloquial/idiomatic figures of speech pass detection."
    }
    
    passage_list = ""
    for i, item in enumerate(top_flagged):
        risks_text = "; ".join(risk_descriptions.get(r, r) for r in item["risks"])
        passage_list += f"\n\nFLAGGED PASSAGE {i+1}:\n\"{item['sentence'][:300]}\"\nISSUES: {risks_text}"
    
    metric_warnings = ""
    for name, data in metrics.items():
        if data["level"] == "RED":
            metric_warnings += f"\n- {name}: {data['value']} {data['unit']} — RED (high detection risk)"
        elif data["level"] == "YELLOW":
            metric_warnings += f"\n- {name}: {data['value']} {data['unit']} — YELLOW (moderate risk)"
    
    # Build rhythm instructions from source profile
    chapter_cv = metrics.get("Sentence length CV", {}).get("value", 0)
    rhythm_block = build_rhythm_instructions(source_profile, chapter_cv)
    
    prompt = f"""You are revising a chapter of fiction to reduce AI detection risk while preserving the voice, content, and literary quality.

IMPORTANT RULES:
- Preserve the original text wherever possible. Only modify flagged passages.
- Do NOT rewrite the entire chapter. Output the COMPLETE chapter with only the flagged passages revised.
- Do NOT add new content, scenes, or dialogue that wasn't there before.
- Do NOT remove scenes, beats, or plot points.
- Revisions should make the prose sound more like someone talking and less like someone writing.
- Prefer colloquial, idiomatic, tossed-off phrasing over composed literary images.
- When replacing an em dash gloss, don't just move the gloss to a new sentence — consider cutting it entirely if the image works without it.
- When cutting "as though" or "the way" constructions, let the physical action stand alone. Trust the reader.
- ELIMINATE all constructed similes and metaphors. Replace "as though [interpretation]", "the way [pronoun] [verb]s when...", "like a [noun] that...", "with the air of a...", "the kind of [noun] who..." with either: (a) a flat direct statement, (b) an idiomatic/colloquial phrase, or (c) nothing — just cut it. The only figurative language that passes detection is dead metaphors and idioms that a person would say without thinking ("a rag that had seen worse" = good; "as though hurrying would remind his body how long it had been" = bad).
- Do NOT introduce new em dashes, "as though" constructions, similes, metaphors, or metacognitive verbs.
{rhythm_block}

CHAPTER METRICS (current scores):
{metric_warnings}

The following passages have been identified as highest detection risk:
{passage_list}

Here is the complete chapter. Revise ONLY the flagged passages. Output the full chapter text with your revisions in place.

<chapter>
{chapter_text}
</chapter>"""
    
    return prompt


def generate_report(score_result, chapter_text, label, pass_num, source_profile=None):
    """
    Generates a Word document report for a scoring pass.
    Contains: summary, metrics table, rhythm comparison, flagged passages with context.
    """
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    
    # Title
    title = doc.add_heading(f'Detection Score Report — {label}', level=1)
    
    # Summary paragraph
    summary = score_result["summary"]
    overall = score_result["overall"]
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    p = doc.add_paragraph()
    p.add_run(f"Generated: {timestamp}\n").bold = False
    p.add_run(f"Words: {score_result['word_count']:,} | "
              f"Sentences: {summary['total_sentences']} | "
              f"Mean sentence length: {summary['mean_sentence_length']} words\n")
    p.add_run(f"Dialogue: {summary['dialogue_word_pct']}% of words | "
              f"Flagged sentences: {summary['flagged_sentences']} of {summary['total_sentences']} "
              f"({round(100*summary['flagged_sentences']/max(summary['total_sentences'],1))}%)\n")
    
    # Overall risk
    p2 = doc.add_paragraph()
    risk_run = p2.add_run(f"OVERALL: {overall}")
    risk_run.bold = True
    risk_run.font.size = Pt(14)
    if overall == "HIGH RISK":
        risk_run.font.color.rgb = RGBColor(200, 0, 0)
    elif overall == "MODERATE RISK":
        risk_run.font.color.rgb = RGBColor(200, 150, 0)
    else:
        risk_run.font.color.rgb = RGBColor(0, 150, 0)
    
    p2.add_run(f"  ({score_result['red_count']} RED, {score_result['yellow_count']} YELLOW, {score_result['green_count']} GREEN)")
    
    # Metrics table
    doc.add_heading('Metrics', level=2)
    
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0].cells
    hdr[0].text = 'Metric'
    hdr[1].text = 'Value'
    hdr[2].text = 'Unit'
    hdr[3].text = 'Status'
    
    for name, data in score_result["metrics"].items():
        row = table.add_row().cells
        row[0].text = name
        row[1].text = str(data["value"])
        row[2].text = data["unit"]
        row[3].text = data["level"]
        # Color the status cell
        for paragraph in row[3].paragraphs:
            for run in paragraph.runs:
                if data["level"] == "RED":
                    run.font.color.rgb = RGBColor(200, 0, 0)
                    run.bold = True
                elif data["level"] == "YELLOW":
                    run.font.color.rgb = RGBColor(200, 150, 0)
                    run.bold = True
                else:
                    run.font.color.rgb = RGBColor(0, 130, 0)
    
    # Counts summary
    doc.add_paragraph(
        f"Em dashes: {summary['em_dashes']} | "
        f"'As though/if': {summary['as_though_total']} | "
        f"'The way': {summary['the_way_total']} | "
        f"Negation-leading: {summary['negation_total']} | "
        f"Constructed similes: {summary['simile_total']}"
    )
    
    # Source rhythm comparison
    if source_profile:
        doc.add_heading('Sentence Rhythm — Source Comparison', level=2)
        ch_cv = score_result["metrics"].get("Sentence length CV", {}).get("value", 0)
        src_cv = source_profile["cv"]
        src_d = source_profile["distribution"]
        src_t = source_profile["transitions"]
        
        rtable = doc.add_table(rows=1, cols=3)
        rtable.style = 'Light Grid Accent 1'
        rhdr = rtable.rows[0].cells
        rhdr[0].text = 'Measure'
        rhdr[1].text = 'Source'
        rhdr[2].text = 'Chapter'
        
        rhythm_rows = [
            ("Sentence CV", str(src_cv), str(ch_cv)),
            ("Mean sentence length", f"{source_profile['mean_length']}w", f"{summary['mean_sentence_length']}w"),
            ("Range (min–max)", f"{source_profile['min_length']}–{source_profile['max_length']}w", "—"),
            ("Ultra-short (≤5w)", f"{src_d['ultra_short_pct']}%", "—"),
            ("Short (6-12w)", f"{src_d['short_pct']}%", "—"),
            ("Medium (13-25w)", f"{src_d['medium_pct']}%", "—"),
            ("Long (26-40w)", f"{src_d['long_pct']}%", "—"),
            ("Very long (>40w)", f"{src_d['very_long_pct']}%", "—"),
            ("Short-after-long transitions", f"{src_t['short_after_long_pct']}%", "—"),
            ("Same-band consecutive", f"{src_t['same_band_pct']}%", "—"),
            ("Paratactic (3+ 'and')", f"{source_profile['paratactic_pct']}%", "—"),
        ]
        for measure, src_val, ch_val in rhythm_rows:
            row = rtable.add_row().cells
            row[0].text = measure
            row[1].text = src_val
            row[2].text = ch_val
        
        gap = round(src_cv - ch_cv, 3)
        if gap > 0.1:
            p_rhythm = doc.add_paragraph()
            r = p_rhythm.add_run(f"Rhythm gap: {gap} — chapter is too uniform compared to source texts.")
            r.font.color.rgb = RGBColor(200, 150, 0)
            r.bold = True
    
    # Flagged passages
    flagged = score_result["flagged"]
    if flagged:
        doc.add_heading(f'Flagged Passages ({len(flagged)})', level=2)
        
        for item in flagged[:25]:
            p = doc.add_paragraph()
            tag_run = p.add_run(f"[{item['risk_count']} flags: {', '.join(item['risks'])}]")
            tag_run.bold = True
            tag_run.font.size = Pt(9)
            tag_run.font.color.rgb = RGBColor(180, 0, 0)
            
            p2 = doc.add_paragraph()
            text_run = p2.add_run(item['sentence'][:400])
            text_run.font.size = Pt(10)
            text_run.italic = True
            
            doc.add_paragraph()  # spacer
    
    # Full chapter text
    doc.add_heading('Full Chapter Text', level=2)
    for para_text in chapter_text.split("\n"):
        if para_text.strip():
            doc.add_paragraph(para_text)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def display_scorecard(score_result, source_profile=None):
    """Renders the scorecard in Streamlit."""
    
    st.markdown("### Detection Risk Scorecard")
    
    overall = score_result["overall"]
    if overall == "HIGH RISK":
        st.error(f"Overall: {overall} — {score_result['red_count']} red, {score_result['yellow_count']} yellow, {score_result['green_count']} green")
    elif overall == "MODERATE RISK":
        st.warning(f"Overall: {overall} — {score_result['red_count']} red, {score_result['yellow_count']} yellow, {score_result['green_count']} green")
    else:
        st.success(f"Overall: {overall} — {score_result['red_count']} red, {score_result['yellow_count']} yellow, {score_result['green_count']} green")
    
    # Metrics table — 4 columns to fit 11 metrics
    col1, col2, col3, col4 = st.columns(4)
    metric_items = list(score_result["metrics"].items())
    
    for i, (name, data) in enumerate(metric_items):
        target_col = [col1, col2, col3, col4][i % 4]
        icon = {"GREEN": "\u2705", "YELLOW": "\u26A0\uFE0F", "RED": "\u274C"}[data["level"]]
        target_col.metric(
            label=f"{icon} {name}",
            value=f"{data['value']} {data['unit']}",
            delta=data["level"],
            delta_color="normal" if data["level"] == "GREEN" else ("off" if data["level"] == "YELLOW" else "inverse")
        )
    
    s = score_result["summary"]
    st.caption(f"{s['total_sentences']} sentences | Mean length: {s['mean_sentence_length']} words | "
               f"{s['flagged_sentences']} flagged ({round(100*s['flagged_sentences']/max(s['total_sentences'],1))}%) | "
               f"Dialogue: {s['dialogue_word_pct']}% | Similes: {s['simile_total']}")
    
    # Source rhythm comparison
    if source_profile:
        ch_cv = score_result["metrics"].get("Sentence length CV", {}).get("value", 0)
        src_cv = source_profile["cv"]
        src_d = source_profile["distribution"]
        gap = round(src_cv - ch_cv, 3)
        
        if gap > 0.1:
            st.warning(
                f"**Rhythm gap:** Chapter CV {ch_cv} vs Source CV {src_cv} (gap: {gap}). "
                f"Source has {src_d['ultra_short_pct']}% ultra-short sentences (≤5w) and "
                f"{src_d['very_long_pct']}% very long (>40w). "
                f"Short-after-long transitions: {source_profile['transitions']['short_after_long_pct']}%."
            )
        elif gap > 0:
            st.info(f"Rhythm close to source: Chapter CV {ch_cv} vs Source CV {src_cv}.")
        else:
            st.success(f"Rhythm matches or exceeds source: Chapter CV {ch_cv} vs Source CV {src_cv}.")
    
    flagged = score_result["flagged"]
    if flagged:
        with st.expander(f"Flagged Passages ({len(flagged)} sentences)", expanded=False):
            for item in flagged[:20]:
                risk_tags = " ".join([f"`{r}`" for r in item["risks"]])
                st.markdown(f"**[{item['risk_count']} flags]** {risk_tags}")
                st.text(item["sentence"][:200])
                st.markdown("---")


# ──────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────

with st.sidebar:
    st.header("Settings")
    api_key = st.text_input("Anthropic API Key", type="password")
    
    model_choice = st.selectbox("Model", [
        "Sonnet",
        "Sonnet Extended Thinking",
        "Opus",
        "Haiku"
    ])
    
    model_map = {
        "Sonnet": "claude-sonnet-4-20250514",
        "Sonnet Extended Thinking": "claude-sonnet-4-20250514",
        "Opus": "claude-opus-4-20250514",
        "Haiku": "claude-haiku-4-5-20251001"
    }
    
    model_id = model_map[model_choice]
    
    temperature = st.slider("Temperature", 0.0, 1.0, 1.0, 0.1)
    max_tokens = st.number_input("Max Tokens", min_value=1000, max_value=128000, value=16000, step=1000)
    
    if model_choice == "Sonnet Extended Thinking":
        thinking_budget = st.number_input("Thinking Budget (tokens)", min_value=1000, max_value=50000, value=10000, step=1000)
        st.caption("Extended thinking requires temperature = 1.0.")
    
    st.markdown("---")
    st.header("Revision Settings")
    auto_revise = st.checkbox("Auto-revise after scoring", value=False)
    max_revision_passes = st.slider("Max revision passes", 1, 3, 2)
    
    st.markdown("---")
    st.header("Revision Model")
    revision_model_choice = st.selectbox("Model for revisions", [
        "Same as writing model",
        "Sonnet",
        "Haiku"
    ])


# ──────────────────────────────────────────────
# MAIN AREA
# ──────────────────────────────────────────────

st.subheader("Source Documents")
col1, col2, col3 = st.columns(3)
with col1:
    source_file = st.file_uploader("Source Texts", type=["txt", "docx"])
with col2:
    char_file = st.file_uploader("Character Profiles / Parts Bin", type=["txt", "docx"])
with col3:
    outline_file = st.file_uploader("Chapter Outline", type=["txt", "docx"])

def read_uploaded(f):
    if f is None:
        return ""
    if f.name.endswith(".txt"):
        return f.read().decode("utf-8")
    elif f.name.endswith(".docx"):
        doc = Document(io.BytesIO(f.read()))
        return "\n".join([p.text for p in doc.paragraphs])
    return ""

prompt_default = """Using the source texts, character profiles, and chapter outline provided, write the chapter in one continuous pass from first sentence to last. Do not draft short and expand."""

prompt = st.text_area("Writing Prompt", value=prompt_default, height=150)

# ── Analyze source texts when uploaded ──
if source_file is not None:
    source_text_for_analysis = read_uploaded(source_file)
    # Reset the file pointer so it can be read again later
    source_file.seek(0)
    if source_text_for_analysis.strip():
        profile = analyze_source_rhythm(source_text_for_analysis)
        if profile:
            st.session_state.source_profile = profile
            with st.expander("Source Rhythm Profile", expanded=False):
                display_source_profile(profile)
        else:
            st.caption("Source text too short for rhythm analysis (need 10+ sentences).")
else:
    st.session_state.source_profile = None

# ──────────────────────────────────────────────
# SESSION STATE
# ──────────────────────────────────────────────

if "chapter_text" not in st.session_state:
    st.session_state.chapter_text = None
if "score_result" not in st.session_state:
    st.session_state.score_result = None
if "revision_history" not in st.session_state:
    st.session_state.revision_history = []
if "current_pass" not in st.session_state:
    st.session_state.current_pass = 0
if "reports" not in st.session_state:
    st.session_state.reports = []
if "source_profile" not in st.session_state:
    st.session_state.source_profile = None


def call_api(client, message_text, is_revision=False):
    """Makes an API call with the current settings."""
    
    if is_revision and revision_model_choice != "Same as writing model":
        rev_model = model_map.get(revision_model_choice, model_id)
        response = client.messages.create(
            model=rev_model,
            max_tokens=max_tokens,
            temperature=temperature,
            messages=[{"role": "user", "content": message_text}]
        )
    elif model_choice == "Sonnet Extended Thinking" and not is_revision:
        response = client.messages.create(
            model=model_id,
            max_tokens=max_tokens,
            temperature=1.0,
            thinking={"type": "enabled", "budget_tokens": thinking_budget},
            messages=[{"role": "user", "content": message_text}]
        )
    else:
        response = client.messages.create(
            model=model_id,
            max_tokens=max_tokens,
            temperature=temperature,
            messages=[{"role": "user", "content": message_text}]
        )
    
    chapter_text = ""
    thinking_text = ""
    for block in response.content:
        if block.type == "thinking":
            thinking_text = block.thinking
        elif block.type == "text":
            chapter_text += block.text
    
    return chapter_text, thinking_text


def make_docx(text):
    """Creates a Word document from text."""
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Georgia'
    style.font.size = Pt(12)
    for para_text in text.split("\n"):
        doc.add_paragraph(para_text)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ──────────────────────────────────────────────
# WRITE BUTTON
# ──────────────────────────────────────────────

if st.button("Write Chapter", type="primary"):
    if not api_key:
        st.error("Enter your API key in the sidebar.")
    elif not outline_file:
        st.error("Upload at least a chapter outline.")
    else:
        source_text = read_uploaded(source_file)
        char_text = read_uploaded(char_file)
        outline_text = read_uploaded(outline_file)
        
        parts = []
        if source_text:
            parts.append(f"<source_texts>\n{source_text}\n</source_texts>")
        if char_text:
            parts.append(f"<character_profiles>\n{char_text}\n</character_profiles>")
        parts.append(f"<chapter_outline>\n{outline_text}\n</chapter_outline>")
        parts.append(prompt)
        
        user_message = "\n\n".join(parts)
        
        char_count = len(user_message)
        token_est = char_count // 4
        st.info(f"Input: ~{char_count:,} characters (~{token_est:,} tokens)")
        
        with st.spinner("Writing chapter..."):
            try:
                client = anthropic.Anthropic(api_key=api_key)
                chapter_text, thinking_text = call_api(client, user_message)
                
                if not chapter_text.strip():
                    st.error("Model returned empty response.")
                else:
                    if thinking_text:
                        with st.expander("Model's thinking process"):
                            st.text(thinking_text)
                    
                    word_count = len(chapter_text.split())
                    st.success(f"Chapter complete — {word_count:,} words")
                    
                    # Store in session
                    st.session_state.chapter_text = chapter_text
                    st.session_state.revision_history = [{"pass": 0, "text": chapter_text, "label": "Original"}]
                    st.session_state.current_pass = 0
                    st.session_state.reports = []
                    
                    # Score it
                    score_result = score_chapter(chapter_text)
                    st.session_state.score_result = score_result
                    
                    # Generate report
                    report_buf = generate_report(score_result, chapter_text, "Original", 0, st.session_state.source_profile)
                    st.session_state.reports.append({"label": "Original", "buffer": report_buf})
                    
                    # Display
                    st.markdown("---")
                    display_scorecard(score_result, st.session_state.source_profile)
                    
                    st.markdown("---")
                    st.markdown("### Chapter Text")
                    st.text(chapter_text)
                    
                    # Downloads
                    dcol1, dcol2 = st.columns(2)
                    with dcol1:
                        buffer = make_docx(chapter_text)
                        st.download_button(
                            label="Download Original (.docx)",
                            data=buffer,
                            file_name="chapter_original.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    with dcol2:
                        st.download_button(
                            label="Download Original Report (.docx)",
                            data=st.session_state.reports[-1]["buffer"],
                            file_name="report_original.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="report_dl_orig"
                        )
                    
                    # Auto-revise if enabled and score is not low risk
                    if auto_revise and score_result["overall"] != "LOW RISK":
                        st.markdown("---")
                        st.markdown("### Auto-Revision")
                        
                        current_text = chapter_text
                        current_score = score_result
                        
                        for pass_num in range(1, max_revision_passes + 1):
                            if current_score["overall"] == "LOW RISK":
                                st.success(f"Reached LOW RISK after {pass_num - 1} revision(s). Stopping.")
                                break
                            
                            if not current_score["flagged"]:
                                st.info("No flagged passages to revise.")
                                break
                            
                            st.info(f"Revision pass {pass_num}/{max_revision_passes}...")
                            
                            revision_prompt = build_revision_prompt(current_text, current_score, st.session_state.source_profile)
                            
                            with st.spinner(f"Revision pass {pass_num}..."):
                                revised_text, rev_thinking = call_api(client, revision_prompt, is_revision=True)
                            
                            if not revised_text.strip():
                                st.warning(f"Pass {pass_num} returned empty. Stopping.")
                                break
                            
                            # Score the revision
                            new_score = score_chapter(revised_text)
                            
                            # Generate report
                            rev_label = f"Revision {pass_num}"
                            report_buf = generate_report(new_score, revised_text, rev_label, pass_num, st.session_state.source_profile)
                            st.session_state.reports.append({"label": rev_label, "buffer": report_buf})
                            
                            # Store history
                            st.session_state.revision_history.append({
                                "pass": pass_num,
                                "text": revised_text,
                                "label": rev_label
                            })
                            
                            # Display comparison
                            st.markdown(f"#### Pass {pass_num} Results")
                            prev_red = current_score["red_count"]
                            new_red = new_score["red_count"]
                            prev_flagged = len(current_score["flagged"])
                            new_flagged = len(new_score["flagged"])
                            
                            rcol1, rcol2, rcol3 = st.columns(3)
                            rcol1.metric("Red metrics", new_red, delta=f"{new_red - prev_red}", delta_color="inverse")
                            rcol2.metric("Flagged sentences", new_flagged, delta=f"{new_flagged - prev_flagged}", delta_color="inverse")
                            rcol3.metric("Overall", new_score["overall"])
                            
                            display_scorecard(new_score, st.session_state.source_profile)
                            
                            # Download buttons for this pass
                            pcol1, pcol2 = st.columns(2)
                            with pcol1:
                                rev_buf = make_docx(revised_text)
                                st.download_button(
                                    label=f"Download {rev_label} (.docx)",
                                    data=rev_buf,
                                    file_name=f"chapter_revision_{pass_num}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"ch_dl_{pass_num}"
                                )
                            with pcol2:
                                st.download_button(
                                    label=f"Download {rev_label} Report (.docx)",
                                    data=st.session_state.reports[-1]["buffer"],
                                    file_name=f"report_revision_{pass_num}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key=f"report_dl_{pass_num}"
                                )
                            
                            current_text = revised_text
                            current_score = new_score
                            st.session_state.chapter_text = current_text
                            st.session_state.score_result = current_score
                        
                        # Final download
                        st.markdown("---")
                        st.markdown("### Final Revised Chapter")
                        rev_word_count = len(current_text.split())
                        st.success(f"Final version — {rev_word_count:,} words after {len(st.session_state.revision_history) - 1} revision(s)")
                        st.text(current_text)
                        
                        buffer = make_docx(current_text)
                        st.download_button(
                            label="Download Final Revised (.docx)",
                            data=buffer,
                            file_name="chapter_final_revised.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
            
            except anthropic.APIError as e:
                st.error(f"API Error: {e}")
            except Exception as e:
                st.error(f"Error: {e}")


# ──────────────────────────────────────────────
# MANUAL SCORE & REVISE (for pasted text)
# ──────────────────────────────────────────────

st.markdown("---")
st.subheader("Score Existing Text")
st.caption("Paste any chapter text below to score it without writing a new one.")

pasted_text = st.text_area("Paste chapter text", height=200, key="paste_score")

if st.button("Score This Text"):
    if pasted_text.strip():
        score_result = score_chapter(pasted_text)
        display_scorecard(score_result, st.session_state.source_profile)
        st.session_state.chapter_text = pasted_text
        st.session_state.score_result = score_result
        
        # Generate and offer report download
        report_buf = generate_report(score_result, pasted_text, "Pasted Text", 0, st.session_state.source_profile)
        st.download_button(
            label="Download Score Report (.docx)",
            data=report_buf,
            file_name="report_pasted.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="report_pasted"
        )
    else:
        st.warning("Paste some text first.")

if st.button("Revise This Text"):
    if not api_key:
        st.error("Enter your API key in the sidebar.")
    elif not pasted_text.strip():
        st.warning("Paste some text first.")
    else:
        score_result = score_chapter(pasted_text)
        if not score_result["flagged"]:
            st.info("No flagged passages to revise.")
        else:
            display_scorecard(score_result, st.session_state.source_profile)
            revision_prompt = build_revision_prompt(pasted_text, score_result, st.session_state.source_profile)
            
            with st.spinner("Revising..."):
                try:
                    client = anthropic.Anthropic(api_key=api_key)
                    revised_text, _ = call_api(client, revision_prompt, is_revision=True)
                    
                    if revised_text.strip():
                        new_score = score_chapter(revised_text)
                        
                        st.markdown("### Revised Version")
                        display_scorecard(new_score, st.session_state.source_profile)
                        st.text(revised_text)
                        
                        rcol1, rcol2 = st.columns(2)
                        with rcol1:
                            buffer = make_docx(revised_text)
                            st.download_button(
                                label="Download Revised (.docx)",
                                data=buffer,
                                file_name="chapter_revised.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        with rcol2:
                            report_buf = generate_report(new_score, revised_text, "Revised (Pasted)", 1, st.session_state.source_profile)
                            st.download_button(
                                label="Download Revised Report (.docx)",
                                data=report_buf,
                                file_name="report_revised.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="report_revised_pasted"
                            )
                except Exception as e:
                    st.error(f"Error: {e}")


# ──────────────────────────────────────────────
# VERSION HISTORY
# ──────────────────────────────────────────────

if st.session_state.revision_history and len(st.session_state.revision_history) > 1:
    st.markdown("---")
    st.subheader("Version History")
    
    for entry in st.session_state.revision_history:
        with st.expander(f"{entry['label']} ({len(entry['text'].split()):,} words)"):
            st.text(entry["text"][:2000] + ("..." if len(entry["text"]) > 2000 else ""))
            dcol1, dcol2 = st.columns(2)
            with dcol1:
                buffer = make_docx(entry["text"])
                st.download_button(
                    label=f"Download {entry['label']}",
                    data=buffer,
                    file_name=f"chapter_{entry['label'].lower().replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{entry['pass']}"
                )
            with dcol2:
                # Find matching report
                matching_reports = [r for r in st.session_state.reports if r["label"] == entry["label"]]
                if matching_reports:
                    st.download_button(
                        label=f"Download {entry['label']} Report",
                        data=matching_reports[0]["buffer"],
                        file_name=f"report_{entry['label'].lower().replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"rdl_{entry['pass']}"
                    )
