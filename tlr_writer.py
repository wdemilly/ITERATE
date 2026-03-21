import streamlit as st
import anthropic
from docx import Document
from docx.shared import Pt
import io
import re
import math
import json

st.set_page_config(page_title="Fiction Writer + Detection Scorer", layout="wide")
st.title("Fiction Chapter Writer")

# ──────────────────────────────────────────────
# DETECTION SCORING ENGINE
# ──────────────────────────────────────────────

def score_chapter(text):
    """
    Scores a chapter against 10 detection metrics derived from
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
    # Handle dialogue quotes, abbreviations, etc.
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
    # Count words inside quotation marks (smart and straight)
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
    
    # ── Scoring: each metric gets GREEN / YELLOW / RED ──
    metrics = {}
    
    def rate(name, value, green_thresh, yellow_thresh, unit, invert=False):
        """Rate a metric. invert=True means higher is better (e.g., dialogue density)."""
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
    
    # ── Flag specific high-risk passages ──
    flagged = []
    
    for i, sent in enumerate(sentences):
        risks = []
        
        # Em dash with interpretation after it
        if '\u2014' in sent or '--' in sent:
            risks.append("em_dash")
        
        # "As though" / "as if"
        if re.search(r'\bas though\b|\bas if\b', sent, re.I):
            risks.append("as_though")
        
        # "The way" characterization
        if re.search(r'\bthe way (?:he|she|they|I|it|men|women|people|a man|a woman|soldiers|hungry)\b', sent, re.I):
            risks.append("the_way")
        
        # "Of a man/woman who"
        if re.search(r'\bof a (?:man|woman|person) who\b', sent, re.I):
            risks.append("of_person_who")
        
        # Metacognitive verbs
        if re.search(r'\b(?:I\s+)?(?:noted|filed|registered|understood|recognised|recognized)\b', sent, re.I):
            risks.append("metacognitive")
        
        # "The fact that"
        if re.search(r'\bthe fact that\b', sent, re.I):
            risks.append("fact_that")
        
        # Negation-as-framing (not X but Y or negation-leading)
        if re.search(r'^(?:It|That|This|I|She|He) (?:was|did|had|could) not\b', sent):
            risks.append("negation_leading")
        
        # Long sentence with multiple commas (potential inventory/run-on)
        comma_count = sent.count(',')
        sent_words = len(sent.split())
        if comma_count >= 4 and sent_words >= 35:
            risks.append("long_compound")
        
        # Observation-interpretation in same sentence (concrete detail + em dash + interpretation)
        if ('\u2014' in sent or '--' in sent) and sent_words > 25:
            if re.search(r'\bas though\b|\bthe way\b|\bwhich (?:was|meant|told)\b', sent, re.I):
                risks.append("obs_interp_coupling")
        
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
            "dialogue_word_pct": round(dialogue_pct, 1),
            "mean_sentence_length": round(mean_len, 1),
            "flagged_sentences": len(flagged),
            "total_sentences": len(sentences)
        }
    }


def build_revision_prompt(chapter_text, score_result):
    """
    Builds a targeted revision prompt that identifies specific flagged
    passages and tells the model how to fix them without destroying
    the writing.
    """
    flagged = score_result["flagged"]
    metrics = score_result["metrics"]
    
    # Only include top flagged passages (limit to 15 most severe)
    top_flagged = flagged[:15]
    
    # Build the risk descriptions
    risk_descriptions = {
        "em_dash": "Contains em dash — often used for interpretive gloss. Replace with a period and a new sentence, or remove the gloss entirely.",
        "as_though": "Contains 'as though' or 'as if' — appends speculative interpretation to concrete observation. Cut the tag. Let the action stand alone.",
        "the_way": "Contains 'the way he/she/they' — embeds interpretation inside description. Replace with direct physical observation.",
        "of_person_who": "Contains 'of a man/woman who' — analytical characterization. Replace with action or a short direct statement.",
        "metacognitive": "Contains metacognitive verb (noted/filed/registered/understood). Show through action, don't narrate the cognitive process.",
        "fact_that": "Contains 'the fact that' — analytical abstraction. Remove the frame; state the fact directly.",
        "negation_leading": "Negation-leading construction — defines by what something is not. Replace with a direct assertion of what it is.",
        "long_compound": "Long compound sentence with many commas — may read as inventory or process narration. Consider breaking into shorter sentences.",
        "obs_interp_coupling": "Observation-interpretation coupling — concrete detail and interpretive gloss fused in one sentence. Separate them or cut the interpretation."
    }
    
    passage_list = ""
    for i, item in enumerate(top_flagged):
        risks_text = "; ".join(risk_descriptions.get(r, r) for r in item["risks"])
        passage_list += f"\n\nFLAGGED PASSAGE {i+1}:\n\"{item['sentence'][:300]}\"\nISSUES: {risks_text}"
    
    # Build metric warnings
    metric_warnings = ""
    for name, data in metrics.items():
        if data["level"] == "RED":
            metric_warnings += f"\n- {name}: {data['value']} {data['unit']} — RED (high detection risk)"
        elif data["level"] == "YELLOW":
            metric_warnings += f"\n- {name}: {data['value']} {data['unit']} — YELLOW (moderate risk)"
    
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
- Break long compound sentences into shorter ones using periods.
- Do NOT introduce new em dashes, "as though" constructions, or metacognitive verbs.

CHAPTER METRICS (current scores):
{metric_warnings}

The following passages have been identified as highest detection risk:
{passage_list}

Here is the complete chapter. Revise ONLY the flagged passages. Output the full chapter text with your revisions in place.

<chapter>
{chapter_text}
</chapter>"""
    
    return prompt


def display_scorecard(score_result):
    """Renders the scorecard in Streamlit."""
    
    st.markdown("### Detection Risk Scorecard")
    
    overall = score_result["overall"]
    if overall == "HIGH RISK":
        st.error(f"Overall: {overall} — {score_result['red_count']} red, {score_result['yellow_count']} yellow, {score_result['green_count']} green")
    elif overall == "MODERATE RISK":
        st.warning(f"Overall: {overall} — {score_result['red_count']} red, {score_result['yellow_count']} yellow, {score_result['green_count']} green")
    else:
        st.success(f"Overall: {overall} — {score_result['red_count']} red, {score_result['yellow_count']} yellow, {score_result['green_count']} green")
    
    # Metrics table
    col1, col2, col3 = st.columns(3)
    metric_items = list(score_result["metrics"].items())
    
    for i, (name, data) in enumerate(metric_items):
        target_col = [col1, col2, col3][i % 3]
        icon = {"GREEN": "\u2705", "YELLOW": "\u26A0\uFE0F", "RED": "\u274C"}[data["level"]]
        target_col.metric(
            label=f"{icon} {name}",
            value=f"{data['value']} {data['unit']}",
            delta=data["level"],
            delta_color="normal" if data["level"] == "GREEN" else ("off" if data["level"] == "YELLOW" else "inverse")
        )
    
    # Summary stats
    s = score_result["summary"]
    st.caption(f"{s['total_sentences']} sentences | Mean length: {s['mean_sentence_length']} words | "
               f"{s['flagged_sentences']} flagged ({round(100*s['flagged_sentences']/max(s['total_sentences'],1))}%) | "
               f"Dialogue: {s['dialogue_word_pct']}% of words")
    
    # Flagged passages
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


def call_api(client, message_text, is_revision=False):
    """Makes an API call with the current settings."""
    
    # Determine model for revisions
    if is_revision and revision_model_choice != "Same as writing model":
        rev_model = model_map.get(revision_model_choice, model_id)
        # Revisions always use standard mode (no extended thinking)
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
                    
                    # Score it
                    score_result = score_chapter(chapter_text)
                    st.session_state.score_result = score_result
                    
                    # Display
                    st.markdown("---")
                    display_scorecard(score_result)
                    
                    st.markdown("---")
                    st.markdown("### Chapter Text")
                    st.text(chapter_text)
                    
                    # Download
                    buffer = make_docx(chapter_text)
                    st.download_button(
                        label="Download Original (.docx)",
                        data=buffer,
                        file_name="chapter_original.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
                            
                            revision_prompt = build_revision_prompt(current_text, current_score)
                            
                            with st.spinner(f"Revision pass {pass_num}..."):
                                revised_text, rev_thinking = call_api(client, revision_prompt, is_revision=True)
                            
                            if not revised_text.strip():
                                st.warning(f"Pass {pass_num} returned empty. Stopping.")
                                break
                            
                            # Score the revision
                            new_score = score_chapter(revised_text)
                            
                            # Store history
                            st.session_state.revision_history.append({
                                "pass": pass_num,
                                "text": revised_text,
                                "label": f"Revision {pass_num}"
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
                            
                            display_scorecard(new_score)
                            
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
                            label="Download Revised (.docx)",
                            data=buffer,
                            file_name="chapter_revised.docx",
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
        display_scorecard(score_result)
        st.session_state.chapter_text = pasted_text
        st.session_state.score_result = score_result
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
            display_scorecard(score_result)
            revision_prompt = build_revision_prompt(pasted_text, score_result)
            
            with st.spinner("Revising..."):
                try:
                    client = anthropic.Anthropic(api_key=api_key)
                    revised_text, _ = call_api(client, revision_prompt, is_revision=True)
                    
                    if revised_text.strip():
                        new_score = score_chapter(revised_text)
                        
                        st.markdown("### Revised Version")
                        display_scorecard(new_score)
                        st.text(revised_text)
                        
                        buffer = make_docx(revised_text)
                        st.download_button(
                            label="Download Revised (.docx)",
                            data=buffer,
                            file_name="chapter_revised.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
            buffer = make_docx(entry["text"])
            st.download_button(
                label=f"Download {entry['label']}",
                data=buffer,
                file_name=f"chapter_{entry['label'].lower().replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_{entry['pass']}"
            )
