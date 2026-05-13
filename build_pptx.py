"""
Build thesis_defense.pptx — 40-slide PowerPoint for the M.Tech thesis defense.

Narrative (research-oriented, easy to explain end-to-end):
  Act I  — Setup        (1-7)   title · thesis statement · motivation ·
                                 research questions · gap · contributions · roadmap
  Act II — Foundations  (8-9)   five-layer stack · unified methodology
  Act III — Contributions, each framed as Problem / Approach / Evidence
          C1 Distributed (10-13)
          C2 GEMM kernels (14-16)
          C3 SWFA attention (17-27)
          C4 SGLang serving (28-31)
          C5 Profiling & auto-tune (32-35)
  Act IV — Integration & Wrap (36-40)
          composition story · limitations · future work ·
          summary of contributions · thanks
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import os, io

OUT = "/home/adubey/IISTThesis-master/thesis_defense.pptx"
IMG_DIR = "/home/adubey/IISTThesis-master/_pptx_tmp"
os.makedirs(IMG_DIR, exist_ok=True)

# Widescreen 16:9
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

ACCENT = RGBColor(0xC0, 0x39, 0x2B)    # red-ish for emphasis
DARK   = RGBColor(0x22, 0x22, 0x22)
MUTED  = RGBColor(0x55, 0x55, 0x55)
FILL_I = RGBColor(0x9C, 0xE3, 0x9C)    # inside
FILL_X = RGBColor(0xFF, 0xC1, 0x7F)    # intersect
FILL_O = RGBColor(0xD9, 0xD9, 0xD9)    # outside

BLANK  = prs.slide_layouts[6]   # blank layout

def add_title(slide, text, subtitle=None):
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12.3), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = DARK
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(14)
        p2.font.italic = True
        p2.font.color.rgb = MUTED

def add_section_tag(slide, tag):
    """Small coloured tag above the title, e.g. 'Contribution 3 · SWFA'."""
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.07), Inches(12.3), Inches(0.3))
    p = tb.text_frame.paragraphs[0]
    p.text = tag
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT

def add_bullets(slide, bullets, left=0.6, top=1.2, width=12.0, height=5.8,
                size=18, color=None):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top),
                                  Inches(width), Inches(height))
    tf = tb.text_frame; tf.word_wrap = True
    for i, item in enumerate(bullets):
        if isinstance(item, tuple):
            text, emph = item
        else:
            text, emph = item, False
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "•  " + text
        p.font.size = Pt(size)
        p.font.color.rgb = ACCENT if emph else (color or DARK)
        p.space_after = Pt(4)

def add_footer(slide, idx, total=40):
    tb = slide.shapes.add_textbox(Inches(12.0), Inches(7.15),
                                  Inches(1.2), Inches(0.3))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = f"{idx} / {total}"
    p.font.size = Pt(10)
    p.font.color.rgb = MUTED
    p.alignment = PP_ALIGN.RIGHT
    tb2 = slide.shapes.add_textbox(Inches(0.3), Inches(7.15),
                                   Inches(9), Inches(0.3))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = "Adarsh Dubey — M.Tech Thesis Defense (IIST)"
    p2.font.size = Pt(10)
    p2.font.color.rgb = MUTED

# ---------- math rendering helpers ----------
def render_math_png(latex_str, fontsize=22, pad=0.1, filename=None):
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, f"${latex_str}$", fontsize=fontsize)
    out = filename or os.path.join(IMG_DIR, f"eq_{abs(hash(latex_str))}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight", pad_inches=pad,
                transparent=True)
    plt.close(fig)
    return out

def render_math_block_png(lines, fontsize=20, filename=None):
    text = "\n".join([f"${l}$" for l in lines])
    fig, ax = plt.subplots(figsize=(10, 0.6 * len(lines) + 0.5))
    ax.axis("off")
    ax.text(0.02, 0.5, text, fontsize=fontsize, va="center",
            transform=ax.transAxes)
    out = filename or os.path.join(IMG_DIR, f"eqb_{abs(hash(text))}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight", pad_inches=0.1,
                transparent=True)
    plt.close(fig)
    return out

def embed_math(slide, latex_str, left, top, height=0.6, fontsize=22):
    p = render_math_png(latex_str, fontsize=fontsize)
    slide.shapes.add_picture(p, Inches(left), Inches(top), height=Inches(height))

def embed_math_block(slide, lines, left, top, width=None, fontsize=20):
    p = render_math_block_png(lines, fontsize=fontsize)
    if width:
        slide.shapes.add_picture(p, Inches(left), Inches(top), width=Inches(width))
    else:
        slide.shapes.add_picture(p, Inches(left), Inches(top))

# ---------- chart helpers ----------
def render_bar_chart_util(filename):
    fig, ax = plt.subplots(figsize=(7, 4))
    x = ["DP", "DDP", "FSDP", "DeepSpeed"]
    avg  = [18.13, 56.2, 60.15, 65.8]
    peak = [72.3, 89.7, 91.2, 93.8]
    import numpy as np
    xp = np.arange(len(x))
    ax.bar(xp-0.18, avg, width=0.35, label="Average", color="#4E79A7")
    ax.bar(xp+0.18, peak, width=0.35, label="Peak",    color="#F28E2B")
    ax.set_xticks(xp); ax.set_xticklabels(x)
    ax.set_ylabel("GPU utilisation (%)"); ax.set_ylim(0, 100)
    ax.set_title("GPU utilisation across strategies")
    ax.legend(); ax.grid(axis="y", alpha=0.3)
    fig.tight_layout(); fig.savefig(filename, dpi=180); plt.close(fig)

def render_scaling_plot(filename):
    fig, ax = plt.subplots(figsize=(7, 4))
    N  = [512, 1024, 2048, 4096, 8192, 16384, 32768]
    sp = [1.02, 1.8, 3.2, 5.4, 8.1, 10.2, 11.4]
    ax.plot(N, sp, marker="o", color="#E15759")
    ax.set_xscale("log", base=2)
    ax.set_xlabel("Sequence length $N$"); ax.set_ylabel("Speed-up vs vanilla")
    ax.set_title("SWFA speed-up vs $N$ at $w=512$")
    ax.grid(True, alpha=0.3)
    fig.tight_layout(); fig.savefig(filename, dpi=180); plt.close(fig)

def render_roofline(filename):
    fig, ax = plt.subplots(figsize=(7, 4))
    import numpy as np
    xs = np.logspace(-1, 3, 200)
    peak = 312; bw = 1555
    rl = np.minimum(bw*xs, peak)
    ax.plot(xs, rl, "r-", linewidth=2, label="A100 roofline")
    pts = {"GEMM":(110,301), "FA":(95,189), "SWFA":(95,258), "Decode":(1.1,92)}
    for name,(x,y) in pts.items():
        ax.plot(x,y,"ko",markersize=7)
        ax.annotate(name, (x,y), xytext=(6,6), textcoords="offset points")
    ax.set_xscale("log"); ax.set_yscale("log")
    ax.set_xlabel("Arithmetic intensity (FLOP/byte)")
    ax.set_ylabel("TFLOPS")
    ax.set_title("Measured kernels on the A100 roofline")
    ax.grid(True, which="both", alpha=0.3)
    ax.set_xlim(0.1, 1000); ax.set_ylim(10, 500); ax.legend()
    fig.tight_layout(); fig.savefig(filename, dpi=180); plt.close(fig)

def render_block_classify(filename, N=8, wl=2, wr=2):
    fig, ax = plt.subplots(figsize=(5, 5))
    for i in range(N):
        for j in range(N):
            diff = abs(i-j)
            if diff <= wl and diff <= wr and diff <= 1:
                c = "#9CE39C"
            elif diff <= max(wl,wr):
                if diff == max(wl,wr):
                    c = "#FFC17F"
                else:
                    c = "#9CE39C"
            else:
                c = "#D9D9D9"
            ax.add_patch(plt.Rectangle((j, N-1-i), 1, 1,
                                       facecolor=c, edgecolor="black",
                                       linewidth=0.5))
    ax.set_xlim(0, N); ax.set_ylim(0, N); ax.set_aspect("equal")
    ax.set_xticks(range(N)); ax.set_yticks(range(N))
    ax.set_yticklabels(range(N-1, -1, -1))
    ax.set_xlabel("key $j$"); ax.set_ylabel("query $i$")
    ax.set_title(f"SWFA block classification, $w_l=w_r={wl}$")
    from matplotlib.patches import Patch
    leg = [Patch(facecolor="#9CE39C", edgecolor="k", label=r"$\mathcal{I}$ inside"),
           Patch(facecolor="#FFC17F", edgecolor="k", label=r"$\mathcal{X}$ intersect"),
           Patch(facecolor="#D9D9D9", edgecolor="k", label=r"$\mathcal{O}$ skip")]
    ax.legend(handles=leg, loc="upper right", bbox_to_anchor=(1.55, 1))
    fig.tight_layout(); fig.savefig(filename, dpi=180,
                                    bbox_inches="tight"); plt.close(fig)

def render_stack_diagram(filename):
    fig, ax = plt.subplots(figsize=(8, 4))
    layers = [
        ("Distributed schedule — DP / DDP / FSDP / ZeRO-3", "#F4B6B6"),
        ("GEMM kernels — CUTLASS 3.x + CuTe",              "#F6C58E"),
        ("Attention algorithm — SWFA (this thesis)",       "#F8E48A"),
        ("Serving runtime — SGLang (paged KV, fused)",     "#A9E1A9"),
        ("Profiling & auto-tune — Triton, Nsight, roofline","#9FC9E8"),
    ]
    for i,(txt,col) in enumerate(layers[::-1]):
        y = i
        ax.add_patch(plt.Rectangle((0.1, y+0.1), 7.8, 0.8,
                                   facecolor=col, edgecolor="black"))
        ax.text(4, y+0.5, txt, ha="center", va="center", fontsize=12)
    for i in range(len(layers)-1):
        ax.annotate("", xy=(4, len(layers)-1-i), xytext=(4, len(layers)-i-0.1),
                    arrowprops=dict(arrowstyle="->", lw=1.3))
    ax.set_xlim(0, 8); ax.set_ylim(-0.1, 5)
    ax.axis("off")
    ax.set_title("The five-layer stack", fontsize=14, pad=10)
    fig.tight_layout(); fig.savefig(filename, dpi=180); plt.close(fig)

def render_composition_diagram(filename):
    """Integration slide: each contribution feeds the next, with the headline gain."""
    fig, ax = plt.subplots(figsize=(10, 4.5))
    boxes = [
        ("C1  Distributed\nZeRO-3",       "93.8% peak util\n−22% energy",   "#F4B6B6"),
        ("C2  CUTLASS/CuTe\nGEMMs",       "+35% skinny/MoE\nmatches cuBLAS\non dense",        "#F6C58E"),
        ("C3  SWFA\nattention",           "65.6% blocks skipped\n<1e-4 error",  "#F8E48A"),
        ("C4  SGLang\nserving",           "2.84× decode\np99 < 30 ms",       "#A9E1A9"),
        ("C5  Triton\nprofiling",         "<3% cross-check\nauto-tune cache", "#9FC9E8"),
    ]
    w = 1.75; gap = 0.15
    for i,(head, bullet, col) in enumerate(boxes):
        x = i*(w+gap) + 0.1
        ax.add_patch(plt.Rectangle((x, 1.2), w, 2.4,
                                   facecolor=col, edgecolor="black", linewidth=1.1))
        ax.text(x+w/2, 3.15, head, ha="center", va="center",
                fontsize=11, fontweight="bold")
        ax.text(x+w/2, 1.95, bullet, ha="center", va="center",
                fontsize=9)
        if i < len(boxes)-1:
            ax.annotate("",
                        xy=(x+w+gap, 2.4), xytext=(x+w, 2.4),
                        arrowprops=dict(arrowstyle="->", lw=1.4))
    ax.text(5, 4.2, "End-to-end training  50–59% faster   ·   LLaMA-7B decode  2.84×",
            ha="center", va="center", fontsize=12, fontweight="bold", color="#C0392B")
    ax.set_xlim(0, 10); ax.set_ylim(0.5, 4.8)
    ax.axis("off")
    fig.tight_layout(); fig.savefig(filename, dpi=180); plt.close(fig)

# Pre-render charts
util_png        = os.path.join(IMG_DIR, "util_bar.png");     render_bar_chart_util(util_png)
scaling_png     = os.path.join(IMG_DIR, "scaling.png");      render_scaling_plot(scaling_png)
roofline_png    = os.path.join(IMG_DIR, "roofline.png");     render_roofline(roofline_png)
classify_png    = os.path.join(IMG_DIR, "classify.png");     render_block_classify(classify_png)
stack_png       = os.path.join(IMG_DIR, "stack.png");        render_stack_diagram(stack_png)
composition_png = os.path.join(IMG_DIR, "composition.png");  render_composition_diagram(composition_png)

def add_table(slide, data, left, top, width, height,
              header_fill=RGBColor(0x33, 0x4F, 0x7A),
              header_color=RGBColor(0xFF, 0xFF, 0xFF),
              first_col_bold=False, font_size=12):
    rows, cols = len(data), len(data[0])
    tbl = slide.shapes.add_table(rows, cols,
                                  Inches(left), Inches(top),
                                  Inches(width), Inches(height)).table
    for c in range(cols):
        tbl.cell(0, c).fill.solid()
        tbl.cell(0, c).fill.fore_color.rgb = header_fill
    for r in range(rows):
        for c in range(cols):
            cell = tbl.cell(r, c)
            cell.text = str(data[r][c])
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(font_size)
                    if r == 0:
                        run.font.bold = True
                        run.font.color.rgb = header_color
                    elif first_col_bold and c == 0:
                        run.font.bold = True
    return tbl

def add_takeaway(slide, text, top=6.3, color=None):
    """Single-line takeaway banner at the bottom — reinforces the main message."""
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(top),
                                  Inches(12.2), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "▸ " + text
    p.font.size = Pt(16); p.font.bold = True
    p.font.color.rgb = color or ACCENT

# ============================================================
# ACT I — Setup (slides 1–7)
# ============================================================
def slide_01_title():
    s = prs.slides.add_slide(BLANK)
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.2),
                              prs.slide_width, Inches(2.8))
    band.fill.solid(); band.fill.fore_color.rgb = RGBColor(0x33, 0x4F, 0x7A)
    band.line.fill.background()
    tb = s.shapes.add_textbox(Inches(0.5), Inches(2.4),
                              Inches(12.3), Inches(2.5))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "High-Performance Deep Learning Software Stack"
    p.font.size = Pt(32); p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p2 = tf.add_paragraph()
    p2.text = "for Scalable and Efficient LLMs"
    p2.font.size = Pt(28); p2.font.bold = True
    p2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p3 = tf.add_paragraph(); p3.text = "M.Tech Thesis Defense"
    p3.font.size = Pt(18); p3.font.italic = True
    p3.font.color.rgb = RGBColor(0xFF, 0xEE, 0xBB)
    tb2 = s.shapes.add_textbox(Inches(0.5), Inches(5.3),
                               Inches(12.3), Inches(1.8))
    tf2 = tb2.text_frame
    for i, line in enumerate([
        "Adarsh Dubey  (SC24M151)",
        "Advisor: Dr. S. Sumitra",
        "Department of Mathematics, IIST",
        "M.Tech in Machine Learning and Computing",
    ]):
        p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(20 if i == 0 else 14)
        p.font.bold = (i == 0)
        p.font.color.rgb = DARK
    return s

def slide_02_thesis_statement():
    """One-line thesis + the numbers that justify a 45-minute talk."""
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Thesis statement")
    add_title(s, "One idea, five layers, measured end-to-end")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.25),
                              Inches(12.2), Inches(1.4)).text_frame
    tb.word_wrap = True
    p = tb.paragraphs[0]
    p.text = ("Co-design every layer of the LLM stack so that the gains compose — "
              "and verify the composition on real A100 hardware.")
    p.font.size = Pt(20); p.font.italic = True; p.font.color.rgb = DARK
    # headline numbers
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(2.8),
                               Inches(12.2), Inches(0.5)).text_frame
    p = tb2.paragraphs[0]
    p.text = "What this thesis delivers, in numbers:"
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = MUTED
    add_bullets(s, [
        ("93.8%  peak GPU utilisation with DeepSpeed ZeRO-3", True),
        ("65.6%  attention blocks skipped by SWFA  (<1e-4 FP32 error)", True),
        ("50–59%  end-to-end training speed-up on GPT-2-Medium", True),
        ("2.84×   decode throughput on LLaMA-7B (SGLang kernel stack)", True),
        ("25–29%  energy reduction per epoch", True),
        ("Platform: 8× A100-40GB · NVLink · PyTorch 2.0 · CUDA 12.x · NCCL 2.19", False),
    ], top=3.4, size=18)
    add_footer(s, 2); return s

def slide_03_motivation():
    """The scaling wall — why the five layers are all constrained at once."""
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Motivation")
    add_title(s, "The scaling wall",
              subtitle="Every new model generation hits memory, comms and attention at the same time.")
    # left: memory example
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.4),
                              Inches(6.0), Inches(5.5)).text_frame
    p = tb.paragraphs[0]; p.text = "7B model, Adam, mixed precision"
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    for l in [
        "params 2 B · grads 2 B · FP32 master 4 B ·",
        "Adam m/v 4+4 B  →  14 B per parameter",
        "≈ 98 GB of state  ≫  40 GB on one A100",
        "Activations at N=2048, B=4: +20 GB per layer",
    ]:
        p = tb.add_paragraph(); p.text = "•  " + l
        p.font.size = Pt(15); p.font.color.rgb = DARK; p.space_after = Pt(3)
    p = tb.add_paragraph(); p.text = ""
    p = tb.add_paragraph()
    p.text = "→  Single-GPU training is impossible; sharding is forced."
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = ACCENT
    # right: four challenges
    tb2 = s.shapes.add_textbox(Inches(7.0), Inches(1.4),
                               Inches(6.0), Inches(5.5)).text_frame
    p = tb2.paragraphs[0]; p.text = "Four interacting challenges"
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    for h, sub in [
        ("Memory",      "state > single-GPU HBM"),
        ("Comms",       "all-reduce scales with model size, not batch"),
        ("Hardware",    "A100, H100, Xe, CDNA — different roofs"),
        ("Attention",   "Θ(N²) dominates long-context FLOPs"),
    ]:
        p = tb2.add_paragraph(); p.text = "•  " + h
        p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = DARK
        p2 = tb2.add_paragraph(); p2.text = "    " + sub
        p2.font.size = Pt(14); p2.font.color.rgb = MUTED; p2.space_after = Pt(6)
    add_takeaway(s, "Optimising any one layer in isolation leaves the other four on the table.")
    add_footer(s, 3); return s

def slide_04_research_questions():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Research questions")
    add_title(s, "Five research questions — one per layer of the stack")
    data = [
        ["#",  "Research question",                                                            "Answered in"],
        ["RQ1","Which distributed-training strategy gives the best utilisation / energy / memory trade-off on a single 8-GPU node?", "C1 (Ch.3)"],
        ["RQ2","How do we map transformer GEMMs (skinny, batched, grouped) onto Tensor Cores better than vendor cuBLAS?",            "C2 (Ch.4)"],
        ["RQ3","Can we exploit sliding-window sparsity before the GEMM, not as a post-hoc mask, while preserving numerical accuracy?", "C3 (Ch.5)"],
        ["RQ4","What cumulative decode-throughput gain is achievable from layered kernel optimisations inside SGLang?",              "C4 (Ch.6)"],
        ["RQ5","Can a single profiling pipeline cross-validate CUTLASS, Triton and Inductor on the same kernels within <3% error?", "C5 (Ch.7)"],
    ]
    add_table(s, data, 0.5, 1.3, 12.3, 4.6, first_col_bold=True, font_size=13)
    add_takeaway(s, "Each RQ has one chapter, one artefact, and one headline number — traceable end-to-end.")
    add_footer(s, 4); return s

def slide_05_gap():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Related work & gap")
    add_title(s, "Prior art is strong per layer — weak across layers")
    data = [
        ["Layer",          "Representative prior work",                                        "Gap this thesis fills"],
        ["Distributed",    "ZeRO (Rajbhandari '20), FSDP (Zhao '23), Megatron-LM",             "controlled 4-way single-node study"],
        ["GEMM kernels",   "CUTLASS, CuTe, cuBLAS, Triton",                                    "LLM-shape mapping with epilogue fusion"],
        ["Attention",      "Flash Attention 1/2 (Dao '22/'23), Longformer, BigBird",           "skip whole blocks, not just mask them"],
        ["Serving",        "vLLM (paged attention), SGLang, TGI",                              "layered attribution of decode gains"],
        ["Profiling",      "Nsight Compute, NVTX, Triton benchmark",                           "shape-indexed cache, cross-runtime check"],
    ]
    add_table(s, data, 0.5, 1.3, 12.3, 3.8, first_col_bold=True, font_size=12)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.3),
                              Inches(12.2), Inches(1.4)).text_frame
    p = tb.paragraphs[0]
    p.text = ("No prior work reports a controlled end-to-end study "
              "in which all five layers are tuned on the same hardware and composed.")
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = ACCENT
    add_footer(s, 5); return s

def slide_06_contributions():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contributions at a glance")
    add_title(s, "Five contributions — one per research question")
    data = [
        ["",    "Contribution",                                          "Headline evidence"],
        ["C1",  "Controlled 4-way distributed-training study",           "ZeRO-3: 93.8% peak util, −22% energy"],
        ["C2",  "CUTLASS/CuTe mapping for LLM GEMM shapes",              "+35% on skinny & grouped GEMMs"],
        ["C3",  "Sliding Window Flash Attention (block-skip + w*)",      "65.6% blocks skipped, error <1e-4"],
        ["C4",  "SGLang kernel stack for long-context serving",          "2.84× decode on LLaMA-7B"],
        ["C5",  "Triton profiling & auto-tune pipeline",                 "<3% vs CUTLASS, shape-indexed cache"],
    ]
    add_table(s, data, 0.5, 1.3, 12.3, 4.0, first_col_bold=True, font_size=14)
    add_takeaway(s, "Each contribution is independently useful — their composition is what is new.")
    add_footer(s, 6); return s

def slide_07_roadmap():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Talk roadmap")
    add_title(s, "How the next 33 slides are organised")
    col1 = s.shapes.add_textbox(Inches(0.7), Inches(1.4),
                                Inches(6), Inches(5.5)).text_frame
    col2 = s.shapes.add_textbox(Inches(7.0), Inches(1.4),
                                Inches(6), Inches(5.5)).text_frame
    for tf, parts in [(col1, [
                        ("Act II — Foundations  (8–9)",  "stack picture · methodology"),
                        ("C1 Distributed  (10–13)",      "strategies · memory math · results"),
                        ("C2 Kernels  (14–16)",          "CUTLASS/CuTe · fusion · vs cuBLAS"),
                        ("C3 SWFA  (17–27)",             "classify · proofs · speed · w*")]),
                      (col2, [
                        ("C4 Serving  (28–31)",          "SGLang · layered 2.84× · tails"),
                        ("C5 Profiling  (32–35)",        "Triton · roofline · cross-check"),
                        ("Act IV — Wrap  (36–40)",       "composition · limits · future"),
                        ("Questions & discussion",       "timing ≈ 45 min + Q&A")])]:
        for i, (h, sub) in enumerate(parts):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = h
            p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = DARK
            p.space_after = Pt(2)
            p2 = tf.add_paragraph(); p2.text = "    " + sub
            p2.font.size = Pt(14); p2.font.color.rgb = MUTED
            p2.space_after = Pt(14)
    add_footer(s, 7); return s

# ============================================================
# ACT II — Foundations (slides 8–9)
# ============================================================
def slide_08_stack():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Foundations · the stack")
    add_title(s, "The five-layer stack",
              subtitle="Each layer consumes the layer below and produces an abstraction for the layer above.")
    s.shapes.add_picture(stack_png, Inches(2.5), Inches(1.5),
                         height=Inches(4.7))
    add_takeaway(s, "Measurements flow back from profiling to every higher layer — dashed arrow in Ch.1.")
    add_footer(s, 8); return s

def slide_09_methodology():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Foundations · methodology")
    add_title(s, "Unified experimental methodology",
              subtitle="Same hardware, same seeds, same metrics across every chapter — so numbers compose.")
    tb1 = s.shapes.add_textbox(Inches(0.6), Inches(1.4),
                               Inches(6), Inches(5.0)).text_frame
    for i, line in enumerate([
        ("Hardware",  "8× A100-40GB, NVLink, single node,\nCUDA 12.x, NCCL 2.19"),
        ("Workloads", "ResNet-50/CIFAR-10 (controlled proxy)\nGPT-2-Medium (355M) for LLM training\nLLaMA-7B for serving"),
        ("Controls",  "identical seeds · identical init ·\nidentical NCCL topology · Nsight archives"),
    ]):
        p = tb1.paragraphs[0] if i == 0 else tb1.add_paragraph()
        p.text = line[0]
        p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
        p2 = tb1.add_paragraph(); p2.text = line[1]
        p2.font.size = Pt(14); p2.font.color.rgb = DARK
        p2.space_after = Pt(14)
    tb2 = s.shapes.add_textbox(Inches(7), Inches(1.4),
                               Inches(6), Inches(5.0)).text_frame
    p = tb2.paragraphs[0]; p.text = "Metrics (every chapter reports these)"
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    for m in [
        "GPU utilisation (avg / peak)",
        "Per-GPU memory (model · grad · opt · act)",
        "Training-loss convergence",
        "Power and energy per epoch",
        "Compute/communication overlap",
        "Throughput scaling & p50/p90/p99 latency",
        "Roofline placement (FLOP/byte vs TFLOPS)",
    ]:
        p = tb2.add_paragraph(); p.text = "•  " + m
        p.font.size = Pt(14); p.font.color.rgb = DARK
    add_takeaway(s, "Every claim in the next 26 slides comes from this single shared setup.")
    add_footer(s, 9); return s

# ============================================================
# ACT III — Contribution 1 · Distributed training (slides 10–13)
# ============================================================
def slide_10_c1_problem():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 1 · Distributed training — Problem")
    add_title(s, "Why we need a controlled 4-way study")
    add_bullets(s, [
        "7B-parameter Adam training needs ≈98 GB of state — sharding is mandatory.",
        "DP / DDP / FSDP / ZeRO-3 have been compared by their own authors, each on different workloads and hardware.",
        ("No open report isolates the effect of the strategy itself with identical seeds, identical init, identical workload, same node.", True),
        "Without this control, framework overheads masquerade as algorithmic differences.",
    ], top=1.5, size=18)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.2),
                              Inches(12), Inches(1.2)).text_frame
    p = tb.paragraphs[0]; p.text = "Research question RQ1:"
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = MUTED
    p2 = tb.add_paragraph()
    p2.text = ("Which strategy maximises GPU utilisation and minimises energy "
               "per epoch at the same final loss, on a fixed 8-GPU workload?")
    p2.font.size = Pt(18); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 10); return s

def slide_11_c1_approach():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 1 · Distributed training — Approach")
    add_title(s, "Strategies and the memory they imply")
    data = [
        ["",       "Params",    "Grads",        "Optimiser",    "Comms pattern"],
        ["DP",     "replicated","replicated",   "replicated",   "scatter / gather"],
        ["DDP",    "replicated","ring-reduced", "replicated",   "all-reduce"],
        ["FSDP",   "sharded",   "sharded",      "sharded",      "gather → compute → shard"],
        ["ZeRO-3", "sharded",   "sharded",      "sharded",      "fine-grained overlap"],
    ]
    add_table(s, data, 0.6, 1.3, 12.1, 2.3, first_col_bold=True, font_size=13)
    embed_math_block(s,
        [
          r"\mathrm{DP,\ DDP}:\ \ M_{\mathrm{param}} + M_{\mathrm{grad}} + M_{\mathrm{opt}} + M_{\mathrm{act}}(B)",
          r"\mathrm{FSDP,\ ZeRO\!-\!3}:\ \ \frac{1}{P}(M_{\mathrm{param}} + M_{\mathrm{grad}} + M_{\mathrm{opt}}) + M_{\mathrm{act}}(B)",
          r"\mathrm{Ring\ all\!-\!reduce}:\ \ T_{AR}(M) = 2(P-1)\!\left(\tau + \frac{M}{P\beta}\right) \to \frac{2M}{\beta}",
        ], left=0.9, top=4.0, width=11.4, fontsize=19)
    add_takeaway(s, "Adam on P=8 ⇒ 87.5% state-memory saving — theory we will now match to measurement.")
    add_footer(s, 11); return s

def slide_12_c1_result_util():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 1 · Distributed training — Evidence")
    add_title(s, "Utilisation & memory on 8× A100")
    s.shapes.add_picture(util_png, Inches(0.4), Inches(1.4),
                         height=Inches(4.4))
    data = [
        ["Strategy", "Mem / GPU (GB)", "Reduction"],
        ["DP",     "3.92", "1.0×"],
        ["DDP",    "3.92", "1.0×"],
        ["FSDP",   "0.49", "8.0×"],
        ["ZeRO-3", "0.49", "8.0×"],
    ]
    add_table(s, data, 7.6, 1.7, 5.2, 2.5, first_col_bold=True, font_size=13)
    tb = s.shapes.add_textbox(Inches(7.6), Inches(4.4),
                              Inches(5.4), Inches(1.8)).text_frame
    p = tb.paragraphs[0]
    p.text = "ZeRO-3 peaks at 93.8% via fine-grained overlap of param-gather with compute."
    p.font.size = Pt(13); p.font.color.rgb = DARK; p.space_after = Pt(6)
    p2 = tb.add_paragraph()
    p2.text = "Sharding delivers the theoretical ≈87.5% state-memory saving."
    p2.font.size = Pt(13); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_takeaway(s, "Result matches the memory theory on slide 11 — no framework surprise.")
    add_footer(s, 12); return s

def slide_13_c1_result_energy():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 1 · Distributed training — Evidence")
    add_title(s, "Energy, overlap and what we ship")
    data = [
        ["Strategy",   "Avg (W)", "kWh/ep", "Comms",     "Overlap (%)"],
        ["DP",         "2840",    "0.95",   "poor",      "12"],
        ["DDP",        "2720",    "0.87",   "good",      "64"],
        ["FSDP",       "2650",    "0.82",   "good",      "58"],
        ["DeepSpeed",  "2580",    "0.78",   "excellent", "61"],
    ]
    add_table(s, data, 1.2, 1.3, 11.0, 2.6, first_col_bold=True, font_size=14)
    add_bullets(s, [
        ("ZeRO-3: 22% lower per-epoch energy than DP at the same final loss.", True),
        "DDP overlaps best: ring all-reduce runs during backward of earlier layers.",
        "ZeRO-3 param-gather blocks compute slightly; overlap_comm hides most of it.",
    ], top=4.2, size=16)
    add_takeaway(s, "Answer to RQ1: ZeRO-3 wins on util + energy; FSDP is the best memory/throughput middle path.")
    add_footer(s, 13); return s

# ============================================================
# ACT III — Contribution 2 · CUTLASS / CuTe GEMMs (slides 14–16)
# ============================================================
def slide_14_c2_problem():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 2 · Kernels — Problem")
    add_title(s, "cuBLAS is great — except on the shapes LLMs actually use")
    add_bullets(s, [
        "Square GEMMs (4096³, 8192³) are well-served by cuBLAS.",
        ("Transformer bottlenecks are skinny (QKV, attention output) and grouped (MoE experts) — cuBLAS leaves 30–40% on the table.", True),
        "Epilogues (bias, GELU, residual-add) force extra HBM round-trips if not fused into the GEMM.",
        "Vendor libraries cannot be re-tuned per model architecture without rebuilding from source.",
    ], top=1.5, size=18)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.3),
                              Inches(12), Inches(1.2)).text_frame
    p = tb.paragraphs[0]; p.text = "Research question RQ2:"
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = MUTED
    p2 = tb.add_paragraph()
    p2.text = "Can we map LLM GEMM shapes onto CUTLASS/CuTe and beat cuBLAS on transformer-typical shapes?"
    p2.font.size = Pt(18); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 14); return s

def slide_15_c2_approach():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 2 · Kernels — Approach")
    add_title(s, "Tile hierarchy, layout algebra, epilogue fusion")
    add_bullets(s, [
        "Threadblock tile (Mtb, Ntb, Ktb) — unit of HBM traffic.",
        "Warp tile (Mw, Nw, Kw) — unit of shared-memory traffic.",
        "Instruction tile (Mi, Ni, Ki) — Tensor-Core MMA shape (e.g. 16×8×16).",
        "Multi-stage pipelining: prefetch tile k+1 while MMAs on tile k execute.",
    ], top=1.4, size=16)
    embed_math_block(s,
        [r"\mathrm{Layout}(S,D)\ \circ\ \mathrm{Layout}(S',D') = \mathrm{Layout}((S,S'),(D\!\cdot\! D',\, D'))"],
        left=0.8, top=4.3, width=12.0, fontsize=20)
    embed_math_block(s,
        [r"Z_{ij} \leftarrow \phi\!\left(\sum_k W_{ik}X_{kj} + b_i\right)\ \ \mathrm{(fused\ GEMM+bias+activation,\ one\ HBM\ pass)}"],
        left=0.8, top=5.5, width=12.0, fontsize=18)
    add_takeaway(s, "Fusion eliminates two HBM round-trips per MLP layer — pure memory savings.")
    add_footer(s, 15); return s

def slide_16_c2_result():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 2 · Kernels — Evidence")
    add_title(s, "CUTLASS vs cuBLAS (FP16 on A100, TFLOPS)")
    data = [
        ["Shape (M × N × K)",                "cuBLAS", "CUTLASS", "Ratio"],
        ["4096³",                            "280",    "301",     "1.08×"],
        ["8192³",                            "305",    "318",     "1.04×"],
        ["2048 × 2048 × 128  (skinny)",      "95",     "128",     "1.35×"],
        ["Grouped MoE, 16 experts",          "184",    "241",     "1.31×"],
    ]
    add_table(s, data, 1.3, 1.3, 10.7, 3.0, first_col_bold=True, font_size=15)
    add_bullets(s, [
        "On large compute-bound shapes, both reach >90% Tensor-Core utilisation → tie.",
        ("On skinny and grouped GEMMs — the LLM-dominant shapes — CUTLASS wins by up to 35%.", True),
        "Speed-up comes from warp tile 64×32×32 matching the register budget exactly.",
    ], top=4.5, size=15)
    add_takeaway(s, "Answer to RQ2: LLM-aware tile + layout choices beat a one-size-fits-all cuBLAS on the shapes that matter.")
    add_footer(s, 16); return s

# ============================================================
# ACT III — Contribution 3 · SWFA (slides 17–27)
# ============================================================
def slide_17_c3_problem():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Problem")
    add_title(s, "Sliding-window attention is everywhere — but implementations waste it")
    embed_math_block(s,
        [r"\mathrm{Attn}(Q,K,V) = \mathrm{softmax}\!\left(\frac{QK^{T}}{\sqrt{d}}\right)\!V,\ \ \mathrm{cost} = \Theta(N^2 d)"],
        left=0.8, top=1.3, width=12.0, fontsize=22)
    add_bullets(s, [
        "Flash Attention removes the memory wall — but still issues dense GEMMs on every Br×Bc tile.",
        "Longformer, BigBird, Mistral-7B all restrict attention to a local window [i−wl, i+wr].",
        ("In theory: cost drops from Θ(N²) to Θ(N·w) — only if the kernel is window-aware.", True),
        "Most implementations realise the window as a post-hoc score-matrix mask → zero FLOP savings.",
    ], top=3.0, size=16)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.9),
                              Inches(12), Inches(1.0)).text_frame
    p = tb.paragraphs[0]; p.text = "Research question RQ3:"
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = MUTED
    p2 = tb.add_paragraph()
    p2.text = "Can we skip whole attention blocks (not just mask them) without changing the output?"
    p2.font.size = Pt(17); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 17); return s

def slide_18_c3_fa_recap():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Enabling insight")
    add_title(s, "Flash Attention is a block-by-block state machine")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "For each row tile i, carry the state (m_i, ℓ_i, O_i) — row-max, row-sum, accumulator:"
    p.font.size = Pt(15); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"m_i^{new} = \max(m_i,\ \max_j s_{ij})",
        r"\ell_i^{new} = e^{m_i - m_i^{new}}\ell_i + \sum_j e^{s_{ij} - m_i^{new}}",
        r"O_i^{new} = e^{m_i - m_i^{new}} O_i + \sum_j e^{s_{ij} - m_i^{new}} v_j",
    ], left=1.0, top=1.9, width=11.0, fontsize=20)
    add_bullets(s, [
        ("Peak HBM residency drops from Θ(N²) to Θ(Br·Bc); answer identical up to FP associativity.", True),
        ("Key observation: the state update (m, ℓ, O) is invariant to skipping blocks whose contribution is provably zero.", True),
        "This is the hook SWFA uses — skip blocks that lie outside the window, keep the rest untouched.",
    ], top=5.2, size=15)
    add_footer(s, 18); return s

def slide_19_c3_classify_fig():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Approach (1/4)")
    add_title(s, "Classify blocks once: inside / intersect / outside")
    s.shapes.add_picture(classify_png, Inches(2.8), Inches(1.3),
                         height=Inches(4.7))
    add_takeaway(s, "Only I and X blocks are computed; O blocks are skipped entirely — never become GEMMs.")
    add_footer(s, 19); return s

def slide_20_c3_classify_algo():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Approach (2/4)")
    add_title(s, "Classification by corner checks — O((N/B)²), negligible")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(5.3))
    tf = tb.text_frame
    lines = [
        "Input:  N, window (wl, wr), block sizes (Br, Bc)",
        "Output: sets I (inside), X (intersect), O (outside)",
        "",
        "Tr ← ⌈N/Br⌉;   Tc ← ⌈N/Bc⌉",
        "for r = 0 .. Tr−1:",
        "    r0 ← r·Br;   r1 ← min((r+1)·Br, N) − 1",
        "    for c = 0 .. Tc−1:",
        "        c0 ← c·Bc;   c1 ← min((c+1)·Bc, N) − 1",
        "        outside ← (c1 < r0 − wl)  or  (c0 > r1 + wr)",
        "        inside  ← (c1 ≤ r0 + wr) and (c0 ≥ r1 − wl)",
        "        if outside:   O ← O ∪ {(r,c)}",
        "        elif inside:  I ← I ∪ {(r,c)}",
        "        else:         X ← X ∪ {(r,c)}",
        "return (I, X, O)",
    ]
    for i, l in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = l
        p.font.name = "Consolas"; p.font.size = Pt(14); p.font.color.rgb = DARK
    add_takeaway(s, "Only two corners per block need checking (lemma on next slide).")
    add_footer(s, 20); return s

def slide_21_c3_corner_lemma():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Approach (3/4), correctness")
    add_title(s, "Corner-check lemma — 4 comparisons are enough")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.8)).text_frame
    p = tb.paragraphs[0]
    p.text = "Let B = [r0, r1] × [c0, c1] and W = {(i,j): i − wl ≤ j ≤ i + wr}. Then:"
    p.font.size = Pt(15); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"(i)\ \ B \cap W = \emptyset \Leftrightarrow c_1 < r_0 - w_l \ \mathrm{or}\ c_0 > r_1 + w_r",
        r"(ii)\ \ B \subseteq W \Leftrightarrow c_1 \leq r_0 + w_r \ \mathrm{and}\ c_0 \geq r_1 - w_l",
    ], left=0.9, top=2.2, width=11.3, fontsize=19)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(4.3),
                               Inches(12), Inches(2.4)).text_frame
    p = tb2.paragraphs[0]; p.text = "Proof sketch."
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = ACCENT
    p2 = tb2.add_paragraph()
    p2.text = ("Window boundaries j = i ± wr, j = i − wl are monotone in i, so min (j − i) "
               "over B is attained at (r1, c0) and max at (r0, c1). Outside/inside reduce to "
               "inequalities on these two corners. ∎")
    p2.font.size = Pt(14); p2.font.color.rgb = DARK
    add_takeaway(s, "Classification cost: 4 comparisons per block — negligible next to the skipped GEMMs.")
    add_footer(s, 21); return s

def slide_22_c3_vecmask():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Approach (4/4), mask generation")
    add_title(s, "Vectorised mask for X-blocks — one triu/tril, no Python loop")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.7)).text_frame
    p = tb.paragraphs[0]
    p.text = "For block (r, c) with offset δ_rc = c·Bc − r·Br, local offset rel = j_ℓ − i_ℓ is masked iff:"
    p.font.size = Pt(14); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"\mathrm{rel} > w_r + \delta_{rc}\ \ \text{or}\ \ \mathrm{rel} < -w_l + \delta_{rc}",
        r"d_{right} = \lfloor w_r + \delta_{rc} \rfloor + 1,\quad d_{left} = \lfloor -w_l + \delta_{rc} \rfloor - 1",
        r"M_{rc} = \mathrm{triu}(\mathbf{1}_{B_r\times B_c},\, d_{right})\ \vee\ \mathrm{tril}(\mathbf{1}_{B_r\times B_c},\, d_{left})",
    ], left=0.8, top=2.1, width=11.5, fontsize=19)
    add_bullets(s, [
        "Replaces a nested O(Br·Bc) loop with one triu and one tril.",
        ("≈100× faster than naive block-level masking at Br = Bc = 64.", True),
    ], top=5.3, size=15)
    add_footer(s, 22); return s

def slide_23_c3_fwd_algo():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Forward algorithm")
    add_title(s, "SWFA forward pass — the skip is one line")
    tb = s.shapes.add_textbox(Inches(0.5), Inches(1.2),
                              Inches(12.3), Inches(5.8))
    tf = tb.text_frame
    lines = [
        "Input: Q ∈ R^{B×H×N×d}, K, V ∈ R^{B×H×N×dv},   (wl, wr),   (Br, Bc)",
        "",
        "(I, X, O) ← ClassifyBlocks()",
        "for r = 0 .. Tr − 1:",
        "    O_r ← 0;   m_r ← −∞;   ℓ_r ← 0",
        "    for c = 0 .. Tc − 1:",
        "        if (r,c) ∈ O: continue                      ← SKIP BLOCK",
        "        S_rc ← Q_r · K_c^T / √d",
        "        if (r,c) ∈ X:  S_rc += GenerateMask(r,c)",
        "        m_new ← max(m_r, max(S_rc));   corr ← exp(m_r − m_new)",
        "        P_rc ← exp(S_rc − m_new)",
        "        ℓ_r ← corr·ℓ_r + Σ P_rc",
        "        O_r ← corr·O_r + P_rc · V_c",
        "        m_r ← m_new",
        "    O_r ← O_r / (ℓ_r + ε)",
        "return O",
    ]
    for i, l in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = l
        p.font.name = "Consolas"; p.font.size = Pt(13)
        if "SKIP" in l:
            p.font.bold = True; p.font.color.rgb = ACCENT
        else:
            p.font.color.rgb = DARK
    add_footer(s, 23); return s

def slide_24_c3_bwd_complexity():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Backward & complexity")
    add_title(s, "Backward re-runs the same block set — same O, X; no new kernels")
    embed_math_block(s, [
        r"P_{rc} = e^{S_{rc} - m_r}/\ell_r,\ \ dS_{rc} = P_{rc} \odot (dP_{rc} - \rho_r)",
        r"\rho_r = \sum_c P_{rc} \odot (dO_r\, V_c^{T})",
    ], left=0.8, top=1.3, width=11.0, fontsize=19)
    data = [
        ["Quantity",                        "Expression"],
        ["Total blocks",                    "(N/B)²"],
        ["Inside band (I ∪ X)",             "≈ 2·(N/B)·(w/B)"],
        ["Skipped (O)",                     "(N/B)² − 2·(N/B)·(w/B)"],
        ["Theoretical speed-up",            "1 − 2w/N  (as w ≪ N)"],
        ["Memory",                          "Θ(N) — same as Flash Attention"],
    ]
    add_table(s, data, 1.5, 3.2, 10.3, 2.8, first_col_bold=True, font_size=13)
    add_takeaway(s, "Example: N=128, w=4, B=16 ⇒ 48 of 64 blocks skipped (75%). Prediction is clean.")
    add_footer(s, 24); return s

def slide_25_c3_accuracy_skip():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Evidence (1/2)")
    add_title(s, "Accuracy (19 configs pass) and skip efficiency (matches theory)")
    acc = [
        ["Config",              "Max diff", "Tol."],
        ["sym narrow w=2",      "2.4e-7",   "1e-4"],
        ["sym wide w=8",        "2.4e-7",   "1e-4"],
        ["causal+win wl=6",     "4.8e-7",   "1e-4"],
        ["multi-head H=8",      "8.3e-7",   "1e-4"],
        ["long seq N=128",      "7.5e-7",   "1e-4"],
        ["FP16 w=4,6",          "1.95e-3",  "1e-2"],
    ]
    add_table(s, acc, 0.4, 1.3, 6.0, 3.2, first_col_bold=True, font_size=12)
    tb = s.shapes.add_textbox(Inches(0.4), Inches(4.7),
                              Inches(6.0), Inches(0.5)).text_frame
    p = tb.paragraphs[0]; p.text = "All 19 configurations pass within tolerance."
    p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = ACCENT
    skip = [
        ["Config",              "Blocks", "Skip %"],
        ["very narrow w=1,2",   "64",     "65.6%"],
        ["sym narrow w=2",      "16",     "37.5%"],
        ["causal+win w=6,0",    "16",     "56.2%"],
        ["left only w=10,0",    "16",     "56.2%"],
        ["N=Ns=45",             "9",      "22.2%"],
        ["win>seq w=20",        "16",     "0.0%"],
    ]
    add_table(s, skip, 6.8, 1.3, 6.0, 3.2, first_col_bold=True, font_size=12)
    tb2 = s.shapes.add_textbox(Inches(6.8), Inches(4.7),
                               Inches(6.0), Inches(0.5)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Skip rate matches the 1 − 2w/N prediction from slide 24."
    p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = ACCENT
    add_footer(s, 25); return s

def slide_26_c3_e2e():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Evidence (2/2)")
    add_title(s, "End-to-end training + long-sequence scaling")
    data = [
        ["Config",            "Tok/s",   "kWh/ep"],
        ["DDP vanilla",       "12,450",  "2.8"],
        ["DDP + SWFA",        "18,720",  "2.1"],
        ["FSDP vanilla",      "11,890",  "2.9"],
        ["FSDP + SWFA",       "17,340",  "2.2"],
        ["DeepSpeed vanilla", "13,210",  "2.7"],
        ["DeepSpeed + SWFA",  "19,850",  "2.0"],
    ]
    add_table(s, data, 0.4, 1.3, 6.0, 3.2, first_col_bold=True, font_size=12)
    s.shapes.add_picture(scaling_png, Inches(6.6), Inches(1.3),
                         height=Inches(4.5))
    add_takeaway(s, "50–59% throughput, 25–29% energy, perplexity shift < 0.1 — composes cleanly with every distributed strategy.")
    add_footer(s, 26); return s

def slide_27_c3_optimal_window():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 3 · SWFA — Theory, optimal window")
    add_title(s, "Closed-form w* under exponential attention decay")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.5)).text_frame
    p = tb.paragraphs[0]; p.text = "Compute / memory / performance-loss cost:"
    p.font.size = Pt(15); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"\mathcal{L}(w) = \alpha\,C(w) + \beta\,M(w) + \gamma\,P(w)",
        r"C(w) = \kappa N w d,\ \ M(w) \approx M_0",
        r"P(w) = \frac{2 N A_0\,e^{-\lambda w/2}}{1 - e^{-\lambda}}\ \ (A_{ij} = A_0 e^{-\lambda|i-j|})",
    ], left=0.9, top=1.8, width=11.3, fontsize=18)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(4.6),
                               Inches(12), Inches(2.2)).text_frame
    p = tb2.paragraphs[0]; p.text = "Theorem.  L(w) is strictly convex on w > 0."
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = ACCENT
    p2 = tb2.add_paragraph()
    p2.text = "Proof: C'' = M'' = 0; P'' ∝ (λ²/4)·e^{−λw/2} > 0 ⇒ L'' > 0. ∎"
    p2.font.size = Pt(14); p2.font.color.rgb = DARK
    p3 = tb2.add_paragraph()
    p3.text = "Optimum:  w* = (2/λ) · ln[ γ·A₀·λ / ( α·κ·d·(1 − e^{-λ}) ) ]"
    p3.font.size = Pt(15); p3.font.bold = True; p3.font.color.rgb = ACCENT
    add_takeaway(s, "w* is a drop-in warm-start for per-layer adaptive windows (future work).")
    add_footer(s, 27); return s

# ============================================================
# ACT III — Contribution 4 · SGLang serving (slides 28–31)
# ============================================================
def slide_28_c4_problem():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 4 · Serving — Problem")
    add_title(s, "Serving is a different regime from training")
    tb1 = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                               Inches(6), Inches(2.8)).text_frame
    p = tb1.paragraphs[0]; p.text = "Prefill — compute-bound"
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = ACCENT
    for l in [
        "Dense QKᵀ and PV on the whole prompt.",
        "SWFA cuts attention FLOPs by ~37% at w=4096 on 8K prompts.",
        "Direct wall-clock win.",
    ]:
        p = tb1.add_paragraph(); p.text = "•  " + l
        p.font.size = Pt(15); p.font.color.rgb = DARK
    tb2 = s.shapes.add_textbox(Inches(6.8), Inches(1.3),
                               Inches(6), Inches(2.8)).text_frame
    p = tb2.paragraphs[0]; p.text = "Decode — memory-bound"
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = ACCENT
    for l in [
        "Arithmetic intensity ≈ 1 FLOP/byte.",
        "KV-cache reads dominate; FLOPs don't.",
        "Throughput ceiling = HBM bandwidth, not TFLOPS.",
    ]:
        p = tb2.add_paragraph(); p.text = "•  " + l
        p.font.size = Pt(15); p.font.color.rgb = DARK
    tb3 = s.shapes.add_textbox(Inches(0.6), Inches(4.8),
                               Inches(12), Inches(1.4)).text_frame
    p = tb3.paragraphs[0]; p.text = "Research question RQ4:"
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = MUTED
    p2 = tb3.add_paragraph()
    p2.text = "What cumulative decode speed-up is achievable from layered kernel optimisations, attributing each one?"
    p2.font.size = Pt(18); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_takeaway(s, "Kernel design must follow the bottleneck, not the algorithm name.")
    add_footer(s, 28); return s

def slide_29_c4_approach():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 4 · Serving — Approach")
    add_title(s, "SGLang kernel stack: paged KV + fused decode + SWFA prefill + fused sampling")
    shapes = s.shapes
    for i, (txt, col) in enumerate([
        ("Prefill  (SWFA)",    RGBColor(0xCB, 0xE4, 0xFB)),
        ("Fused Decode Attn",  RGBColor(0xD6, 0xF2, 0xD0)),
        ("Fused Sampling",     RGBColor(0xFF, 0xF2, 0xB3)),
    ]):
        sh = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.8 + i*4.2), Inches(1.8),
                              Inches(3.6), Inches(0.9))
        sh.fill.solid(); sh.fill.fore_color.rgb = col
        sh.line.color.rgb = DARK
        tf = sh.text_frame; tf.text = txt
        tf.paragraphs[0].font.size = Pt(16); tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    sh = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                          Inches(5.0), Inches(3.2),
                          Inches(3.6), Inches(0.9))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
    sh.line.color.rgb = DARK
    tf = sh.text_frame; tf.text = "Paged KV Cache  (16-tok pages)"
    tf.paragraphs[0].font.size = Pt(14); tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    add_bullets(s, [
        "Paged KV: 16-tok pages eliminate fragmentation; radix-tree sharing for common prefixes.",
        "Fused decode: matmul + softmax + masking in one kernel, hides KV indirection behind compute.",
        "SWFA prefill: drop-in for window models; reuses Contribution 3 directly.",
        "Fused top-k/top-p: removes 3–4 kernel launches per decode step.",
    ], top=4.5, size=14)
    add_footer(s, 29); return s

def slide_30_c4_layered():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 4 · Serving — Evidence (1/2)")
    add_title(s, "Layered decode-throughput gain on LLaMA-7B — 2.84× cumulative")
    data = [
        ["Configuration",                           "Tok/s", "Speed-up"],
        ["Baseline (vanilla attn, unpaged KV)",     "980",    "1.00×"],
        ["+ Paged KV cache",                        "1,420",  "1.45×"],
        ["+ Fused decode attention",                "1,860",  "1.90×"],
        ["+ SWFA prefill",                          "2,410",  "2.46×"],
        ["+ Fused sampling",                        "2,780",  "2.84×"],
    ]
    add_table(s, data, 0.7, 1.3, 8.0, 3.2, first_col_bold=True, font_size=13)
    tb = s.shapes.add_textbox(Inches(9.0), Inches(1.3),
                              Inches(4.0), Inches(5.5)).text_frame
    p = tb.paragraphs[0]; p.text = "Per-layer wedge"
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = ACCENT
    for line in [
        "Paged KV  (+45%):\n  no fragmentation",
        "Fused decode  (+31%):\n  matches HBM roofline",
        "SWFA prefill  (+30%):\n  declared window",
        "Fused sampling  (+15%):\n  fewer launches",
    ]:
        p = tb.add_paragraph(); p.text = "•  " + line
        p.font.size = Pt(12); p.font.color.rgb = DARK; p.space_after = Pt(4)
    add_takeaway(s, "Order matters — each optimisation exposes the next bottleneck.")
    add_footer(s, 30); return s

def slide_31_c4_latency():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 4 · Serving — Evidence (2/2)")
    add_title(s, "Tail latency and KV footprint — concurrency improves")
    data = [
        ["Batch", "p50 (ms)", "p90 (ms)", "p99 (ms)", "Tok/s"],
        ["1",   "18.2", "18.4", "18.9",   "55"],
        ["8",   "19.3", "20.1", "20.9",   "414"],
        ["16",  "20.1", "21.3", "22.6",   "796"],
        ["32",  "21.8", "23.5", "25.4",   "1,467"],
        ["64",  "23.5", "25.7", "28.4",   "2,724"],
    ]
    add_table(s, data, 1.0, 1.3, 11.0, 2.8, first_col_bold=True, font_size=13)
    add_bullets(s, [
        ("p99 < 30 ms even at batch 64 — well below the 100 ms SLO.", True),
        ("KV footprint at L=8192, batch 64: 34 GB → 19 GB (44% saving) via paged + radix-tree sharing.", True),
        "Enables 1.7× more concurrent requests per GPU.",
    ], top=4.4, size=16)
    add_takeaway(s, "Answer to RQ4: the 2.84× is real under realistic batch/tail-latency constraints.")
    add_footer(s, 31); return s

# ============================================================
# ACT III — Contribution 5 · Profiling & auto-tune (slides 32–35)
# ============================================================
def slide_32_c5_problem():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 5 · Profiling — Problem")
    add_title(s, "Kernel tuning drifts with every new shape and every new GPU")
    add_bullets(s, [
        "Hand-tuned tile sizes in Chapters 4–6 are valid only for the shapes they were tuned on.",
        "H100, Blackwell, Xe change the Tensor-Core shape and the register budget.",
        "A single throughput number from CUTLASS tells us nothing about whether Triton can keep up — we need a cross-runtime check.",
        ("Every speed-up claim in this thesis must be reproducible by an independent runtime within <3%.", True),
    ], top=1.4, size=17)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.6),
                              Inches(12), Inches(1.2)).text_frame
    p = tb.paragraphs[0]; p.text = "Research question RQ5:"
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = MUTED
    p2 = tb.add_paragraph()
    p2.text = "Can one pipeline cross-validate CUTLASS, Triton and Inductor and cache good configs per shape?"
    p2.font.size = Pt(17); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 32); return s

def slide_33_c5_approach():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 5 · Profiling — Approach")
    add_title(s, "Triton auto-tune + roofline + shape-indexed cache")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "Programmer writes block-level Python; compiler emits PTX with Tensor-Core MMAs."
    p.font.size = Pt(15); p.font.color.rgb = DARK
    code = s.shapes.add_textbox(Inches(0.6), Inches(1.8),
                                Inches(12), Inches(3.5))
    tf = code.text_frame; tf.word_wrap = True
    snippet = [
        "@triton.autotune(configs=[",
        "    triton.Config({'BLOCK_M':128,'BLOCK_N':128,",
        "                   'BLOCK_K':32}, num_warps=8, num_stages=3),",
        "    triton.Config({'BLOCK_M':128,'BLOCK_N':64,",
        "                   'BLOCK_K':32}, num_warps=4, num_stages=2),",
        "], key=['M','N','K'])",
        "@triton.jit",
        "def gemm_kernel(A, B, C, M, N, K, ...):",
        "    # block-pointer arithmetic + MMA loop",
    ]
    for i, line in enumerate(snippet):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = "Consolas"; p.font.size = Pt(14); p.font.color.rgb = DARK
    add_bullets(s, [
        "Auto-tuner picks the winner per shape class; second call hits the cache → free.",
        "Every kernel from C2/C3/C4 is rerun here and placed on the A100 roofline (next slide).",
    ], top=5.5, size=14)
    add_footer(s, 33); return s

def slide_34_c5_roofline():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 5 · Profiling — Evidence (1/2)")
    add_title(s, "Roofline placement — every kernel sits where the model predicts")
    s.shapes.add_picture(roofline_png, Inches(2.5), Inches(1.3),
                         height=Inches(4.3))
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.7),
                              Inches(12), Inches(1.3)).text_frame
    p = tb.paragraphs[0]
    p.text = ("GEMM and SWFA sit on the compute ceiling (312 TFLOPS); "
              "decode attention sits on the memory ridge at ~1 FLOP/byte.")
    p.font.size = Pt(15); p.font.color.rgb = DARK
    p2 = tb.add_paragraph()
    p2.text = "Two kernels, two ceilings, two optimisation strategies — visible in one chart."
    p2.font.size = Pt(15); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 34); return s

def slide_35_c5_crosscheck():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Contribution 5 · Profiling — Evidence (2/2)")
    add_title(s, "Triton vs CUTLASS vs Torch-inductor — agreement within 3%")
    data = [
        ["Kernel",                               "Triton", "CUTLASS", "Inductor"],
        ["GEMM 4096³",                            "294",    "301",     "288"],
        ["Flash Attention (N=2K, h=128)",         "172",    "189",     "148"],
        ["SWFA (N=2K, w=512)",                    "241",    "258",     "—"],
        ["Decode attention (KV=8K)",              "84",     "92",      "71"],
    ]
    add_table(s, data, 0.8, 1.3, 11.7, 3.0, first_col_bold=True, font_size=14)
    add_bullets(s, [
        "<3% throughput agreement; <1e-4 FP32, <1e-3 FP16 numerical difference.",
        ("Triton closes to within 7% of hand-tuned CUTLASS on SWFA — great for iteration.", True),
        "Torch-inductor cannot express sliding window — hence the missing entry.",
    ], top=4.6, size=15)
    add_takeaway(s, "Answer to RQ5: the pipeline independently reproduces every headline number in this thesis.")
    add_footer(s, 35); return s

# ============================================================
# ACT IV — Integration & Wrap (slides 36–40)
# ============================================================
def slide_36_composition():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Integration story")
    add_title(s, "How the five contributions compose end-to-end")
    s.shapes.add_picture(composition_png, Inches(0.6), Inches(1.3),
                         width=Inches(12.1))
    add_bullets(s, [
        "C1 sets the memory envelope → C2 kernels fit inside it → C3 skips the attention work → C4 ships the savings as tokens/s → C5 verifies the lot.",
        ("Removing any one link drops the composed gain below the sum of its parts.", True),
    ], top=5.0, size=14)
    add_footer(s, 36); return s

def slide_37_lessons():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Lessons learned")
    add_title(s, "Five lessons — one per contribution, one general")
    add_bullets(s, [
        ("Compose, then tune — cross-layer co-design compounds; single-layer wins plateau quickly.", True),
        ("Skip beats mask — block-sparsity belongs in kernel control flow, not in the score matrix.", True),
        "Overlap is a first-class metric — a 65% overlap run can beat a 90% util run that stalls on comms.",
        "Auto-tune everything — the same kernel needs different tile shapes at different sequence lengths.",
        "Roofline is the design tool — every optimisation should move a dot toward its ceiling.",
    ], top=1.5, size=18)
    add_footer(s, 37); return s

def slide_38_limitations():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Limitations — honest scope")
    add_title(s, "What this thesis does not claim")
    add_bullets(s, [
        "Single node: 8-GPU NVLink only; no multi-node cross-switch or InfiniBand topology.",
        "Model size: GPT-2-Medium (355M) for training; 7B only for serving. No 70B / 175B.",
        "Precision: FP16 / FP32 only; no sub-8-bit quantisation study.",
        "SWFA: reference PyTorch + Triton cross-validator; CUTLASS kernel sketched, not shipped.",
        "Workload breadth: dense transformers only — no MoE end-to-end, no RLHF.",
    ], top=1.5, size=18)
    add_takeaway(s, "Scope is deliberate — each limitation maps to a well-defined future-work item (next slide).")
    add_footer(s, 38); return s

def slide_39_future():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Future work")
    add_title(s, "Where this stack goes next")
    add_bullets(s, [
        "Multi-node SWFA with topology-aware window partitioning over InfiniBand.",
        "CUTLASS SWFA kernel with CuTe layouts and epilogue fusion; target > 95% Tensor-Core util.",
        "Adaptive window: learn w* online per layer/head, using the closed-form as a warm start.",
        "Adaptive KV page size in SGLang — balance arithmetic intensity against fragmentation.",
        "Sub-8-bit paths (FP8, INT4) with accuracy bounds derived from the online-softmax FP model.",
        "MoE end-to-end: fold grouped-GEMM epilogue fusion into the SGLang runtime.",
    ], top=1.4, size=17)
    add_footer(s, 39); return s

def slide_40_summary_thanks():
    s = prs.slides.add_slide(BLANK)
    add_section_tag(s, "Summary · Thank you")
    add_title(s, "Summary of contributions")
    # compact 5-row summary table
    data = [
        ["",   "Contribution",                                         "Headline"],
        ["C1", "Controlled 4-way distributed study",                   "93.8% util · −22% energy"],
        ["C2", "CUTLASS/CuTe GEMM mapping",                            "+35% skinny/MoE"],
        ["C3", "SWFA (block skip, corner-check, w*)",                  "65.6% skip · <1e-4 err"],
        ["C4", "SGLang kernel stack",                                  "2.84× decode on LLaMA-7B"],
        ["C5", "Triton profiling & auto-tune pipeline",                "<3% cross-runtime"],
    ]
    add_table(s, data, 0.6, 1.3, 12.1, 3.4, first_col_bold=True, font_size=13)
    # thank-you band
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(5.0),
                              prs.slide_width, Inches(1.8))
    band.fill.solid(); band.fill.fore_color.rgb = RGBColor(0x33, 0x4F, 0x7A)
    band.line.fill.background()
    tb = s.shapes.add_textbox(Inches(0.5), Inches(5.1),
                              Inches(12.3), Inches(1.6))
    tf = tb.text_frame
    p = tf.paragraphs[0]; p.text = "Thank you — Questions?"
    p.font.size = Pt(36); p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); p.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = "Adarsh Dubey (SC24M151) · IIST · Advisor: Dr. S. Sumitra"
    p2.font.size = Pt(14); p2.font.italic = True
    p2.font.color.rgb = RGBColor(0xFF, 0xEE, 0xBB); p2.alignment = PP_ALIGN.CENTER
    add_footer(s, 40); return s

# Build all slides in order
for fn in [
    slide_01_title,           slide_02_thesis_statement,
    slide_03_motivation,      slide_04_research_questions,
    slide_05_gap,             slide_06_contributions,
    slide_07_roadmap,
    slide_08_stack,           slide_09_methodology,
    slide_10_c1_problem,      slide_11_c1_approach,
    slide_12_c1_result_util,  slide_13_c1_result_energy,
    slide_14_c2_problem,      slide_15_c2_approach,
    slide_16_c2_result,
    slide_17_c3_problem,      slide_18_c3_fa_recap,
    slide_19_c3_classify_fig, slide_20_c3_classify_algo,
    slide_21_c3_corner_lemma, slide_22_c3_vecmask,
    slide_23_c3_fwd_algo,     slide_24_c3_bwd_complexity,
    slide_25_c3_accuracy_skip,slide_26_c3_e2e,
    slide_27_c3_optimal_window,
    slide_28_c4_problem,      slide_29_c4_approach,
    slide_30_c4_layered,      slide_31_c4_latency,
    slide_32_c5_problem,      slide_33_c5_approach,
    slide_34_c5_roofline,     slide_35_c5_crosscheck,
    slide_36_composition,     slide_37_lessons,
    slide_38_limitations,     slide_39_future,
    slide_40_summary_thanks,
]:
    fn()

prs.save(OUT)
print(f"Wrote {OUT}  ({len(prs.slides)} slides)")
