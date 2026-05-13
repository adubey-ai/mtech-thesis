"""
Build thesis_defense.pptx — 40-slide PowerPoint mirroring the Beamer deck.

Strategy:
  * Use python-pptx for layout + text.
  * Render each slide's math-heavy / equation region as a PNG via matplotlib
    (with full LaTeX through mathtext) and embed the PNG where Beamer renders
    a block equation.
  * Tables are native python-pptx tables.
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
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12.3), Inches(0.8))
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
        p2.font.color.rgb = MUTED

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
    # Name on the left
    tb2 = slide.shapes.add_textbox(Inches(0.3), Inches(7.15),
                                   Inches(8), Inches(0.3))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = "Adarsh Dubey — M.Tech Thesis Defense (IIST)"
    p2.font.size = Pt(10)
    p2.font.color.rgb = MUTED

# ---------- math rendering helpers ----------
def render_math_png(latex_str, fontsize=22, pad=0.1, filename=None):
    """Render a LaTeX math expression (mathtext) to PNG, return path."""
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, f"${latex_str}$", fontsize=fontsize)
    out = filename or os.path.join(IMG_DIR, f"eq_{abs(hash(latex_str))}.png")
    fig.savefig(out, dpi=300, bbox_inches="tight", pad_inches=pad,
                transparent=True)
    plt.close(fig)
    return out

def render_math_block_png(lines, fontsize=20, filename=None):
    """Render a multi-line math block as a single PNG."""
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
    # legend
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

# Pre-render charts
util_png     = os.path.join(IMG_DIR, "util_bar.png");      render_bar_chart_util(util_png)
scaling_png  = os.path.join(IMG_DIR, "scaling.png");       render_scaling_plot(scaling_png)
roofline_png = os.path.join(IMG_DIR, "roofline.png");      render_roofline(roofline_png)
classify_png = os.path.join(IMG_DIR, "classify.png");      render_block_classify(classify_png)
stack_png    = os.path.join(IMG_DIR, "stack.png");         render_stack_diagram(stack_png)

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

# ============================================================
# Slide builders
# ============================================================
def slide_01_title():
    s = prs.slides.add_slide(BLANK)
    # Title band
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.4),
                              prs.slide_width, Inches(2.6))
    band.fill.solid(); band.fill.fore_color.rgb = RGBColor(0x33, 0x4F, 0x7A)
    band.line.fill.background()
    tb = s.shapes.add_textbox(Inches(0.5), Inches(2.6),
                              Inches(12.3), Inches(2.3))
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
    # Author
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

def slide_02_glance():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "At a Glance")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(1.0))
    p = tb.text_frame.paragraphs[0]
    p.text = "One-line thesis: co-design every layer of the LLM stack so the savings compose, then measure it."
    p.font.size = Pt(16); p.font.italic = True; p.font.color.rgb = MUTED
    add_bullets(s, [
        ("93.8%  peak GPU utilisation with DeepSpeed ZeRO-3", True),
        ("65.6%  attention blocks skipped by SWFA (<1e-4 FP32 error)", True),
        ("50-59%  end-to-end training speedup on GPT-2-Medium", True),
        ("2.84×  decode throughput on LLaMA-7B (SGLang kernels)", True),
        ("25-29%  energy savings per epoch", True),
        ("8× A100-40GB, NVLink, PyTorch 2.0, CUDA 12.x, NCCL 2.19", False),
    ], top=2.3, size=20)
    add_footer(s, 2); return s

def slide_03_problem():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "The Problem")
    add_bullets(s, [
        "Schedule: ~10³ accelerators must stay in lockstep.",
        "Kernel: GEMMs must saturate Tensor Cores.",
        "Algorithm: attention must scale sub-quadratically with N.",
        "Runtime: decode must turn FLOPs savings into tokens/s under SLOs.",
        "Tooling: every kernel needs a roofline and an auto-tuner.",
    ], top=1.3, size=20)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.5),
                              Inches(12), Inches(1.2))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Isolated optimisations at any one layer leave the gains of the others on the table."
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    p2 = tf.add_paragraph()
    p2.text = "This thesis builds an end-to-end stack in which the five layers are designed to compose."
    p2.font.size = Pt(16); p2.font.color.rgb = DARK
    add_footer(s, 3); return s

def slide_04_stack():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "The Five-Layer Stack")
    s.shapes.add_picture(stack_png, Inches(2.5), Inches(1.3),
                         height=Inches(5.3))
    add_footer(s, 4); return s

def slide_05_contributions():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Contributions")
    add_bullets(s, [
        "Controlled 8-GPU study of DP / DDP / FSDP / ZeRO-3 — identical seeds and workload.",
        ("CUTLASS/CuTe mapping: up to 35% speedup on skinny and grouped GEMMs.", False),
        ("SWFA: block-sparse Sliding Window Flash Attention with corner-check classifier, vectorised mask, and a closed-form optimal window w* (strict convexity proved).", True),
        ("SGLang kernel stack: paged KV + fused decode + SWFA prefill + fused sampling = 2.84× on LLaMA-7B.", True),
        "Triton profiling & auto-tune pipeline; cross-validates CUTLASS / Torch-inductor within <3%.",
    ], top=1.3, size=18)
    add_footer(s, 5); return s

def slide_06_roadmap():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Talk Roadmap")
    col1 = s.shapes.add_textbox(Inches(0.7), Inches(1.4),
                                Inches(6), Inches(5.5)).text_frame
    col2 = s.shapes.add_textbox(Inches(7.0), Inches(1.4),
                                Inches(6), Inches(5.5)).text_frame
    for tf, parts in [(col1, [("Part I — Foundations (§2–3)", "literature · distributed training"),
                              ("Part II — Kernels (§4)",       "CUTLASS + CuTe GEMMs"),
                              ("Part III — Algorithm (§5)",    "Sliding Window Flash Attention")]),
                      (col2, [("Part IV — Serving (§6)",       "SGLang kernel stack"),
                              ("Part V — Tooling (§7)",        "Triton profiling & auto-tune"),
                              ("Part VI — Wrap (§8)",          "lessons, limitations, future work")])]:
        for i, (h, sub) in enumerate(parts):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = h
            p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = DARK
            p.space_after = Pt(2)
            p2 = tf.add_paragraph()
            p2.text = sub
            p2.font.size = Pt(14); p2.font.color.rgb = MUTED
            p2.space_after = Pt(18)
    add_footer(s, 6); return s

def slide_07_priorart():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Prior Art at Each Layer")
    data = [
        ["Layer", "Representative prior work"],
        ["Distributed",  "ZeRO (Rajbhandari '20), FSDP (Zhao '23), Megatron-LM"],
        ["GEMM kernels", "CUTLASS, CuTe, cuBLAS, Triton"],
        ["Attention",    "FlashAttention 1/2 (Dao '22/'23), Longformer, BigBird"],
        ["Serving",      "vLLM (paged attention), SGLang, TGI"],
        ["Profiling",    "Nsight Compute, NVTX, Triton benchmark"],
    ]
    add_table(s, data, 1.3, 1.3, 10.7, 3.5, first_col_bold=True, font_size=14)
    tb = s.shapes.add_textbox(Inches(0.8), Inches(5.3),
                              Inches(12), Inches(1.3)).text_frame
    p = tb.paragraphs[0]
    p.text = "Each layer has strong isolated work, but the interfaces between them are ad-hoc."
    p.font.size = Pt(16); p.font.color.rgb = DARK
    p2 = tb.add_paragraph()
    p2.text = "No prior work reports a controlled five-layer study on the same hardware."
    p2.font.size = Pt(16); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 7); return s

def slide_08_position():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Where This Thesis Sits")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(5.5)).text_frame
    p = tb.paragraphs[0]; p.text = "What is new:"
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = ACCENT
    for item in [
        "Controlled 4-way distributed study on one node (identical seeds, workload, init).",
        "Block-classification step in front of Flash Attention that skips whole tiles.",
        "Closed-form optimal window w* with a convexity proof.",
        "Layered evaluation of 5 SGLang kernel optimisations that attributes the cumulative 2.84×.",
    ]:
        p = tb.add_paragraph(); p.text = "  • " + item
        p.font.size = Pt(16); p.font.color.rgb = DARK; p.space_after = Pt(4)
    p = tb.add_paragraph(); p.text = ""; p.font.size = Pt(6)
    p = tb.add_paragraph(); p.text = "What is deliberately out of scope:"
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = MUTED
    p = tb.add_paragraph()
    p.text = "  multi-node scaling, sub-8-bit quantisation, RLHF workloads."
    p.font.size = Pt(16); p.font.color.rgb = DARK
    add_footer(s, 8); return s

def slide_09_fourstrat():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Four Distributed Training Strategies")
    data = [
        ["", "Params", "Grads", "Optimiser", "Comms"],
        ["DP",     "replicated", "replicated",   "replicated",    "scatter/gather"],
        ["DDP",    "replicated", "ring-reduced", "replicated",    "all-reduce"],
        ["FSDP",   "sharded",    "sharded",      "sharded",       "gather→compute→shard"],
        ["ZeRO-3", "sharded",    "sharded",      "sharded",       "fine-grained overlap"],
    ]
    add_table(s, data, 1.0, 1.3, 11.3, 3.0, first_col_bold=True, font_size=13)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(4.8),
                              Inches(12), Inches(2.2)).text_frame
    p = tb.paragraphs[0]
    p.text = "Trade-off: replication → fewer comms, higher memory; sharding → P× memory reduction, extra gather/scatter."
    p.font.size = Pt(16); p.font.color.rgb = DARK
    p2 = tb.add_paragraph()
    p2.text = "Goal: isolate the effect of the strategy itself, not the framework or workload."
    p2.font.size = Pt(16); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 9); return s

def slide_10_memmath():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Memory Partitioning Math")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "Let M = parameter memory, P = GPUs, B = per-GPU batch:"
    p.font.size = Pt(16); p.font.color.rgb = DARK
    embed_math_block(s,
        [
          r"\mathrm{DP,\ DDP}:\ \ M_{param} + M_{grad} + M_{opt} + M_{act}(B)",
          r"\mathrm{FSDP}:\ \ \frac{1}{P}(M_{param} + M_{grad} + M_{opt}) + M_{act}(B)",
          r"\mathrm{ZeRO}\text{-}3:\ \ \frac{1}{P}(M_{param} + M_{grad} + M_{opt}) + M_{act}(B) + M_{bucket}",
        ], left=1.1, top=2.1, width=11.0, fontsize=22)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(5.0),
                               Inches(12), Inches(1.8)).text_frame
    p = tb2.paragraphs[0]
    p.text = "For Adam on P=8: 87.5% savings on param+grad+opt; activation memory unchanged."
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = ACCENT
    p2 = tb2.add_paragraph()
    p2.text = "Ring all-reduce time: T_AR(M) = 2(P-1)(τ + M/(Pβ)) → 2M/β  (as P → ∞)."
    p2.font.size = Pt(16); p2.font.color.rgb = DARK
    add_footer(s, 10); return s

def slide_11_method():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Experimental Methodology")
    tb1 = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                               Inches(6), Inches(5.5)).text_frame
    for i, line in enumerate([
        ("Hardware", "8× A100-40GB, NVLink, single node,\nCUDA 12.x, NCCL 2.19"),
        ("Workload (training)", "ResNet-50 / CIFAR-10 (controlled proxy)\nGPT-2 Medium (355M) for LLM runs"),
        ("Seeds", "fixed across strategies, framework, kernels"),
    ]):
        p = tb1.paragraphs[0] if i == 0 else tb1.add_paragraph()
        p.text = line[0]
        p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
        p.space_after = Pt(2)
        p2 = tb1.add_paragraph(); p2.text = line[1]
        p2.font.size = Pt(15); p2.font.color.rgb = DARK
        p2.space_after = Pt(16)
    tb2 = s.shapes.add_textbox(Inches(7), Inches(1.3),
                               Inches(6), Inches(5.5)).text_frame
    p = tb2.paragraphs[0]; p.text = "Metrics"
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    for m in [
        "GPU utilisation (avg / peak)",
        "Per-GPU memory (model / grad / opt)",
        "Training-loss convergence",
        "Power and energy per epoch",
        "Compute / communication overlap",
        "Throughput scaling",
    ]:
        p = tb2.add_paragraph(); p.text = "• " + m
        p.font.size = Pt(15); p.font.color.rgb = DARK
    add_footer(s, 11); return s

def slide_12_resultsutil():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Results — GPU Utilisation & Memory")
    s.shapes.add_picture(util_png, Inches(0.4), Inches(1.4),
                         height=Inches(4.4))
    data = [
        ["Strategy", "Mem/GPU (GB)", "Reduction"],
        ["DP",     "3.92", "1.0×"],
        ["DDP",    "3.92", "1.0×"],
        ["FSDP",   "0.49", "8.0×"],
        ["ZeRO-3", "0.49", "8.0×"],
    ]
    add_table(s, data, 7.6, 1.7, 5.2, 2.5, first_col_bold=True, font_size=13)
    tb = s.shapes.add_textbox(Inches(7.6), Inches(4.5),
                              Inches(5.4), Inches(2.5)).text_frame
    p = tb.paragraphs[0]
    p.text = "ZeRO-3 peaks at 93.8% via fine-grained overlap of param gather with compute."
    p.font.size = Pt(14); p.font.color.rgb = DARK; p.space_after = Pt(6)
    p2 = tb.add_paragraph()
    p2.text = "Sharding delivers the theoretical ≈87.5% memory saving."
    p2.font.size = Pt(14); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 12); return s

def slide_13_energy():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Results — Energy & Communication Overlap")
    data = [
        ["Strategy", "Avg (W)", "kWh/ep", "Comms", "Overlap (%)"],
        ["DP",        "2840", "0.95", "poor",      "12"],
        ["DDP",       "2720", "0.87", "good",      "64"],
        ["FSDP",      "2650", "0.82", "good",      "58"],
        ["DeepSpeed", "2580", "0.78", "excellent", "61"],
    ]
    add_table(s, data, 1.2, 1.3, 11.0, 2.6, first_col_bold=True, font_size=14)
    add_bullets(s, [
        ("ZeRO-3: 22% lower per-epoch energy than DP at the same final loss.", True),
        "DDP overlaps best: ring all-reduce runs during backward of earlier layers.",
        "ZeRO-3's param-gather phase blocks compute slightly; overlap_comm hides it partly.",
    ], top=4.3, size=16)
    add_footer(s, 13); return s

def slide_14_cutlass():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "CUTLASS & CuTe — The Tile Hierarchy")
    add_bullets(s, [
        "Threadblock tile (Mtb, Ntb, Ktb) — unit of HBM traffic.",
        "Warp tile (Mw, Nw, Kw) — unit of shared-memory traffic.",
        "Instruction tile (Mi, Ni, Ki) — Tensor Core MMA shape, e.g. 16×8×16.",
    ], top=1.3, size=18)
    embed_math_block(s,
        [r"\mathrm{Layout}(S,D)\ \circ\ \mathrm{Layout}(S',D') = \mathrm{Layout}((S,S'),(D \cdot D', D'))"],
        left=0.8, top=3.6, width=12.0, fontsize=22)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(4.8),
                              Inches(12), Inches(2)).text_frame
    p = tb.paragraphs[0]
    p.text = "Multi-stage software pipelining: prefetch tile k+1 while MMAs on tile k execute."
    p.font.size = Pt(16); p.font.color.rgb = DARK
    p2 = tb.add_paragraph()
    p2.text = "Square GEMM → 3 stages; skinny GEMM → 2 stages (short K can't hide a third)."
    p2.font.size = Pt(15); p2.font.italic = True; p2.font.color.rgb = MUTED
    add_footer(s, 14); return s

def slide_15_epilogue():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Epilogue Fusion")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]; p.text = "Without fusion — three HBM round-trips:"
    p.font.size = Pt(16); p.font.color.rgb = DARK
    embed_math_block(s,
        [r"Y \leftarrow W X; \ \ Z \leftarrow Y + b; \ \ Z \leftarrow \phi(Z)"],
        left=1.2, top=1.9, width=10.0, fontsize=22)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(3.0),
                               Inches(12), Inches(0.6)).text_frame
    p = tb2.paragraphs[0]; p.text = "With fusion — one HBM round-trip:"
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = ACCENT
    embed_math_block(s,
        [r"Z_{ij} \leftarrow \phi\left(\, \sum_k W_{ik} X_{kj} + b_i \,\right)"],
        left=1.2, top=3.6, width=10.0, fontsize=22)
    add_bullets(s, [
        ("GPT-2 MLP with bias+GELU fused into FC1: 18% faster at N=2048.", True),
        ("Residual-add fused into FC2: 22% faster at N=4096.", True),
        "Savings come entirely from eliminated HBM passes.",
    ], top=5.0, size=16)
    add_footer(s, 15); return s

def slide_16_cutlass_res():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Results — CUTLASS vs cuBLAS (FP16, TFLOPS)")
    data = [
        ["Shape (M×N×K)", "cuBLAS", "CUTLASS", "Ratio"],
        ["4096³",                        "280", "301", "1.08×"],
        ["8192³",                        "305", "318", "1.04×"],
        ["2048×2048×128 (skinny)",       "95",  "128", "1.35×"],
        ["Grouped MoE, 16 experts",      "184", "241", "1.31×"],
    ]
    add_table(s, data, 1.5, 1.3, 10.3, 2.8, first_col_bold=True, font_size=15)
    add_bullets(s, [
        "Matches cuBLAS on large shapes (both compute-bound, TC util >90%).",
        ("Beats cuBLAS on skinny and grouped GEMMs — the shapes that dominate attention & MoE.", True),
        "35% win on 2048×2048×128 from warp tile 64×32×32 matching register budget.",
    ], top=4.4, size=16)
    add_footer(s, 16); return s

def slide_17_swfa_motiv():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "SWFA — Motivation")
    embed_math_block(s,
        [r"\mathrm{Attn}(Q,K,V) = \mathrm{softmax}\left( \frac{QK^{T}}{\sqrt{d}} \right) V, \ \ \mathrm{cost} = \Theta(N^2 d)"],
        left=0.8, top=1.3, width=12.0, fontsize=22)
    add_bullets(s, [
        "Flash Attention fixes the memory wall — still issues dense GEMMs on every Br×Bc tile.",
        "Long-context models (Longformer, BigBird, Mistral-7B) restrict to a local window [i-wl, i+wr].",
        "Theory: cost drops from Θ(N²) to Θ(Nw) — only if the kernel is window-aware.",
    ], top=3.0, size=17)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.8),
                              Inches(12), Inches(0.8)).text_frame
    p = tb.paragraphs[0]
    p.text = "SWFA = Flash Attention  +  explicit block classification  +  vectorised mask."
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = ACCENT
    add_footer(s, 17); return s

def slide_18_onlinesoftmax():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Flash Attention — Online Softmax")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "For each row tile i, maintain (m_i, ℓ_i, O_i) = (row-max, row-sum, accumulator):"
    p.font.size = Pt(16); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"m_i^{new} = \max(m_i,\ \max_j s_{ij})",
        r"\ell_i^{new} = e^{m_i - m_i^{new}} \ell_i + \sum_j e^{s_{ij} - m_i^{new}}",
        r"O_i^{new} = e^{m_i - m_i^{new}} O_i + \sum_j e^{s_{ij} - m_i^{new}} v_j",
    ], left=1.0, top=1.9, width=11.0, fontsize=20)
    add_bullets(s, [
        ("Peak HBM residency drops from Θ(N²) to Θ(Br·Bc); same answer up to FP associativity.", True),
        ("Key insight: identical (m, ℓ, O) state-machine block-by-block — so we can SKIP blocks.", True),
    ], top=5.4, size=16)
    add_footer(s, 18); return s

def slide_19_classify_fig():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Block Classification: I / X / O")
    s.shapes.add_picture(classify_png, Inches(2.8), Inches(1.3),
                         height=Inches(5.3))
    add_footer(s, 19); return s

def slide_20_classify_algo():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Classification Algorithm (Corner Check)")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(5.5))
    tf = tb.text_frame
    lines = [
        "Input:  N, window (wl, wr), block sizes (Br, Bc)",
        "Output: sets I, X, O",
        "",
        "Tr ← ⌈N/Br⌉;  Tc ← ⌈N/Bc⌉",
        "for r = 0 .. Tr−1:",
        "    r0 ← r·Br;  r1 ← min((r+1)·Br, N) − 1",
        "    for c = 0 .. Tc−1:",
        "        c0 ← c·Bc;  c1 ← min((c+1)·Bc, N) − 1",
        "        outside ← (c1 < r0 − wl) or (c0 > r1 + wr)",
        "        inside  ← (c1 ≤ r0 + wr) and (c0 ≥ r1 − wl)",
        "        if outside:   O ← O ∪ {(r,c)}",
        "        elif inside:  I ← I ∪ {(r,c)}",
        "        else:         X ← X ∪ {(r,c)}",
        "return (I, X, O)",
    ]
    for i, l in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = l
        p.font.name = "Consolas"
        p.font.size = Pt(14)
        p.font.color.rgb = DARK
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(6.3),
                               Inches(12), Inches(0.8)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Cost: O(Tr·Tc) = O((N/B)²), dominated by the O(Nw/B) non-skipped kernel work."
    p.font.size = Pt(14); p.font.italic = True; p.font.color.rgb = MUTED
    add_footer(s, 20); return s

def slide_21_corner_lemma():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Corner-Check Sufficiency (Lemma)")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(1.0)).text_frame
    p = tb.paragraphs[0]
    p.text = "Let B = [r0, r1] × [c0, c1] be a block, W = {(i,j): i − wl ≤ j ≤ i + wr}. Then:"
    p.font.size = Pt(16); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"(i)\ \ B \cap W = \emptyset \Leftrightarrow c_1 < r_0 - w_l \ \mathrm{or}\ c_0 > r_1 + w_r",
        r"(ii)\ \ B \subseteq W \Leftrightarrow c_1 \leq r_0 + w_r \ \mathrm{and}\ c_0 \geq r_1 - w_l",
    ], left=0.9, top=2.4, width=11.3, fontsize=20)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(4.5),
                               Inches(12), Inches(2.4)).text_frame
    p = tb2.paragraphs[0]; p.text = "Proof sketch."
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = ACCENT
    p2 = tb2.add_paragraph()
    p2.text = "Window boundaries j = i ± wr, j = i − wl are monotone in i, so min (j−i) over B is attained at (r1, c0) and max at (r0, c1). Outside/inside reduce to inequalities on these two corners. ∎"
    p2.font.size = Pt(15); p2.font.color.rgb = DARK
    p3 = tb2.add_paragraph(); p3.text = ""; p3.font.size = Pt(6)
    p4 = tb2.add_paragraph()
    p4.text = "Consequence: classification needs only 4 comparisons per block — trivial overhead."
    p4.font.size = Pt(16); p4.font.bold = True; p4.font.color.rgb = ACCENT
    add_footer(s, 21); return s

def slide_22_vecmask():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Vectorised Mask Generation for X")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(0.7)).text_frame
    p = tb.paragraphs[0]
    p.text = "For block (r,c), let δ_rc = c·Bc − r·Br. A local offset rel = j_ℓ − i_ℓ is masked iff:"
    p.font.size = Pt(15); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"\mathrm{rel} > w_r + \delta_{rc} \ \ \mathrm{or}\ \ \mathrm{rel} < -w_l + \delta_{rc}",
        r"d_{right} = \lfloor w_r + \delta_{rc} \rfloor + 1,\ \ d_{left} = \lfloor -w_l + \delta_{rc} \rfloor - 1",
        r"M_{rc} = \mathrm{triu}(\mathbf{1}_{B_r \times B_c},\ d_{right})\ \vee\ \mathrm{tril}(\mathbf{1}_{B_r \times B_c},\ d_{left})",
    ], left=0.8, top=2.1, width=11.5, fontsize=20)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(5.7),
                               Inches(12), Inches(1.3)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Replaces nested O(Br·Bc) Python loop by a single triu/tril call."
    p.font.size = Pt(16); p.font.color.rgb = DARK
    p2 = tb2.add_paragraph()
    p2.text = "≈100× faster than naive masking at the block level."
    p2.font.size = Pt(16); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 22); return s

def slide_23_fwd_algo():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "SWFA Forward Algorithm")
    tb = s.shapes.add_textbox(Inches(0.5), Inches(1.2),
                              Inches(12.3), Inches(5.8))
    tf = tb.text_frame
    lines = [
        "Input: Q ∈ R^{B×H×N×d}, K, V ∈ R^{B×H×N×dv},  (wl, wr),  (Br, Bc)",
        "",
        "(I, X, O) ← ClassifyBlocks()",
        "for r = 0 .. Tr − 1:",
        "    O_r ← 0;  m_r ← −∞;  ℓ_r ← 0",
        "    for c = 0 .. Tc − 1:",
        "        if (r,c) ∈ O: continue              ← SKIP BLOCK",
        "        S_rc ← Q_r · K_c^T / √d",
        "        if (r,c) ∈ X: S_rc += GenerateMask(r,c)",
        "        m_new ← max(m_r, max(S_rc));  corr ← exp(m_r − m_new)",
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

def slide_24_backward_complexity():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Backward Pass & Complexity")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "Backward re-runs the forward loop on the same block set I ∪ X:"
    p.font.size = Pt(16); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"P_{rc} = e^{S_{rc} - m_r} / \ell_r,\ \ dS_{rc} = P_{rc} \odot (dP_{rc} - \rho_r)",
        r"\rho_r = \sum_c P_{rc} \odot (dO_r\, V_c^{T})",
    ], left=0.8, top=1.9, width=11.0, fontsize=20)
    data = [
        ["Total blocks",                   "(N/B)²"],
        ["Inside band (I ∪ X)",            "≈ 2·(N/B)·(w/B)"],
        ["Skipped (O)",                    "(N/B)² − 2·(N/B)·(w/B)"],
        ["Theoretical speedup",            "1 − 2w/N  (as w ≪ N)"],
        ["Memory",                         "Θ(N)  — same as Flash Attention"],
    ]
    add_table(s, [["Quantity", "Expression"]] + data,
              1.5, 3.8, 10.3, 2.8, first_col_bold=True, font_size=13)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(6.7),
                               Inches(12), Inches(0.6)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Example: N=128, w=4, B=16 ⇒ 48 of 64 blocks skipped (75%)."
    p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = ACCENT
    add_footer(s, 24); return s

def slide_25_accuracy_skip():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Accuracy & Block-Skip Efficiency")
    acc = [
        ["Config", "Max diff", "Tol."],
        ["sym narrow w=2",    "2.4e-7",  "1e-4"],
        ["sym wide w=8",      "2.4e-7",  "1e-4"],
        ["causal+win wl=6",   "4.8e-7",  "1e-4"],
        ["multi-head H=8",    "8.3e-7",  "1e-4"],
        ["long seq N=128",    "7.5e-7",  "1e-4"],
        ["FP16 w=4,6",        "1.95e-3", "1e-2"],
    ]
    add_table(s, acc, 0.4, 1.3, 6.0, 3.2, first_col_bold=True, font_size=12)
    tb = s.shapes.add_textbox(Inches(0.4), Inches(4.7),
                              Inches(6.0), Inches(0.5)).text_frame
    p = tb.paragraphs[0]; p.text = "All 19 configurations pass."
    p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = ACCENT
    skip = [
        ["Config", "Blocks", "Skip %"],
        ["very narrow w=1,2",   "64",  "65.6%"],
        ["sym narrow  w=2",     "16",  "37.5%"],
        ["causal+win  w=6,0",   "16",  "56.2%"],
        ["left only   w=10,0",  "16",  "56.2%"],
        ["N=Ns=45",             "9",   "22.2%"],
        ["win>seq w=20",        "16",  "0.0%"],
    ]
    add_table(s, skip, 6.8, 1.3, 6.0, 3.2, first_col_bold=True, font_size=12)
    tb2 = s.shapes.add_textbox(Inches(6.8), Inches(4.7),
                               Inches(6.0), Inches(0.5)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Matches the 1 − 2w/N prediction."
    p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = ACCENT
    add_footer(s, 25); return s

def slide_26_e2e_scaling():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "End-to-End Training & Long-Sequence Scaling")
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
    tb = s.shapes.add_textbox(Inches(0.4), Inches(4.7),
                              Inches(6.0), Inches(1.5)).text_frame
    p = tb.paragraphs[0]
    p.text = "50–59% throughput, 25–29% energy; perplexity shift <0.1."
    p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = ACCENT
    s.shapes.add_picture(scaling_png, Inches(6.6), Inches(1.3),
                         height=Inches(4.5))
    add_footer(s, 26); return s

def slide_27_optimal_window():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Optimal Window: Cost Model & Convexity")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "Compute / memory / performance-loss cost:"
    p.font.size = Pt(16); p.font.color.rgb = DARK
    embed_math_block(s, [
        r"\mathcal{L}(w) = \alpha\,C(w) + \beta\,M(w) + \gamma\,P(w)",
        r"C(w) = \kappa\,N\,w\,d,\ \ M(w) \approx M_0",
        r"P(w) = \frac{2\,N\,A_0\,e^{-\lambda w/2}}{1 - e^{-\lambda}}\ \ (\mathrm{decay}:\ A_{ij} = A_0 e^{-\lambda|i-j|})",
    ], left=0.9, top=1.9, width=11.3, fontsize=19)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(4.5),
                               Inches(12), Inches(2.2)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Theorem (strict convexity): L(w) is strictly convex for w > 0."
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = ACCENT
    p2 = tb2.add_paragraph()
    p2.text = "Proof: C''=M''=0; P'' ∝ (λ²/4)·e^{−λw/2} > 0 ⇒ L'' > 0. ∎"
    p2.font.size = Pt(14); p2.font.color.rgb = DARK
    p3 = tb2.add_paragraph(); p3.text = ""
    p4 = tb2.add_paragraph()
    p4.text = "Closed-form optimum:  w* = (2/λ) · ln[ γ·A₀·λ / ( α·κ·d·(1 − e^{-λ}) ) ]"
    p4.font.size = Pt(15); p4.font.bold = True; p4.font.color.rgb = ACCENT
    add_footer(s, 27); return s

def slide_28_sglang():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "SGLang — Architecture & Execution")
    add_bullets(s, [
        "Prefill (compute-bound): processes full prompt, builds KV cache.",
        "Decode (memory-bound):   one token at a time; reads all prior KV.",
    ], top=1.3, size=17)
    # pipeline shapes
    shapes = s.shapes
    x0 = Inches(0.5)
    for i, (txt, col) in enumerate([
        ("Prefill (SWFA)",    RGBColor(0xCB, 0xE4, 0xFB)),
        ("Fused Decode Attn", RGBColor(0xD6, 0xF2, 0xD0)),
        ("Fused Sampling",    RGBColor(0xFF, 0xF2, 0xB3)),
    ]):
        sh = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.8 + i*4.2), Inches(3.2),
                              Inches(3.6), Inches(0.9))
        sh.fill.solid(); sh.fill.fore_color.rgb = col
        sh.line.color.rgb = DARK
        tf = sh.text_frame; tf.text = txt
        tf.paragraphs[0].font.size = Pt(16)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    # KV cache below
    sh = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                          Inches(5.0), Inches(4.6),
                          Inches(3.6), Inches(0.9))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
    sh.line.color.rgb = DARK
    tf = sh.text_frame; tf.text = "Paged KV Cache (16-tok pages)"
    tf.paragraphs[0].font.size = Pt(14); tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    add_bullets(s, [
        "Paged KV: 16-tok pages eliminate fragmentation.",
        "Fused decode: matmul + softmax + masking in one kernel.",
        "SWFA prefill: drop-in for window models.",
        "Fused top-k/top-p: removes 3–4 kernel launches per step.",
    ], top=5.7, size=14)
    add_footer(s, 28); return s

def slide_29_layered():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Layered Decode-Throughput Gain on LLaMA-7B")
    data = [
        ["Configuration",                         "Tok/s", "Speed-up"],
        ["Baseline (vanilla attn, unpaged KV)",    "980",    "1.00×"],
        ["+ Paged KV cache",                       "1,420",  "1.45×"],
        ["+ Fused decode attention",               "1,860",  "1.90×"],
        ["+ SWFA prefill",                         "2,410",  "2.46×"],
        ["+ Fused sampling",                       "2,780",  "2.84×"],
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
        p = tb.add_paragraph(); p.text = "• " + line
        p.font.size = Pt(12); p.font.color.rgb = DARK; p.space_after = Pt(4)
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(4.9),
                               Inches(8.0), Inches(1.2)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Order matters — each layer exposes the next bottleneck."
    p.font.size = Pt(14); p.font.italic = True; p.font.color.rgb = MUTED
    add_footer(s, 29); return s

def slide_30_tail():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Tail Latency & KV Footprint")
    data = [
        ["Batch", "p50 (ms)", "p90 (ms)", "p99 (ms)", "Tok/s"],
        ["1",   "18.2", "18.4", "18.9",   "55"],
        ["8",   "19.3", "20.1", "20.9",  "414"],
        ["16",  "20.1", "21.3", "22.6",  "796"],
        ["32",  "21.8", "23.5", "25.4", "1,467"],
        ["64",  "23.5", "25.7", "28.4", "2,724"],
    ]
    add_table(s, data, 1.0, 1.3, 11.0, 2.8, first_col_bold=True, font_size=13)
    add_bullets(s, [
        ("p99 < 30 ms even at batch 64 — well below 100 ms SLO.", True),
        ("KV footprint @ L=8192, batch 64: 34 GB → 19 GB (44% saving) via paged + radix-tree sharing.", True),
        "Enables 1.7× more concurrent requests per GPU.",
    ], top=4.4, size=16)
    add_footer(s, 30); return s

def slide_31_prefill_decode():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Prefill vs Decode — Why Both Matter")
    tb1 = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                               Inches(6), Inches(2.6)).text_frame
    p = tb1.paragraphs[0]; p.text = "Prefill — compute-bound"
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = ACCENT
    for l in [
        "Dense QKᵀ and PV on whole prompt.",
        "SWFA cuts attention FLOPs by ~37% at w=4096 on 8K prompts.",
        "Direct wall-clock win.",
    ]:
        p = tb1.add_paragraph(); p.text = "• " + l
        p.font.size = Pt(15); p.font.color.rgb = DARK
    tb2 = s.shapes.add_textbox(Inches(6.8), Inches(1.3),
                               Inches(6), Inches(2.6)).text_frame
    p = tb2.paragraphs[0]; p.text = "Decode — memory-bound"
    p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = ACCENT
    for l in [
        "Arithmetic intensity ≈ 1 FLOP/byte.",
        "Fused kernels hide KV-page indirection behind compute.",
        "Throughput ceiling = HBM bandwidth, not FLOPs.",
    ]:
        p = tb2.add_paragraph(); p.text = "• " + l
        p.font.size = Pt(15); p.font.color.rgb = DARK
    tb3 = s.shapes.add_textbox(Inches(0.6), Inches(4.8),
                               Inches(12), Inches(1.8)).text_frame
    p = tb3.paragraphs[0]
    p.text = "The same building blocks appear in both phases, but the bottleneck changes."
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    p2 = tb3.add_paragraph()
    p2.text = "Kernel design must follow the bottleneck, not the algorithm."
    p2.font.size = Pt(16); p2.font.color.rgb = DARK
    add_footer(s, 31); return s

def slide_32_triton():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Triton — Programming Model & Auto-tune")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2),
                              Inches(12), Inches(0.6)).text_frame
    p = tb.paragraphs[0]
    p.text = "Programmer writes block-level Python; compiler emits PTX with Tensor-Core MMAs."
    p.font.size = Pt(15); p.font.color.rgb = DARK
    code = s.shapes.add_textbox(Inches(0.6), Inches(1.8),
                                Inches(12), Inches(4.2))
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
        p.font.name = "Consolas"; p.font.size = Pt(14)
        p.font.color.rgb = DARK
    tb2 = s.shapes.add_textbox(Inches(0.6), Inches(6.1),
                               Inches(12), Inches(0.8)).text_frame
    p = tb2.paragraphs[0]
    p.text = "Auto-tuner picks the winner per shape class and caches it — second call is free."
    p.font.size = Pt(15); p.font.italic = True; p.font.color.rgb = MUTED
    add_footer(s, 32); return s

def slide_33_roofline():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Roofline Analysis on A100")
    s.shapes.add_picture(roofline_png, Inches(2.5), Inches(1.3),
                         height=Inches(4.3))
    tb = s.shapes.add_textbox(Inches(0.6), Inches(5.8),
                              Inches(12), Inches(1.3)).text_frame
    p = tb.paragraphs[0]
    p.text = "GEMM and SWFA sit on the compute ceiling (312 TFLOPS); decode attention sits on the memory ridge at ~1 FLOP/byte."
    p.font.size = Pt(15); p.font.color.rgb = DARK
    p2 = tb.add_paragraph()
    p2.text = "Two kernels, two ceilings, two optimisation strategies."
    p2.font.size = Pt(15); p2.font.bold = True; p2.font.color.rgb = ACCENT
    add_footer(s, 33); return s

def slide_34_triton_table():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Triton vs CUTLASS vs Torch-inductor (FP16, TFLOPS)")
    data = [
        ["Kernel",                                "Triton", "CUTLASS", "Inductor"],
        ["GEMM 4096³",                              "294",   "301",     "288"],
        ["Flash Attention (N=2K, h=128)",           "172",   "189",     "148"],
        ["SWFA (N=2K, w=512)",                      "241",   "258",     "—"],
        ["Decode attention (KV=8K)",                 "84",    "92",      "71"],
    ]
    add_table(s, data, 0.8, 1.3, 11.7, 3.0, first_col_bold=True, font_size=14)
    add_bullets(s, [
        "Agreement: <3% throughput deviation; <1e-4 FP32, <1e-3 FP16 numerical difference.",
        ("Triton closes the gap to within 7% of hand-tuned CUTLASS on SWFA.", True),
        "Torch-inductor cannot express sliding window — hence the missing entry.",
    ], top=4.7, size=15)
    tb = s.shapes.add_textbox(Inches(0.6), Inches(6.5),
                              Inches(12), Inches(0.5)).text_frame
    p = tb.paragraphs[0]
    p.text = "Triton for iteration speed; CUTLASS for the last 5–10%."
    p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = ACCENT
    add_footer(s, 34); return s

def slide_35_lessons():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Lessons Learned")
    add_bullets(s, [
        ("Compose, then tune: cross-layer co-design compounds — single-layer wins plateau quickly.", True),
        ("Skip beats mask: block-sparsity should be encoded in kernel control flow, not in the score matrix.", True),
        "Overlap is a first-class metric: a 65% overlap run beats a 90% util run that stalls on comms.",
        "Auto-tune everything: the same kernel needs different tile shapes at different sequence lengths.",
        "Roofline is the design tool: every optimisation should move a dot toward its ceiling.",
    ], top=1.5, size=18)
    add_footer(s, 35); return s

def slide_36_limits():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Limitations")
    add_bullets(s, [
        "Single node: 8-GPU NVLink only; no multi-node cross-switch, no InfiniBand topology.",
        "Model size: GPT-2 Medium (355M) for training; 7B only for serving. No 70B/175B.",
        "Precision: FP16/FP32 only; no sub-8-bit quantisation study.",
        "SWFA: reference PyTorch + Triton cross-validator; CUTLASS kernel sketched but not shipped.",
        "Workload breadth: dense transformers only — no MoE end-to-end, no RLHF.",
    ], top=1.5, size=18)
    add_footer(s, 36); return s

def slide_37_future():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Future Work")
    add_bullets(s, [
        "Multi-node SWFA with topology-aware window partitioning over InfiniBand.",
        "CUTLASS SWFA kernel with CuTe layouts and epilogue fusion; target >95% TC util.",
        "Adaptive window: learn w* online per layer/head using the closed-form as warm start.",
        "Adaptive KV page size in SGLang — balance arithmetic intensity against fragmentation.",
        "Sub-8-bit paths: FP8 and INT4 kernels with accuracy bounds from the online-softmax FP model.",
        "MoE end-to-end: fold grouped-GEMM epilogue fusion into the SGLang runtime.",
    ], top=1.4, size=17)
    add_footer(s, 37); return s

def slide_38_summary():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Summary of Contributions")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(5.8)).text_frame
    for hdr, items in [
        ("Algorithmic", [
            "SWFA with corner-check classifier and vectorised mask.",
            "Closed-form optimal window w* under exponential decay; strict convexity proved.",
        ]),
        ("Systems", [
            "Controlled 4-way distributed study: ZeRO-3 peaks at 93.8% GPU util, 22% lower energy.",
            "SGLang kernel stack: 2.84× decode on LLaMA-7B.",
            "CUTLASS/CuTe mapping: ≤35% speedup on skinny and MoE GEMMs.",
        ]),
        ("Reproducibility", [
            "Reference PyTorch SWFA + Triton cross-validator.",
            "19 accuracy configurations, all pass within tolerance.",
            "Full proofs of block-count, convexity and w* in Appendix B.",
        ]),
    ]:
        p = tb.add_paragraph() if tb.paragraphs[0].text else tb.paragraphs[0]
        p.text = hdr
        p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = ACCENT
        p.space_after = Pt(2); p.space_before = Pt(6)
        for item in items:
            p = tb.add_paragraph(); p.text = "• " + item
            p.font.size = Pt(14); p.font.color.rgb = DARK
    add_footer(s, 38); return s

def slide_39_pubs():
    s = prs.slides.add_slide(BLANK)
    add_title(s, "Publications & Reproducibility")
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.3),
                              Inches(12), Inches(5.5)).text_frame
    p = tb.paragraphs[0]; p.text = "Artefacts in progress:"
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    for x in [
        "“Sliding Window Flash Attention with Block Classification” — preprint in preparation.",
        "Open-source reference PyTorch + Triton implementation of SWFA.",
        "Benchmark harness for 4-way distributed comparison (ResNet-50 ∪ GPT-2 Medium).",
    ]:
        p = tb.add_paragraph(); p.text = "• " + x
        p.font.size = Pt(14); p.font.color.rgb = DARK
    p = tb.add_paragraph(); p.text = ""; p.font.size = Pt(8)
    p = tb.add_paragraph(); p.text = "Reproducibility:"
    p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = ACCENT
    for x in [
        "All random seeds logged; NCCL topology captured; Nsight reports archived.",
        "Shape-indexed auto-tune cache shipped with kernels.",
        "Triton pipeline reproduces CUTLASS numbers within <3%.",
    ]:
        p = tb.add_paragraph(); p.text = "• " + x
        p.font.size = Pt(14); p.font.color.rgb = DARK
    p = tb.add_paragraph(); p.text = ""; p.font.size = Pt(6)
    p = tb.add_paragraph()
    p.text = "Environment: CUDA 12.x · NCCL 2.19 · PyTorch 2.0 · CUTLASS 3.x · Triton 2.x · A100-40GB × 8."
    p.font.size = Pt(13); p.font.italic = True; p.font.color.rgb = MUTED
    add_footer(s, 39); return s

def slide_40_thanks():
    s = prs.slides.add_slide(BLANK)
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.5),
                              prs.slide_width, Inches(2.5))
    band.fill.solid(); band.fill.fore_color.rgb = RGBColor(0x33, 0x4F, 0x7A)
    band.line.fill.background()
    tb = s.shapes.add_textbox(Inches(0.5), Inches(2.7),
                              Inches(12.3), Inches(2.0))
    tf = tb.text_frame
    p = tf.paragraphs[0]; p.text = "Thank you"
    p.font.size = Pt(56); p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph(); p2.text = "Questions?"
    p2.font.size = Pt(28); p2.font.italic = True
    p2.font.color.rgb = RGBColor(0xFF, 0xEE, 0xBB)
    p2.alignment = PP_ALIGN.CENTER
    tb2 = s.shapes.add_textbox(Inches(0.5), Inches(5.5),
                               Inches(12.3), Inches(2.0))
    tf2 = tb2.text_frame
    for i, line in enumerate([
        "Adarsh Dubey — SC24M151",
        "Department of Mathematics, IIST · Advisor: Dr. S. Sumitra",
        "adarsh.dubey.sc24@iist.ac.in",
    ]):
        p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(18 if i == 0 else 14)
        p.font.bold = (i == 0)
        p.font.color.rgb = DARK
        p.alignment = PP_ALIGN.CENTER
    return s

# Build all slides in order
for fn in [
    slide_01_title, slide_02_glance, slide_03_problem, slide_04_stack,
    slide_05_contributions, slide_06_roadmap, slide_07_priorart,
    slide_08_position, slide_09_fourstrat, slide_10_memmath,
    slide_11_method, slide_12_resultsutil, slide_13_energy,
    slide_14_cutlass, slide_15_epilogue, slide_16_cutlass_res,
    slide_17_swfa_motiv, slide_18_onlinesoftmax, slide_19_classify_fig,
    slide_20_classify_algo, slide_21_corner_lemma, slide_22_vecmask,
    slide_23_fwd_algo, slide_24_backward_complexity,
    slide_25_accuracy_skip, slide_26_e2e_scaling,
    slide_27_optimal_window, slide_28_sglang, slide_29_layered,
    slide_30_tail, slide_31_prefill_decode, slide_32_triton,
    slide_33_roofline, slide_34_triton_table, slide_35_lessons,
    slide_36_limits, slide_37_future, slide_38_summary,
    slide_39_pubs, slide_40_thanks,
]:
    fn()

prs.save(OUT)
print(f"Wrote {OUT}  ({len(prs.slides)} slides)")
