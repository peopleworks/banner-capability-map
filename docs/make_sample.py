"""Render the sample chart for the README, from invented data.

The numbers below are FICTION. They belong to a college that does not exist, and they
were deliberately written to differ from any real institution's results. No database is
touched and none is needed: this imports the chart renderer and hands it made-up rows.

The generator lives in the repository on purpose. A screenshot in a README is a claim,
and a reader should be able to check what it is a picture of. Run it yourself:

    python docs/make_sample.py

It writes docs/sample-chart.svg. The PNG in the README is that SVG, rendered.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from bcm_report import _chart_svg  # noqa: E402

esc = lambda s: (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

#  A fictional college. name, verdict, people who can reach it, records held.
#  The shape is what matters: a healthy diagonal, and a crowded strip along the
#  bottom where the keys were handed out and nothing was ever written.
SAMPLE = [
    ("Requisitions / e-Procurement",   "In use",          388,   214_770),
    ("Time Entry",                     "In use",          402, 1_866_004),
    ("Mass Communication",             "In use",           61, 4_412_889),
    ("Population Selection",           "In use",          177,   903_512),
    ("Personnel e-Forms",              "In use",           44,     6_120),
    ("Benefits Enrollment",            "In use",           37,    18_455),
    ("Grants / Research Accounting",   "In use",          126,     1_408),
    ("Event Management",               "Abandoned",       144,        23),
    ("Workflow Engine",                "Owned, not used",  11,         1),
    ("Fixed Assets",                   "Owned, not used",  57,         0),
    ("Recruitment / Applicant Track",  "Owned, not used",  33,         0),
    ("Skills Inventory",               "Owned, not used",  48,         0),
    ("Effort Certification",           "Owned, not used",  26,         0),
    ("Budget Planner",                 "Custom-built",     19,         0),
    ("Residence Life",                 "Expected absent",  40,         0),
]

data = [{"e": {"name": n}, "verdict": v, "access": a, "rows": r}
        for n, v, a, r in SAMPLE]

#  In the report the chart inherits its styling from the page. Standing alone it has
#  none, and an SVG line with no stroke simply does not draw, so the rules travel with
#  it. Same values as the report's stylesheet, plus a white background: a transparent
#  chart is unreadable against a dark README.
STYLE = """<style>
 .chart-bg{fill:#ffffff}
 .cgrid{stroke:#eef1f4;stroke-width:1}
 .cgridv{stroke:#f4f6f8;stroke-width:1}
 .cax{font:10.5px -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;fill:#9aa5b1}
 .cax.ct{font-size:11px;fill:#5B6B7C;font-weight:600}
 .cax.cz{fill:#B03A2E;font-weight:700}
 .clbl{font:10.5px -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-weight:600}
 .clead{stroke:#c3ccd6;stroke-width:.8}
 .ttl{font:700 15px -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;fill:#1F3864}
</style>"""

svg = _chart_svg(data, esc)
svg = svg[svg.index("<svg"):svg.index("</svg>") + 6]
head = svg[:svg.index(">") + 1]
body = svg[svg.index(">") + 1:]
svg = (head + STYLE
       + '<rect class="chart-bg" x="0" y="0" width="100%" height="100%"/>'
       + '<text x="40" y="26" class="ttl">Keys handed out, against what is inside '
         'the room</text>'
       + body)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample-chart.svg")
with open(out, "w", encoding="utf-8") as f:
    f.write(svg)
print(f"wrote {out}   ({len(SAMPLE)} invented capabilities)")
