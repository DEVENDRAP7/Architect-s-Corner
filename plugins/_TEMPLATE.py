# Plugin template for ArchTools.  Copy to plugins/<command>.py (no leading _).
# Files starting with "_" are ignored by the loader, so this template is safe.
#
# A plugin adds ONE new command. It automatically gets --building / --all and
# (if a building stores one) the right sheet bounding box, exactly like the
# built-in DXF commands. `args.file` is already resolved to the .dxf path.
#
# `ctx` hands you the core helpers so you don't re-import anything:
#   ctx.load_dxf(path) -> ezdxf document     ctx.fmt_table(rows, headers) -> str
#   ctx.all_text_entities(msp) -> (text, x, y, layer)
#   ctx.entity_text(e)         ctx.Counter   ctx.defaultdict
#   ctx.ezdxf   ctx.re   ctx.math   ctx.parse_levels   ctx.open_workbook

META = {
    "command": "example",                 # what the user types
    "summary": "one-line description",     # shown in `tools` and to the AI
    "args": [                              # extra options beyond --building/--all
        {"name": "--layer", "type": "str", "required": False,
         "help": "an optional layer name"},
        # types: "str" | "int" | "float" | "flag"   (flag = on/off switch)
    ],
}


def run(args, ctx):
    doc = ctx.load_dxf(args.file)          # args.file already resolved
    msp = doc.modelspace()
    n = sum(1 for _ in msp)
    print(f"{n} entities in modelspace.")
