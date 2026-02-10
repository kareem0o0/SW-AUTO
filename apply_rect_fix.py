from pathlib import Path
path = Path('SwAutomationApp/test1.cs')
text = path.read_text()
old = """            swSketchManager.FullyDefineSketch(true, true,\n                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Horizontal |\n                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Vertical,\n                true, 1, null, 1, null, 1, 1);\n"""
new = """            swSketchManager.FullyDefineSketch(true, true,\n                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Horizontal |\n                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Vertical |\n                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Distance,\n                true, 1, null, 1, null, 1, 1);\n"""
if old not in text:
    raise SystemExit('pattern not found')
path.write_text(text.replace(old, new, 1))
