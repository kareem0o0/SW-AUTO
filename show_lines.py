from pathlib import Path
lines = Path('SwAutomationApp/test1.cs').read_text().splitlines()
for idx,line in enumerate(lines):
    if 'FullyDefineSketch' in line:
        for j in range(idx-2, min(idx+3, len(lines))):
            print(f"{j+1:04d}: {lines[j]}")
        break
else:
    print('no match')
