import sys, json
sys.stdout.reconfigure(encoding='utf-8')
with open('section_mapping_ch7.json', 'r', encoding='utf-8') as f:
    data = json.load(f)
    print('Mappings:')
    for m in data['mappings']:
        sr = m.get('slide_range', '?')
        wh = m['word_heading'][:40]
        fst = m.get('first_slide_title', '')[:30]
        print(f"  Slides {sr}: {fst} <-> {wh}")
    print(f"\nChanges: {len(data['changes'])}")
    for c in data['changes']:
        st = c.get('slide_title','')[:30]
        wh = c['word_heading'][:30]
        print(f"  {c['tag']} - {st}... -> {wh}...")
