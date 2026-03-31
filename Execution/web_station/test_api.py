import requests, json

resp = requests.post(
    'http://localhost:5000/api/append-build',
    json={'build_path': 'c:/Project/DoU/Execution', 'build_label': 'TEST_DRY'}
)
d = resp.json()
print('success:', d.get('success'))

if d.get('steps'):
    for s in d['steps']:
        print(f"  Step {s['id']}: [{s['status']}] {s['name']}")
        if s.get('detail'):
            print(f"    -> {s['detail']}")

if not d.get('success'):
    print('error:', d.get('error'))
else:
    r = d['result']
    print(f"\nDoU Current : {r['dou_current']}")
    print(f"Proj DoU    : {r['proj_dou']}")
    print(f"Delta Proj  : {r['dou_delta_proj']}")
    print(f"Comparison rows: {len(d['comparison'])}")
    print("\nFirst comparison row:")
    print(json.dumps(d['comparison'][0], indent=2))
    print("\nDoU Current row:")
    dou_row = next((x for x in d['comparison'] if x['is_dou']), None)
    if dou_row:
        print(json.dumps(dou_row, indent=2))