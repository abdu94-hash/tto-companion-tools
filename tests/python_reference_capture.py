"""Capture the values computed by the independent Python reference check (tto_calc_reference_check_v1.1.py for CALC-SPEC 1.1)
for every numeric vector, so tests/run_vectors.js can confirm that the JavaScript engine and the
reference agree value by value, not only on PASS/FAIL.

Usage: python python_reference_capture.py <tto_calc_reference_check_v1.1.py> <tto_calc_test_vectors_v1.1.json>
Prints JSON: {"V01": [[got, want], ...], ...} in the order the reference compares them.
The reference script itself is not modified; its comparison function is wrapped while it runs.
"""
import importlib.util
import json
import sys


def main(ref_path, vectors_path):
    spec = importlib.util.spec_from_file_location('tto_ref', ref_path)
    ref = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ref)
    original_close = ref.close
    state = {'depth': 0, 'pairs': []}

    def recording_close(a, b, tol):
        top = state['depth'] == 0
        state['depth'] += 1
        try:
            ok = original_close(a, b, tol)
        finally:
            state['depth'] -= 1
        if top:
            state['pairs'].append([a, b])
        return ok

    ref.close = recording_close
    data = json.load(open(vectors_path, encoding='utf-8'))
    out = {}
    for v in data['vectors']:
        if v['type'] == 'behavioral':
            continue
        state['pairs'] = []
        problems = ref.run_vector(v, data['tolerance'])
        out[v['id']] = {'pairs': state['pairs'], 'problems': problems}
    print(json.dumps(out))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
