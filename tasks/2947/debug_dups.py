import importlib.util
import sys

spec = importlib.util.spec_from_file_location('mc', r'clawhunt/pipeline/tasks/2947/make_content.py')
m = importlib.util.module_from_spec(spec)
sys.modules['mc'] = m
try:
    spec.loader.exec_module(m)
except AssertionError:
    pass
names = {}
for r in m.rows[7:]:
    for n in (r[6], r[7]):
        names.setdefault(n, 0)
        names[n] += 1
dups = {k: v for k, v in names.items() if v > 1 and k != '—'}
print('dups:', dups)
print('leaves:', len(m.rows) - 7)
