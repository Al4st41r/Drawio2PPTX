import inspect
import pptx.dml.line
from pptx.dml.line import LineFormat

print("LineFormat members:")
for m in dir(LineFormat):
    if not m.startswith('_'):
        print(m)

# Check if we can find MSO_LINE_END anywhere
import pptx.enum.dml
print("\npptx.enum.dml members:")
print(dir(pptx.enum.dml))
