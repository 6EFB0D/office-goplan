#!/usr/bin/env python3
import json
import re
import urllib.parse
import urllib.request

url = "https://docs.google.com/forms/d/1NpXzk1kyUn2LhUzQhhMHq_tnT1oOGAsv561L-7nMfos/viewform"
html = urllib.request.urlopen(url).read().decode("utf-8")
m = re.search(r"FB_PUBLIC_LOAD_DATA_ = (.+?);\s*</script>", html)
data = json.loads(m.group(1))
qs = data[1][1]
print("desc:", data[1][0])
print("title:", data[3] if isinstance(data[3], str) else data[1][8])
for q in qs:
    print("---")
    print("title:", q[1], "type:", q[3])
    entry = q[4][0][0]
    print("entry:", entry)
    opts_node = q[4][0]
    if len(opts_node) > 1 and isinstance(opts_node[1], list):
        print("options:", [o[0] for o in opts_node[1]])

# build sample prefill
form_action = re.search(r'action="([^"]+/formResponse)"', html)
print("action:", form_action.group(1) if form_action else None)
eid = re.search(r"/forms/d/e/(1FAIpQLS[^/\"]+)/", html)
print("eid:", eid.group(1) if eid else None)
