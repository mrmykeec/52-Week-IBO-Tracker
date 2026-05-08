"""Verify linkifyTaskText logic against curriculum text and tricky edge cases.
Uses Python regex which is functionally equivalent to JS for these patterns."""

import re
import sys

TASK_LINK_URLS = {
    'championleadership1.com': 'https://www.championleadership1.com/',
    'weareamway.com':          'https://www.weareamway.com/',
    'ibofacts.com':            'https://www.ibofacts.com/',
    'iboai.com':               'https://www.iboai.com/',
    'amway.com':               'https://www.amway.com/',
}
TASK_LINK_REGEX = re.compile(
    r'\b(?:championleadership1\.com|weareamway\.com|ibofacts\.com|iboai\.com|amway\.com)\b'
)

def linkify(text):
    return TASK_LINK_REGEX.sub(
        lambda m: f'<a href="{TASK_LINK_URLS[m.group(0)]}" target="_blank" rel="noopener" onclick="event.stopPropagation()">{m.group(0)}</a>',
        text
    )

tests = [
    ('Standard single URL',
     'Watch: "Activating New IBOs" — championleadership1.com › Training Videos (Password: vision3)',
     ['championleadership1.com']),
    ('wwdb.com no longer linked (app-first)',
     'Listen: "Howie Danzik Wrap-Up" — WWG App / wwdb.com › WWG Store › Search "151"',
     []),
    ('amway.com after slash',
     'Watch: "Customer Sales Incentive (CSI)" — Amway App / amway.com › Education › Search',
     ['amway.com']),
    ('Three URLs comma-separated (trickiest)',
     'Learn about Amway — weareamway.com, iboai.com, ibofacts.com',
     ['weareamway.com', 'iboai.com', 'ibofacts.com']),
    ('No URL — should not change',
     'Clear Topics Daily',
     []),
    ('Word boundary: amway.com inside weareamway.com should NOT double-match',
     'Visit weareamway.com only',
     ['weareamway.com']),
    ('Word boundary: amway.com adjacent to weareamway.com',
     'Compare amway.com vs weareamway.com pages',
     ['amway.com', 'weareamway.com']),
    ('Same URL twice',
     'See championleadership1.com or championleadership1.com again',
     ['championleadership1.com', 'championleadership1.com']),
    ('URL at end of sentence',
     'Open amway.com.',
     ['amway.com']),
    ('URL inside parentheses',
     '(see amway.com)',
     ['amway.com']),
]

passed, failed = 0, 0
for name, text, expected in tests:
    out = linkify(text)
    actual = re.findall(r'>([^<]+)</a>', out)
    if actual == expected:
        passed += 1
        print(f"PASS  {name:55} -> {len(actual)} link(s)")
    else:
        failed += 1
        print(f"FAIL  {name}")
        print(f"      Input:    {text!r}")
        print(f"      Expected: {expected}")
        print(f"      Got:      {actual}")
        print(f"      Output:   {out!r}")

# Real curriculum sample: pull a few task text strings from the file and dump linked output
print("\n--- Real curriculum samples (first 5 with URLs) ---")
with open(r'C:\Users\mrmyk\Desktop\My First Claude Folder\52-Week Training\52-week-tracker.html', encoding='utf-8') as f:
    content = f.read()
task_re = re.compile(r"text:\s*['\"](.*?)['\"]\s*[},]", re.DOTALL)
samples_with_urls = []
for m in task_re.finditer(content):
    t = m.group(1)
    if TASK_LINK_REGEX.search(t):
        samples_with_urls.append(t)
        if len(samples_with_urls) >= 5:
            break

for i, t in enumerate(samples_with_urls, 1):
    print(f"\n[{i}] BEFORE:  {t}")
    print(f"    AFTER:   {linkify(t)}")

# Stats: how many tasks contain at least one URL?
all_tasks = task_re.findall(content)
with_url = sum(1 for t in all_tasks if TASK_LINK_REGEX.search(t))
print(f"\n--- Stats ---")
print(f"Total task entries scanned: {len(all_tasks)}")
print(f"Tasks containing a linkable URL: {with_url}")
print(f"\nResults: {passed} passed, {failed} failed")
sys.exit(0 if failed == 0 else 1)
