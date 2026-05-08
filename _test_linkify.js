// Smoke test for linkifyTaskText against actual task text strings.
// Runs the real regex/replacement and verifies expected behavior on tricky cases.

const TASK_LINK_URLS = {
  'championleadership1.com': 'https://www.championleadership1.com/',
  'weareamway.com':          'https://www.weareamway.com/',
  'ibofacts.com':            'https://www.ibofacts.com/',
  'iboai.com':               'https://www.iboai.com/',
  'amway.com':               'https://www.amway.com/'
};
const TASK_LINK_REGEX = /\b(?:championleadership1\.com|weareamway\.com|ibofacts\.com|iboai\.com|amway\.com)\b/g;

function linkifyTaskText(text) {
  return text.replace(TASK_LINK_REGEX, (m) =>
    `<a href="${TASK_LINK_URLS[m]}" target="_blank" rel="noopener" onclick="event.stopPropagation()">${m}</a>`
  );
}

const tests = [
  // Standard single URL
  { text: 'Watch: "Activating New IBOs" — championleadership1.com › Training Videos (Password: vision3)',
    expectedLinks: ['championleadership1.com'] },
  // wwdb.com is no longer linked (app-first; web search returns no results for audio numbers)
  { text: 'Listen: "Howie Danzik Wrap-Up" — WWG App / wwdb.com › WWG Store › Search "151"',
    expectedLinks: [] },
  // amway.com after slash
  { text: 'Watch: "Customer Sales Incentive (CSI)" — Amway App / amway.com › Education › Search',
    expectedLinks: ['amway.com'] },
  // Multiple URLs in one string (the trickiest case)
  { text: 'Learn about Amway — weareamway.com, iboai.com, ibofacts.com',
    expectedLinks: ['weareamway.com', 'iboai.com', 'ibofacts.com'] },
  // No match: should return unchanged
  { text: 'Clear Topics Daily',
    expectedLinks: [] },
  // Word-boundary check: "amway.com" inside "weareamway.com" should NOT match separately
  { text: 'Visit weareamway.com only',
    expectedLinks: ['weareamway.com'] },
  // Multiple of same URL in one task
  { text: 'See championleadership1.com or championleadership1.com again',
    expectedLinks: ['championleadership1.com', 'championleadership1.com'] },
];

let pass = 0, fail = 0;
for (const t of tests) {
  const out = linkifyTaskText(t.text);
  const matches = (out.match(/<a href="([^"]+)"[^>]*>([^<]+)<\/a>/g) || []);
  const linked = matches.map(m => m.match(/>([^<]+)<\/a>/)[1]);
  const expected = t.expectedLinks;
  const ok = linked.length === expected.length && linked.every((v, i) => v === expected[i]);
  if (ok) {
    pass++;
    console.log(`✓ "${t.text.slice(0, 50)}${t.text.length > 50 ? '...' : ''}" → ${linked.length} link(s)`);
  } else {
    fail++;
    console.log(`✗ FAIL: "${t.text}"`);
    console.log(`   Expected: [${expected.join(', ')}]`);
    console.log(`   Got:      [${linked.join(', ')}]`);
    console.log(`   Output:   ${out}`);
  }
}

// Edge case: verify generated HTML structure is valid
console.log('\n--- Sample output ---');
console.log(linkifyTaskText('Visit amway.com and weareamway.com'));

console.log(`\nResults: ${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);
