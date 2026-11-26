#!/bin/bash
# Usage: ./deploy.sh "commit message"   (run with github credentials)
set -e
cd /workspace/work/site
B=/workspace/built/sites/a568459b-9dc5-4005-b45b-39a4cb2d317c
cp $B/index.html $B/hark.css $B/hark-hero.css $B/hark.js .
python3 - <<'PY'
p="index.html"; s=open(p).read()
head='    <link rel="manifest" href="manifest.webmanifest" />\n  <link rel="apple-touch-icon" sizes="180x180" href="icon-180.png" />\n'
if 'manifest.webmanifest' not in s: s=s.replace('</head>', head+'</head>',1)
if 'native-bridge.js' not in s: s=s.replace('<head>', '<head>\n  <script src="native-bridge.js"></script>',1)
reg=open("sw-register.html").read()
i=s.rfind('</body>'); s=s[:i]+reg+s[i:]
open(p,"w").write(s)
PY
V=$(date -u +%Y%m%d%H%M%S)
sed "s/__VERSION__/$V/" sw.template.js > sw.js
git add -A
git -c user.name=shinesdags-bot -c user.email=338882197+shinesdags-bot@users.noreply.github.com commit -q -m "$1"
git pull -q --rebase origin main
git push -q origin main
echo "deployed $V $(git log --oneline -1)"
