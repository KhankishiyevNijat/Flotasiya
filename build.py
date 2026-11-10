#!/usr/bin/env python3
"""Builds the installable app (docs/index.html) from flotasiya-kalkulyator.html."""
from pathlib import Path

root = Path(__file__).parent
src = (root / "flotasiya-kalkulyator.html").read_text(encoding="utf-8")
marker = '<div class="wrap">'
head, body = src.split(marker, 1)

page = f"""<!doctype html>
<html lang="az">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#0f6e5c">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Flotasiya">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icon-192.png">
<link rel="apple-touch-icon" href="apple-touch-icon.png">
<style>
:root {{ color-scheme: light; padding: env(safe-area-inset-top, 0px) 0 env(safe-area-inset-bottom, 0px); }}
body {{ margin: 0; font: 14px system-ui, sans-serif; }}
img {{ max-width: 100%; }}
[hidden] {{ display: none !important; }}
</style>
{head.strip()}
</head>
<body>
{marker}{body.rstrip()}
<script>
if ("serviceWorker" in navigator) {{
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {{}}));
}}
</script>
</body>
</html>
"""
(root / "docs" / "index.html").write_text(page, encoding="utf-8")
print("docs/index.html written,", len(page), "bytes")
