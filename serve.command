#!/bin/bash
cd "$(dirname "$0")"
echo "Todo app running at http://localhost:8765/todo.html"
echo "Keep this window open. Close it to stop."
python3 -m http.server 8765 2>/dev/null || python -m SimpleHTTPServer 8765
