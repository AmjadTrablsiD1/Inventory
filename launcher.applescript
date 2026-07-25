-- Inventory Manager launcher (starts the app with no Terminal window)
--
-- Build it into a double-clickable app with:
--     ./build_launcher.sh
-- which produces "Inventory Manager.app" next to app.py.
--
-- Double-clicking the built app starts app.py in the background; app.py then
-- opens your browser. Clicking it again while running just reopens the browser.
on run
	set appPosix to POSIX path of (path to me)
	set projectDir to do shell script "d=" & quoted form of appPosix & "; d=\"${d%/}\"; dirname \"$d\""

	set sh to "cd " & quoted form of projectDir & " || exit 1

# Find a usable python3
PY=''
for c in python3 /usr/bin/python3 /usr/local/bin/python3 /opt/homebrew/bin/python3; do
  if command -v \"$c\" >/dev/null 2>&1; then PY=\"$c\"; break; fi
done
if [ -z \"$PY\" ]; then
  osascript -e 'display alert \"Python 3 not found\" message \"Install Python 3 from python.org, then try again.\"'
  exit 1
fi

# Make sure Flask is available
if ! \"$PY\" -c 'import flask' >/dev/null 2>&1; then
  \"$PY\" -m pip install --user flask >/tmp/inventory_pip.log 2>&1 || \"$PY\" -m pip install flask >/tmp/inventory_pip.log 2>&1
fi

# If it's already running, just open the browser instead of starting a 2nd copy
if curl -s -o /dev/null http://127.0.0.1:8765/api/state; then
  open http://127.0.0.1:8765
  exit 0
fi

nohup \"$PY\" app.py >/tmp/inventory_manager.log 2>&1 </dev/null &"

	do shell script sh
end run
