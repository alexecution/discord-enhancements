# Discord Enhancements Add-on for NVDA
# Command handler implementations
#
# Each public function corresponds to one layer command.
# Handlers use the UIA helpers from uia.py for element access and
# report results via ui.message().  All handlers are COMError-safe.
# Tree walking only happens when a command is explicitly invoked —
# NEVER in event handlers or overlay identification.

import time
from comtypes import COMError
from logHandler import log
import api
import controlTypes
import tones
import ui
import speech

from . import uia

import os
import wx


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_app():
	"""Return the Discord appModule instance, or None."""
	try:
		focus = api.getFocusObject()
		if focus and focus.appModule:
			return focus.appModule
	except Exception:
		pass
	return None


# ---------------------------------------------------------------------------
# Internal state for message navigation
# ---------------------------------------------------------------------------

_message_cursor = -1
_last_messages = []
_last_current_msg_time = 0.0


def _refresh_messages():
	"""Refresh the internal message list from UIA."""
	global _last_messages
	_last_messages = uia.get_messages()
	return _last_messages


def _clamp_cursor(index, count):
	if count == 0:
		return -1
	return max(0, min(index, count - 1))


def _speak_message(index, messages=None):
	global _message_cursor, _last_current_msg_time
	if messages is None:
		messages = _last_messages
	if not messages:
		ui.message("No messages available.")
		return
	index = _clamp_cursor(index, len(messages))
	_message_cursor = index
	msg = messages[index]
	content = uia.read_message_content(msg)
	position = "Message %d of %d" % (index + 1, len(messages))
	ui.message("%s. %s" % (position, content))
	_last_current_msg_time = time.time()


def _spell_current_message():
	if not _last_messages or _message_cursor < 0:
		ui.message("No current message.")
		return
	index = _clamp_cursor(_message_cursor, len(_last_messages))
	msg = _last_messages[index]
	content = uia.read_message_content(msg)
	speech.speakSpelling(content)


# ---------------------------------------------------------------------------
# General commands
# ---------------------------------------------------------------------------

def cmd_activeNow():
	"""[ A — Announce the contents of the 'Active Now' section."""
	panel = uia.find_active_now()
	if not panel:
		ui.message("Active Now section not found. "
				   "Try navigating to the Friends page with Control+1 twice.")
		return
	parts = []
	for child in uia._iter_children(panel):
		name = uia.safe_name(child)
		if name:
			parts.append(name)
	if parts:
		ui.message("Active Now: " + ", ".join(parts))
	else:
		ui.message("Active Now section is empty.")


def cmd_voiceServers():
	"""[ V — Report which visible servers have active voice channels."""
	servers = uia.get_server_items()
	if not servers:
		ui.message("No servers found in the server list.")
		return
	active = []
	for srv in servers:
		name = uia.safe_name(srv)
		if not name:
			continue
		participants = uia.get_voice_participants_from_server(srv)
		if participants:
			active.append("%s: %s" % (name, ", ".join(participants)))
	if active:
		ui.message("Voice activity. " + ". ".join(active))
	else:
		# Fallback: check if WE are in a voice channel and report that
		info = uia.get_voice_connection_info()
		if info and info["channel"]:
			ui.message("You are connected to voice in: %s" % info["channel"])
		else:
			ui.message("No voice activity detected on visible servers.")


def cmd_listButtons():
	"""[ B — List buttons for activation (deduplicated)."""
	import wx
	buttons = uia.get_all_buttons()
	if not buttons:
		ui.message("No buttons found.")
		return
	names = [b[0] for b in buttons]

	def _show_dialog():
		dlg = wx.SingleChoiceDialog(
			None,
			"Select a button to activate:",
			"Discord Buttons",
			names,
		)
		if dlg.ShowModal() == wx.ID_OK:
			selection = dlg.GetSelection()
			if 0 <= selection < len(buttons):
				_, btn_obj = buttons[selection]
				try:
					btn_obj.doAction()
					ui.message("Activated: %s" % names[selection])
				except (COMError, Exception):
					log.debugWarning("Error activating button", exc_info=True)
					ui.message("Could not activate the button.")
		dlg.Destroy()

	wx.CallAfter(_show_dialog)


# ---------------------------------------------------------------------------
# Navigation commands
# ---------------------------------------------------------------------------

def cmd_messageInput():
	"""[ E — Move focus to the message input box."""
	edit = uia.find_message_input()
	if edit:
		if uia.focus_element(edit):
			ui.message("Message input")
		else:
			ui.message("Message input found but could not focus it.")
	else:
		ui.message("Message input not found.")


def cmd_serverList():
	"""[ S — Move focus to the server list."""
	srv = uia.find_server_list()
	if srv:
		if uia.focus_element(srv):
			ui.message("Server list")
		else:
			ui.message("Server list found but could not focus it.")
	else:
		ui.message("Server list not found.")


# Navigation area cycling — stored index avoids slow parent-chain walk
_nav_area_index = 0


def cmd_navigateAreas():
	"""[ N — Cycle focus among major Discord areas."""
	global _nav_area_index
	areas = uia.find_all_areas()

	if not areas:
		ui.message("No navigable area found.")
		return

	# Wrap index if areas changed
	if _nav_area_index >= len(areas):
		_nav_area_index = 0

	# Try each area starting from the stored index
	for attempt in range(len(areas)):
		target_idx = (_nav_area_index + attempt) % len(areas)
		name, region = areas[target_idx]
		if uia.focus_element(region):
			ui.message(name)
			# Advance index for next invocation
			_nav_area_index = (target_idx + 1) % len(areas)
			return

	ui.message("No navigable area found.")


def cmd_userArea():
	"""[ U — Move focus to the user area."""
	area = uia.find_user_area()
	if area:
		if uia.focus_element(area):
			ui.message("User area")
		else:
			ui.message("User area found but could not focus it.")
	else:
		ui.message("User area not found.")


# ---------------------------------------------------------------------------
# Chat message navigation
# ---------------------------------------------------------------------------

def cmd_firstMessage():
	"""[ H / [ Home — Jump to the first available message."""
	messages = _refresh_messages()
	if messages:
		_speak_message(0, messages)
	else:
		ui.message("No messages available.")


def cmd_prevMessage():
	"""[ J / [ Left — Previous message."""
	global _message_cursor
	messages = _refresh_messages()
	if not messages:
		ui.message("No messages available.")
		return
	if _message_cursor <= 0:
		tones.beep(200, 40)
		_message_cursor = 0
		_speak_message(0, messages)
	else:
		_speak_message(_message_cursor - 1, messages)


def cmd_currentMessage():
	"""[ K / [ Numpad5 — Read current message. Double-tap to spell."""
	global _last_current_msg_time
	now = time.time()
	messages = _refresh_messages()
	if not messages:
		ui.message("No messages available.")
		return
	if now - _last_current_msg_time < 0.5:
		_spell_current_message()
		return
	if _message_cursor < 0:
		_speak_message(len(messages) - 1, messages)
	else:
		_speak_message(
			_clamp_cursor(_message_cursor, len(messages)),
			messages,
		)


def cmd_nextMessage():
	"""[ L / [ Right — Next message."""
	global _message_cursor
	messages = _refresh_messages()
	if not messages:
		ui.message("No messages available.")
		return
	if _message_cursor >= len(messages) - 1:
		tones.beep(200, 40)
		_message_cursor = len(messages) - 1
		_speak_message(_message_cursor, messages)
	else:
		_speak_message(_message_cursor + 1, messages)


def cmd_lastMessage():
	"""[ ; / [ End — Jump to the last available message."""
	messages = _refresh_messages()
	if messages:
		_speak_message(len(messages) - 1, messages)
	else:
		ui.message("No messages available.")


def cmd_unreadMarker():
	"""[ Shift+H / [ Shift+Home — Jump to the 'Unread' marker."""
	messages = _refresh_messages()
	marker_index = uia.find_unread_marker()
	if marker_index >= 0 and messages:
		target = _clamp_cursor(marker_index, len(messages))
		_speak_message(target, messages)
		ui.message("Unread messages start here.")
	else:
		ui.message("No unread marker found.")


def cmd_focusCurrentMessage():
	"""[ Shift+K — Move real focus to the current message."""
	messages = _refresh_messages()
	if not messages or _message_cursor < 0:
		ui.message("No current message to focus.")
		return
	index = _clamp_cursor(_message_cursor, len(messages))
	msg = messages[index]
	try:
		msg.setFocus()
		ui.message("Focused message %d of %d" % (index + 1, len(messages)))
	except (COMError, Exception):
		log.debugWarning("Error focusing message", exc_info=True)
		ui.message("Could not focus the message.")


# ---------------------------------------------------------------------------
# Digit commands — recent messages
# ---------------------------------------------------------------------------

def cmd_recentMessage(n):
	"""[ 1–0 — Read the Nth most recent message."""
	messages = _refresh_messages()
	if not messages:
		ui.message("No messages available.")
		return
	index = len(messages) - n
	if index < 0:
		ui.message("Only %d messages available." % len(messages))
		return
	_speak_message(index, messages)


# ---------------------------------------------------------------------------
# Voice / call management
# ---------------------------------------------------------------------------

def cmd_disconnect():
	"""[ D — Press the Disconnect button."""
	btn = uia.find_disconnect_button()
	if btn:
		try:
			btn.doAction()
			ui.message("Disconnected.")
		except (COMError, Exception):
			log.debugWarning("Error pressing disconnect", exc_info=True)
			ui.message("Could not disconnect.")
	else:
		ui.message("Disconnect button not found. "
				   "You may not be in a voice channel.")


def cmd_ping():
	"""[ P — Report voice connection ping/latency."""
	info = uia.get_voice_connection_info()
	if info is None:
		ui.message("Voice connection not found. "
				   "You may not be in a voice channel.")
		return

	if info["latency"]:
		ui.message("Ping: %s" % info["latency"])
	else:
		ui.message("Connected, but latency not available.")


# ---------------------------------------------------------------------------
# Chat information commands
# ---------------------------------------------------------------------------

def cmd_typing():
	"""[ T — Announce who is typing and slow-mode status."""
	indicator = uia.find_typing_indicator()
	if indicator:
		name = uia.safe_name(indicator)
		if name:
			ui.message(name)
		else:
			ui.message("Typing indicator visible but content unreadable.")
	else:
		ui.message("No one is typing.")


def cmd_channelInfo():
	"""[ W — Announce DM / channel information."""
	root = uia.get_foreground()
	if not root:
		ui.message("Could not read channel information.")
		return
	context = uia.get_window_context(root)
	parts = []
	if context["channel"]:
		parts.append(context["channel"])
	topic = uia.get_channel_topic(root)
	if topic:
		parts.append("Topic: %s" % topic)
	if context["server"]:
		parts.append("Server: %s" % context["server"])
	if parts:
		ui.message(". ".join(parts))
	else:
		ui.message("Channel information not available.")


# ---------------------------------------------------------------------------
# Pinned messages and threads
# ---------------------------------------------------------------------------

def cmd_pinnedMessages():
	"""[ Shift+P — Press the Pinned Messages button."""
	btn = uia.find_pinned_messages_button()
	if btn:
		try:
			btn.doAction()
			ui.message("Pinned messages")
		except (COMError, Exception):
			log.debugWarning("Error opening pinned messages", exc_info=True)
			ui.message("Could not open pinned messages.")
	else:
		ui.message("Pinned Messages button not found.")


def cmd_threadList():
	"""[ Shift+T — Open or close the thread list."""
	btn = uia.find_threads_button()
	if btn:
		try:
			btn.doAction()
			ui.message("Threads")
		except (COMError, Exception):
			log.debugWarning("Error toggling threads", exc_info=True)
			ui.message("Could not toggle the thread list.")
	else:
		ui.message("Threads button not found.")


# ---------------------------------------------------------------------------
# Diagnostic command
# ---------------------------------------------------------------------------

def cmd_diagnostic():
	"""[ Ctrl+E — Dump the Discord accessibility tree for debugging.

	The output is displayed in a browseable message window and also
	saved to a file on the desktop so it can be copied and shared.
	"""
	import wx
	import os

	ui.message("Scanning Discord accessibility tree, please wait...")
	dump = uia.dump_tree(max_depth=6)

	# Append a message-list diagnostic section
	dump += "\n\n" + "=" * 70 + "\n"
	dump += "MESSAGE LIST DIAGNOSTIC\n"
	dump += "=" * 70 + "\n"
	try:
		msg_list = uia.find_message_list()
		if msg_list:
			ml_name = uia.safe_name(msg_list) or "(no name)"
			ml_role = uia.safe_role(msg_list)
			dump += "find_message_list() returned: name=%r  role=%s\n" % (ml_name, ml_role)
			msgs = uia.get_messages(msg_list)
			dump += "get_messages() found %d items:\n" % len(msgs)
			for i, m in enumerate(msgs[:15]):
				mn = uia.safe_name(m) or "(no name)"
				mr = uia.safe_role(m)
				preview = mn[:120] + "..." if len(mn) > 120 else mn
				dump += "  [%d] role=%s  name=%s\n" % (i, mr, preview)
			if len(msgs) > 15:
				dump += "  ... and %d more\n" % (len(msgs) - 15)
		else:
			dump += "find_message_list() returned None!\n"
			dump += "Candidates in depth-1/2 cache:\n"
			from . import uia as _uia
			patterns = _uia._MESSAGE_NAMES
			for name, child in _uia._get_depth1():
				if name and _uia._name_matches(name, patterns):
					role = _uia.safe_role(child)
					dump += "  D1: name=%r  role=%s\n" % (name, role)
			for name, child, parent in _uia._get_depth2():
				if name and _uia._name_matches(name, patterns):
					role = _uia.safe_role(child)
					dump += "  D2: name=%r  role=%s  parent=%r\n" % (name, role, parent)
	except Exception as e:
		dump += "Error during message diagnostic: %s\n" % e

	# Save to a file on the Desktop as a reliable fallback
	dump_path = ""
	try:
		desktop = os.path.join(os.path.expanduser("~"), "Desktop")
		dump_path = os.path.join(desktop, "discord_tree_dump.txt")
		with open(dump_path, "w", encoding="utf-8") as f:
			f.write(dump)
	except Exception:
		log.debugWarning("Could not write dump file", exc_info=True)

	# Show in a browseable window (must be on GUI thread)
	def _show():
		try:
			ui.browseableMessage(dump, "Discord Accessibility Tree Dump")
		except Exception:
			log.debugWarning("browseableMessage failed", exc_info=True)

	wx.CallAfter(_show)

	if dump_path:
		ui.message("Dump also saved to %s" % dump_path)


def cmd_messageDebug():
	"""[ Ctrl+M — Quick diagnostic showing what find_message_list() finds.

	Much faster than a full tree dump; just reports the message list
	element and the first few messages inside it.  If the message list
	is NOT found, lists all depth-1 and depth-2 elements so we can
	see what Discord exposes.
	"""
	import wx
	import os

	lines = []

	# Always show depth-1 structure first
	lines.append("=== DEPTH-1 ELEMENTS ===")
	for name, child in uia._get_depth1():
		role = uia.safe_role(child)
		lines.append("  D1: name=%s  role=%s" % (repr(name[:80]) if name else '(none)', role))

	lines.append("")
	lines.append("=== DEPTH-2 ELEMENTS ===")
	for gc_name, gc_child, parent_name in uia._get_depth2():
		role = uia.safe_role(gc_child)
		lines.append("  D2: name=%s  role=%s  parent=%s" % (
			repr(gc_name[:60]) if gc_name else '(none)',
			role,
			repr(parent_name[:40]) if parent_name else '(none)',
		))

	lines.append("")
	lines.append("=== MESSAGE LIST SEARCH ===")
	msg_list = uia.find_message_list()
	if not msg_list:
		lines.append("find_message_list() returned: NONE")
		lines.append("No message list could be found.")
	else:
		ml_name = uia.safe_name(msg_list) or "(no name)"
		ml_role = uia.safe_role(msg_list)
		msgs = uia.get_messages(msg_list)

		lines.append("find_message_list() returned: name=%r  role=%s" % (ml_name[:80], ml_role))
		lines.append("%d messages found." % len(msgs))
		# Show first few and last message
		preview_count = min(5, len(msgs))
		for i in range(preview_count):
			mn = uia.safe_name(msgs[i]) or "(no name)"
			mr = uia.safe_role(msgs[i])
			lines.append("  [%d] role=%s  name=%s" % (i + 1, mr, mn[:120]))
		if len(msgs) > 5:
			lines.append("  ...")
			mn = uia.safe_name(msgs[-1]) or "(no name)"
			mr = uia.safe_role(msgs[-1])
			lines.append("  [%d] role=%s  name=%s" % (len(msgs), mr, mn[:120]))

	dump = "\n".join(lines)

	# Save to desktop
	try:
		desktop = os.path.join(os.path.expanduser("~"), "Desktop")
		dump_path = os.path.join(desktop, "discord_msg_debug.txt")
		with open(dump_path, "w", encoding="utf-8") as f:
			f.write(dump)
	except Exception:
		dump_path = ""

	# Show in browseable window
	def _show():
		try:
			ui.browseableMessage(dump, "Message List Diagnostic")
		except Exception:
			pass
	wx.CallAfter(_show)

	if msg_list:
		ml_name = uia.safe_name(msg_list) or "(no name)"
		msgs = uia.get_messages(msg_list)
		ui.message("Message list: %s. %d messages. Saved to desktop." % (ml_name[:60], len(msgs)))
	else:
		ui.message("Message list not found. Diagnostic saved to desktop.")


# ---------------------------------------------------------------------------
# Event logging diagnostic
# ---------------------------------------------------------------------------

def cmd_eventLog():
	"""Start logging accessibility events from Discord for 15 seconds.

	Logs ALL Win32 accessibility events to discord_events.log on the
	Desktop.  This helps identify which events Discord fires when a
	new chat message arrives.
	"""
	import ctypes
	from ctypes import wintypes

	app = _get_app()
	if app is None:
		ui.message("Not in Discord")
		return

	# If already logging, stop
	if getattr(app, '_eventLogActive', False):
		app._eventLogActive = False
		ui.message("Event logging stopped")
		return

	pid = app.processID
	desktop = os.path.join(os.path.expanduser("~"), "Desktop")
	log_path = os.path.join(desktop, "discord_events.log")

	try:
		f = open(log_path, "w", encoding="utf-8")
	except Exception:
		ui.message("Cannot create log file")
		return

	f.write("Discord event log - PID %d\n" % pid)
	f.write("Started at %s\n" % time.strftime("%H:%M:%S"))
	f.write("Send a message in Discord from another device/user and check this file.\n")
	f.write("=" * 70 + "\n\n")
	f.flush()

	# Event name table for readability
	EVENT_NAMES = {
		0x0001: "SYSTEM_SOUND",
		0x0002: "SYSTEM_ALERT",
		0x0003: "SYSTEM_FOREGROUND",
		0x8000: "OBJECT_CREATE",
		0x8001: "OBJECT_DESTROY",
		0x8002: "OBJECT_SHOW",
		0x8003: "OBJECT_HIDE",
		0x8004: "OBJECT_REORDER",
		0x8005: "OBJECT_FOCUS",
		0x8006: "OBJECT_SELECTION",
		0x800A: "OBJECT_STATECHANGE",
		0x800B: "OBJECT_LOCATIONCHANGE",
		0x800C: "OBJECT_NAMECHANGE",
		0x800D: "OBJECT_DESCRIPTIONCHANGE",
		0x800E: "OBJECT_VALUECHANGE",
		0x800F: "OBJECT_PARENTCHANGE",
		0x8013: "OBJECT_INVOKED",
		0x8014: "OBJECT_TEXTSELECTIONCHANGED",
		0x8015: "OBJECT_CONTENTSCROLLED",
		0x8019: "OBJECT_LIVEREGIONCHANGED",
	}

	counter = [0]  # mutable for closure

	# Install a broad WinEvent hook that captures ALL events
	_WINEVENTPROC = ctypes.WINFUNCTYPE(
		None,
		wintypes.HANDLE,
		wintypes.DWORD,
		wintypes.HWND,
		ctypes.c_long,
		ctypes.c_long,
		wintypes.DWORD,
		wintypes.DWORD,
	)
	_user32 = ctypes.windll.user32

	def _logCallback(hHook, event, hwnd, idObject, idChild, tid, ts):
		if not getattr(app, '_eventLogActive', False):
			return
		try:
			ename = EVENT_NAMES.get(event, "0x%04X" % event)
			f.write("%s %-30s hwnd=%-8s obj=%-6d child=%-4d\n" % (
				time.strftime("%H:%M:%S"),
				ename, hwnd, idObject, idChild,
			))
			counter[0] += 1
			if counter[0] % 100 == 0:
				f.flush()
		except Exception:
			pass

	cb = _WINEVENTPROC(_logCallback)

	# Hook ALL events from the Discord process
	hook = _user32.SetWinEventHookW(
		0x0001,    # EVENT_MIN
		0x8022,    # EVENT_OBJECT_END (covers all event types)
		0,         # no DLL
		cb,
		pid,
		0,         # all threads
		0x0000,    # WINEVENT_OUTOFCONTEXT
	)

	if not hook:
		f.close()
		ui.message("Failed to install event hook")
		return

	app._eventLogActive = True

	# Auto-stop after 15 seconds
	def _stop():
		if not getattr(app, '_eventLogActive', False):
			return
		app._eventLogActive = False
		try:
			_user32.UnhookWinEvent(hook)
		except Exception:
			pass
		f.write("\n" + "=" * 70 + "\n")
		f.write("Stopped at %s. Total events: %d\n" % (
			time.strftime("%H:%M:%S"), counter[0]))
		f.flush()
		f.close()
		ui.message(
			"Event logging complete. %d events saved to discord_events.log on desktop."
			% counter[0]
		)

	# The cb and hook refs must stay alive -- store on app module
	app._eventLogCB = cb
	app._eventLogHook = hook
	app._eventLogStop = _stop

	wx.CallLater(15000, _stop)

	ui.message("Logging Discord events for 15 seconds. Send a message now.")
