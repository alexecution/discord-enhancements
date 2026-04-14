# Discord Enhancements Add-on for NVDA
# AppModule for discord.exe
#
# This is the main entry point.  It:
#   - Disables browse mode by default for Discord
#   - Provides a configurable command layer (prefix key) via _captureFunc
#   - Appends "window" to the Discord title for clarity
#   - Registers overlay classes for enhanced speech
#   - Announces incoming chat messages via live-region events
#
# ARCHITECTURE NOTES:
#   The command layer uses inputCore.manager._captureFunc rather than
#   gesture binding.  This is the same mechanism NVDA's lock-screen
#   module uses.  _captureFunc is called BEFORE any gesture/script
#   resolution, which means it works regardless of browse-mode state.
#
#   Return False from the captor -> block the gesture (swallow it)
#   Return True  from the captor -> continue normal processing
#
# PERFORMANCE RULES (must be observed to avoid lag):
#   - event_NVDAObject_init  -- O(1), no tree walks
#   - chooseNVDAObjectOverlayClasses -- O(1), no tree walks
#   - event_liveRegionChange -- O(1), reads name/children of one object
#   - _discordCaptor -- O(1), just key comparisons
#   - Tree walking ONLY happens inside explicit command handlers

import comtypes
import contextlib
import ctypes
import ctypes.wintypes
import re
import threading
import time
from comtypes import COMError
from logHandler import log
import appModuleHandler
import api
import config
import controlTypes
import core
import inputCore
import speech
import tones
import ui

from . import commands
from . import overlays
from . import uia


# ---------------------------------------------------------------------------
# Tone constants (frequency Hz, duration ms)
# ---------------------------------------------------------------------------
_TONE_ENTER = (800, 25)     # layer entered
_TONE_EXIT = (400, 25)      # layer exited / command executed
_TONE_WRAP = (200, 40)      # Tab exploration wrapped around
_TONE_ERROR = (150, 60)     # unknown key / error


# ---------------------------------------------------------------------------
# UIA polling constants
# ---------------------------------------------------------------------------
_POLL_INTERVAL_MS = 500  # milliseconds between UIA polls

# UIA property/type constants
_UIA_NamePropertyId = 30005
_UIA_ControlTypePropertyId = 30003
_UIA_ListControlTypeId = 50008
_UIA_TreeScope_Descendants = 4

# Filtering: status suffixes to ignore in announcements
_STATUS_SUFFIXES_LOWER = (
	', online', ', offline', ', idle',
	', do not disturb', ', streaming',
)
# Standalone timestamp pattern like "9:04 AM"
_TIMESTAMP_RE = re.compile(r'^\d{1,2}:\d{2}\s*(AM|PM)?$', re.IGNORECASE)


# ---------------------------------------------------------------------------
# Thread-safe command execution
# ---------------------------------------------------------------------------
# _captureFunc runs on NVDA's keyboard hook thread.  UIA COM calls
# must happen on the main thread.  core.callLater posts directly to
# NVDA's main loop which is faster than wx.CallAfter.

def _run_on_main(handler):
	"""Execute *handler* on the main NVDA thread via core.callLater."""
	def _do():
		try:
			handler()
		except Exception:
			log.error("Command handler error", exc_info=True)
			tones.beep(*_TONE_ERROR)
	core.callLater(0, _do)


# ---------------------------------------------------------------------------
# Command registry
# ---------------------------------------------------------------------------
# Each entry: (key_name, handler_callable, description)
# key_name uses NVDA mainKeyName values (lower-case).
# Modifiers are prepended: "shift+h", "control+e", etc.

COMMAND_REGISTRY = [
	# --- General ---
	("a",       commands.cmd_activeNow,       "Announce Active Now section"),
	("v",       commands.cmd_voiceServers,     "Report servers with active voice"),
	("b",       commands.cmd_listButtons,      "List and activate buttons"),
	# --- Navigation ---
	("e",       commands.cmd_messageInput,     "Move focus to message input"),
	("s",       commands.cmd_serverList,        "Move focus to server list"),
	("n",       commands.cmd_navigateAreas,     "Cycle focus among Discord areas"),
	("u",       commands.cmd_userArea,          "Move focus to user area"),
	# --- Chat messages (home-row) ---
	("h",       commands.cmd_firstMessage,     "First message"),
	("j",       commands.cmd_prevMessage,      "Previous message"),
	("k",       commands.cmd_currentMessage,   "Current message (double-tap to spell)"),
	("l",       commands.cmd_nextMessage,      "Next message"),
	("oem_1",   commands.cmd_lastMessage,      "Last message"),  # semicolon
	# --- Chat messages (arrow/nav keys) ---
	("home",        commands.cmd_firstMessage,     "First message"),
	("leftarrow",   commands.cmd_prevMessage,      "Previous message"),
	("numpad5",     commands.cmd_currentMessage,   "Current message (double-tap to spell)"),
	("rightarrow",  commands.cmd_nextMessage,      "Next message"),
	("end",         commands.cmd_lastMessage,      "Last message"),
	# --- Shift variants ---
	("shift+h",         commands.cmd_unreadMarker,        "Jump to unread marker"),
	("shift+home",      commands.cmd_unreadMarker,        "Jump to unread marker"),
	("shift+k",         commands.cmd_focusCurrentMessage, "Focus current message"),
	("shift+numpad5",   commands.cmd_focusCurrentMessage, "Focus current message"),
	("shift+p",         commands.cmd_pinnedMessages,      "Open pinned messages"),
	("shift+t",         commands.cmd_threadList,           "Toggle thread list"),
	("shift+d",         commands.cmd_toggleAnnounce,       "Toggle message announcements"),
	# --- Voice ---
	("d",       commands.cmd_disconnect,       "Disconnect from voice"),
	("p",       commands.cmd_ping,             "Report ping / latency"),
	# --- Information ---
	("t",       commands.cmd_typing,           "Who is typing"),
	("w",       commands.cmd_channelInfo,       "Channel / DM information"),
	# --- Diagnostic ---
	("control+e",  commands.cmd_diagnostic,    "Dump accessibility tree"),
	("control+m",  commands.cmd_messageDebug,  "Message list diagnostic"),
	("control+l",  commands.cmd_eventLog,      "Log accessibility events (15s)"),
]

# Digit commands: [ 1-9, 0
for _digit in range(1, 10):
	COMMAND_REGISTRY.append((
		str(_digit),
		lambda n=_digit: commands.cmd_recentMessage(n),
		"Read %s most recent message" % (
			{1: "1st", 2: "2nd", 3: "3rd"}.get(_digit, "%dth" % _digit)
		),
	))
COMMAND_REGISTRY.append((
	"0",
	lambda: commands.cmd_recentMessage(10),
	"Read 10th most recent message",
))

# Build the lookup dict and exploration list
_COMMAND_MAP = {}
for _key, _handler, _desc in COMMAND_REGISTRY:
	_COMMAND_MAP[_key] = (_handler, _desc)

_EXPLORE_LIST = [
	(_key, _desc) for _key, _handler, _desc in COMMAND_REGISTRY
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _getConfigPrefix():
	"""Return the current command prefix key from config."""
	try:
		return config.conf["discordAddon"]["commandPrefix"]
	except (KeyError, Exception):
		return "["


def _isEditField(obj):
	"""Return True if *obj* is an editable text field."""
	try:
		role = obj.role
		if role in (
			controlTypes.Role.EDITABLETEXT,
			controlTypes.Role.DOCUMENT,
			controlTypes.Role.TEXTFRAME,
			controlTypes.Role.TERMINAL,
		):
			return True
		states = obj.states or set()
		if controlTypes.State.EDITABLE in states:
			return True
	except (COMError, AttributeError, Exception):
		pass
	return False


# ---------------------------------------------------------------------------
# AppModule
# ---------------------------------------------------------------------------

class AppModule(appModuleHandler.AppModule):
	"""NVDA AppModule for the Discord desktop application."""

	# Disable the tree interceptor (virtual buffer) entirely.
	# The virtual buffer constantly monitors DOM changes which causes
	# lag in Chromium/Electron apps.  All navigation is handled by
	# our command layer instead.
	disableBrowseModeByDefault = True

	# Timeout in seconds for the command layer auto-cancel
	LAYER_TIMEOUT = 5.0

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._layerActive = False
		self._layerStartTime = 0.0
		self._exploreIndex = -1
		self._lastExplored = None

		# UIA polling state (background thread owns the poll loop;
		# main-thread Alt+N shortcuts use these for their own cache)
		self._lastPollText = ""
		self._cachedMsgList = None
		self._cachedMsgListName = None
		self._discordHwnd = 0
		self._terminated = False
		self._lastInteractionTime = 0.0

		# Background polling thread — UIA reads happen off the main
		# thread so they never block keyboard/speech processing.
		self._pollStopEvent = threading.Event()
		self._pollThread = threading.Thread(
			target=self._pollLoop, daemon=True, name="DiscordUiaPoll")
		self._pollThread.start()
		log.info("Discord Enhancements appModule loaded")

	def terminate(self):
		self._terminated = True
		self._pollStopEvent.set()  # wake thread so it exits
		if self._pollThread:
			self._pollThread.join(timeout=2.0)
		if inputCore.manager._captureFunc == self._discordCaptor:
			inputCore.manager._captureFunc = None
		self._layerActive = False
		super().terminate()

	# ------------------------------------------------------------------
	# AppModule focus events
	# ------------------------------------------------------------------

	def event_appModule_gainFocus(self):
		"""Discord gained focus -- activate the capture function."""
		inputCore.manager._captureFunc = self._discordCaptor

	def event_appModule_loseFocus(self):
		"""Discord lost focus -- deactivate capture, exit layer."""
		self._exitCommandLayer(silent=True)
		if inputCore.manager._captureFunc == self._discordCaptor:
			inputCore.manager._captureFunc = None

	# ------------------------------------------------------------------
	# Window title enrichment
	# ------------------------------------------------------------------

	def event_NVDAObject_init(self, obj):
		"""Append ' window' to the Discord top-level window title.

		PERFORMANCE: only checks immediate properties -- O(1).
		"""
		try:
			cls = getattr(obj, "windowClassName", "")
			if cls != "Chrome_WidgetWin_1":
				return
			parent = getattr(obj, "parent", None)
			if not parent:
				return
			# Top-level window: parent is the Desktop (has no parent
			# itself, or its role is PANE / WINDOW).
			try:
				is_top_level = parent.parent is None
			except Exception:
				is_top_level = False
			if not is_top_level:
				try:
					parent_role = parent.role
					is_top_level = parent_role in (
						controlTypes.Role.PANE,
						controlTypes.Role.WINDOW,
					)
				except Exception:
					pass
			if is_top_level:
				name = obj.name or ""
				if name and "discord" in name.lower() and not name.endswith("window"):
					obj.name = name + " window"
		except (COMError, AttributeError, Exception):
			pass

	# ------------------------------------------------------------------
	# Enhanced title script (NVDA+T override)
	# ------------------------------------------------------------------

	def script_title(self, gesture):
		"""Report an enriched window title for Discord."""
		root = uia.get_foreground()
		if root is None:
			gesture.send()
			return
		context = uia.get_window_context(root)
		parts = []

		# Server and channel
		server = context.get("server", "")
		channel = context.get("channel", "")
		dm_name = context.get("dm_name", "")
		if channel and server:
			parts.append("%s in %s" % (channel, server))
		elif channel:
			parts.append(channel)
		elif dm_name:
			parts.append("DM with %s" % dm_name)
		elif server:
			parts.append(server)
		else:
			parts.append(uia.safe_name(root) or "Discord")

		# Status indicators
		if context.get("muted"):
			parts.append("muted")
		if context.get("deafened"):
			parts.append("deafened")
		if context.get("voice_channel"):
			parts.append("voice: %s" % context["voice_channel"])

		# Badges (unread counts)
		for badge in context.get("badges", []):
			parts.append(badge)

		# Alerts
		for alert in context.get("alerts", []):
			parts.append(alert)

		title = ", ".join(parts)
		if not title.lower().endswith("window"):
			title += " window"
		ui.message(title)

	script_title.__doc__ = "Report enhanced Discord window title"
	script_title.category = "Discord Enhancements"

	__gestures = {
		"kb:NVDA+t": "title",
		"kb:alt+1": "readMessage1",
		"kb:alt+2": "readMessage2",
		"kb:alt+3": "readMessage3",
		"kb:alt+4": "readMessage4",
		"kb:alt+5": "readMessage5",
		"kb:alt+6": "readMessage6",
		"kb:alt+7": "readMessage7",
		"kb:alt+8": "readMessage8",
		"kb:alt+9": "readMessage9",
		"kb:alt+0": "readMessage10",
	}

	# ------------------------------------------------------------------
	# Overlay classes -- O(1) identification
	# ------------------------------------------------------------------

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""Add Discord overlay classes when appropriate.

		PERFORMANCE: identify_overlay_class is O(1) -- it only checks
		obj.role.  No tree walking, no parent checks.
		"""
		overlay = overlays.identify_overlay_class(obj)
		if overlay is not None:
			clsList.insert(0, overlay)

	# ------------------------------------------------------------------
	# Incoming message announcements
	# ------------------------------------------------------------------
	# Hybrid approach: UIA polling every 500ms as the primary mechanism
	# (reliable for Chromium/Electron), plus reactive event handlers
	# as a fast-path for immediate detection.  Both paths funnel
	# through _filterAndAnnounce for deduplication and filtering.

	_lastAnnouncedText = ""
	_lastAnnouncedTime = 0.0

	def _shouldAnnounce(self):
		try:
			return config.conf["discordAddon"]["announceChatMessages"]
		except (KeyError, Exception):
			return True

	# ------------------------------------------------------------------
	# UIA polling — background thread
	# ------------------------------------------------------------------
	# All UIA COM calls for polling happen on a daemon thread with its
	# own COM client.  Only the final speech call is posted back to
	# NVDA's main thread via core.callLater.
	#
	# PERFORMANCE DESIGN:
	#   The message list element is cached after the first successful
	#   lookup.  On subsequent polls only GetLastChildElement +
	#   GetCurrentPropertyValue run — both O(1).  The expensive
	#   FindAll(TreeScope_Descendants) only fires on cache miss
	#   (app start, channel switch) and runs on this background
	#   thread, so NVDA's main thread is never blocked.

	def _pollLoop(self):
		"""Background thread: poll Discord's UIA tree for new messages."""
		# MTA — no message pump required (STA would deadlock without one)
		_COINIT_MULTITHREADED = 0x0
		ctypes.windll.ole32.CoInitializeEx(None, _COINIT_MULTITHREADED)
		uia_client = None
		walker = None
		cached_msg_list = None
		cached_msg_list_name = None
		list_condition = None

		try:
			while not self._terminated:
				# Sleep for the poll interval; wake early if terminated
				self._pollStopEvent.wait(_POLL_INTERVAL_MS / 1000.0)
				if self._terminated:
					break
				if not self._shouldAnnounce():
					continue

				# Back off while the user is actively navigating so our
				# cross-process UIA calls don't contend with NVDA's
				# main-thread reads and cause speech delay.
				if time.time() - self._lastInteractionTime < 1.0:
					continue

				# Check foreground via pure Win32 (no NVDA API needed)
				try:
					fg_hwnd = ctypes.windll.user32.GetForegroundWindow()
					pid = ctypes.wintypes.DWORD()
					ctypes.windll.user32.GetWindowThreadProcessId(
						fg_hwnd, ctypes.byref(pid))
					if pid.value != self.processID:
						continue
				except Exception:
					continue

				# Store hwnd for main-thread Alt+N usage too
				if fg_hwnd and not self._discordHwnd:
					self._discordHwnd = fg_hwnd

				# Lazy-init UIA client on this thread
				if uia_client is None:
					try:
						from comInterfaces import UIAutomationClient as UIA
						try:
							uia_client = comtypes.CoCreateInstance(
								UIA.CUIAutomation8._reg_clsid_,
								interface=UIA.IUIAutomation,
								clsctx=comtypes.CLSCTX_INPROC_SERVER,
							)
						except Exception:
							uia_client = comtypes.CoCreateInstance(
								UIA.CUIAutomation._reg_clsid_,
								interface=UIA.IUIAutomation,
								clsctx=comtypes.CLSCTX_INPROC_SERVER,
							)
						walker = uia_client.RawViewWalker
						list_condition = uia_client.CreatePropertyCondition(
							_UIA_ControlTypePropertyId,
							_UIA_ListControlTypeId,
						)
						log.info("Discord poll thread: UIA client created")
					except Exception:
						log.error(
							"Discord poll thread: UIA client creation failed",
							exc_info=True,
						)
						continue

				# Validate message list cache
				if cached_msg_list:
					try:
						n = cached_msg_list.GetCurrentPropertyValue(
							_UIA_NamePropertyId) or ""
						if not (n and n == cached_msg_list_name):
							cached_msg_list = None
							cached_msg_list_name = None
					except Exception:
						cached_msg_list = None
						cached_msg_list_name = None

				# Find message list (cache miss only — rare)
				if not cached_msg_list:
					try:
						root = uia_client.ElementFromHandle(fg_hwnd)
						if not root:
							continue
						lists = root.FindAll(
							_UIA_TreeScope_Descendants, list_condition)
						if not lists or lists.Length == 0:
							continue
						found = False
						for i in range(lists.Length):
							elem = lists.GetElement(i)
							n = elem.GetCurrentPropertyValue(
								_UIA_NamePropertyId) or ""
							if "messages in" in n.lower():
								cached_msg_list = elem
								cached_msg_list_name = n
								found = True
								break
						if not found:
							continue
					except Exception:
						continue

				# Read the latest message — O(1) via cached list
				# Minimise cross-process COM calls to avoid contention
				# with NVDA's main-thread UIA reads.
				try:
					child = walker.GetLastChildElement(cached_msg_list)

					if not child:
						cached_msg_list = None
						cached_msg_list_name = None
						continue

					name = child.GetCurrentPropertyValue(
						_UIA_NamePropertyId) or ""

					# Message content may be nested one level deeper
					if not name:
						gchild = walker.GetLastChildElement(child)
						if gchild:
							name = gchild.GetCurrentPropertyValue(
								_UIA_NamePropertyId) or ""

					# Last child might be a separator; try previous sibling
					if not name:
						prev = walker.GetPreviousSiblingElement(child)
						if prev:
							name = prev.GetCurrentPropertyValue(
								_UIA_NamePropertyId) or ""
							if not name:
								gchild = walker.GetLastChildElement(prev)
								if gchild:
									name = gchild.GetCurrentPropertyValue(
										_UIA_NamePropertyId) or ""

					if name and name != self._lastPollText:
						self._lastPollText = name
						core.callLater(0, self._filterAndAnnounce, name)

				except COMError:
					cached_msg_list = None
					cached_msg_list_name = None
				except Exception:
					cached_msg_list = None
					cached_msg_list_name = None

		except Exception:
			log.error("Discord poll loop crashed", exc_info=True)
		finally:
			with contextlib.suppress(Exception):
				ctypes.windll.ole32.CoUninitialize()

	# ------------------------------------------------------------------
	# Main-thread UIA helpers (for Alt+N history reading)
	# ------------------------------------------------------------------

	def _getMsgListViaUIA(self, uia_client):
		"""Find and cache the Discord message list UIA element (main thread).

		Used by Alt+N shortcuts.  The background poll thread has its own
		independent cache so COM objects are never shared across threads.
		"""
		if self._cachedMsgList:
			cache_valid = False
			with contextlib.suppress(Exception):
				n = self._cachedMsgList.GetCurrentPropertyValue(
					_UIA_NamePropertyId) or ""
				cache_valid = bool(n and n == self._cachedMsgListName)
			if cache_valid:
				return self._cachedMsgList
			self._cachedMsgList = None
			self._cachedMsgListName = None

		hwnd = self._discordHwnd
		if not hwnd:
			try:
				fg = api.getForegroundObject()
				if fg and fg.appModule is self:
					hwnd = fg.windowHandle
					self._discordHwnd = hwnd
			except Exception:
				pass
		if not hwnd:
			return None

		try:
			root = uia_client.ElementFromHandle(hwnd)
		except Exception:
			return None
		if not root:
			return None

		condition = uia_client.CreatePropertyCondition(
			_UIA_ControlTypePropertyId, _UIA_ListControlTypeId
		)
		lists = root.FindAll(_UIA_TreeScope_Descendants, condition)
		if not lists or lists.Length == 0:
			return None

		for i in range(lists.Length):
			elem = lists.GetElement(i)
			n = elem.GetCurrentPropertyValue(_UIA_NamePropertyId) or ""
			if "messages in" in n.lower():
				self._cachedMsgList = elem
				self._cachedMsgListName = n
				return elem
		return None

	def _getMessagesViaUIA(self, count=10):
		"""Walk the UIA tree and return up to *count* recent messages (oldest first)."""
		try:
			import UIAHandler
			uia_client = UIAHandler.handler.clientObject
			if not uia_client:
				return []

			msgList = self._getMsgListViaUIA(uia_client)
			if not msgList:
				return []

			walker = uia_client.RawViewWalker
			child = walker.GetLastChildElement(msgList)

			if not child and self._cachedMsgList is not None:
				self._cachedMsgList = None
				self._cachedMsgListName = None
				msgList = self._getMsgListViaUIA(uia_client)
				if not msgList:
					return []
				child = walker.GetLastChildElement(msgList)

			candidates = []
			limit = count * 4
			iterations = 0
			while child and iterations < limit:
				iterations += 1
				name = child.GetCurrentPropertyValue(_UIA_NamePropertyId) or ""
				if name:
					candidates.append(name)
				else:
					grandchild = walker.GetLastChildElement(child)
					while grandchild:
						gname = grandchild.GetCurrentPropertyValue(_UIA_NamePropertyId) or ""
						if gname:
							candidates.append(gname)
							break
						grandchild = walker.GetPreviousSiblingElement(grandchild)
				child = walker.GetPreviousSiblingElement(child)

			messages = [m for m in candidates if self._isValidMessage(m)]
			messages = messages[:count]
			messages.reverse()  # oldest first
			return messages
		except Exception as e:
			log.warning("Discord Enhancements: getMessages error: %s" % e)
			return []

	# ------------------------------------------------------------------
	# Filtering and announcement
	# ------------------------------------------------------------------

	def _isValidMessage(self, name):
		"""Return True if *name* looks like a real chat message."""
		lower = name.lower()
		if any(lower.endswith(s) for s in _STATUS_SUFFIXES_LOWER):
			return False
		if 'is typing' in lower or 'are typing' in lower:
			return False
		if ' , ' in name:
			parts = name.split(' , ')
			if ':' not in parts[-1]:
				return False
			body = ' , '.join(parts[1:-1]).strip() if len(parts) >= 3 else ""
			return bool(body)
		return len(name) >= 3 and not _TIMESTAMP_RE.match(name.strip())

	def _filterAndAnnounce(self, name):
		"""Filter out non-message text and announce if new."""
		lower = name.lower()

		if any(lower.endswith(s) for s in _STATUS_SUFFIXES_LOWER):
			return
		if 'is typing' in lower or 'are typing' in lower:
			return

		if ' , ' in name:
			# IAccessible format: "username , body , HH:MM AM"
			parts = name.split(' , ')
			if ':' not in parts[-1]:
				return
			body = ' , '.join(parts[1:-1]).strip() if len(parts) >= 3 else ""
			if not body:
				return
			self._scheduleAnnounce(name)
			return

		# Plain-text UIA — skip timestamps and very short UI labels
		if len(name) < 3 or _TIMESTAMP_RE.match(name.strip()):
			return
		self._scheduleAnnounce(name)

	def _scheduleAnnounce(self, text):
		"""Announce *text* if it differs from the last announced text."""
		now = time.time()
		if text == self._lastAnnouncedText and now - self._lastAnnouncedTime < 1.0:
			return
		if not self._shouldAnnounce():
			return
		self._lastAnnouncedText = text
		self._lastAnnouncedTime = now
		self._lastPollText = text
		self._doAnnounce(text)

	def _doAnnounce(self, text):
		"""Format and speak the message text at highest priority."""
		# IAccessible format → "username: body"
		if ' , ' in text:
			parts = text.split(' , ')
			formatted = (
				"%s: %s" % (parts[0], ' , '.join(parts[1:-1]))
				if len(parts) >= 3
				else parts[0]
			)
		else:
			formatted = text
		try:
			speech.speak([formatted], priority=speech.Spri.NOW)
		except Exception as e:
			log.warning("Discord Enhancements: speech error: %s" % e)

	# ------------------------------------------------------------------
	# History reading via raw UIA (Alt+1 through Alt+0)
	# ------------------------------------------------------------------

	def _readNthLastMessage(self, n):
		"""Speak the Nth-last message (1 = most recent)."""
		try:
			fg = api.getForegroundObject()
			if not fg or fg.appModule is not self:
				return
		except Exception:
			return
		messages = self._getMessagesViaUIA(count=10)
		if not messages:
			ui.message("No messages found")
			return
		idx = len(messages) - n
		if idx < 0:
			ui.message("Message %d not available" % n)
			return
		self._doAnnounce(messages[idx])

	def script_readMessage1(self, gesture):
		"""Read the most recent message."""
		self._readNthLastMessage(1)

	def script_readMessage2(self, gesture):
		"""Read the 2nd most recent message."""
		self._readNthLastMessage(2)

	def script_readMessage3(self, gesture):
		"""Read the 3rd most recent message."""
		self._readNthLastMessage(3)

	def script_readMessage4(self, gesture):
		"""Read the 4th most recent message."""
		self._readNthLastMessage(4)

	def script_readMessage5(self, gesture):
		"""Read the 5th most recent message."""
		self._readNthLastMessage(5)

	def script_readMessage6(self, gesture):
		"""Read the 6th most recent message."""
		self._readNthLastMessage(6)

	def script_readMessage7(self, gesture):
		"""Read the 7th most recent message."""
		self._readNthLastMessage(7)

	def script_readMessage8(self, gesture):
		"""Read the 8th most recent message."""
		self._readNthLastMessage(8)

	def script_readMessage9(self, gesture):
		"""Read the 9th most recent message."""
		self._readNthLastMessage(9)

	def script_readMessage10(self, gesture):
		"""Read the 10th most recent message."""
		self._readNthLastMessage(10)

	# ------------------------------------------------------------------
	# Toggle announcement
	# ------------------------------------------------------------------

	# ------------------------------------------------------------------
	# Reactive event handlers (fast-path, supplement polling)
	# ------------------------------------------------------------------

	def event_liveRegionChange(self, obj, nextHandler):
		"""Announce live-region updates (incoming chat messages)."""
		if self._shouldAnnounce():
			text = uia.safe_name(obj) or ""
			if not text:
				text = uia.read_message_content(obj)
			if text and text != "(empty message)":
				self._filterAndAnnounce(text)
		nextHandler()

	def event_alert(self, obj, nextHandler):
		"""Announce alert elements (notifications, toasts)."""
		if self._shouldAnnounce():
			text = uia.safe_name(obj) or uia.safe_value(obj) or ""
			if text:
				self._filterAndAnnounce(text)
		nextHandler()

	# ------------------------------------------------------------------
	# Master capture function -- handles ALL command-layer logic
	# ------------------------------------------------------------------

	def _discordCaptor(self, gesture):
		"""Called for EVERY input gesture while Discord is focused.

		This runs before any script/gesture resolution, so it works
		regardless of browse-mode state.

		Return False -> block the gesture (swallow it).
		Return True  -> continue normal processing.
		"""
		# Track interaction time so the poll thread can back off
		# during active navigation (avoids UIA COM contention).
		self._lastInteractionTime = time.time()
		try:
			# If the command layer is active, handle the follow-up key
			if self._layerActive:
				# Check for timeout (auto-cancel)
				if time.time() - self._layerStartTime > self.LAYER_TIMEOUT:
					tones.beep(*_TONE_ERROR)
					self._exitCommandLayer(silent=True)
					return True  # Timed out, let the key through
				self._handleLayerKey(gesture)
				return False  # Always swallow keys while the layer is active

			# Check if this is the prefix key (no modifiers)
			if not self._isPrefixGesture(gesture):
				return True  # Not our key, pass through

			# Enter the command layer (works everywhere, even in
			# edit fields -- press the prefix key twice to type it)
			self._enterCommandLayer()
			return False  # Swallow the prefix key
		except Exception:
			log.error("Discord captor error", exc_info=True)
			return True  # On error, let the key through

	# ------------------------------------------------------------------
	# Prefix gesture detection
	# ------------------------------------------------------------------

	def _isPrefixGesture(self, gesture):
		"""Check if a gesture matches the configured prefix key."""
		try:
			from keyboardHandler import KeyboardInputGesture
			if not isinstance(gesture, KeyboardInputGesture):
				return False
			# Must have no modifiers (just the bare key)
			if gesture.modifierNames:
				return False
			prefix = _getConfigPrefix()
			# Primary check: direct mainKeyName comparison
			if gesture.mainKeyName == prefix:
				return True
			# Fallback: compare normalised gesture identifiers
			expected = inputCore.normalizeGestureIdentifier("kb:%s" % prefix)
			for gid in gesture.normalizedIdentifiers:
				if gid == expected:
					return True
		except Exception:
			pass
		return False

	# ------------------------------------------------------------------
	# Layer lifecycle
	# ------------------------------------------------------------------

	def _enterCommandLayer(self):
		"""Activate the command layer."""
		self._layerActive = True
		self._layerStartTime = time.time()
		self._exploreIndex = -1
		self._lastExplored = None
		tones.beep(*_TONE_ENTER)

	def _exitCommandLayer(self, silent=False):
		"""Deactivate the command layer."""
		if not self._layerActive and not silent:
			return
		self._layerActive = False
		self._exploreIndex = -1
		self._lastExplored = None
		if not silent:
			tones.beep(*_TONE_EXIT)

	# ------------------------------------------------------------------
	# Keystroke handler for the active command layer
	# ------------------------------------------------------------------

	def _handleLayerKey(self, gesture):
		"""Process a keystroke while the command layer is active."""
		# Reset the auto-cancel timeout
		self._layerStartTime = time.time()

		# We only handle keyboard gestures
		try:
			from keyboardHandler import KeyboardInputGesture
			if not isinstance(gesture, KeyboardInputGesture):
				return
			key_name = gesture.mainKeyName
			mod_names = gesture.modifierNames or []
		except Exception:
			self._exitCommandLayer()
			return

		if not key_name:
			return

		key_lower = key_name.lower()

		# --- Escape: cancel ---
		if key_lower == "escape":
			self._exitCommandLayer()
			return

		# --- Prefix key again: type the literal prefix character ---
		if self._isPrefixGesture(gesture):
			self._exitCommandLayer()
			# Send the actual keystroke so it types in edit fields
			try:
				gesture.send()
			except Exception:
				pass
			return

		# --- Tab / Shift+Tab: explore commands ---
		if key_lower == "tab":
			shift = "shift" in [m.lower() for m in mod_names]
			if shift:
				self._explorePrev()
			else:
				self._exploreNext()
			return

		# --- Enter: execute last explored command ---
		if key_lower in ("return", "enter"):
			if self._lastExplored is not None:
				handler, desc = self._lastExplored
				self._exitCommandLayer()
				_run_on_main(handler)
			else:
				tones.beep(*_TONE_ERROR)
			return

		# --- Build the full key name with modifiers ---
		full_key = key_lower
		if mod_names:
			mods = sorted([m.lower() for m in mod_names if m.lower() != key_lower])
			if mods:
				full_key = "+".join(mods) + "+" + key_lower

		# --- Look up the command ---
		entry = _COMMAND_MAP.get(full_key)
		if entry is None and full_key != key_lower:
			# Fallback: try without modifiers
			entry = _COMMAND_MAP.get(key_lower)
		if entry is not None:
			handler, desc = entry
			self._exitCommandLayer()
			_run_on_main(handler)
			return

		# Unknown key
		tones.beep(*_TONE_ERROR)

	# ------------------------------------------------------------------
	# Tab exploration
	# ------------------------------------------------------------------

	def _exploreNext(self):
		if not _EXPLORE_LIST:
			return
		self._exploreIndex += 1
		if self._exploreIndex >= len(_EXPLORE_LIST):
			self._exploreIndex = 0
			tones.beep(*_TONE_WRAP)
		key, desc = _EXPLORE_LIST[self._exploreIndex]
		self._lastExplored = _COMMAND_MAP.get(key)
		ui.message("%s: %s" % (key, desc))

	def _explorePrev(self):
		if not _EXPLORE_LIST:
			return
		self._exploreIndex -= 1
		if self._exploreIndex < 0:
			self._exploreIndex = len(_EXPLORE_LIST) - 1
			tones.beep(*_TONE_WRAP)
		key, desc = _EXPLORE_LIST[self._exploreIndex]
		self._lastExplored = _COMMAND_MAP.get(key)
		ui.message("%s: %s" % (key, desc))
