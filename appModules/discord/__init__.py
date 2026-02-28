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

import time
from comtypes import COMError
from logHandler import log
import appModuleHandler
import api
import config
import controlTypes
import inputCore
import tones
import ui

from . import commands
from . import overlays
from . import uia


# ---------------------------------------------------------------------------
# Thread-safe command execution
# ---------------------------------------------------------------------------
# _captureFunc runs on NVDA's keyboard hook thread.  UIA COM calls and
# wx dialogs must happen on the main GUI thread.  This helper schedules
# the handler via wx.CallAfter.

def _run_on_main(handler):
	"""Execute *handler* on the main wx GUI thread."""
	import wx

	def _do():
		try:
			handler()
		except Exception:
			log.error("Command handler error", exc_info=True)
			tones.beep(*_TONE_ERROR)

	wx.CallAfter(_do)


# ---------------------------------------------------------------------------
# Tone constants (frequency Hz, duration ms)
# ---------------------------------------------------------------------------
_TONE_ENTER = (800, 25)     # layer entered
_TONE_EXIT = (400, 25)      # layer exited / command executed
_TONE_WRAP = (200, 40)      # Tab exploration wrapped around
_TONE_ERROR = (150, 60)     # unknown key / error


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
		log.info("Discord Enhancements appModule loaded")

	def terminate(self):
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
	# Incoming message announcements (passive event handlers)
	# ------------------------------------------------------------------
	# These handlers are called BY NVDA when it detects live-region or
	# alert events.  Zero cost when no events fire.  No polling, no
	# timers, no hooks -- purely reactive.

	_lastAnnouncedText = ""
	_lastAnnouncedTime = 0.0

	def _shouldAnnounce(self):
		try:
			return config.conf["discordAddon"]["announceChatMessages"]
		except (KeyError, Exception):
			return True

	def _announceIfNew(self, text):
		"""Speak text if it's not a duplicate of the last announcement."""
		if not text:
			return
		now = time.time()
		if text == self._lastAnnouncedText and now - self._lastAnnouncedTime < 1.0:
			return
		self._lastAnnouncedText = text
		self._lastAnnouncedTime = now
		ui.message(text)

	def event_liveRegionChange(self, obj, nextHandler):
		"""Announce live-region updates (incoming chat messages)."""
		if self._shouldAnnounce():
			text = uia.safe_name(obj) or ""
			if not text:
				text = uia.read_message_content(obj)
			if text and text != "(empty message)":
				lower = text.lower()
				if "typing" not in lower:
					self._announceIfNew(text)
		nextHandler()

	def event_alert(self, obj, nextHandler):
		"""Announce alert elements (notifications, toasts)."""
		if self._shouldAnnounce():
			text = uia.safe_name(obj) or uia.safe_value(obj) or ""
			if text:
				self._announceIfNew(text)
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
