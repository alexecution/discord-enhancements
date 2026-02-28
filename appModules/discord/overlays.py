# Discord Enhancements Add-on for NVDA
# Overlay classes for enhanced speech output
#
# PERFORMANCE CRITICAL: identify_overlay_class() is called for EVERY
# NVDA object created while Discord is focused.  It MUST be O(1) —
# only checking the object's own immediate properties (role, name).
# NEVER walk parents, siblings, or children here.

from comtypes import COMError
from logHandler import log
import controlTypes
import ui
from NVDAObjects.UIA import UIA
from . import uia


# ---------------------------------------------------------------------------
# Server tree item overlay
# ---------------------------------------------------------------------------

class DiscordServerItem(UIA):
	"""Overlay for server/guild items in the server navigation list."""

	def _get_name(self):
		base = uia.safe_name(self) or ""
		try:
			if uia.has_voice_activity(self):
				return base + ", voice active"
		except Exception:
			pass
		return base


# ---------------------------------------------------------------------------
# Channel list item overlay
# ---------------------------------------------------------------------------

class DiscordChannelItem(UIA):
	"""Overlay for channel items in the channel sidebar."""

	def _get_name(self):
		return uia.safe_name(self) or ""


# ---------------------------------------------------------------------------
# Chat message item overlay
# ---------------------------------------------------------------------------

class DiscordMessageItem(UIA):
	"""Overlay for individual chat messages in the message list."""

	def _get_name(self):
		return uia.read_message_content(self)


# ---------------------------------------------------------------------------
# Members list item overlay
# ---------------------------------------------------------------------------

class DiscordMemberItem(UIA):
	"""Overlay for items in the Members sidebar.

	Announces the heading above the current member when focused.
	"""

	def event_gainFocus(self):
		try:
			prev = self.previous
			attempts = 0
			while prev and attempts < 5:
				role = uia.safe_role(prev)
				if role in (controlTypes.Role.HEADING, controlTypes.Role.GROUPING):
					heading = uia.safe_name(prev)
					if heading:
						ui.message(heading)
					break
				prev_next = prev.previous
				if prev_next is prev:
					break
				prev = prev_next
				attempts += 1
		except (COMError, Exception):
			pass
		super().event_gainFocus()


# ---------------------------------------------------------------------------
# Identification — O(1), no tree walking
# ---------------------------------------------------------------------------

def identify_overlay_class(obj):
	"""Determine which overlay class (if any) should apply to *obj*.

	PERFORMANCE: Only checks obj.role and obj.name — never walks
	parents, children, or siblings.  Returns None for most objects.
	"""
	try:
		role = obj.role
	except (COMError, AttributeError, Exception):
		return None

	# Only consider list items and articles — the vast majority of
	# Discord objects are other types and can skip immediately.
	if role not in (
		controlTypes.Role.LISTITEM,
		controlTypes.Role.TREEVIEWITEM,
		controlTypes.Role.ARTICLE,
	):
		return None

	# We can't cheaply determine WHICH region the item is in without
	# walking parents.  Instead, return DiscordMessageItem for
	# ARTICLE elements (Discord messages use role="article" in
	# recent versions) and None for everything else.
	if role == controlTypes.Role.ARTICLE:
		return DiscordMessageItem

	return None
