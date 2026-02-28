# Discord Enhancements Add-on for NVDA
# UIA helper utilities for navigating Discord's accessibility tree
#
# Discord is an Electron (Chromium) app.  Chromium exposes its UI via
# Microsoft UI Automation (UIA).  NVDA wraps UIA elements in its own
# NVDAObject classes, but the standard NVDA child-enumeration methods
# (firstChild/next, obj.children) are UNRELIABLE for Chromium.
#
# This module therefore uses the RAW UIA API via UIAHandler to walk
# the accessibility tree.  A lightweight _UIAElementWrapper class
# provides the same interface that safe_* functions and commands.py
# expect (.name, .role, .states, .setFocus(), .doAction(), etc.)
# without the overhead of full NVDA UIA object creation.

import time
from collections import deque
from comtypes import COMError
from logHandler import log
import api
import controlTypes

# ---------------------------------------------------------------------------
# Raw UIA access  (imported lazily to avoid startup issues)
# ---------------------------------------------------------------------------

_HAS_UIA = False
_walker = None

try:
	import UIAHandler
	_HAS_UIA = True
except ImportError:
	log.warning("Discord addon: UIAHandler not available")


def _get_walker():
	"""Return a cached RawViewWalker for tree enumeration."""
	global _walker
	if not _HAS_UIA:
		return None
	try:
		if _walker is None:
			_walker = UIAHandler.handler.clientObject.RawViewWalker
		return _walker
	except (COMError, AttributeError, Exception):
		return None


def _get_uia_element(obj):
	"""Extract a raw IUIAutomationElement from an NVDA object.

	Only returns an element if the object NATIVELY has one (UIA objects
	or _UIAElementWrapper).  Does NOT use ElementFromHandle, because
	Discord's Chromium puts all names/roles in the IA2 layer — the
	UIA view from the HWND is a skeleton with no names.
	"""
	try:
		elem = getattr(obj, "UIAElement", None)
		if elem is not None:
			return elem
		elem = getattr(obj, "_element", None)
		if elem is not None:
			return elem
	except (COMError, Exception):
		pass

	return None


# ---------------------------------------------------------------------------
# UIA control-type → NVDA Role mapping
# ---------------------------------------------------------------------------

_CT_TO_ROLE = {
	50000: controlTypes.Role.BUTTON,
	50002: controlTypes.Role.CHECKBOX,
	50003: controlTypes.Role.COMBOBOX,
	50004: controlTypes.Role.EDITABLETEXT,
	50005: controlTypes.Role.LINK,
	50006: controlTypes.Role.GRAPHIC,
	50007: controlTypes.Role.LISTITEM,
	50008: controlTypes.Role.LIST,
	50009: controlTypes.Role.MENU,
	50010: controlTypes.Role.MENUBAR,
	50011: controlTypes.Role.MENUITEM,
	50012: controlTypes.Role.PROGRESSBAR,
	50013: controlTypes.Role.RADIOBUTTON,
	50014: controlTypes.Role.SCROLLBAR,
	50015: controlTypes.Role.SLIDER,
	50017: controlTypes.Role.STATUSBAR,
	50018: controlTypes.Role.TABCONTROL,
	50019: controlTypes.Role.TAB,
	50020: controlTypes.Role.STATICTEXT,
	50021: controlTypes.Role.TOOLBAR,
	50022: controlTypes.Role.TOOLTIP,
	50023: controlTypes.Role.TREEVIEW,
	50024: controlTypes.Role.TREEVIEWITEM,
	50025: controlTypes.Role.UNKNOWN,       # Custom
	50026: controlTypes.Role.GROUPING,      # Group
	50028: controlTypes.Role.TABLE,         # DataGrid
	50029: controlTypes.Role.LISTITEM,      # DataItem
	50030: controlTypes.Role.DOCUMENT,
	50031: controlTypes.Role.SPLITBUTTON,
	50032: controlTypes.Role.WINDOW,
	50033: controlTypes.Role.PANE,
	50034: controlTypes.Role.GROUPING,      # Header
	50036: controlTypes.Role.TABLE,
	50037: controlTypes.Role.TITLEBAR,
	50038: controlTypes.Role.SEPARATOR,
}


# ---------------------------------------------------------------------------
# Lightweight wrapper around raw IUIAutomationElement
# ---------------------------------------------------------------------------

class _UIAElementWrapper:
	"""Makes a raw IUIAutomationElement behave like an NVDA object.

	This is MUCH cheaper than creating a full NVDAObjects.UIA.UIA
	instance for every node during a tree walk.  It implements just
	enough of the NVDA-object interface for safe_* functions,
	_iter_children, setFocus, and doAction to work.
	"""

	def __init__(self, element):
		self._element = element
		# Expose UIAElement so _get_uia_element() finds it
		self.UIAElement = element

	# --- Properties used by safe_* functions and commands.py ---

	@property
	def name(self):
		try:
			return self._element.CurrentName or ""
		except (COMError, Exception):
			return ""

	@property
	def role(self):
		try:
			ct = self._element.CurrentControlType
			return _CT_TO_ROLE.get(ct, controlTypes.Role.UNKNOWN)
		except (COMError, Exception):
			return controlTypes.Role.UNKNOWN

	@property
	def states(self):
		result = set()
		try:
			if self._element.CurrentIsKeyboardFocusable:
				result.add(controlTypes.State.FOCUSABLE)
			if not self._element.CurrentIsEnabled:
				result.add(controlTypes.State.UNAVAILABLE)
		except (COMError, Exception):
			pass
		return result

	@property
	def value(self):
		try:
			# UIA ValuePattern value
			return self._element.GetCurrentPropertyValue(30045) or ""
		except (COMError, Exception):
			return ""

	@property
	def description(self):
		try:
			return self._element.GetCurrentPropertyValue(30159) or ""
		except (COMError, Exception):
			return ""

	@property
	def windowClassName(self):
		try:
			return self._element.CurrentClassName or ""
		except (COMError, Exception):
			return ""

	@property
	def UIAAutomationId(self):
		try:
			return self._element.CurrentAutomationId or ""
		except (COMError, Exception):
			return ""

	@property
	def IA2Attributes(self):
		return {}

	@property
	def childCount(self):
		"""Approximate child count (0 means unknown, not necessarily empty)."""
		return 0

	@property
	def parent(self):
		walker = _get_walker()
		if walker is None:
			return None
		try:
			parent_elem = walker.GetParentElement(self._element)
			if parent_elem:
				return _UIAElementWrapper(parent_elem)
		except (COMError, Exception):
			pass
		return None

	@property
	def treeInterceptor(self):
		return None

	@property
	def firstChild(self):
		return None

	@property
	def next(self):
		return None

	@property
	def children(self):
		return list(_iter_children(self))

	# --- Actions ---

	def setFocus(self):
		"""Move keyboard focus to this element."""
		try:
			self._element.SetFocus()
		except (COMError, Exception) as e:
			# Fallback: try creating a real NVDA object
			try:
				from NVDAObjects.UIA import UIA as UIAObj
				obj = UIAObj(UIAElement=self._element)
				obj.setFocus()
			except Exception:
				raise e

	def doAction(self):
		"""Invoke the default action (click a button, etc.)."""
		# Try Invoke pattern (UIA_InvokePatternId = 10000)
		try:
			pattern = self._element.GetCurrentPattern(10000)
			if pattern:
				from comInterfaces.UIAutomationClient import IUIAutomationInvokePattern
				invoke = pattern.QueryInterface(IUIAutomationInvokePattern)
				invoke.Invoke()
				return
		except (COMError, ImportError, Exception):
			pass
		# Try Toggle pattern (UIA_TogglePatternId = 10015)
		try:
			pattern = self._element.GetCurrentPattern(10015)
			if pattern:
				from comInterfaces.UIAutomationClient import IUIAutomationTogglePattern
				toggle = pattern.QueryInterface(IUIAutomationTogglePattern)
				toggle.Toggle()
				return
		except (COMError, ImportError, Exception):
			pass
		# Fallback: create a real NVDA object and delegate
		try:
			from NVDAObjects.UIA import UIA as UIAObj
			obj = UIAObj(UIAElement=self._element)
			obj.doAction()
			return
		except Exception:
			pass
		raise COMError(-1, "No actionable pattern available", ())

	def __repr__(self):
		return "<Wrapper name=%r role=%s>" % (self.name[:40], self.role)


# ---------------------------------------------------------------------------
# Safe property access  (works on NVDA objects AND _UIAElementWrapper)
# ---------------------------------------------------------------------------

def safe_name(obj):
	try:
		return obj.name or ""
	except (COMError, AttributeError, Exception):
		return ""


def safe_role(obj):
	try:
		return obj.role
	except (COMError, AttributeError, Exception):
		return controlTypes.Role.UNKNOWN


def safe_states(obj):
	try:
		return obj.states or set()
	except (COMError, AttributeError, Exception):
		return set()


def safe_value(obj):
	try:
		return obj.value or ""
	except (COMError, AttributeError, Exception):
		return ""


def safe_description(obj):
	try:
		return obj.description or ""
	except (COMError, AttributeError, Exception):
		return ""


def safe_automation_id(obj):
	try:
		return obj.UIAAutomationId or ""
	except (COMError, AttributeError, Exception):
		return ""


def safe_class_name(obj):
	try:
		return obj.windowClassName or ""
	except (COMError, AttributeError, Exception):
		return ""


def safe_child_count(obj):
	try:
		return obj.childCount or 0
	except (COMError, AttributeError, Exception):
		return 0


def safe_ia2_attrs(obj):
	try:
		attrs = obj.IA2Attributes
		if isinstance(attrs, dict):
			return attrs
	except (COMError, AttributeError, Exception):
		pass
	return {}


# ---------------------------------------------------------------------------
# Focusing helpers
# ---------------------------------------------------------------------------

def focus_element(obj):
	"""Move focus to *obj* or its first focusable descendant.

	IA2 landmark/region containers (Servers sidebar, User area, etc.)
	can't receive keyboard focus directly — they're structural.
	This function tries:
	  1. obj.setFocus() directly
	  2. First focusable child via simpleFirstChild/simpleNext
	  3. Set the NVDA navigator object + review position

	Returns True if focus was moved, False otherwise.
	"""
	# Try 1: direct focus
	try:
		states = safe_states(obj)
		if controlTypes.State.FOCUSABLE in states:
			obj.setFocus()
			return True
	except (COMError, Exception):
		pass

	# Try 2: find first focusable child (shallow, max 20 children)
	count = 0
	for child in _iter_children(obj):
		count += 1
		if count > 20:
			break
		try:
			child_states = safe_states(child)
			if controlTypes.State.FOCUSABLE in child_states:
				child.setFocus()
				return True
		except (COMError, Exception):
			continue

	# Try 3: set navigator object (this always works for NVDA objects)
	try:
		import review
		api.setNavigatorObject(obj)
		review.handleCaretMove(obj)
		# Speak the navigator object so the user knows where they are
		return True
	except (COMError, Exception):
		pass

	# Try 4: brute force setFocus on the container itself
	try:
		obj.setFocus()
		return True
	except (COMError, Exception):
		pass

	return False


# ---------------------------------------------------------------------------
# Child iteration  —  simpleNav for IA2, UIA walker for wrappers
# ---------------------------------------------------------------------------

def _iter_children(obj):
	"""Yield children of *obj*.

	Discord (Chromium) uses IA2.  Named elements are ONLY accessible
	via NVDA's simpleFirstChild/simpleNext ("simple navigation").
	The raw DOM (firstChild/next) and obj.children return unnamed
	div containers that waste time.  The UIA walker is only useful
	for _UIAElementWrapper objects (which we rarely use).

	Removing the UIA walker check for regular IA2 objects saves
	~20ms per element — a huge speedup when walking hundreds of nodes.
	"""
	# --- _UIAElementWrapper: use raw UIA walker ---
	if isinstance(obj, _UIAElementWrapper):
		element = getattr(obj, '_element', None)
		walker = _get_walker()
		if element is not None and walker is not None:
			try:
				child_elem = walker.GetFirstChildElement(element)
				while child_elem is not None:
					yield _UIAElementWrapper(child_elem)
					try:
						child_elem = walker.GetNextSiblingElement(child_elem)
					except (COMError, Exception):
						break
			except (COMError, Exception):
				pass
		return

	# --- All NVDA objects: simpleFirstChild / simpleNext ONLY ---
	try:
		child = obj.simpleFirstChild
		if child is not None:
			visited = set()
			while child is not None:
				cid = id(child)
				if cid in visited:
					break
				visited.add(cid)
				yield child
				try:
					child = child.simpleNext
				except (COMError, AttributeError, Exception):
					break
	except (COMError, AttributeError, Exception):
		pass


# ---------------------------------------------------------------------------
# Raw UIA child iteration  (element-level, for dump_tree efficiency)
# ---------------------------------------------------------------------------

def _raw_uia_children(element):
	"""Yield child IUIAutomationElements using raw UIA walker."""
	walker = _get_walker()
	if walker is None:
		return
	try:
		child = walker.GetFirstChildElement(element)
		while child is not None:
			yield child
			try:
				child = walker.GetNextSiblingElement(child)
			except (COMError, Exception):
				break
	except (COMError, Exception):
		pass


def _raw_uia_name(element):
	try:
		return element.CurrentName or ""
	except (COMError, Exception):
		return ""


def _raw_uia_control_type(element):
	try:
		return element.CurrentControlType
	except (COMError, Exception):
		return 0


def _raw_uia_class_name(element):
	try:
		return element.CurrentClassName or ""
	except (COMError, Exception):
		return ""


def _raw_uia_automation_id(element):
	try:
		return element.CurrentAutomationId or ""
	except (COMError, Exception):
		return ""


def _raw_uia_aria_role(element):
	"""Read the AriaRole property (UIA_AriaRolePropertyId = 30101)."""
	try:
		return element.GetCurrentPropertyValue(30101) or ""
	except (COMError, Exception):
		return ""


def _raw_uia_landmark_type(element):
	"""Read the LandmarkType property (UIA_LandmarkTypePropertyId = 30157)."""
	try:
		return element.GetCurrentPropertyValue(30157) or 0
	except (COMError, Exception):
		return 0


# ---------------------------------------------------------------------------
# Time-bounded, depth-limited tree walk  (DFS)
# ---------------------------------------------------------------------------

WALK_TIMEOUT = 10.0  # seconds
WALK_MAX_DEPTH = 10

# Roles whose children should NEVER be explored during tree walks.
# These elements are "leaves" — exploring their children wastes
# ~50ms per COM call with zero useful results.
_LEAF_ROLES = frozenset({
	controlTypes.Role.BUTTON,
	controlTypes.Role.LINK,
	controlTypes.Role.STATICTEXT,
	controlTypes.Role.TOGGLEBUTTON,
	controlTypes.Role.SPLITBUTTON,
	controlTypes.Role.GRAPHIC,
	controlTypes.Role.SEPARATOR,
	controlTypes.Role.MENUITEM,
	controlTypes.Role.CHECKBOX,
	controlTypes.Role.RADIOBUTTON,
	controlTypes.Role.SLIDER,
	controlTypes.Role.PROGRESSBAR,
	controlTypes.Role.SCROLLBAR,
	controlTypes.Role.TOOLTIP,
	controlTypes.Role.HEADING,
})


def _role_label(role):
	"""Return a human-readable role label like 'BUTTON' or 'LANDMARK'."""
	try:
		name = getattr(role, 'name', None)
		if name:
			return name
	except Exception:
		pass
	return str(role)


def walk_descendants(root, max_depth=WALK_MAX_DEPTH, timeout=WALK_TIMEOUT):
	"""Yield descendant NVDA objects in breadth-first order.

	BFS is critical for Discord because major regions (servers,
	channels, messages) are siblings at the same depth.  DFS would
	dive into 'Servers' (128 children) before reaching 'channels'.

	Leaf-type roles (buttons, links, etc.) have their children
	pruned to avoid wasteful COM calls.
	"""
	if root is None:
		return
	start = time.time()

	# BFS using deque
	queue = deque()
	queue.append((root, 0))
	while queue:
		if time.time() - start > timeout:
			return
		obj, depth = queue.popleft()
		if depth > 0:
			yield obj
		if depth < max_depth:
			# Prune: don't explore children of leaf-type elements
			if depth > 0 and safe_role(obj) in _LEAF_ROLES:
				continue
			for child in _iter_children(obj):
				queue.append((child, depth + 1))


# ---------------------------------------------------------------------------
# Name matching
# ---------------------------------------------------------------------------

def _name_matches(obj_name, patterns):
	"""Case-insensitive substring match against any pattern."""
	if not obj_name:
		return False
	lower = obj_name.lower()
	for pat in patterns:
		if pat.lower() in lower:
			return True
	return False


def _name_matches_message(obj_name):
	"""Match _MESSAGE_NAMES but reject _MESSAGE_EXCLUDE hits."""
	if not obj_name:
		return False
	lower = obj_name.lower()
	# Reject exclusion patterns first
	for excl in _MESSAGE_EXCLUDE:
		if excl in lower:
			return False
	# Then check positive patterns
	for pat in _MESSAGE_NAMES:
		if pat.lower() in lower:
			return True
	return False


# ---------------------------------------------------------------------------
# Getting the Discord foreground and content root
# ---------------------------------------------------------------------------

def get_foreground():
	"""Return the Discord foreground window object, or None."""
	try:
		return api.getForegroundObject()
	except (COMError, Exception):
		return None


def get_content_root():
	"""Return the root of Discord's web content area.

	Strategy 1: Walk UP from focus, return deepest DOCUMENT/APPLICATION,
	            or the highest non-WINDOW ancestor.
	Strategy 2: Walk DOWN from foreground looking for the Chromium
	            render-widget host.
	Strategy 3: Fall back to the foreground window.
	"""
	# --- Strategy 1: walk UP from focus ---
	try:
		focus = api.getFocusObject()
		if focus:
			chain = []
			obj = focus
			depth = 0
			while obj and depth < 40:
				cls = safe_class_name(obj)
				role = safe_role(obj)
				if cls == "Chrome_WidgetWin_1":
					break
				if role == controlTypes.Role.WINDOW:
					break
				chain.append(obj)
				try:
					obj = obj.parent
				except (COMError, Exception):
					break
				depth += 1

			if chain:
				# Prefer a DOCUMENT, APPLICATION, or PANE in the chain
				for ancestor in reversed(chain):
					r = safe_role(ancestor)
					if r in (
						controlTypes.Role.DOCUMENT,
						controlTypes.Role.APPLICATION,
						controlTypes.Role.PANE,
					):
						return ancestor
				# Otherwise return the highest ancestor
				return chain[-1]
	except (COMError, Exception):
		pass

	# --- Strategy 2: walk DOWN from foreground ---
	fg = get_foreground()
	if fg is None:
		return None
	try:
		for child in _iter_children(fg):
			cls = safe_class_name(child)
			if "Chrome_RenderWidget" in cls or "Intermediate" in cls:
				for grandchild in _iter_children(child):
					return grandchild
				return child
			role = safe_role(child)
			if role in (
				controlTypes.Role.DOCUMENT,
				controlTypes.Role.APPLICATION,
				controlTypes.Role.PANE,
				controlTypes.Role.GROUPING,
			):
				return child
	except (COMError, Exception):
		pass

	# --- Strategy 3: foreground itself ---
	return fg


# ---------------------------------------------------------------------------
# Element finders
# ---------------------------------------------------------------------------

# Broad name patterns (case-insensitive substring)

_SERVER_NAMES = [
	"servers", "guilds", "guild", "server navigation",
	"server sidebar", "guild sidebar",
]

_CHANNEL_NAMES = [
	"channels", "channel list", "private channels",
	"text channels", "voice channels", "channel sidebar",
	"text & voice",
]

_MESSAGE_NAMES = [
	"messages in", "chat messages", "message list", "messages",
]

# Names that contain 'messages' but are sidebar navigation, NOT chat
_MESSAGE_EXCLUDE = [
	"direct messages", "private messages", "group messages",
	"message requests",
]

_MESSAGE_INPUT_NAMES = [
	"message @", "message #", "message ",
	"chat message", "type a message", "reply to",
]

_MEMBERS_NAMES = [
	"members", "member list", "people",
]

_USER_AREA_NAMES = [
	"user area", "user panel", "account",
	"user status", "status and settings",
]

_ACTIVE_NOW_NAMES = [
	"active now",
]

_DISCONNECT_NAMES = [
	"disconnect", "leave voice", "leave call",
	"disconnect quietly",
]

_PINNED_NAMES = [
	"pinned messages", "pinned", "pins",
]

_THREADS_NAMES = [
	"threads", "thread list", "show threads",
]


def _log_missing(what):
	log.debugWarning(
		"Discord element not found: '%s'. "
		"Discord's UI structure may have changed.", what,
	)


def _find_by_name(patterns, root=None, roles=None, max_depth=WALK_MAX_DEPTH):
	"""Find the first descendant whose name matches a pattern.

	Key optimisation: safe_role() is ONLY called when the name matches.
	This eliminates one COM call (~50ms) per non-matching element.
	With 50+ siblings at depth 1, this saves ~2.5 seconds.
	"""
	if root is None:
		root = get_content_root()
	if root is None:
		return None

	start = time.time()
	first_name_match = None
	current_level = [root]

	for _depth in range(max_depth):
		if time.time() - start > WALK_TIMEOUT:
			break
		next_level = []
		for parent in current_level:
			if time.time() - start > WALK_TIMEOUT:
				break
			for child in _iter_children(parent):
				name = safe_name(child)
				if _name_matches(name, patterns):
					if not roles:
						return child
					role = safe_role(child)  # only check role when name matches
					if role in roles:
						return child
					if first_name_match is None:
						first_name_match = child
				next_level.append(child)

		if first_name_match is not None:
			return first_name_match

		current_level = next_level
		if not current_level:
			break

	return first_name_match


def _find_by_role(target_roles, root=None, max_depth=WALK_MAX_DEPTH):
	"""Find the first named descendant with one of the target roles."""
	if root is None:
		root = get_content_root()
	if root is None:
		return None
	for obj in walk_descendants(root, max_depth=max_depth):
		if safe_role(obj) in target_roles and safe_name(obj):
			return obj
	return None


# ---------------------------------------------------------------------------
# Depth-1 cache  —  the single most important optimisation
# ---------------------------------------------------------------------------
# Scanning depth-1 children of the content root requires one COM call
# per element (~50-100ms each).  With 50+ depth-1 children, this takes
# 5-7 seconds.  Caching the results for 3 seconds means all finder
# functions can reuse the same scan.

_d1_cache = None      # depth-1 children: [(name, obj), ...]
_d2_cache = None      # depth-2 children: [(name, obj, parent_name), ...]
_d1_cache_time = 0.0
_D1_CACHE_TTL = 3.0  # seconds


def _get_depth1(root=None):
	"""Return cached (name, obj) tuples for all depth-1 children.

	The cache is valid for 3 seconds.  After that, a fresh scan is
	performed.  This eliminates the #1 performance bottleneck: every
	finder function was independently scanning the same 50+ children.
	"""
	global _d1_cache, _d2_cache, _d1_cache_time
	now = time.time()
	if _d1_cache is not None and now - _d1_cache_time < _D1_CACHE_TTL:
		return _d1_cache

	if root is None:
		root = get_content_root()
	if root is None:
		_d1_cache = []
		_d2_cache = []
		_d1_cache_time = now
		return _d1_cache

	children = []
	d2_children = []
	for child in _iter_children(root):
		name = safe_name(child)
		children.append((name, child))
		# Also scan depth 2 for non-leaf elements
		if name and safe_role(child) not in _LEAF_ROLES:
			for grandchild in _iter_children(child):
				gc_name = safe_name(grandchild)
				if gc_name:
					d2_children.append((gc_name, grandchild, name))

	_d1_cache = children
	_d2_cache = d2_children
	_d1_cache_time = now
	log.debug("Discord cache: %d d1 + %d d2 children scanned in %.1fs",
			  len(children), len(d2_children), time.time() - now)
	return _d1_cache


def _get_depth2():
	"""Return cached depth-2 children. Triggers depth-1 scan if needed."""
	if _d2_cache is None or time.time() - _d1_cache_time >= _D1_CACHE_TTL:
		_get_depth1()  # refresh both caches
	return _d2_cache or []


def _d1_find_name(patterns):
	"""Find the first depth-1 child whose name matches *patterns*."""
	for name, child in _get_depth1():
		if _name_matches(name, patterns):
			return child
	return None


def _d2_find_name(patterns):
	"""Find at depth-1 first, then depth-2 if not found."""
	result = _d1_find_name(patterns)
	if result:
		return result
	for name, child, _parent in _get_depth2():
		if _name_matches(name, patterns):
			return child
	return None


def _d1_find_landmark(landmark_text):
	"""Find a depth-1 child whose name contains *landmark_text*."""
	lower = landmark_text.lower()
	for name, child in _get_depth1():
		if name and lower in name.lower():
			return child
	return None


# --- Region finders ---

def find_server_list(root=None):
	"""'Servers sidebar' is at depth 1."""
	result = _d2_find_name(_SERVER_NAMES)
	if not result:
		_log_missing("server list")
	return result


def find_channel_list(root=None):
	"""'Channels' list is inside the '(server)' landmark at depth 2."""
	# Strategy 1: check depth-1 and depth-2 cache
	result = _d2_find_name(_CHANNEL_NAMES)
	if result:
		return result
	# Strategy 2: search within the server landmark (deeper)
	server = _d1_find_landmark("(server)")
	if server:
		result = _find_by_name(_CHANNEL_NAMES, server, max_depth=4)
		if result:
			return result
	_log_missing("channel list")
	return None


def find_message_list(root=None):
	"""Find the actual chat message list (not the DM sidebar).

	Strategy 1: Look inside landmarks (server, DM conversation, etc.)
	            where the real message area lives.
	Strategy 2: Search depth-1/2 cache with role + exclusion filtering.
	Strategy 3: Fall back to any depth-1 container that looks like it
	            holds messages (by checking children for message-like
	            content).

	The _MESSAGE_EXCLUDE list prevents matching sidebar navigation
	elements like 'Direct Messages' or 'Private Messages'.
	"""
	_MSG_ROLES = frozenset({
		controlTypes.Role.LIST,
		controlTypes.Role.GROUPING,
		controlTypes.Role.TREEVIEW,
		controlTypes.Role.DOCUMENT,
		controlTypes.Role.SECTION,
		controlTypes.Role.APPLICATION,
	})

	# Strategy 1: Look inside landmarks at depth 1
	# In server view: "ServerName (server)" landmark
	# In DM view: there may be a different landmark or a large
	# container that holds the chat area
	for name, child in _get_depth1():
		if not name:
			continue
		# Skip leaf roles — they can't contain a message list
		if safe_role(child) in _LEAF_ROLES:
			continue
		# Search inside this depth-1 container
		for gc_name, gc_child, parent_name in _get_depth2():
			if parent_name and parent_name == name:
				if _name_matches_message(gc_name):
					role = safe_role(gc_child)
					if role in _MSG_ROLES:
						return gc_child

	# Strategy 1b: deeper scan inside the server landmark specifically
	server = _d1_find_landmark("(server)")
	if server:
		result = _find_by_name(_MESSAGE_NAMES, server, roles=_MSG_ROLES, max_depth=5)
		if result:
			n = safe_name(result) or ""
			if not any(excl in n.lower() for excl in _MESSAGE_EXCLUDE):
				return result

	# Strategy 1c: deeper scan inside ALL non-leaf depth-1 containers
	for name, child in _get_depth1():
		if not name:
			continue
		role = safe_role(child)
		if role in _LEAF_ROLES:
			continue
		# Skip known sidebar lists
		if any(excl in name.lower() for excl in _MESSAGE_EXCLUDE):
			continue
		# Skip the server list itself
		if _name_matches(name, _SERVER_NAMES):
			continue
		result = _find_by_name(_MESSAGE_NAMES, child, roles=_MSG_ROLES, max_depth=4)
		if result:
			n = safe_name(result) or ""
			if not any(excl in n.lower() for excl in _MESSAGE_EXCLUDE):
				return result

	# Strategy 2: depth-1 / depth-2 with exclusion + role filtering
	for name, child in _get_depth1():
		if _name_matches_message(name):
			role = safe_role(child)
			if role in _MSG_ROLES:
				return child
	for name, child, _parent in _get_depth2():
		if _name_matches_message(name):
			role = safe_role(child)
			if role in _MSG_ROLES:
				return child

	_log_missing("message list")
	return None


def find_message_input(root=None):
	"""Find the chat message input box.

	The input is typically an EDITABLETEXT element inside the
	'(channel)' landmark with a name like 'Message @User' or
	'Message #channel'.  We search by role first (most reliable),
	then fall back to name matching.
	"""
	_EDIT_ROLES = frozenset({
		controlTypes.Role.EDITABLETEXT,
		controlTypes.Role.DOCUMENT,
	})

	# Strategy 1: Look inside (channel) landmark for an editable text
	channel = _d1_find_landmark("(channel)")
	if channel:
		# Check by name first (fast, uses cache)
		for gc_name, gc_child, parent_name in _get_depth2():
			if parent_name and "(channel)" in parent_name.lower():
				if _name_matches(gc_name, _MESSAGE_INPUT_NAMES):
					role = safe_role(gc_child)
					if role in _EDIT_ROLES:
						return gc_child
		# Deeper scan by name + role
		result = _find_by_name(_MESSAGE_INPUT_NAMES, channel, roles=_EDIT_ROLES, max_depth=5)
		if result:
			return result
		# Any focusable edit control inside (channel)
		for obj in walk_descendants(channel, max_depth=5, timeout=5.0):
			role = safe_role(obj)
			if role in _EDIT_ROLES:
				states = safe_states(obj)
				if controlTypes.State.FOCUSABLE in states:
					return obj

	# Strategy 2: Look inside (server) landmark
	server = _d1_find_landmark("(server)")
	if server:
		result = _find_by_name(_MESSAGE_INPUT_NAMES, server, roles=_EDIT_ROLES, max_depth=5)
		if result:
			return result
		# Any focusable edit control inside (server)
		for obj in walk_descendants(server, max_depth=5, timeout=5.0):
			role = safe_role(obj)
			if role in _EDIT_ROLES:
				states = safe_states(obj)
				if controlTypes.State.FOCUSABLE in states:
					return obj

	# Strategy 3: Search all depth-2 elements by name + role
	for gc_name, gc_child, _parent in _get_depth2():
		if _name_matches(gc_name, _MESSAGE_INPUT_NAMES):
			role = safe_role(gc_child)
			if role in _EDIT_ROLES:
				return gc_child

	# Strategy 4: Any focusable edit control anywhere
	content = get_content_root()
	if content:
		for obj in walk_descendants(content, max_depth=5, timeout=5.0):
			role = safe_role(obj)
			if role in _EDIT_ROLES:
				states = safe_states(obj)
				if controlTypes.State.FOCUSABLE in states:
					# Skip search boxes
					name = safe_name(obj) or ""
					if "search" not in name.lower():
						return obj

	_log_missing("message input")
	return None


def find_members_list(root=None):
	"""Members list may be at depth 1-2 or inside the server landmark."""
	result = _d2_find_name(_MEMBERS_NAMES)
	if result:
		return result
	server = _d1_find_landmark("(server)")
	if server:
		result = _find_by_name(_MEMBERS_NAMES, server, max_depth=4)
		if result:
			return result
	_log_missing("members list")
	return None


def find_user_area(root=None):
	"""'User status and settings' is at depth 1."""
	result = _d2_find_name(_USER_AREA_NAMES)
	if not result:
		_log_missing("user area")
	return result


def find_active_now(root=None):
	"""Active Now may be at depth 1-2, or inside a page landmark."""
	result = _d2_find_name(_ACTIVE_NOW_NAMES)
	if result:
		return result
	# Deeper search through all depth-1 landmarks
	for name, child in _get_depth1():
		if name and safe_role(child) not in _LEAF_ROLES:
			result = _find_by_name(_ACTIVE_NOW_NAMES, child, max_depth=3)
			if result:
				return result
	_log_missing("Active Now")
	return None


def find_all_areas(root=None):
	"""Find all major Discord areas using the depth-1 cache.

	Pass 1: Check cached depth-1 children (instant if already cached).
	Pass 2: Search within the server landmark for missing areas.

	Returns a list of (label, obj) tuples.
	"""
	# Core areas to search for at depth 1
	area_defs = [
		("Server list", _SERVER_NAMES),
		("User area", _USER_AREA_NAMES),
	]
	found = {}  # area_label -> obj
	server_landmark = None

	# Pass 1: depth-1 cache scan (instant after first call)
	for name, child in _get_depth1():
		if not name:
			continue
		if "(server)" in name.lower():
			server_landmark = child
		for area_label, patterns in area_defs:
			if area_label not in found and _name_matches(name, patterns):
				found[area_label] = child

	# Pass 1b: also check depth-2 cache for areas at depth 1
	for name, child, _parent in _get_depth2():
		for area_label, patterns in area_defs:
			if area_label not in found and _name_matches(name, patterns):
				found[area_label] = child

	# Always include the server landmark as a navigable area
	if server_landmark is not None and "Server" not in found:
		sname = safe_name(server_landmark)
		# Use the server name (e.g. "The Insane Comedy Group")
		label = sname.replace(" (server)", "") if sname else "Server"
		found[label] = server_landmark

	# Pass 2: search inside server landmark for deeper areas (use d2 cache)
	if server_landmark is not None:
		inner_defs = [
			("Channel list", _CHANNEL_NAMES),
			("Messages", _MESSAGE_NAMES),
			("Members", _MEMBERS_NAMES),
		]
		sname = safe_name(server_landmark) or ""
		# Check d2 cache entries whose parent matches the server landmark
		for gc_name, gc_child, parent_name in _get_depth2():
			if parent_name and parent_name == sname:
				for area_label, patterns in inner_defs:
					if area_label not in found and _name_matches(gc_name, patterns):
						found[area_label] = gc_child
		# If still missing, do a live scan
		missing_inner = [
			(label, pats) for label, pats in inner_defs
			if label not in found
		]
		if missing_inner:
			for child in _iter_children(server_landmark):
				name = safe_name(child)
				if not name:
					continue
				for area_label, patterns in missing_inner:
					if area_label not in found and _name_matches(name, patterns):
						found[area_label] = child
				if all(a in found for a, _ in missing_inner):
					break

	# Return in a fixed, logical order
	ordered_labels = [
		"Server list", "Channel list", "Messages",
		"Members", "User area",
	]
	result = []
	# Add server landmark right after server list
	for label in ordered_labels:
		if label in found:
			result.append((label, found[label]))
		if label == "Server list":
			# Insert server-specific landmark
			for k, v in found.items():
				if k not in ordered_labels:
					result.append((k, v))
	return result


# --- Specific element finders ---

def find_button_by_name(name_patterns, root=None):
	"""Find a button by name. Checks depth-1 cache first."""
	# Fast pass: check depth-1 cache
	for name, child in _get_depth1():
		if _name_matches(name, name_patterns):
			return child
	# Slower pass: search within server landmark
	server = _d1_find_landmark("(server)")
	if server:
		result = _find_by_name(name_patterns, server, max_depth=4)
		if result:
			return result
	return None


def find_disconnect_button(root=None):
	btn = find_button_by_name(_DISCONNECT_NAMES, root)
	if not btn:
		log.debugWarning("Disconnect button not found")
	return btn


def find_pinned_messages_button(root=None):
	btn = find_button_by_name(_PINNED_NAMES, root)
	if not btn:
		log.debugWarning("Pinned Messages button not found")
	return btn


def find_threads_button(root=None):
	btn = find_button_by_name(_THREADS_NAMES, root)
	if not btn:
		log.debugWarning("Threads button not found")
	return btn


def find_typing_indicator(root=None):
	# Check depth-1 first
	result = _d1_find_name(["typing"])
	if result:
		return result
	# Then check inside server landmark
	server = _d1_find_landmark("(server)")
	if server:
		return _find_by_name(["typing"], server, max_depth=4)
	return None


def get_voice_connection_info():
	"""Return a dict with voice connection details.

	When connected to a voice channel, Discord shows several sibling
	elements inside the user area:
	  - Latency: e.g. '78 ms'  (role STATICTEXT)
	  - Status:  e.g. 'Voice Details Voice Connected'  (role BUTTON)
	  - Channel: e.g. 'The Assabet River / Server Name'  (role HEADING)
	  - Extras:  e.g. 'Noise Suppression powered by Krisp'

	Returns a dict with keys: 'latency', 'status', 'channel', 'extras'.
	All values are strings (empty if not found).
	Returns None if no voice connection is detected at all.
	"""
	import re

	info = {"latency": "", "status": "", "channel": "", "extras": []}
	found_anything = False

	# Scan depth-2 children of the user area
	user_area_name = None
	for name, child in _get_depth1():
		if name and _name_matches(name, _USER_AREA_NAMES):
			user_area_name = name
			break

	if not user_area_name:
		return None

	for gc_name, gc_child, parent_name in _get_depth2():
		if parent_name != user_area_name:
			continue
		if not gc_name:
			continue
		lower = gc_name.lower()

		# Latency: matches patterns like '78 ms', '120ms', '45 ms'
		if re.search(r'\d+\s*ms', lower):
			info["latency"] = gc_name.strip()
			found_anything = True
		# Voice connection status button
		elif "voice" in lower and ("connected" in lower or "details" in lower):
			info["status"] = gc_name.strip()
			found_anything = True
		# Channel / server name (often contains ' / ')
		elif "/" in gc_name and not any(
			kw in lower for kw in ("mute", "deafen", "settings", "profile",
									"camera", "screen", "activity", "soundboard")
		):
			info["channel"] = gc_name.strip()
			found_anything = True
		# Noise suppression or similar extras
		elif any(kw in lower for kw in ("noise", "suppression", "krisp", "echo")):
			info["extras"].append(gc_name.strip())
			found_anything = True

	if found_anything:
		return info

	# Fallback: check depth-1 for voice connection elements
	for name, child in _get_depth1():
		if not name:
			continue
		lower = name.lower()
		if "voice connected" in lower or "voice details" in lower:
			info["status"] = name.strip()
			return info

	log.debugWarning("Voice connection info not found")
	return None


# ---------------------------------------------------------------------------
# Data extraction helpers
# ---------------------------------------------------------------------------

def get_messages(message_list=None):
	if message_list is None:
		message_list = find_message_list()
	if message_list is None:
		return []
	messages = []
	for child in _iter_children(message_list):
		role = safe_role(child)
		if role in (controlTypes.Role.LISTITEM, controlTypes.Role.GROUPING,
					controlTypes.Role.ARTICLE, controlTypes.Role.TREEVIEWITEM):
			messages.append(child)
		else:
			name = safe_name(child)
			if name:
				messages.append(child)
	return messages


def read_message_content(msg_obj):
	name = safe_name(msg_obj)
	if name:
		return name
	parts = []
	for child in _iter_children(msg_obj):
		child_name = safe_name(child)
		if child_name:
			parts.append(child_name)
		else:
			val = safe_value(child)
			if val:
				parts.append(val)
	return " ".join(parts) if parts else "(empty message)"


def get_server_items(server_list=None):
	if server_list is None:
		server_list = find_server_list()
	if server_list is None:
		return []
	items = []
	for child in _iter_children(server_list):
		role = safe_role(child)
		if role in (controlTypes.Role.TREEVIEWITEM, controlTypes.Role.LISTITEM,
					controlTypes.Role.BUTTON, controlTypes.Role.LINK):
			items.append(child)
		else:
			name = safe_name(child)
			if name:
				items.append(child)
	return items


def has_voice_activity(server_obj):
	"""Check if a server icon has voice/live activity indicators.

	Discord lists voice participants as child elements of the server
	item — they're just usernames like '343florida2014'.  We detect
	this by filtering out known non-participant children (badges,
	the server name itself, etc.) and checking if any remain.
	"""
	server_name = (safe_name(server_obj) or "").lower()

	# Patterns for children that are NOT voice participants
	_BADGE_PATTERNS = (
		"new", "mention", "mentions", "unread", "level",
	)

	participant_names = []
	for child in _iter_children(server_obj):
		child_name = safe_name(child) or ""
		if not child_name:
			continue
		lower = child_name.lower().strip()
		if not lower:
			continue

		# Skip if it's the server name itself
		if lower == server_name or server_name.startswith(lower):
			continue

		# Skip known badge / status text
		if any(bp in lower for bp in _BADGE_PATTERNS):
			continue

		# Skip purely numeric text like "7 of 128"
		if lower.replace(" ", "").replace("of", "").isdigit():
			continue

		# Skip very short text (single chars, icons)
		if len(lower) < 2:
			continue

		# This looks like a voice participant username
		participant_names.append(child_name)

	return len(participant_names) > 0


def get_voice_participants_from_server(server_obj):
	"""Return list of voice participant names from a server item."""
	server_name = (safe_name(server_obj) or "").lower()

	_BADGE_PATTERNS = (
		"new", "mention", "mentions", "unread", "level",
	)

	participants = []
	for child in _iter_children(server_obj):
		child_name = safe_name(child) or ""
		if not child_name:
			continue
		lower = child_name.lower().strip()
		if not lower:
			continue
		if lower == server_name or server_name.startswith(lower):
			continue
		if any(bp in lower for bp in _BADGE_PATTERNS):
			continue
		if lower.replace(" ", "").replace("of", "").isdigit():
			continue
		if len(lower) < 2:
			continue
		participants.append(child_name)

	return participants


def get_voice_participants(channel_obj):
	participants = []
	for child in _iter_children(channel_obj):
		name = safe_name(child)
		if not name:
			continue
		role = safe_role(child)
		if role in (controlTypes.Role.LISTITEM, controlTypes.Role.BUTTON,
					controlTypes.Role.TREEVIEWITEM):
			status_parts = [name]
			for sub in _iter_children(child):
				sub_name = safe_name(sub)
				if sub_name:
					lower = sub_name.lower()
					if "mute" in lower:
						status_parts.append("muted")
					elif "deafen" in lower:
						status_parts.append("deafened")
			participants.append(" ".join(status_parts))
	return participants


def find_unread_marker(message_list=None):
	if message_list is None:
		message_list = find_message_list()
	if message_list is None:
		return -1
	for i, child in enumerate(_iter_children(message_list)):
		name = safe_name(child)
		if name:
			lower = name.lower()
			if "new" in lower or "unread" in lower:
				role = safe_role(child)
				if role in (controlTypes.Role.SEPARATOR, controlTypes.Role.GROUPING,
							controlTypes.Role.STATICTEXT, controlTypes.Role.HEADING):
					return i
	return -1


def get_all_buttons(root=None):
	"""Get all named buttons. Uses cache for depth-1 buttons."""
	seen = set()
	buttons = []

	# Fast: depth-1 buttons from cache
	for name, child in _get_depth1():
		if name and name not in seen:
			# Check class name for buttons (avoids safe_role COM call)
			cls_name = type(child).__name__
			if "Button" in cls_name or "Toggle" in cls_name or "Split" in cls_name:
				seen.add(name)
				buttons.append((name, child))

	# Also check inside server landmark
	server = _d1_find_landmark("(server)")
	if server:
		for child in _iter_children(server):
			cls_name = type(child).__name__
			if "Button" in cls_name or "Toggle" in cls_name or "Split" in cls_name:
				name = safe_name(child)
				if name and name not in seen:
					seen.add(name)
					buttons.append((name, child))
					if len(buttons) >= 50:
						break

	return buttons


def get_channel_topic(root=None):
	result = _find_by_name(["topic", "description"], root, max_depth=8)
	if result:
		return safe_name(result)
	return ""


def get_window_context(root=None):
	"""Gather comprehensive status information from Discord's UI.

	Scans the accessibility tree for:
	- Server / channel name (from window title)
	- Mute / deafen status (buttons in user area)
	- Voice channel (from voice connection info)
	- Badge counts (unread indicators on servers)
	- Alerts (visible notification elements)

	Returns a dict suitable for building an enhanced title.
	"""
	import re

	context = {
		"server": "",
		"channel": "",
		"dm_name": "",
		"voice_channel": "",
		"muted": False,
		"deafened": False,
		"badges": [],
		"alerts": [],
	}
	if root is None:
		root = get_foreground()
	if root is None:
		return context

	# --- Parse the window title for server/channel ---
	win_name = safe_name(root) or ""
	if win_name:
		# Discord titles look like:
		#   "#channel — Server — Discord"
		#   "@username — Discord"
		#   "Server — Discord"
		#   "Discord"
		cleaned = win_name
		# Strip trailing " — Discord" or " - Discord"
		for sep in (" \u2014 Discord", " - Discord"):
			if cleaned.endswith(sep):
				cleaned = cleaned[:-len(sep)]
				break
		# Now split remaining by " — " or " - "
		for sep in (" \u2014 ", " - "):
			if sep in cleaned:
				parts = cleaned.split(sep, 1)
				context["channel"] = parts[0].strip()
				context["server"] = parts[1].strip() if len(parts) > 1 else ""
				break
		else:
			# No separator — might be just a server name or DM
			if cleaned and cleaned.lower() != "discord":
				if cleaned.startswith("@"):
					context["dm_name"] = cleaned
				else:
					context["server"] = cleaned

	# --- Scan user area for mute/deafen status ---
	try:
		user_area_name = None
		for name, child in _get_depth1():
			if name and _name_matches(name, _USER_AREA_NAMES):
				user_area_name = name
				break

		if user_area_name:
			for gc_name, gc_child, parent_name in _get_depth2():
				if parent_name != user_area_name:
					continue
				if not gc_name:
					continue
				lower = gc_name.lower()

				# Mute button: "Mute" or "Unmute" or "Muted"
				if "mute" in lower and "unmute" not in lower:
					# Check if it indicates currently muted
					# The button label is "Mute" when NOT muted, "Unmute" when muted
					# But some Discord versions show "Muted"
					if "unmute" in lower or "muted" in lower:
						context["muted"] = True
				elif "unmute" in lower:
					context["muted"] = True

				# Deafen button: similar logic
				if "deafen" in lower:
					if "undeafen" in lower or "deafened" in lower:
						context["deafened"] = True
				elif "undeafen" in lower:
					context["deafened"] = True
	except (COMError, Exception):
		pass

	# --- Voice channel info ---
	try:
		voice_info = get_voice_connection_info()
		if voice_info:
			if voice_info.get("channel"):
				context["voice_channel"] = voice_info["channel"]
			elif voice_info.get("status"):
				context["voice_channel"] = "Connected"
	except (COMError, Exception):
		pass

	# --- Badge counts on servers ---
	try:
		server_list = find_server_list()
		if server_list:
			for child in _iter_children(server_list):
				name = safe_name(child)
				if not name:
					continue
				lower = name.lower()
				# Look for unread/mention badges like "3 mentions" or "1 unread"
				if re.search(r'\d+\s*(mention|unread|notification)', lower):
					context["badges"].append(name.strip())
				# Some servers show "NEW" or badge count in name
				elif ", " in name:
					# E.g. "My Server, 3 mentions"
					parts_list = name.split(", ")
					for part in parts_list[1:]:
						pl = part.lower().strip()
						if re.search(r'\d+\s*(mention|unread|notification)', pl):
							context["badges"].append(part.strip())
						elif pl == "new":
							context["badges"].append("%s: new" % parts_list[0])
	except (COMError, Exception):
		pass

	# --- Alerts (toasts, banners) ---
	try:
		for name, child in _get_depth1():
			if not name:
				continue
			role = safe_role(child)
			if role == controlTypes.Role.ALERT:
				context["alerts"].append(name.strip())
	except (COMError, Exception):
		pass

	return context


# ---------------------------------------------------------------------------
# Diagnostic dump  —  [ Ctrl+E
# ---------------------------------------------------------------------------

def dump_tree(root=None, max_depth=8):
	"""Dump Discord's accessibility tree using BFS.

	Shows the focus chain, content root, and a breadth-first walk
	of the tree with all named elements.
	"""
	lines = []
	lines.append("=== Discord Accessibility Tree Dump ===")
	lines.append("")

	# --- Focus info ---
	lines.append("--- Focus ---")
	try:
		focus = api.getFocusObject()
		if focus:
			lines.append("  role=%s name=%r type=%s" % (
				safe_role(focus), safe_name(focus)[:60],
				type(focus).__name__,
			))
		else:
			lines.append("  Focus is None!")
	except Exception as e:
		lines.append("  Error: %s" % e)

	# --- Content root ---
	cr = get_content_root()
	lines.append("--- Content root ---")
	if cr:
		lines.append("  role=%s name=%r type=%s" % (
			safe_role(cr), safe_name(cr)[:80], type(cr).__name__,
		))
	else:
		lines.append("  NOT FOUND")
	lines.append("")

	# --- BFS tree walk ---
	lines.append("--- Tree walk (BFS, depth %d, timeout 8s) ---" % max_depth)
	walk_root = root if root else (cr if cr else get_foreground())
	if walk_root is None:
		lines.append("  No root available.")
		return "\n".join(lines)

	start = time.time()
	count = 0
	named_count = 0
	pruned_count = 0

	queue = deque()
	queue.append((walk_root, 0))
	while queue:
		if time.time() - start > 8.0:
			lines.append("... (timed out after 8s)")
			break
		obj, depth = queue.popleft()
		if depth > 0:
			count += 1
			name = safe_name(obj)
			role = safe_role(obj)
			indent = "  " * min(depth, 8)
			if name:
				named_count += 1
				lines.append("%sd%d %s: %r" % (
					indent, depth, _role_label(role), name[:100],
				))
		if depth < max_depth:
			# Prune leaf roles
			if depth > 0 and safe_role(obj) in _LEAF_ROLES:
				pruned_count += 1
				continue
			for child in _iter_children(obj):
				queue.append((child, depth + 1))
		if count > 500:
			lines.append("... (stopped at 500 elements)")
			break

	elapsed = time.time() - start
	lines.append("")
	lines.append("=== %d total, %d named, %d pruned, %.1fs ===" % (
		count, named_count, pruned_count, elapsed,
	))
	return "\n".join(lines)
