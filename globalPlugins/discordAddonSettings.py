# Discord Enhancements Add-on for NVDA
# Global plugin — configuration registration and settings panel
#
# This global plugin:
#   1. Registers the [discordAddon] config spec section on load
#   2. Provides a SettingsPanel in NVDA Preferences for changing the
#      command prefix key and toggle options
#   3. Shows a conflict warning when the chosen prefix key is already
#      bound to another gesture

from io import StringIO
from logHandler import log
import config
import globalPluginHandler
import gui
from gui import guiHelper
from gui.settingsDialogs import NVDASettingsDialog, SettingsPanel
import inputCore
import wx


# ---------------------------------------------------------------------------
# Configuration specification
# ---------------------------------------------------------------------------

_CONFIG_SPEC = """
[discordAddon]
	commandPrefix = string(default="[")
	announceChatMessages = boolean(default=True)
	verbosityLevel = integer(default=1, min=0, max=3)
"""


def _registerConfig():
	"""Merge the add-on's config spec into NVDA's global config."""
	try:
		confspec = StringIO(_CONFIG_SPEC)
		config.conf.spec.update(config.ConfigObj(confspec))
	except Exception:
		log.error("Failed to register discordAddon config spec", exc_info=True)


# ---------------------------------------------------------------------------
# Settings panel
# ---------------------------------------------------------------------------

class DiscordAddonSettingsPanel(SettingsPanel):
	"""NVDA settings panel for the Discord Enhancements add-on."""

	# Translators: title of the settings category in NVDA Preferences
	title = "Discord Enhancements"

	# Panel description shown at the top
	panelDescription = (
		"Configure the Discord Enhancements add-on.  "
		"Changes to the command prefix key take effect immediately."
	)

	def makeSettings(self, sizer):
		sHelper = guiHelper.BoxSizerHelper(self, sizer=sizer)

		# --- Command prefix key ---
		self.prefixEdit = sHelper.addLabeledControl(
			# Translators: label for the command prefix key field
			"Command &prefix key:",
			wx.TextCtrl,
		)
		self.prefixEdit.SetValue(
			config.conf["discordAddon"]["commandPrefix"]
		)
		self.prefixEdit.SetMaxLength(20)

		# --- Announce chat messages ---
		self.announceChatCheckBox = sHelper.addItem(
			wx.CheckBox(
				self,
				# Translators: label for the announce chat messages checkbox
				label="Announce incoming &chat messages",
			)
		)
		self.announceChatCheckBox.SetValue(
			config.conf["discordAddon"]["announceChatMessages"]
		)

		# --- Verbosity level ---
		verbosityChoices = [
			"Minimal",       # 0
			"Normal",        # 1
			"Verbose",       # 2
			"Extra verbose", # 3
		]
		self.verbosityChoice = sHelper.addLabeledControl(
			# Translators: label for the verbosity selector
			"&Verbosity level:",
			wx.Choice,
			choices=verbosityChoices,
		)
		self.verbosityChoice.SetSelection(
			config.conf["discordAddon"]["verbosityLevel"]
		)

	def isValid(self):
		"""Validate settings before saving."""
		prefix = self.prefixEdit.GetValue().strip()
		if not prefix:
			# Translators: error when prefix key is blank
			gui.messageBox(
				"The command prefix key cannot be empty.",
				"Validation Error",
				wx.OK | wx.ICON_ERROR,
				self,
			)
			return False
		return True

	def onSave(self):
		"""Write settings to config and notify the app module."""
		new_prefix = self.prefixEdit.GetValue().strip()
		old_prefix = config.conf["discordAddon"]["commandPrefix"]

		# --- Conflict warning ---
		if new_prefix != old_prefix:
			if self._checkGestureConflict(new_prefix):
				result = gui.messageBox(
					'The key "%s" is already assigned to another gesture '
					"or NVDA command.  Using it as the Discord command "
					"prefix may cause conflicts.\n\n"
					"Do you want to use it anyway?" % new_prefix,
					"Key Conflict Warning",
					wx.YES_NO | wx.NO_DEFAULT | wx.ICON_WARNING,
					self,
				)
				if result != wx.YES:
					# Revert the text field and don't save the prefix
					self.prefixEdit.SetValue(old_prefix)
					return

		config.conf["discordAddon"]["commandPrefix"] = new_prefix
		config.conf["discordAddon"]["announceChatMessages"] = (
			self.announceChatCheckBox.IsChecked()
		)
		config.conf["discordAddon"]["verbosityLevel"] = (
			self.verbosityChoice.GetSelection()
		)

	@staticmethod
	def _checkGestureConflict(key):
		"""Return True if 'kb:<key>' is bound to a script somewhere."""
		gesture_id = "kb:%s" % key
		try:
			normalised = inputCore.normalizeGestureIdentifier(gesture_id)
		except Exception:
			return False

		# Check user gesture map
		try:
			user_map = inputCore.manager.userGestureMap
			for cls_name, module, gesture, script_name in user_map.getAll():
				try:
					if inputCore.normalizeGestureIdentifier(gesture) == normalised:
						return True
				except Exception:
					continue
		except Exception:
			pass

		# Check locale gesture map
		try:
			locale_map = inputCore.manager.localeGestureMap
			for cls_name, module, gesture, script_name in locale_map.getAll():
				try:
					if inputCore.normalizeGestureIdentifier(gesture) == normalised:
						return True
				except Exception:
					continue
		except Exception:
			pass

		return False


# ---------------------------------------------------------------------------
# Global plugin
# ---------------------------------------------------------------------------

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	"""Global plugin that registers the Discord add-on config and settings."""

	def __init__(self):
		super().__init__()
		_registerConfig()
		NVDASettingsDialog.categoryClasses.append(DiscordAddonSettingsPanel)
		log.info("Discord Enhancements global plugin loaded")

	def terminate(self):
		try:
			NVDASettingsDialog.categoryClasses.remove(
				DiscordAddonSettingsPanel,
			)
		except ValueError:
			pass
		log.info("Discord Enhancements global plugin unloaded")
