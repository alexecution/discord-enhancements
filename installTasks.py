# Discord Enhancements Add-on for NVDA
# Install / Uninstall tasks


def onInstall():
	"""Called when the add-on is installed or updated."""
	pass


def onUninstall():
	"""Called when the add-on is uninstalled.

	Resets the command prefix key to the default value to prevent
	orphaned gesture bindings that could interfere with normal
	keyboard input after the add-on is removed.
	"""
	try:
		import config
		if "discordAddon" in config.conf:
			config.conf["discordAddon"]["commandPrefix"] = "["
			config.conf.save()
	except Exception:
		pass
