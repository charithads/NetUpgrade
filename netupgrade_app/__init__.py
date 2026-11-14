"""NetUpgrade application package.

Phase 1 of the restructure: the storage layer (``netupgrade_app.storage``) replaces the JSON files the
application used to write. The rest of the application still lives in ``ios_xe_upgrade_manager.py``
and ``launcher.py`` and will be split into modules of this package in a later phase.
"""
