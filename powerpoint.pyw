import switcher3
import platform

os_version = platform.win32_ver()[0]

window = switcher3.Window("PPTFrameClass", "- Microsoft PowerPoint", "POWERPNT")

window.activate_app()
window.debug()

