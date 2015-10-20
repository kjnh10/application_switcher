import switcher3
import platform

os_version = platform.win32_ver()[0]

if os_version == "8":
    #for excel 2013
    window = switcher3.Window("XLMAIN", "- Excel","EXCEL")
else:
    # #for excel 2007
    # window = switcher3.Window("XLMAIN", " Excel -","EXCEL")

    #for excel 2013
    window = switcher3.Window("XLMAIN", "- Excel","EXCEL")

window.activate_app()
#window.debug()

